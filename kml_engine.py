"""KML/KMZ parsing, validation, repair, and export for land parcel review."""

from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass, field
from html import unescape
from html.parser import HTMLParser
from pathlib import Path
import copy
import io
import math
import random
import re
import uuid
import xml.etree.ElementTree as ET
import zipfile

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

KML_NS = "http://www.opengis.net/kml/2.2"
GX_NS = "http://www.google.com/kml/ext/2.2"
NS = {"k": KML_NS}
ET.register_namespace("", KML_NS)
ET.register_namespace("gx", GX_NS)

UUID_RE = re.compile(
    r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[1-5][0-9a-fA-F]{3}-[89abAB][0-9a-fA-F]{3}-[0-9a-fA-F]{12}$"
)
ANY_UUID_RE = re.compile(
    r"\b[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[1-5][0-9a-fA-F]{3}-[89abAB][0-9a-fA-F]{3}-[0-9a-fA-F]{12}\b"
)

FIELDS = [
    ("id", "كود القطعة"),
    ("holding_number", "رقم الحيازة"),
    ("parcel_count", "عدد القطع بالحيازة"),
    ("holder_name", "اسم الحائز"),
    ("owner_name", "اسم المالك"),
    ("national_id", "الرقم القومي للحائز"),
    ("association_name", "اسم الجمعية"),
    ("association_code", "كود الجمعية"),
    ("basin_name", "اسم الحوض"),
    ("basin_code", "كود الحوض"),
    ("parcel_number", "رقم القطعة"),
    ("feddan", "فدان"),
    ("qirat", "قيراط"),
    ("sahm", "سهم"),
    ("land_use", "نوع الاستخدام"),
    ("crop_type", "نوع المحصول"),
    ("growth_stage", "مراحل النمو"),
    ("notes", "ملاحظات علي الحيازة (من فريق العمل)"),
]

LABELS = {
    "id": ["id", "ID", "كود القطعة"],
    "holding_number": ["رقم الحيازة"],
    "parcel_count": ["عدد القطع بالحيازة", "عدد القطع"],
    "holder_name": ["اسم الحائز"],
    "owner_name": ["اسم المالك"],
    "national_id": ["الرقم القومي للحائز", "الرقم القومي"],
    "association_name": ["اسم الجمعية"],
    "association_code": ["كود الجمعية"],
    "basin_name": ["اسم الحوض"],
    "basin_code": ["كود الحوض"],
    "parcel_number": ["رقم القطعة", "رقم الأرض"],
    "feddan": ["فدان"],
    "qirat": ["قيراط", "قراط"],
    "sahm": ["سهم"],
    "land_use": ["نوع الاستخدام"],
    "crop_type": ["نوع المحصول"],
    "growth_stage": ["مراحل النمو", "مرحلة النمو"],
    "notes": ["الملاحظات"],
    "area": ["المساحة"],
}
REQUIRED = {
    "id",
    "holding_number",
    "parcel_count",
    "holder_name",
    "owner_name",
    "national_id",
    "basin_name",
    "basin_code",
    "parcel_number",
    "feddan",
    "qirat",
    "sahm",
    "land_use",
}
OPTIONAL = {
    "crop_type",
    "growth_stage",
    "notes",
    "association_name",
    "association_code",
}
GROWTH_STAGES = ("مرحلة النمو الخضري", "مرحلة الإنبات", "مرحلة الإزهار والإثمار")
FRUITING_CROPS = ("عنب", "برتقال", "كاكا", "كاكاو", "جوافه")
ARABIC_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")


class TextExtractor(HTMLParser):
    def __init__(self):
        super().__init__()
        self.parts: list[str] = []

    def handle_data(self, data: str):
        self.parts.append(data)

    def handle_starttag(self, tag, attrs):
        if tag.lower() in {"br", "p", "div", "tr", "li", "td", "th"}:
            self.parts.append("\n")

    def text(self) -> str:
        return (
            " ".join(" ".join(self.parts).split())
            if not any("\n" in p for p in self.parts)
            else "".join(self.parts)
        )


def clean_description(value: str | None) -> str:
    raw = unescape(value or "").replace("\xa0", " ")
    parser = TextExtractor()
    try:
        parser.feed(raw)
        raw = parser.text()
    except Exception:
        pass
    return re.sub(r"[ \t]+", " ", raw).strip()


def _label_pattern() -> re.Pattern:
    all_labels = sorted(
        {label for values in LABELS.values() for label in values}, key=len, reverse=True
    )
    return re.compile(
        r"(?P<label>" + "|".join(re.escape(x) for x in all_labels) + r")\s*[:：]\s*",
        re.I,
    )


LABEL_PATTERN = _label_pattern()
LABEL_TO_KEY = {
    label.casefold(): key for key, labels in LABELS.items() for label in labels
}


def _tidy(value: str) -> str:
    return value.strip(" \t\n\r|;؛،,")


def normalize_arabic(value: str) -> str:
    """Normalize values only for matching rules; preserve original export text."""
    text = (value or "").translate(ARABIC_DIGITS).casefold()
    text = re.sub(r"[ًٌٍَُِّْـ]", "", text)
    return (
        text.replace("أ", "ا")
        .replace("إ", "ا")
        .replace("آ", "ا")
        .replace("ة", "ه")
        .strip()
    )


def is_unallocated(data: dict[str, str]) -> bool:
    holding = normalize_arabic(data.get("holding_number", "")).replace(" ", "")
    # The source normally includes "غير محيز" in notes, but the holding number
    # itself is the authoritative business indicator.
    return holding in {"0", "-1", "0.0", "-1.0"}


def is_agricultural(data: dict[str, str]) -> bool:
    return "زراع" in normalize_arabic(data.get("land_use", ""))


def must_use_fruiting_stage(data: dict[str, str]) -> bool:
    crop = normalize_arabic(data.get("crop_type", ""))
    return any(crop == name or crop.startswith(name + " ") for name in FRUITING_CROPS)


def parse_description(
    description: str | None, extended: dict[str, str] | None = None
) -> dict[str, str]:
    text = clean_description(description)
    if extended:
        text += "\n" + "\n".join(
            f"{key}: {value}" for key, value in extended.items() if value
        )
    result = {key: "" for key, _ in FIELDS}
    matches = list(LABEL_PATTERN.finditer(text))
    for i, match in enumerate(matches):
        key = LABEL_TO_KEY[match.group("label").casefold()]
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        value = _tidy(text[match.end() : end])
        if key == "area":
            area_matches = list(
                re.finditer(
                    r"(فدان|قيراط|قراط|سهم)\s*[:：]?\s*([0-9٠-٩.,]+)", value, re.I
                )
            )
            for area in area_matches:
                area_key = {
                    "فدان": "feddan",
                    "قيراط": "qirat",
                    "قراط": "qirat",
                    "سهم": "sahm",
                }[area.group(1)]
                result[area_key] = _tidy(area.group(2))
        elif key in result and value:
            result[key] = value
    if not result["id"]:
        found = ANY_UUID_RE.search(text)
        if found:
            result["id"] = found.group(0)
    result["id"] = result["id"].lower().strip()
    # Area can be one phrase ("0 فدان، 4 قيراط، 0 سهم") or separate fields.
    # Resolve it from the complete text after generic parsing so Arabic commas are safe.
    for key, tokens in {
        "feddan": ["فدان"],
        "qirat": ["قيراط", "قراط"],
        "sahm": ["سهم"],
    }.items():
        for token in tokens:
            number = re.search(
                rf"([0-9٠-٩]+(?:[.,][0-9٠-٩]+)?)\s*{re.escape(token)}\b", text
            )
            if not number:
                number = re.search(
                    rf"{re.escape(token)}\s*[:：]\s*([0-9٠-٩]+(?:[.,][0-9٠-٩]+)?)", text
                )
            if number:
                result[key] = number.group(1)
                break
    return result


def is_uuid(value: str) -> bool:
    return bool(UUID_RE.fullmatch((value or "").strip()))


def local_name(tag: str) -> str:
    """Read an XML tag name without making assumptions about its namespace."""
    return tag.rsplit("}", 1)[-1].casefold()


def direct_child(element: ET.Element, tag: str) -> ET.Element | None:
    expected = tag.casefold()
    return next((child for child in element if local_name(child.tag) == expected), None)


def title_id(title: str) -> str:
    return title.strip().lower() if is_uuid(title) else ""


def element_text(element: ET.Element, tag: str) -> str:
    node = direct_child(element, tag)
    return (node.text or "").strip() if node is not None else ""


def element_markup(element: ET.Element, tag: str) -> str:
    """Return all content inside a KML node, including nested author HTML/XML.

    KML descriptions are usually CDATA, but some exporters write real child nodes
    (`table`, `div`, etc.) inside ``description``. Reading only ``node.text`` drops
    those child nodes and makes a populated parcel look empty.
    """
    node = direct_child(element, tag)
    if node is None:
        return ""
    parts = [node.text or ""]
    parts.extend(ET.tostring(child, encoding="unicode") for child in node)
    return "".join(parts).strip()


def set_element_text(element: ET.Element, tag: str, value: str) -> None:
    node = element.find(f"k:{tag}", NS)
    if node is None:
        # Replace a foreign-namespace counterpart so output always uses the
        # standard KML field understood by Google Earth and later imports.
        foreign_node = direct_child(element, tag)
        if foreign_node is not None:
            element.remove(foreign_node)
        node = ET.SubElement(element, f"{{{KML_NS}}}{tag}")
    # A description may contain nested author-controlled HTML/XML. Remove it before
    # writing normalized direct text so no hidden markup survives in the export.
    for child in list(node):
        node.remove(child)
    node.text = value


def extract_extended_data(placemark: ET.Element) -> dict[str, str]:
    values: dict[str, str] = {}
    for data in placemark.iter():
        tag = local_name(data.tag)
        if tag == "data":
            name = data.get("name", "")
            value = direct_child(data, "value")
            content = element_markup(data, "value") if value is not None else data.get("value", "")
            if name and content:
                values[name] = clean_description(content)
        elif tag == "simpledata" and data.get("name") and "".join(data.itertext()).strip():
            values[data.get("name", "")] = clean_description("".join(data.itertext()))
    return values


def fallback_description(placemark: ET.Element) -> str:
    """Collect description-like fields stored under non-standard KML namespaces.

    Several GIS exporters place popup content in a namespaced `description`,
    `snippet`, or value field. Google Earth renders it, while a namespace-specific
    reader sees an empty description. This fallback intentionally excludes geometry
    and titles and only collects fields that can contain human-readable attributes.
    """
    fragments: list[str] = []
    for node in placemark.iter():
        if node is placemark or local_name(node.tag) not in {"description", "snippet", "value", "simpledata"}:
            continue
        content = "".join(node.itertext()).strip()
        if content:
            fragments.append(content)
    return "\n".join(dict.fromkeys(fragments))


def element_attribute(element: ET.Element, name: str) -> str:
    """Get an attribute whether it is plain or namespace-qualified."""
    expected = name.casefold()
    for key, value in element.attrib.items():
        if local_name(key) == expected:
            return value
    return ""


def descendant_text(element: ET.Element, tag: str) -> str:
    expected = tag.casefold()
    for node in element.iter():
        if local_name(node.tag) == expected:
            return "".join(node.itertext()).strip()
    return ""


def balloon_text_by_style(root: ET.Element) -> dict[str, str]:
    """Map KML/GX style IDs to popup text stored in BalloonStyle.

    Google Earth Projects commonly stores a parcel's attribute text on the
    ``gx:CascadingStyle`` referenced by ``Placemark/styleUrl``. The popup looks
    populated in Google Earth even though the Placemark itself has no description.
    """
    styles: dict[str, str] = {}
    for container in root.iter():
        if local_name(container.tag) not in {"style", "cascadingstyle"}:
            continue
        style_id = element_attribute(container, "id")
        if not style_id:
            continue
        balloon = next((node for node in container.iter() if local_name(node.tag) == "balloonstyle"), None)
        text = descendant_text(balloon, "text") if balloon is not None else ""
        if text:
            styles[style_id.lstrip("#")] = text
            # Google Earth Projects references the shared base ID from Placemark
            # while defining the actual popup on `<base>_normal` and
            # `<base>_highlight` cascading styles.
            for suffix in ("_normal", "_highlight"):
                if style_id.endswith(suffix):
                    styles.setdefault(style_id[: -len(suffix)].lstrip("#"), text)
    return styles


def geometry_coordinates(placemark: ET.Element) -> str:
    coords = [
        " ".join((node.text or "").split())
        for node in placemark.findall(".//k:Polygon//k:coordinates", NS)
    ]
    return " | ".join(c for c in coords if c)


def _dms(value: float, positive: str, negative: str) -> str:
    direction = positive if value >= 0 else negative
    absolute = abs(value)
    degrees = int(absolute)
    minutes_float = (absolute - degrees) * 60
    minutes = int(minutes_float)
    seconds = round((minutes_float - minutes) * 60, 2)
    if seconds >= 60:
        seconds, minutes = 0, minutes + 1
    if minutes >= 60:
        minutes, degrees = 0, degrees + 1
    return f'{degrees}°{minutes}\'{seconds:05.2f}"{direction}'


def search_coordinate(coordinates: str) -> str:
    """Return a Google-Earth-friendly DMS point at the first polygon's centroid."""
    first_ring = (coordinates or "").split("|", 1)[0].strip()
    points: list[tuple[float, float]] = []
    for raw_point in first_ring.split():
        parts = raw_point.split(",")
        if len(parts) < 2:
            continue
        try:
            points.append((float(parts[0]), float(parts[1])))  # longitude, latitude
        except ValueError:
            continue
    if len(points) > 1 and points[0] == points[-1]:
        points.pop()
    if not points:
        return ""
    if len(points) < 3:
        longitude = sum(point[0] for point in points) / len(points)
        latitude = sum(point[1] for point in points) / len(points)
    else:
        area_twice = sum(
            points[index][0] * points[(index + 1) % len(points)][1]
            - points[(index + 1) % len(points)][0] * points[index][1]
            for index in range(len(points))
        )
        if abs(area_twice) < 1e-12:
            longitude = sum(point[0] for point in points) / len(points)
            latitude = sum(point[1] for point in points) / len(points)
        else:
            longitude = sum(
                (points[index][0] + points[(index + 1) % len(points)][0])
                * (points[index][0] * points[(index + 1) % len(points)][1]
                   - points[(index + 1) % len(points)][0] * points[index][1])
                for index in range(len(points))
            ) / (3 * area_twice)
            latitude = sum(
                (points[index][1] + points[(index + 1) % len(points)][1])
                * (points[index][0] * points[(index + 1) % len(points)][1]
                   - points[(index + 1) % len(points)][0] * points[index][1])
                for index in range(len(points))
            ) / (3 * area_twice)
    return f"{_dms(latitude, 'N', 'S')} {_dms(longitude, 'E', 'W')}"


def has_polygon(placemark: ET.Element) -> bool:
    return placemark.find(".//k:Polygon", NS) is not None


def canonical_description(data: dict[str, str]) -> str:
    labels = dict(FIELDS)
    lines = []
    for key, _ in FIELDS:
        value = data.get(key, "").strip()
        if value:
            lines.append(f"{labels[key]}: {value}")
    return "\n".join(lines)


@dataclass
class Parcel:
    index: int
    element: ET.Element
    title: str
    description: str
    data: dict[str, str]
    coordinates: str
    status: str = ""
    reason: str = ""
    missing: list[str] = field(default_factory=list)
    review_required: bool = False
    review_confirmed: bool = False
    retained_review: bool = False
    auto_fixed: bool = False
    validation_skipped: bool = False

    @property
    def id(self) -> str:
        return self.data.get("id", "")

    @property
    def search_coordinate(self) -> str:
        return search_coordinate(self.coordinates)

    @property
    def blocking(self) -> bool:
        if self.status == "skipped":
            return False
        return self.status in {
            "empty",
            "no_description",
            "incomplete",
            "unreadable",
        } or (self.review_required and not self.review_confirmed)


STATUS_LABELS = {
    "valid": "سليمة",
    "unallocated": "غير محيزة",
    "title_mismatch": "عنوان غير مطابق",
    "empty": "فارغة",
    "no_description": "بلا وصف",
    "incomplete": "بيانات ناقصة",
    "unreadable": "غير قابلة للقراءة",
    "skipped": "تم تجاوز النواقص",
    "duplicate_different": "ID مكرر وبيانات مختلفة",
    "duplicate_identical": "ID مكرر وبيانات متطابقة",
}


class KmlDocument:
    def __init__(
        self,
        source: Path,
        root: ET.Element,
        parcels: list[Parcel],
        archive: dict[str, bytes] | None = None,
        kml_member: str | None = None,
    ):
        self.source, self.root, self.parcels = source, root, parcels
        self.archive, self.kml_member = archive, kml_member

    @classmethod
    def load(cls, path: str | Path) -> "KmlDocument":
        source = Path(path)
        archive = None
        member = None
        if source.suffix.lower() == ".kmz":
            with zipfile.ZipFile(source) as zf:
                archive = {name: zf.read(name) for name in zf.namelist()}
            candidates = [name for name in archive if name.lower().endswith(".kml")]
            if not candidates:
                raise ValueError("ملف KMZ لا يحتوي على ملف KML.")
            member = next(
                (x for x in candidates if Path(x).name.lower() == "doc.kml"),
                candidates[0],
            )
            raw = archive[member]
        else:
            raw = source.read_bytes()
        try:
            root = ET.fromstring(raw)
        except ET.ParseError as exc:
            raise ValueError(f"ملف KML غير صالح: {exc}") from exc
        style_texts = balloon_text_by_style(root)
        parcels = []
        for index, mark in enumerate(root.findall(".//k:Placemark", NS), start=1):
            if not has_polygon(mark):
                continue
            title = element_text(mark, "name")
            description = element_markup(mark, "description")
            fallback = fallback_description(mark)
            if fallback and fallback not in description:
                description = "\n".join(part for part in [description, fallback] if part)
            style_id = element_text(mark, "styleUrl").lstrip("#")
            balloon_text = style_texts.get(style_id, "")
            if balloon_text and balloon_text not in description:
                description = "\n".join(part for part in [description, balloon_text] if part)
            data = parse_description(description, extract_extended_data(mark))
            parcels.append(
                Parcel(
                    index, mark, title, description, data, geometry_coordinates(mark)
                )
            )
        document = cls(source, root, parcels, archive, member)
        document.analyze()
        return document

    def analyze(self) -> None:
        for parcel in self.parcels:
            retain_review = parcel.retained_review and not parcel.review_confirmed
            parcel.missing = [
                dict(FIELDS)[key]
                for key in REQUIRED
                if not parcel.data.get(key, "").strip()
            ]
            parcel.review_required = False
            if not parcel.title.strip() and not parcel.description.strip():
                parcel.status, parcel.reason = "empty", "لا يوجد عنوان أو وصف للقطعة."
            elif not parcel.description.strip():
                parcel.status, parcel.reason = (
                    "no_description",
                    "يوجد عنوان لكن لا توجد بيانات وصف.",
                )
            elif not parcel.id:
                parcel.status, parcel.reason = "incomplete", "لا يوجد ID صالح في الوصف."
            elif not is_uuid(parcel.id):
                parcel.status, parcel.reason = "unreadable", "ID الوصف ليس UUID صالحًا."
            elif parcel.missing:
                parcel.status, parcel.reason = (
                    "incomplete",
                    "حقول إلزامية ناقصة: " + "، ".join(parcel.missing),
                )
            elif is_unallocated(parcel.data):
                parcel.status, parcel.reason = (
                    "unallocated",
                    "بيانات مكتملة — قطعة غير محيزة.",
                )
            elif title_id(parcel.title) != parcel.id:
                parcel.status, parcel.reason = (
                    "title_mismatch",
                    "عنوان القطعة لا يطابق ID الموجود في الوصف.",
                )
            else:
                parcel.status, parcel.reason = "valid", "البيانات والعنوان متطابقان."

            # A reviewer may explicitly allow a record through with blanks.  Keep
            # the missing-field list visible for traceability, but do not let it
            # block export or keep the parcel in the blocking-review queue.
            if parcel.validation_skipped and parcel.status in {
                "empty",
                "no_description",
                "incomplete",
                "unreadable",
            }:
                parcel.status, parcel.reason = (
                    "skipped",
                    "تم تجاوز النواقص يدويًا؛ ستبقى الحقول الناقصة فارغة.",
                )
            if retain_review and parcel.status == "valid":
                parcel.status, parcel.reason = (
                    "duplicate_identical",
                    "بيانات مكررة سابقًا؛ يلزم اعتماد المراجعة اليدوية.",
                )
                parcel.review_required = True

        groups: dict[str, list[Parcel]] = defaultdict(list)
        for parcel in self.parcels:
            if is_uuid(parcel.id) and not parcel.missing:
                groups[parcel.id].append(parcel)
        for _, group in groups.items():
            if len(group) < 2:
                continue
            fingerprint = lambda p: tuple(
                (key, p.data.get(key, "").strip()) for key, _ in FIELDS if key != "id"
            )
            identical = len({fingerprint(p) for p in group}) == 1
            for parcel in group[1:]:
                parcel.status = (
                    "duplicate_identical" if identical else "duplicate_different"
                )
                parcel.reason = "ID مكرر وبيانات الوصف " + (
                    "متطابقة؛ يلزم اعتماد المراجعة."
                    if identical
                    else "مختلفة؛ سيُنشأ ID جديد."
                )
                parcel.review_required = identical
                parcel.retained_review = identical

    def apply_safe_fixes(self) -> int:
        changed = 0
        groups: dict[str, list[Parcel]] = defaultdict(list)
        for parcel in self.parcels:
            if is_uuid(parcel.id) and not parcel.missing:
                groups[parcel.id].append(parcel)
        for _, group in groups.items():
            fingerprint = lambda p: tuple(
                (key, p.data.get(key, "").strip()) for key, _ in FIELDS if key != "id"
            )
            identical = len(group) > 1 and len({fingerprint(p) for p in group}) == 1
            for parcel in group[1:]:
                if identical:
                    parcel.retained_review = True
                self._set_id(parcel, str(uuid.uuid4()))
                parcel.auto_fixed = True
                changed += 1
        for parcel in self.parcels:
            if (
                is_uuid(parcel.id)
                and not parcel.missing
                and title_id(parcel.title) != parcel.id
            ):
                parcel.title = parcel.id
                set_element_text(parcel.element, "name", parcel.id)
                parcel.auto_fixed = True
                changed += 1
        self.analyze()
        # Duplicate identical records retain an explicit review requirement after IDs are made unique.
        for parcel in self.parcels:
            if parcel.auto_fixed and parcel.status == "valid":
                parcel.reason = "تم إصلاحها تلقائيًا."
        return changed

    def _set_id(self, parcel: Parcel, value: str) -> None:
        parcel.data["id"] = value.lower()
        parcel.title = value.lower()
        parcel.description = canonical_description(parcel.data)
        set_element_text(parcel.element, "name", parcel.title)
        set_element_text(parcel.element, "description", parcel.description)

    def update_parcel(
        self,
        parcel: Parcel,
        values: dict[str, str],
        confirmed: bool = False,
        skip_missing: bool = False,
    ) -> None:
        parcel.data.update({key: values.get(key, "").strip() for key, _ in FIELDS})
        # A standard save re-enables validation.  The bypass is deliberate and
        # only persists when the reviewer chooses the dedicated action.
        parcel.validation_skipped = skip_missing
        parcel.title = parcel.data.get("id", "")
        parcel.description = canonical_description(parcel.data)
        set_element_text(parcel.element, "name", parcel.title)
        set_element_text(parcel.element, "description", parcel.description)
        self.analyze()
        refreshed = self.parcels[self.parcels.index(parcel)]
        if confirmed and not refreshed.missing and is_uuid(refreshed.id):
            refreshed.retained_review = False
            refreshed.review_required = False
            refreshed.review_confirmed = True
            refreshed.status, refreshed.reason = "valid", "تم اعتماد المراجعة اليدوية."

    def apply_association(
        self, name: str, code: str, selected: list[Parcel] | None = None
    ) -> None:
        targets = selected if selected is not None else self.parcels
        for parcel in targets:
            if name.strip():
                parcel.data["association_name"] = name.strip()
            if code.strip():
                parcel.data["association_code"] = code.strip()
            parcel.description = canonical_description(parcel.data)
            set_element_text(parcel.element, "description", parcel.description)
        self.analyze()

    def distribute_growth_stages(
        self,
        percentages: dict[str, str | float | int],
        rng: random.Random | None = None,
    ) -> int:
        """Replace extracted growth stages using the per-file user distribution."""
        ratios: dict[str, float] = {}
        for stage in GROWTH_STAGES:
            raw = (
                str(percentages.get(stage, ""))
                .translate(ARABIC_DIGITS)
                .replace("٪", "")
                .strip()
            )
            try:
                ratios[stage] = float(raw)
            except ValueError as exc:
                raise ValueError("أدخل نسبة رقمية لكل مرحلة نمو.") from exc
            if ratios[stage] < 0:
                raise ValueError("نسب مراحل النمو لا يمكن أن تكون سالبة.")
        if not math.isclose(sum(ratios.values()), 100.0, abs_tol=0.01):
            raise ValueError("يجب أن يكون مجموع نسب مراحل النمو 100٪.")

        agricultural = [p for p in self.parcels if is_agricultural(p.data)]
        variable = [p for p in agricultural if not must_use_fruiting_stage(p.data)]
        for parcel in self.parcels:
            if not is_agricultural(parcel.data):
                parcel.data["growth_stage"] = ""
        for parcel in agricultural:
            if must_use_fruiting_stage(parcel.data):
                parcel.data["growth_stage"] = GROWTH_STAGES[2]

        raw_counts = {stage: len(variable) * ratios[stage] / 100 for stage in GROWTH_STAGES}
        counts = {stage: math.floor(value) for stage, value in raw_counts.items()}
        remaining = len(variable) - sum(counts.values())
        ranking = sorted(
            GROWTH_STAGES,
            key=lambda stage: raw_counts[stage] - counts[stage],
            reverse=True,
        )
        for stage in ranking[:remaining]:
            counts[stage] += 1
        assignments = [stage for stage in GROWTH_STAGES for _ in range(counts[stage])]
        (rng or random.SystemRandom()).shuffle(assignments)
        for parcel, stage in zip(variable, assignments):
            parcel.data["growth_stage"] = stage
        for parcel in self.parcels:
            parcel.description = canonical_description(parcel.data)
            set_element_text(parcel.element, "description", parcel.description)
        self.analyze()
        return len(agricultural)

    def summary(self) -> Counter:
        counts = Counter(parcel.status for parcel in self.parcels)
        counts["total"] = len(self.parcels)
        counts["blocking"] = sum(p.blocking for p in self.parcels)
        counts["auto_fixed"] = sum(p.auto_fixed for p in self.parcels)
        return counts

    def _style_id(self, parcel: Parcel) -> str:
        if parcel.status in {"empty", "no_description", "incomplete", "unreadable"}:
            return "parcel-red"
        if parcel.status == "skipped":
            return "parcel-yellow"
        if parcel.review_required and not parcel.review_confirmed:
            return "parcel-yellow"
        assoc = parcel.data.get("association_name", "")
        if "إصلاح" in assoc or "اصلاح" in assoc:
            return "parcel-cyan"
        if "ائتمان" in assoc:
            return "parcel-blue"
        return "parcel-blue"

    def _install_styles(self) -> None:
        container = self.root.find("k:Document", NS) or self.root
        for old in list(container.findall("k:Style", NS)):
            if old.get("id", "").startswith("parcel-"):
                container.remove(old)
        styles = {
            # KML colour format is aabbggrr. Only blocking red is 75% opaque.
            "parcel-blue": ("00ff6600", "ffcc6600"),
            "parcel-cyan": ("00ffff00", "ffffff00"),
            "parcel-yellow": ("0000ffff", "ff00ffff"),
            "parcel-red": ("bf0000ff", "ff0000ff"),
        }
        for style_id, (fill, line) in styles.items():
            style = ET.Element(f"{{{KML_NS}}}Style", {"id": style_id})
            poly = ET.SubElement(style, f"{{{KML_NS}}}PolyStyle")
            ET.SubElement(poly, f"{{{KML_NS}}}color").text = fill
            ET.SubElement(poly, f"{{{KML_NS}}}fill").text = "1"
            ET.SubElement(poly, f"{{{KML_NS}}}outline").text = "1"
            line_style = ET.SubElement(style, f"{{{KML_NS}}}LineStyle")
            ET.SubElement(line_style, f"{{{KML_NS}}}color").text = line
            ET.SubElement(line_style, f"{{{KML_NS}}}width").text = "1"
            container.insert(0, style)

    def export_kml(self, output: str | Path) -> None:
        self._install_styles()
        for parcel in self.parcels:
            style = parcel.element.find("k:styleUrl", NS)
            if style is None:
                style = ET.SubElement(parcel.element, f"{{{KML_NS}}}styleUrl")
            style.text = "#" + self._style_id(parcel)
            set_element_text(
                parcel.element, "name", parcel.data.get("id", parcel.title)
            )
            set_element_text(
                parcel.element, "description", canonical_description(parcel.data)
            )
        raw = ET.tostring(self.root, encoding="utf-8", xml_declaration=True)
        output = Path(output)
        if output.suffix.lower() == ".kmz":
            archive = copy.copy(self.archive) if self.archive else {}
            archive[self.kml_member or "doc.kml"] = raw
            with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
                for name, data in archive.items():
                    zf.writestr(name, data)
        else:
            output.write_bytes(raw)

    def export_excel(self, output: str | Path) -> None:
        if any(parcel.blocking for parcel in self.parcels):
            raise ValueError("لا يمكن بناء Excel قبل معالجة كل القطع المانعة.")
        book = Workbook()
        sheet = book.active
        sheet.title = "القطع"
        sheet.sheet_view.rightToLeft = True
        header_fill = PatternFill("solid", fgColor="1F4E78")
        for col, (_, label) in enumerate(FIELDS, start=1):
            cell = sheet.cell(1, col, label)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")
        for row, parcel in enumerate(self.parcels, start=2):
            for col, (key, _) in enumerate(FIELDS, start=1):
                sheet.cell(row, col, parcel.data.get(key, ""))
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for col in range(1, len(FIELDS) + 1):
            letter = get_column_letter(col)
            sheet.column_dimensions[letter].width = min(
                38,
                max(
                    14,
                    max(
                        len(str(sheet.cell(r, col).value or ""))
                        for r in range(1, sheet.max_row + 1)
                    )
                    + 2,
                ),
            )
        book.save(output)
