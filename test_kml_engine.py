import random
import zipfile

from kml_engine import KmlDocument, parse_description, search_coordinate


FULL = """id: 7bb79feb-f029-4adb-a56c-33b8b3c5db0f |
رقم الحيازة: 151، عدد القطع: 1 |
اسم الجمعية: البهو فريك-الائتمان الزراعي |
اسم الحائز: محمد حامد سالم سلامه |
اسم المالك: ورثة محححح |
الرقم القومي: 26307301201652 |
اسم الحوض: الشميلى الشرقى اول، كود الحوض: 06323900003230000017 |
رقم الأرض: 13267150 |
المساحة: 0 فدان، 4 قيراط، 0 سهم |
نوع الاستخدام: زراعة |
نوع المحصول: فول |
مرحلة النمو: - |
الملاحظات: ملاحظة"""


def kml(*items):
    marks = "".join(
        f"<Placemark><name>{title}</name><description><![CDATA[{description}]]></description>"
        "<Polygon><outerBoundaryIs><LinearRing><coordinates>30,30,0 31,30,0 31,31,0 30,30,0</coordinates>"
        "</LinearRing></outerBoundaryIs></Polygon></Placemark>"
        for title, description in items
    )
    return f'<?xml version="1.0" encoding="UTF-8"?><kml xmlns="http://www.opengis.net/kml/2.2"><Document>{marks}</Document></kml>'


def test_parses_full_description_and_area():
    data = parse_description(FULL)
    assert (data["feddan"], data["qirat"], data["sahm"]) == ("0", "4", "0")
    assert data["parcel_count"] == "1"


def test_building_allows_empty_crop_and_growth():
    data = parse_description(FULL.replace("نوع الاستخدام: زراعة |\nنوع المحصول: فول |\nمرحلة النمو: - |\n", "نوع الاستخدام: مباني |\n"))
    assert data["land_use"] == "مباني"
    assert data["crop_type"] == ""


def test_building_and_fallow_clear_crop_and_growth_on_load_and_export(tmp_path):
    building = FULL.replace("نوع الاستخدام: زراعة", "نوع الاستخدام: مباني")
    fallow = FULL.replace("نوع الاستخدام: زراعة", "نوع الاستخدام: بور").replace(
        "7bb79feb-f029-4adb-a56c-33b8b3c5db0f", "d6d12fe7-bbb1-4fca-a9c8-ffb6cccbb4dc"
    )
    source = tmp_path / "non-agricultural.kml"
    source.write_text(
        kml(
            ("7bb79feb-f029-4adb-a56c-33b8b3c5db0f", building),
            ("d6d12fe7-bbb1-4fca-a9c8-ffb6cccbb4dc", fallow),
        ),
        encoding="utf-8",
    )
    document = KmlDocument.load(source)
    assert all(not parcel.data["crop_type"] for parcel in document.parcels)
    assert all(not parcel.data["growth_stage"] for parcel in document.parcels)
    output = tmp_path / "non-agricultural-out.kml"
    document.export_kml(output)
    text = output.read_text(encoding="utf-8")
    assert "نوع المحصول" not in text and "مراحل النمو" not in text


def test_html_description_and_excel_export(tmp_path):
    html = "<table><tr><td>ID:</td><td>7bb79feb-f029-4adb-a56c-33b8b3c5db0f</td></tr>" + "<tr><td>رقم الحيازة:</td><td>151</td></tr></table>" + FULL.split("رقم الحيازة: 151، عدد القطع: 1", 1)[1]
    # The direct parser receives mixed HTML/text; its text cleanup keeps Arabic labels readable.
    data = parse_description(html)
    assert data["id"] == "7bb79feb-f029-4adb-a56c-33b8b3c5db0f"
    file = tmp_path / "source.kml"
    file.write_text(kml(("bad title", FULL)), encoding="utf-8")
    doc = KmlDocument.load(file)
    doc.apply_safe_fixes()
    output = tmp_path / "parcels.xlsx"
    doc.export_excel(output)
    assert output.exists() and output.stat().st_size > 0


def test_reads_description_written_as_nested_xml_not_cdata(tmp_path):
    """Some KML exporters use child HTML/XML nodes inside description."""
    nested = FULL.replace("|", "</div><div>")
    source = tmp_path / "nested.kml"
    source.write_text(
        "<?xml version='1.0'?><kml xmlns='http://www.opengis.net/kml/2.2'><Document>"
        "<Placemark><name>5f1e5212-eacd-42c3-b2e1-2d7b9561e089</name>"
        f"<description><div>{nested}</div></description>"
        "<Polygon><outerBoundaryIs><LinearRing><coordinates>30,30,0 31,30,0 31,31,0 30,30,0</coordinates>"
        "</LinearRing></outerBoundaryIs></Polygon></Placemark></Document></kml>",
        encoding="utf-8",
    )
    document = KmlDocument.load(source)
    assert document.parcels[0].data["id"] == "7bb79feb-f029-4adb-a56c-33b8b3c5db0f"
    assert document.parcels[0].data["holder_name"] == "محمد حامد سالم سلامه"
    document.export_kml(tmp_path / "clean.kml")
    clean = (tmp_path / "clean.kml").read_text(encoding="utf-8")
    assert "<div" not in clean


def test_reads_popup_data_from_a_nonstandard_namespace(tmp_path):
    screenshot_style = """ID: 32896017-8d84-45b3-8cfa-9561a32752d2;
رقم الحيازة: 129; عدد القطع: 1;
اسم الجمعية: البهو فريك-الائتمان الزراعي;
اسم الحائز: نهاد محمد فؤاد ابراهيم;
اسم المالك: نهاد محمد فؤاد ابراهيم;
الرقم القومي: 25505211201486;
اسم الحوض: البر الايمن; كود الحوض: 06323900003230000005;
رقم الأرض: 14163225;
المساحة: فدان: 0; قيراط: 5; سهم: 4;
نوع الاستخدام: زراعة"""
    source = tmp_path / "foreign-description.kml"
    source.write_text(
        "<kml xmlns='http://www.opengis.net/kml/2.2' xmlns:vendor='urn:vendor'><Document>"
        "<Placemark><name>32896017-8d84-45b3-8cfa-9561a32752d2</name>"
        f"<vendor:description><![CDATA[{screenshot_style}]]></vendor:description>"
        "<Polygon><outerBoundaryIs><LinearRing><coordinates>30,30,0 31,30,0 31,31,0 30,30,0</coordinates>"
        "</LinearRing></outerBoundaryIs></Polygon></Placemark></Document></kml>",
        encoding="utf-8",
    )
    document = KmlDocument.load(source)
    parcel = document.parcels[0]
    assert parcel.status == "valid"
    assert parcel.data["qirat"] == "5"
    output = tmp_path / "standard-description.kml"
    document.export_kml(output)
    text = output.read_text(encoding="utf-8")
    assert "vendor:description" not in text
    assert "نهاد محمد فؤاد ابراهيم" in text


def test_reads_google_earth_balloon_style_referenced_by_placemark(tmp_path):
    balloon = """ID: 32896017-8d84-45b3-8cfa-9561a32752d2;
رقم الحيازة: 129; عدد القطع: 1; اسم الحائز: نهاد محمد فؤاد ابراهيم;
اسم المالك: نهاد محمد فؤاد ابراهيم; الرقم القومي: 25505211201486;
اسم الحوض: داير الناحيه; كود الحوض: 06323000003239000005;
رقم الأرض: 14163225; المساحة: فدان: 0; قيراط: 5; سهم: 4; نوع الاستخدام: زراعة;"""
    source = tmp_path / "balloon-style.kml"
    source.write_text(
        "<kml xmlns='http://www.opengis.net/kml/2.2' xmlns:gx='http://www.google.com/kml/ext/2.2'>"
        "<Document><gx:CascadingStyle id='parcel-style_highlight'><Style><BalloonStyle>"
        f"<text><![CDATA[{balloon}]]></text></BalloonStyle></Style></gx:CascadingStyle>"
        "<Placemark><name>32896017-8d84-45b3-8cfa-9561a32752d2</name><styleUrl>#parcel-style</styleUrl>"
        "<Polygon><outerBoundaryIs><LinearRing><coordinates>30,30,0 31,30,0 31,31,0 30,30,0</coordinates>"
        "</LinearRing></outerBoundaryIs></Polygon></Placemark></Document></kml>",
        encoding="utf-8",
    )
    document = KmlDocument.load(source)
    parcel = document.parcels[0]
    assert parcel.status == "valid"
    assert parcel.data["holding_number"] == "129"
    output = tmp_path / "normalized.kml"
    document.export_kml(output)
    assert "<description>" in output.read_text(encoding="utf-8")


def test_duplicate_gets_unique_id_and_requires_review(tmp_path):
    file = tmp_path / "source.kml"
    file.write_text(kml(("7bb79feb-f029-4adb-a56c-33b8b3c5db0f", FULL), ("7bb79feb-f029-4adb-a56c-33b8b3c5db0f", FULL)), encoding="utf-8")
    doc = KmlDocument.load(file)
    doc.apply_safe_fixes()
    assert doc.parcels[0].id != doc.parcels[1].id
    assert doc.parcels[1].review_required
    assert doc._style_id(doc.parcels[1]) == "parcel-yellow"


def test_kmz_keeps_other_members(tmp_path):
    source = tmp_path / "source.kmz"
    with zipfile.ZipFile(source, "w") as archive:
        archive.writestr("doc.kml", kml(("wrong", FULL)))
        archive.writestr("files/icon.png", b"asset")
    doc = KmlDocument.load(source)
    doc.apply_safe_fixes()
    output = tmp_path / "review.kmz"
    doc.export_kml(output)
    with zipfile.ZipFile(output) as archive:
        assert archive.read("files/icon.png") == b"asset"
        assert b"parcel-blue" in archive.read("doc.kml")


def test_unallocated_holding_is_complete_when_notes_confirm_it(tmp_path):
    description = FULL.replace("رقم الحيازة: 151", "رقم الحيازة: -1").replace("الملاحظات: ملاحظة", "الملاحظات: غير محيز")
    source = tmp_path / "unallocated.kml"
    source.write_text(kml(("7bb79feb-f029-4adb-a56c-33b8b3c5db0f", description)), encoding="utf-8")
    parcel = KmlDocument.load(source).parcels[0]
    assert parcel.status == "unallocated"
    assert "غير محيزة" in parcel.reason


def test_unallocated_holding_does_not_require_notes_or_matching_title(tmp_path):
    description = FULL.replace("رقم الحيازة: 151", "رقم الحيازة: 0").replace("الملاحظات: ملاحظة", "")
    source = tmp_path / "unallocated-without-notes.kml"
    source.write_text(kml(("عنوان غير مطابق", description)), encoding="utf-8")
    parcel = KmlDocument.load(source).parcels[0]
    assert parcel.status == "unallocated"
    assert not parcel.blocking


def test_growth_distribution_overrides_values_and_forces_fruiting_crops(tmp_path):
    ids = [f"00000000-0000-4000-8000-{number:012d}" for number in range(1, 6)]
    crops = ["عنب", "برتقال", "فول", "قمح", ""]
    descriptions = []
    for parcel_id, crop in zip(ids, crops):
        description = FULL.replace("7bb79feb-f029-4adb-a56c-33b8b3c5db0f", parcel_id).replace("نوع المحصول: فول", f"نوع المحصول: {crop}").replace("مرحلة النمو: -", "مرحلة النمو: قيمة قديمة")
        descriptions.append((parcel_id, description))
    source = tmp_path / "growth.kml"; source.write_text(kml(*descriptions), encoding="utf-8")
    document = KmlDocument.load(source)
    document.distribute_growth_stages({"مرحلة النمو الخضري": 50, "مرحلة الإنبات": 50, "مرحلة الإزهار والإثمار": 0}, rng=random.Random(7))
    stages = [parcel.data["growth_stage"] for parcel in document.parcels]
    assert stages[0] == "مرحلة الإزهار والإثمار" and stages[1] == "مرحلة الإزهار والإثمار"
    assert stages[2:].count("مرحلة النمو الخضري") == 2
    assert stages[2:].count("مرحلة الإنبات") == 1


def test_export_styles_are_transparent_except_red(tmp_path):
    source = tmp_path / "styles.kml"
    source.write_text(kml(("7bb79feb-f029-4adb-a56c-33b8b3c5db0f", FULL), ("bad", "")), encoding="utf-8")
    document = KmlDocument.load(source); output = tmp_path / "styles-out.kml"; document.export_kml(output)
    text = output.read_text(encoding="utf-8")
    assert "00ff6600" in text and "00ffff00" in text and "0000ffff" in text
    assert "bf0000ff" in text


def test_search_coordinate_uses_polygon_centroid_in_dms_format():
    point = search_coordinate("31,30,0 32,30,0 32,31,0 31,31,0 31,30,0")
    assert point == '30°30\'00.00"N 31°30\'00.00"E'


def test_missing_fields_can_be_saved_or_explicitly_skipped(tmp_path):
    source = tmp_path / "missing-owner.kml"
    missing_owner = FULL.replace("اسم المالك: ورثة محححح |\n", "")
    source.write_text(
        kml(("7bb79feb-f029-4adb-a56c-33b8b3c5db0f", missing_owner)),
        encoding="utf-8",
    )
    document = KmlDocument.load(source)
    parcel = document.parcels[0]
    assert parcel.status == "incomplete" and parcel.blocking

    document.update_parcel(parcel, parcel.data, skip_missing=True)
    assert parcel.status == "skipped" and not parcel.blocking
    output = tmp_path / "skipped.xlsx"
    document.export_excel(output)
    assert output.exists()

    values = dict(parcel.data)
    values["owner_name"] = "مالك تمت مراجعته"
    document.update_parcel(parcel, values)
    assert parcel.status == "valid" and not parcel.blocking
