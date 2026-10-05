import zipfile

from kml_engine import KmlDocument, parse_description


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
