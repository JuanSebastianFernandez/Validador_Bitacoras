from dataclasses import replace
from copy import copy
from datetime import datetime
from io import BytesIO
import zipfile
from xml.etree import ElementTree as ET

from openpyxl import load_workbook

import bitacora as b
from template_drawings import _sheet_paths, _relationships_path, _target, REL_NS
from template_drawings import X14_NS, XM_NS
from test_bitacora import make_flow20_workbook, make_workbook


def sample_source_and_validations():
    wb = load_workbook(BytesIO(make_flow20_workbook()))
    for name in (b.SHEET_NEW, b.SHEET_OLD):
        wb[name]["A2"] = datetime(2026, 11, 23)
        wb[name]["H2"] = datetime(2026, 11, 23)
    wb[b.SHEET_NEW]["J2"] = "1000160001;1000160002"
    stream = BytesIO()
    wb.save(stream)
    wb.close()
    source, rows = b.extract_workbook_rows(stream.getvalue())
    validations = [
        replace(
            item, result=b.STATUS_MATCH,
            arrival_datetime=datetime(2026, 11, 23, 8, 15),
            exit_datetime=datetime(2026, 11, 23, 12, 30),
            odm_datetime=datetime(2026, 11, 23, 10) if item.sheet == b.SHEET_FLOW20 else None,
            drawing_datetime=datetime(2026, 11, 23, 12, 30) if item.sheet == b.SHEET_FLOW20 else None,
            count_result=b.COUNT_UPDATED if b.validation_requires_element_counts(item) else b.COUNT_NOT_APPLICABLE,
            export_pel=3 if b.validation_requires_element_counts(item) else None,
            export_tratas=2 if b.validation_requires_element_counts(item) else None,
            export_pel_breakdown="3;0" if item.qr == "QR-N" else None,
            export_tratas_breakdown="0;2" if item.qr == "QR-N" else None,
        )
        for item in b.validate_rows(rows, {}, {})
    ]
    return source, validations


def sample_output():
    source, validations = sample_source_and_validations()
    return b.build_upload_workbook(
        source, validations, template_bytes=b.default_template_bytes(),
        audit_entries=[{"Etapa": "Ejemplo", "Decisión": "Datos ficticios para verificar el formato"}],
    )


def drawing_bytes(archive, sheet):
    path = _sheet_paths(archive)[sheet]
    rels = ET.fromstring(archive.read(_relationships_path(path)))
    rel = next(r for r in rels if r.attrib["Type"] == f"{REL_NS}/drawing")
    return archive.read(_target(path, rel.attrib["Target"]))


def test_default_template_is_the_new_company_format():
    assert b.BUNDLED_TEMPLATE_PATH.name == b.TEMPLATE_NAME
    assert b.default_template_bytes() == b.BUNDLED_TEMPLATE_PATH.read_bytes()
    wb = load_workbook(BytesIO(b.default_template_bytes()))
    assert wb[b.SHEET_FLOW20]["Q1"].value.startswith(b.TEMPLATE_CODE)
    assert wb[b.SHEET_NEW]["C3"].value == "QR"
    wb.close()


def test_new_company_output_preserves_layout_and_typed_dates():
    output = sample_output()
    original = load_workbook(BytesIO(b.default_template_bytes()))
    wb = load_workbook(BytesIO(output))
    expected = [(b.SHEET_NEW, 17, "Nuevo"), (b.SHEET_OLD, 18, "Viejo"), (b.SHEET_FLOW20, 20, "Flujo20")]
    for name, last_col, table in expected:
        ws = wb[name]
        code = ws.cell(1, last_col)
        assert code.value == original[name].cell(1, last_col - (3 if name == b.SHEET_FLOW20 else 2)).value
        assert code.value.startswith(b.TEMPLATE_CODE)
        assert ws.cell(2, last_col).value == "BV_C2.2_Contractual"
        assert ws["B1"].value == original[name]["B1"].value
        assert copy(ws["B1"].font) == copy(original[name]["B1"].font)
        assert copy(ws["C3"].fill) == copy(original[name]["C3"].fill)
        assert ws["C4"].value in {"QR-N", "QR-V", "QR-20"}
        assert ws.tables[table].ref.startswith("A3:")
        assert ws.print_title_rows == "$1:$3"
        assert "A1:A2" in {str(r) for r in ws.merged_cells.ranges}
        for header in b.FLOW20_TIMING_HEADERS if name == b.SHEET_FLOW20 else b.TIMING_HEADERS:
            col = b._header_map(ws, 3)[b.normalize_header(header)]
            assert isinstance(ws.cell(4, col).value, datetime)
            assert ws.cell(4, col).number_format == b.EXCEL_DATETIME_FORMAT
    assert wb[b.SHEET_NEW]["A4"].value == datetime(2026, 11, 23)
    assert wb[b.SHEET_NEW]["H4"].value == datetime(2026, 11, 23)
    assert wb[b.SHEET_NEW]["L4"].value == "3;0"
    assert wb[b.SHEET_NEW]["M4"].value == "0;2"
    assert wb[b.SHEET_NEW]["N4"].value == 5
    assert wb[b.SHEET_FLOW20]["K4"].value == "ODM-20"
    assert wb[b.SHEET_FLOW20]["G4"].value == b.FLOW20_SECONDARY_COMPLETED
    for name in ("Correo-Rotulacion", "Odm", "PNC"):
        assert b._detect_header_row(wb[name], {b.OPERATIONAL_SHEET_KEYS[name]}) == 3
        assert wb[name]["B1"].value == original[name]["B1"].value
    assert wb["DATOS"].sheet_state == "hidden"
    assert wb["Dinamicas"].sheet_state == "hidden"
    assert wb["INVR"].sheet_state == "hidden"
    assert wb[b.SHEET_GUIDED_AUDIT]["A1"].value == "Etapa"
    assert wb[b.SHEET_GUIDED_AUDIT]["B1"].value != b.TEMPLATE_CODE
    wb.close()
    original.close()


def test_native_vector_logos_remain_a_lossless_fallback(monkeypatch, tmp_path):
    monkeypatch.setattr(b, "BUNDLED_LOGO_PATH", tmp_path / "no-logo.png")
    output = sample_output()
    with zipfile.ZipFile(BytesIO(b.default_template_bytes())) as original, zipfile.ZipFile(BytesIO(output)) as result:
        for name in b.CODED_HEADER_SHEETS:
            assert drawing_bytes(result, name) == drawing_bytes(original, name)
        assert not any(name.startswith("xl/media/") for name in result.namelist())  # No OCA PNG fallback.
        assert b'<Types xmlns="' in result.read("[Content_Types].xml")
        types = ET.fromstring(result.read("[Content_Types].xml"))
        overrides = {entry.attrib.get("PartName") for entry in types}
        assert len([p for p in overrides if p and "/template_logo_" in p]) == 7
        # All internal drawing references resolve, including tables/comments.
        for path in _sheet_paths(result).values():
            rel_path = _relationships_path(path)
            if rel_path in result.namelist():
                for rel in ET.fromstring(result.read(rel_path)):
                    if rel.attrib.get("Type") == f"{REL_NS}/drawing":
                        assert _target(path, rel.attrib["Target"]) in result.namelist()


def test_supplied_transparent_png_is_bundled_and_keeps_its_aspect_ratio():
    output = sample_output()
    wb = load_workbook(BytesIO(output))
    for name in b.CODED_HEADER_SHEETS:
        assert len(wb[name]._images) == 1
        image = wb[name]._images[0]
        assert image.anchor._from.col == image.anchor._from.row == 0
        assert image.anchor._from.colOff > 0
        assert abs(image.anchor.ext.cx / image.anchor.ext.cy - 109 / 135) < 0.0001
    assert not wb[b.SHEET_GUIDED_AUDIT]._images
    wb.close()
    with zipfile.ZipFile(BytesIO(output)) as result:
        images = [result.read(n) for n in result.namelist() if n.startswith("xl/media/")]
        assert len(images) == 7
        assert all(data == b.BUNDLED_LOGO_PATH.read_bytes() for data in images)
        assert not any("template_logo_" in n for n in result.namelist())  # No duplicated vector logo.


def test_template_dropdowns_are_not_lost_on_import():
    wb = load_workbook(BytesIO(sample_output()))
    with zipfile.ZipFile(BytesIO(b.default_template_bytes())) as original:
        for name, path in _sheet_paths(original).items():
            expected = []
            for rule in ET.fromstring(original.read(path)).iter(f"{{{X14_NS}}}dataValidation"):
                expected.append((rule.findtext(f"{{{XM_NS}}}sqref"), rule.attrib["type"],
                                 rule.findtext(f"{{{X14_NS}}}formula1/{{{XM_NS}}}f")))
            actual = {(str(v.sqref), v.type, v.formula1) for v in wb[name].data_validations.dataValidation}
            assert set(expected).issubset(actual), name
    wb.close()


def test_new_template_keeps_empty_flow20_coded_for_old_input():
    source, rows = b.extract_workbook_rows(make_workbook())
    output = b.build_upload_workbook(source, b.validate_rows(rows, {}, {}), template_bytes=b.default_template_bytes())
    wb = load_workbook(BytesIO(output))
    assert wb[b.SHEET_FLOW20]["T1"].value.startswith(b.TEMPLATE_CODE)
    assert wb[b.SHEET_FLOW20]["C4"].value is None
    assert wb[b.SHEET_FLOW20].tables["Flujo20"].ref == "A3:T4"
    wb.close()


def test_input_can_already_use_the_new_coded_format():
    source = load_workbook(BytesIO(b.default_template_bytes()))
    source[b.SHEET_NEW]["A4"] = datetime(2026, 11, 23)
    source[b.SHEET_NEW]["C4"] = "QR-CODIFICADO"
    source[b.SHEET_NEW]["E4"] = "Ana"
    source[b.SHEET_NEW]["F4"] = "drawing_review"
    source[b.SHEET_NEW]["J4"] = "1000160001"
    stream = BytesIO()
    source.save(stream)
    source.close()
    prepared, rows = b.extract_workbook_rows(stream.getvalue())
    assert len(rows) == 1
    assert rows[0]["excel_row"] == 4
    validations = [replace(v, result=b.STATUS_MATCH) for v in b.validate_rows(rows, {}, {})]
    output = b.build_upload_workbook(prepared, validations, template_bytes=b.default_template_bytes())
    wb = load_workbook(BytesIO(output))
    assert wb[b.SHEET_NEW]["C4"].value == "QR-CODIFICADO"
    assert wb[b.SHEET_NEW]["A4"].value == datetime(2026, 11, 23)
    assert wb[b.SHEET_NEW]["A4"].number_format == "dd/mm/yyyy"
    assert wb[b.SHEET_NEW]["Q1"].value.startswith(b.TEMPLATE_CODE)
    wb.close()
