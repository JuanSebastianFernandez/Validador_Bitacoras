from __future__ import annotations

from datetime import date, datetime
from dataclasses import replace
from io import BytesIO, StringIO
from pathlib import Path
import zipfile

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill
from openpyxl.worksheet.table import Table
from streamlit.testing.v1 import AppTest
from streamlit.proto.Common_pb2 import FileURLs
from streamlit.runtime.scriptrunner import get_script_run_ctx
from streamlit.runtime.uploaded_file_manager import UploadedFile, UploadedFileRec
import streamlit as st

import bitacora as b


def workbook_with_formatting_tail(*, value_at_end=False):
    wb = Workbook()
    ws = wb.active
    ws.title = b.SHEET_NEW
    ws.append(["QR", "ESTADO DE LA OT", "USUARIO", "PROTOCOLO"])
    ws.append(["QR-N", "drawing_review", "Ana", "1000147997"])
    old = wb.create_sheet(b.SHEET_OLD)
    old.append(["QR", "ESTADO DE LA OT"])
    email = wb.create_sheet("Correo-Rotulacion")
    email.append([
        "ORIGEN DE ACTUALIZACIÓN", "ASUNTO", "ESTADO DE LA OT",
        "PROTOCOLO", "CANTIDAD PEL", "CANTIDAD TRATAS",
    ])
    email.append(["Solicitud por correo", "Asunto", "Finalizado", "1000147997", 0, 2])
    email["A3"] = False
    email["B4"] = '=""'
    email["C5"] = "#N/A"
    email["B2"].comment = Comment("Conservar comentario", "Prueba")
    email.add_table(Table(displayName="Correos", ref="A1:F5"))
    fill = PatternFill("solid", fgColor="FF0000")
    email["E3"].fill = fill  # Blank formatting inside the data span must survive.
    for row in range(6, 11):
        for col in range(1, 7):
            email.cell(row, col).fill = fill
    if value_at_end:
        email["A10"] = "Dato al final: no truncar"
    data = wb.create_sheet("DATOS")
    data.sheet_state = "hidden"
    data["A1"] = "Ana"
    result = BytesIO()
    wb.save(result)
    wb.close()
    return result.getvalue()


def enable_small_compaction(monkeypatch):
    monkeypatch.setattr(b, "WORKSHEET_COMPACTION_MIN_BYTES", 0)
    monkeypatch.setattr(b, "WORKSHEET_COMPACTION_MIN_ROWS", 3)


def test_compaction_preserves_values_formulas_styles_and_other_parts(monkeypatch):
    enable_small_compaction(monkeypatch)
    original = workbook_with_formatting_tail()
    cleaned, entries = b.compact_workbook_formatting(original)
    assert entries[0]["Hoja/Flujo"] == "Correo-Rotulacion"
    assert "5 filas" in entries[0]["Detalle"]
    with zipfile.ZipFile(BytesIO(original)) as before, zipfile.ZipFile(BytesIO(cleaned)) as after:
        assert before.namelist() == after.namelist()
        for name in before.namelist():
            if name != "xl/worksheets/sheet3.xml":
                assert before.read(name) == after.read(name)
        # Tail removal must not alter anything following sheetData, e.g. tables.
        assert before.read("xl/worksheets/sheet3.xml").split(b"</sheetData>")[1] == after.read("xl/worksheets/sheet3.xml").split(b"</sheetData>")[1]
    wb = load_workbook(BytesIO(cleaned))
    email = wb["Correo-Rotulacion"]
    assert email.max_row == 5
    assert email["E2"].value == 0
    assert email["A3"].value is False
    assert email["B4"].value == '=""'
    assert email["C5"].value == "#N/A"
    assert email["B2"].comment.text == "Conservar comentario"
    assert email["E3"].fill.fgColor.rgb == "00FF0000"
    assert email.tables["Correos"].ref == "A1:F5"
    assert wb["DATOS"].sheet_state == "hidden"
    wb.close()
    assert b.compact_workbook_formatting(cleaned) == (cleaned, [])


def test_compaction_does_not_cut_off_a_late_real_record(monkeypatch):
    enable_small_compaction(monkeypatch)
    original = workbook_with_formatting_tail(value_at_end=True)
    assert b.compact_workbook_formatting(original) == (original, [])


def test_extraction_returns_the_compact_working_copy(monkeypatch):
    enable_small_compaction(monkeypatch)
    original = workbook_with_formatting_tail()
    entries = []
    cleaned, rows = b.extract_workbook_rows(original, cleanup_entries=entries)
    assert len(entries) == 1
    assert rows[0]["qr"] == "QR-N"
    assert rows[0]["excel_row"] == 2
    assert rows[0]["protocol"] == "1000147997"
    email = b.evaluate_email_quantities(cleaned, {})
    assert len(email) == 1
    assert email[0].excel_row == 2
    assert email[0].reported_pel == 0
    assert email[0].reported_tratas == 2


def test_log_chunks_keep_results_and_reusable_streams():
    log = StringIO(
        "request_code;initial_state;final_state;date_created;username;user_created_last_name;comments\n"
        "QR-N;;drawing_request;2026-11-09 08:00:00;;Ana;\n"
        "QR-N;drawing_request;assigned_draftsman;2026-11-09 09:00:00;;Ana;\n"
        "QR-N;assigned_draftsman;drawing_review;2026-11-09 10:00:00;;Ana;Cierre\n"
        "QR-X;assigned_draftsman;verify_formats;2026-11-09 11:00:00;;Ana;Devolución\n"
    )
    day = date(2026, 11, 9)
    options = dict(include_timelines=True, include_rejection_comments=True, include_qr_aliases=True)
    small = b.read_log_validation_data(log, {"QR-N"}, day, chunksize=1, **options)
    default = b.read_log_validation_data(log, {"QR-N"}, day, **options)
    assert small == default
    assert not log.closed
    assert len(b.read_full_log_timelines(log, {"QR-N"})["QR-N"]) == 3
    assert not log.closed


def test_state_to_unreported_transition_does_not_keep_source_uploads(monkeypatch):
    enable_small_compaction(monkeypatch)
    cleaned, rows = b.extract_workbook_rows(workbook_with_formatting_tail())
    validations = [replace(item, result=b.STATUS_MATCH) for item in b.validate_rows(rows, {}, {})]
    case = b.UnreportedCase(
        qr="QR-X", state="verify_formats", occurred_at=datetime(2026, 11, 9, 11),
        comment="Devolución", actor="Ana",
    )
    app = AppTest.from_file(str(Path(__file__).resolve().parents[1] / "app.py"))
    app.session_state["audit_context"] = {
        "production_date": date(2026, 11, 9),
        "workbook_bytes": cleaned,
        "validations": validations,
        "latest_on_day": {},
        "timelines": {},
        "unreported_cases": [case],
        "audit_entries": [],
        "source_names": {"Bitácora": "Produccion 09-11-2026 OCA.xlsx", "Log PMS": "log.csv"},
    }
    app.session_state["audit_step"] = 1
    app.run(timeout=30)
    assert not app.exception
    assert len(app.get("file_uploader")) == 1  # Only SAP, disabled before step 3.
    assert app.get("file_uploader")[0].disabled
    next(button for button in app.button if button.label == "Cerrar estados y continuar").click().run(timeout=30)
    assert not app.exception
    assert app.session_state["audit_step"] == 2
    assert any("Casos no reportados" in item.value for item in app.markdown)
    next(radio for radio in app.radio if radio.label == "Decisión").set_value("No agregar").run()
    next(button for button in app.button if button.label == "Guardar decisión").click().run()
    next(button for button in app.button if button.label == "Cerrar casos no reportados y continuar").click().run(timeout=30)
    assert not app.exception
    assert app.session_state["audit_step"] == 3
    assert not app.get("file_uploader")[0].disabled


def test_successful_initialization_releases_only_its_session_uploads(monkeypatch):
    enable_small_compaction(monkeypatch)
    sources = {
        "1. Bitácora del coordinador": ("Produccion 09-11-2026.xlsx", workbook_with_formatting_tail()),
        "2. Log del PMS": ("log.csv", (
            "request_code;initial_state;final_state;date_created;user_created_last_name\n"
            "QR-N;assigned_draftsman;drawing_review;2026-11-09 10:00:00;Ana\n"
        ).encode()),
        "3. Exported Request del PMS": ("request.csv", b"request_code;contract_drawing_completed\nQR-N;true\n"),
    }
    captured = {}
    original_uploader = st.file_uploader
    original_button = st.button

    def source_uploader(label, *args, **kwargs):
        if label not in sources:
            return original_uploader(label, *args, **kwargs)
        name, data = sources[label]
        ctx = get_script_run_ctx()
        captured.update(manager=ctx.uploaded_file_mgr, session=ctx.session_id)
        record = UploadedFileRec(label, name, "application/octet-stream", data)
        ctx.uploaded_file_mgr.add_file(ctx.session_id, record)
        # Another session's upload must not be affected by the cleanup.
        ctx.uploaded_file_mgr.add_file("other-session", record)
        return UploadedFile(record, FileURLs())

    def button(label, *args, **kwargs):
        clicked = original_button(label, *args, **kwargs)
        return clicked or label == "Iniciar auditoría" and not kwargs.get("disabled")

    monkeypatch.setattr(st, "file_uploader", source_uploader)
    monkeypatch.setattr(st, "button", button)
    app = AppTest.from_file(str(Path(__file__).resolve().parents[1] / "app.py"))
    app.run(timeout=30)
    assert not app.exception
    assert app.session_state["audit_step"] == 1
    context = app.session_state["audit_context"]
    assert context["validations"][0].result == b.STATUS_MATCH
    assert context["workbook_cleanup"]
    assert len(context["source_names"]) == 3
    assert captured["manager"].get_files(captured["session"], list(sources)) == []
    assert len(captured["manager"].get_files("other-session", list(sources))) == 3
    assert len(app.get("file_uploader")) == 1
