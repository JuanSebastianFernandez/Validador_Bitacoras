from dataclasses import replace
from datetime import date
from io import BytesIO
from pathlib import Path

from openpyxl import load_workbook
from streamlit.testing.v1 import AppTest
import pytest

import bitacora as b
from test_bitacora import make_workbook


BAD = "1000168365-1000168367;1000168368-1000168369"
CORRECTED = "1000168365;1000168367;1000169366;1000169379"


def context_with_bad_protocol(*, ready=True, protocol=BAD):
    wb = load_workbook(BytesIO(make_workbook()))
    wb[b.SHEET_NEW]["J2"] = protocol
    wb[b.SHEET_OLD]["J2"] = "1000169999"
    source = BytesIO()
    wb.save(source)
    wb.close()
    data, rows = b.extract_workbook_rows(source.getvalue())
    validations = [replace(item, result=b.STATUS_MATCH) for item in b.validate_rows(rows, {}, {})]
    return {
        "production_date": date(2026, 11, 24),
        "working_bytes": data,
        "working_validations": validations,
        "protocol_prevalidation_done": ready,
        "protocol_decisions": {},
        "audit_entries": [],
    }


def start(context):
    app = AppTest.from_file(str(Path(__file__).resolve().parents[1] / "app.py"))
    app.session_state["audit_context"] = context
    app.session_state["audit_step"] = 3
    app.run(timeout=30)
    return app


def button(app, label):
    return next(item for item in app.button if item.label == label)


def set_protocol(app, value=CORRECTED):
    next(item for item in app.text_input if item.label.startswith("Protocolo(s), separados")).set_value(value)
    button(app, "Guardar decisión").click().run(timeout=30)


def assert_3a(app):
    assert not app.exception
    assert any(item.value.startswith("3A.") for item in app.subheader)
    assert not any(item.value.startswith("3B.") for item in app.subheader)
    assert not app.error


@pytest.mark.parametrize("malformed", [BAD, "1000168365-1000168367"])
def test_mixed_separators_reopen_3a_even_when_session_was_already_ready(malformed):
    app = start(context_with_bad_protocol(ready=True, protocol=malformed))
    assert_3a(app)
    assert not app.session_state["audit_context"]["protocol_prevalidation_done"]
    set_protocol(app)
    assert_3a(app)
    button(app, "Aplicar protocolos y preparar listado").click().run(timeout=30)
    assert not app.exception
    assert not app.error
    assert any(item.value.startswith("3B.") for item in app.subheader)
    context = app.session_state["audit_context"]
    assert context["protocol_prevalidation_done"]
    assert next(v.protocol for v in context["working_validations"] if v.qr == "QR-N") == CORRECTED
    assert len(context["audit_entries"]) == 1
    wb = load_workbook(BytesIO(context["working_bytes"]))
    assert wb[b.SHEET_NEW]["J2"].value == CORRECTED
    wb.close()


def test_stale_invalid_saved_decision_is_not_counted_as_reviewed():
    context = context_with_bad_protocol()
    key = "Flujo Nuevo|2|QR-N"
    context["protocol_decisions"][key] = {
        "action": "set", "current_protocol": BAD, "protocol": BAD,
        "sheet": b.SHEET_NEW, "excel_row": 2, "reference": "QR-N",
    }
    app = start(context)
    assert_3a(app)
    assert not any(item.label == "Aplicar protocolos y preparar listado" for item in app.button)
    assert app.get("progress")[0].proto.value == 0


def test_3a_does_not_accept_a_correction_with_remaining_hyphen_groups():
    app = start(context_with_bad_protocol())
    set_protocol(app, BAD)
    assert not app.exception
    assert any(item.value.startswith("3A.") for item in app.subheader)
    assert app.error
    assert app.session_state["audit_context"]["protocol_decisions"] == {}


def test_3a_does_not_replay_old_set_decisions_on_other_references():
    context = context_with_bad_protocol()
    old = next(v for v in context["working_validations"] if v.qr == "QR-V")
    context["protocol_decisions"][b.validation_audit_key(old)] = {
        "action": "set", "sheet": old.sheet, "excel_row": old.excel_row,
        "reference": old.qr, "current_protocol": "", "protocol": "1000169001",
    }
    app = start(context)
    assert_3a(app)
    set_protocol(app)
    button(app, "Aplicar protocolos y preparar listado").click().run(timeout=30)
    assert not app.exception
    context = app.session_state["audit_context"]
    assert next(v.protocol for v in context["working_validations"] if v.qr == "QR-V") == "1000169999"
    assert len(context["audit_entries"]) == 1


def test_pending_no_aplica_must_be_applied_and_traced_before_3b():
    app = start(context_with_bad_protocol(ready=False))
    assert_3a(app)
    app.radio[0].set_value("No aplica")
    button(app, "Guardar decisión").click().run(timeout=30)
    assert_3a(app)
    assert not app.session_state["audit_context"]["protocol_applied_decisions"]
    button(app, "Aplicar protocolos y preparar listado").click().run(timeout=30)
    assert not app.exception
    assert not app.error
    context = app.session_state["audit_context"]
    assert context["audit_entries"][0]["Decisión"] == "No aplica"
    assert len(context["protocol_applied_decisions"]) == 1
    assert any(item.value.startswith("3B.") for item in app.subheader)
    app.run(timeout=30)
    assert not any(item.value.startswith("3A.") for item in app.subheader)


def test_correcting_3a_refreshes_an_existing_export_without_reuploading():
    context = context_with_bad_protocol()
    context["protocol_counts"] = {
        "1000168365": b.ProtocolElementCounts(pel_ids=frozenset({"P-1"}), tratta_ids=frozenset({"T-1"})),
        "1000169999": b.ProtocolElementCounts(pel_ids=frozenset(), tratta_ids=frozenset()),
    }
    context["missing_protocols"] = [BAD]
    context["missing_protocol_decisions"] = {BAD: {"action": "no_apply", "justification": "Anterior"}}
    app = start(context)
    assert_3a(app)
    set_protocol(app, "1000168365")
    button(app, "Aplicar protocolos y preparar listado").click().run(timeout=30)
    assert not app.exception
    assert not app.error
    context = app.session_state["audit_context"]
    item = next(v for v in context["working_validations"] if v.qr == "QR-N")
    assert (item.export_pel, item.export_tratas) == (1, 1)
    assert context["missing_protocols"] == []
    assert context["missing_protocol_decisions"] == {}
    assert not context.get("protocol_counts_need_refresh")
