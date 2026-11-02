from __future__ import annotations

import base64
import zipfile
from datetime import date, datetime
from io import BytesIO, StringIO

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.worksheet.table import Table

from bitacora import (
    BitacoraError,
    COUNT_CONFIRMED,
    COUNT_AUDIT_NOT_APPLICABLE,
    COUNT_INVALID_PROTOCOL_PREFIX,
    COUNT_INVALID_PROTOCOL_SEPARATOR,
    COUNT_PROTOCOLS_INCOMPLETE,
    COUNT_UPDATED,
    DRAWING_DATETIME_HEADER,
    FLOW20_ODM_STATE,
    FLOW20_SECONDARY_COMPLETED,
    FLOW20_SECONDARY_PENDING,
    ODM_DATETIME_HEADER,
    RequestSnapshot,
    SHEET_FLOW20,
    LogCoverage,
    LogState,
    OBSERVATION_UPDATED,
    ProtocolElementCounts,
    STATUS_AUDIT_CORRECTED,
    STATUS_AUDIT_EXCLUDED,
    STATUS_AUDIT_KEEP,
    STATUS_FLOW20_AMBIGUOUS,
    STATUS_MATCH,
    STATUS_MISMATCH,
    STATUS_NO_EVENT,
    STATUS_PROFESSIONAL_MISMATCH,
    UnreportedCase,
    UPDATE_ORIGINS,
    add_unreported_audit_rows,
    apply_qr_alias_corrections,
    apply_protocol_audit_decisions,
    apply_protocol_audit_results,
    apply_protocol_counts_to_validations,
    apply_state_audit_decisions,
    build_upload_workbook,
    build_protocol_inventory_workbook,
    compare_protocol_quantities,
    evaluate_email_quantities,
    extract_workbook_rows,
    has_invalid_agp_separator,
    has_invalid_protocol_prefix,
    invalid_protocol_prefixes,
    is_closed_state_professional,
    same_professional,
    parse_production_date,
    protocol_counts_with_audit_decisions,
    protocol_quantity_breakdown,
    read_export_protocol_counts,
    read_full_log_timelines,
    read_log_validation_data,
    read_request_snapshots,
    read_log_states_with_coverage,
    read_latest_log_states,
    require_log_date_coverage,
    replace_protocol_references,
    resolve_protocol_counts,
    validate_rows,
    validation_audit_key,
    _clean_operational_sheets,
)


def make_workbook() -> bytes:
    wb = Workbook()
    ws_new = wb.active
    ws_new.title = "Flujo Nuevo"
    ws_new.append(
        [
            "FECHA DE LLEGADA",
            "ORIGEN DE ACTUALIZACIÓN",
            "QR",
            "ELEMENTOS ESTIMADOS",
            "USUARIO",
            "ESTADO DE LA OT",
            "ESTADO SECUNDARIO",
            "FECHA ENTREGADO",
            "CÓDIGO DEL INCREMENTO (AGP)",
            "PROTOCOLO",
            "ODM",
            "OBSERVACIONES",
            "CANTIDAD PEL",
            "CANTIDAD TRATAS",
            "TOTAL ELEMENTOS",
            "CONTROL ANS",
        ]
    )
    ws_new.append([None, "PQR", "QR-N", 2, "Ana", "drawing_review", None, None, None, 1000147997, None, "N/A", 3, 0, None, None])
    ws_new.append([None, "PQR", "QR-X", 2, "Ana", "verify_formats", None, None, None, None, None, "S/M", 1, 1, None, None])
    ws_new.add_table(Table(displayName="Nuevo", ref="A1:P3"))

    ws_old = wb.create_sheet("Flujo Viejo")
    ws_old.append(
        [
            "FECHA DE LLEGADA",
            "ORIGEN DE ACTUALIZACIÓN",
            "QR",
            "ELEMENTOS ESTIMADOS",
            "USUARIO",
            "ESTADO DE LA OT",
            "ESTADO SECUNDARIO",
            "FECHA ENTREGADO",
            "CÓDIGO DEL INCREMENTO (AGP)",
            "PROTOCOLO",
            "ODM",
            "NÚMERO DE TICKET BLOQUEO (INC)",
            "OBSERVACIONES",
            "CANTIDAD PEL",
            "CANTIDAD TRATAS",
            "TOTAL ELEMENTOS",
            "CONTROL ANS",
        ]
    )
    ws_old.append([None, "E1", "QR-V", 1, "Luis", "2.7-Dibujo finalizado", None, None, None, None, None, None, None, 2, 3, None, None])
    ws_old.add_table(Table(displayName="Viejo", ref="A1:Q2"))
    output = BytesIO()
    wb.save(output)
    return output.getvalue()


def make_flow20_workbook(
    *,
    secondary: str = FLOW20_SECONDARY_COMPLETED,
    professional: str = "Manuel Fernando Sanchez Mondragon",
) -> bytes:
    wb = load_workbook(BytesIO(make_workbook()), data_only=False)
    ws = wb.create_sheet(SHEET_FLOW20)
    ws.append(
        [
            "ORIGEN DE ACTUALIZACIÓN",
            "QR",
            "ELEMENTOS ESTIMADOS",
            "USUARIO",
            "ESTADO DE LA OT",
            "ESTADO SECUNDARIO",
            "CÓDIGO DEL INCREMENTO (AGP)",
            "PROTOCOLO",
            "ODM",
            "NÚMERO DE TICKET BLOQUEO (INC)",
            "OBSERVACIONES",
            "CANTIDAD PEL",
            "CANTIDAD TRATAS",
            "TOTAL ELEMENTOS",
        ]
    )
    ws.append(
        [
            "PQRs",
            "QR-20",
            3,
            professional,
            FLOW20_ODM_STATE,
            secondary,
            "2000140001",
            "1000147997",
            "ODM-20",
            None,
            None,
            None,
            None,
            None,
        ]
    )
    output = BytesIO()
    wb.save(output)
    wb.close()
    return output.getvalue()


def make_export_workbook() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(
        [
            "MANDT",
            "CHIAVE_RICERCA",
            "PROTOCOLLO_AP",
            "GIS_GUID",
            "ID_AGUI",
            "ID_CONTA",
            "ID_VALORE",
            "NOTA",
            "PEL_CAPTION",
            "TRATTA_CAPTION",
            "TIPO_PEL_O_TRATT",
            "GIS_GUID_AGUI",
            "TIPO_JSON",
        ]
    )
    for guid in range(101, 107):
        for item_type in ("CABINA", "CIRCUITO", "NODO"):
            ws.append(
                [200, f"KEY-{guid}", 1000147997, guid, item_type, 1, "X", None, guid, None, "P", None, None]
            )
    for guid in range(201, 203):
        for item_type in ("CABINA", "CIRCUITO", "TRONCO"):
            ws.append(
                [200, f"KEY-{guid}", 1000147997, guid, item_type, 1, "X", None, None, guid, "T", None, None]
            )
    wb.create_sheet("Sheet2")
    wb.create_sheet("Sheet3")
    output = BytesIO()
    wb.save(output)
    return output.getvalue()


def make_template_workbook() -> bytes:
    wb = Workbook()
    wb.remove(wb.active)
    sheet_headers = {
        "Flujo Nuevo": (
            "Nuevo",
            [
                "ORIGEN DE ACTUALIZACIÓN",
                "QR",
                "ELEMENTOS ESTIMADOS",
                "USUARIO",
                "ESTADO DE LA OT",
                "ESTADO SECUNDARIO",
                "CÓDIGO DEL INCREMENTO (AGP)",
                "PROTOCOLO",
                "OBSERVACIONES",
                "CANTIDAD PEL",
                "CANTIDAD TRATAS",
                "TOTAL ELEMENTOS",
            ],
        ),
        "Flujo Viejo": (
            "Viejo",
            [
                "ORIGEN DE ACTUALIZACIÓN",
                "QR",
                "ELEMENTOS ESTIMADOS",
                "USUARIO",
                "ESTADO DE LA OT",
                "ESTADO SECUNDARIO",
                "CÓDIGO DEL INCREMENTO (AGP)",
                "PROTOCOLO",
                "NÚMERO DE TICKET BLOQUEO (INC)",
                "OBSERVACIONES",
                "CANTIDAD PEL",
                "CANTIDAD TRATAS",
                "TOTAL ELEMENTOS",
            ],
        ),
        "Odm": (
            "Correos6",
            [
                "ORIGEN DE ACTUALIZACIÓN",
                "QR",
                "ODM",
                "ELEMENTOS ESTIMADOS",
                "USUARIO",
                "ESTADO DE LA OT",
                "OBSERVACIONES",
            ],
        ),
        "Correo-Rotulacion": (
            "Correos",
            [
                "ORIGEN DE ACTUALIZACIÓN",
                "ASUNTO",
                "ELEMENTOS ESTIMADOS",
                "USUARIO",
                "ESTADO DE LA OT",
                "ESTADO SECUNDARIO",
                "CÓDIGO DEL INCREMENTO (AGP)",
                "PROTOCOLO",
                "NÚMERO DE TICKET BLOQUEO (INC)",
                "OBSERVACIONES",
                "CANTIDAD PEL",
                "CANTIDAD TRATAS",
                "TOTAL ELEMENTOS",
            ],
        ),
        "PNC": (
            "Tabla4",
            [
                "FECHA DE LLEGADA",
                "ORIGEN DE ACTUALIZACIÓN",
                "QR/ASUNTO",
                "USUARIO QUE DEVUELVE",
                "MOTIVO DE DEVOLUCIÓN",
                "RESULTADO REVISIÓN",
                "REVISADO POR",
                "CONCEPTO DE REVISIÓN",
                "DESTINO/ACCIÓN POSTERIOR",
                "FECHA DE REVISIÓN",
            ],
        ),
    }
    for index, (sheet_name, (table_name, headers)) in enumerate(sheet_headers.items(), start=1):
        ws = wb.create_sheet(sheet_name)
        ws["A1"] = "#VALUE!"
        ws["C1"] = "FORMATO"
        ws["K1"] = "F-AGP-04"
        ws["K2"] = "Rev.02"
        ws["C3"] = "BITACORA SEGUIMIENTO GESTIÓN"
        ws["K3"] = "Fecha: 14/07/2026"
        ws["K4"] = f"Hoja {index} de 6"
        for column, header in enumerate(headers, start=1):
            ws.cell(5, column, header)
        ws.append([None] * len(headers))
        from openpyxl.utils import get_column_letter

        ws.add_table(Table(displayName=table_name, ref=f"A5:{get_column_letter(len(headers))}6"))

    wb.create_sheet("DATOS")
    wb.create_sheet("Control de Cambios")
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    one_pixel_png = base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8"
        "/x8AAwMCAO+/p9sAAAAASUVORK5CYII="
    )
    enriched = BytesIO()
    with zipfile.ZipFile(output, "r") as source, zipfile.ZipFile(enriched, "w") as target:
        for item in source.infolist():
            target.writestr(item, source.read(item.filename))
        target.writestr("xl/media/image1.png", one_pixel_png)
    return enriched.getvalue()


def coded_header_sheets() -> tuple[str, ...]:
    return (
        "Flujo Nuevo",
        "Flujo Viejo",
        "Odm",
        "Correo-Rotulacion",
        "PNC",
        "Control de Cambios",
    )


def test_parse_production_date_ignores_export_log_prefix():
    assert parse_production_date("produccion 18-06-2026.xlsx") == date(2026, 6, 18)


def test_latest_state_is_chronologically_last_inside_day():
    csv = StringIO(
        "request_code;initial_state;final_state;date_created\n"
        "QR-N;;assigned_draftsman;2026-06-18 08:00:00\n"
        "QR-N;;drawing_review;2026-06-18 10:00:00\n"
        "QR-N;;assigned_draftsman;2026-06-18 12:00:00\n"
        "QR-N;;verify_formats;2026-06-19 08:00:00\n"
    )
    on_day, as_of = read_latest_log_states(csv, {"QR-N"}, date(2026, 6, 18), chunksize=2)
    assert on_day["QR-N"].state == "assigned_draftsman"
    assert as_of["QR-N"].state == "assigned_draftsman"


def test_log_coverage_is_computed_from_all_rows_and_rejects_an_older_snapshot():
    csv = StringIO(
        "request_code;final_state;date_created\n"
        "OTHER;drawing_review;2026-08-01 08:00:00\n"
        "QR-N;drawing_review;2026-08-08 17:49:04\n"
    )
    on_day, as_of, coverage = read_log_states_with_coverage(
        csv,
        {"QR-N"},
        date(2026, 8, 11),
        chunksize=1,
    )
    assert not on_day
    assert as_of["QR-N"].state == "drawing_review"
    assert coverage.first_date == date(2026, 8, 1)
    assert coverage.last_date == date(2026, 8, 8)
    with pytest.raises(BitacoraError, match="termina el 08/08/2026"):
        require_log_date_coverage(coverage, date(2026, 8, 11))


def test_log_coverage_accepts_the_production_date():
    coverage = LogCoverage(
        first_datetime=datetime(2021, 7, 26, 17, 13),
        last_datetime=datetime(2026, 8, 13, 10, 0),
    )
    require_log_date_coverage(coverage, date(2026, 8, 11))


def test_unique_missing_final_qr_character_is_reconciled_without_becoming_unreported():
    csv = StringIO(
        "request_code;initial_state;final_state;date_created;comments;username;user_created_last_name\n"
        "QR-VD;2-Valorizado;2.5-Enviado a dibujo;2026-09-23 14:50:30;;system;System.PM\n"
        "QR-VD;2.5-Enviado a dibujo;2.6-Dibujante asignado;2026-09-23 15:06:00;;CO1014301686;Quiroga Lozada.Juan Pablo\n"
        "QR-VD;2.6-Dibujante asignado;2.7-Dibujo finalizado;2026-09-23 16:41:56;;1023014624;Porras Cano.Jenny Lorena\n"
    )
    on_day, as_of, _, unreported, timelines, aliases = read_log_validation_data(
        csv,
        {"QR-V"},
        date(2026, 9, 23),
        chunksize=1,
        allowed_states_by_qr={"QR-V": {"2.5-Enviado a dibujo", "2.6-Dibujante asignado", "2.7-Dibujo finalizado"}},
        include_timelines=True,
        include_qr_aliases=True,
    )

    assert aliases == {"QR-V": "QR-VD"}
    assert on_day["QR-VD"].state == "2.7-Dibujo finalizado"
    assert as_of["QR-VD"].occurred_at == datetime(2026, 9, 23, 16, 41, 56)
    assert [event.state for event in timelines["QR-VD"]] == [
        "2.5-Enviado a dibujo",
        "2.6-Dibujante asignado",
        "2.7-Dibujo finalizado",
    ]
    assert not unreported


def test_ambiguous_missing_final_qr_character_is_not_corrected_automatically():
    csv = StringIO(
        "request_code;initial_state;final_state;date_created;username;user_created_last_name\n"
        "QR-VA;2.6-Dibujante asignado;2.7-Dibujo finalizado;2026-09-23 16:00:00;1023014624;Porras Cano.Jenny Lorena\n"
        "QR-VB;2.6-Dibujante asignado;2.7-Dibujo finalizado;2026-09-23 17:00:00;1023014624;Porras Cano.Jenny Lorena\n"
    )
    *_, aliases = read_log_validation_data(
        csv,
        {"QR-V"},
        date(2026, 9, 23),
        chunksize=1,
        allowed_states_by_qr={"QR-V": {"2.7-Dibujo finalizado"}},
        include_qr_aliases=True,
    )

    assert aliases == {}


def test_existing_exact_qr_is_not_replaced_by_a_longer_reference():
    csv = StringIO(
        "request_code;initial_state;final_state;date_created;username;user_created_last_name\n"
        "QR-V;0-Por Valorizar;1-Enviado a territorio;2026-09-01 08:00:00;otro;Persona Externa\n"
        "QR-VD;2.6-Dibujante asignado;2.7-Dibujo finalizado;2026-09-23 16:00:00;1023014624;Porras Cano.Jenny Lorena\n"
    )
    *_, aliases = read_log_validation_data(
        csv,
        {"QR-V"},
        date(2026, 9, 23),
        chunksize=1,
        allowed_states_by_qr={"QR-V": {"2.7-Dibujo finalizado"}},
        include_qr_aliases=True,
    )

    assert aliases == {}


def test_qr_alias_correction_updates_workbook_rows_and_audit_trail():
    workbook_bytes, rows = extract_workbook_rows(make_workbook())
    corrected_bytes, corrected_rows, audit_entries = apply_qr_alias_corrections(
        workbook_bytes,
        rows,
        {"QR-V": "QR-VD"},
    )

    corrected_old = next(row for row in corrected_rows if row["sheet"] == "Flujo Viejo")
    assert corrected_old["qr"] == "QR-VD"
    wb = load_workbook(BytesIO(corrected_bytes), data_only=False)
    assert wb["Flujo Viejo"]["C2"].value == "QR-VD"
    wb.close()
    assert audit_entries[0]["Etapa"] == "Prevalidación QR"
    assert audit_entries[0]["QR/Referencia"] == "QR-V → QR-VD"


def test_unreported_cases_only_include_target_states_from_the_cutoff_day():
    csv = StringIO(
        "request_code;final_state;date_created;comments;username;user_created_last_name\n"
        "QR-IN;drawing_review;2026-08-11 08:00:00;Incluido;CO1049796769;Sanchez Mondragon.Manuel Fernando\n"
        "QR-MISS;verify_formats;2026-08-11 09:00:00;Corregir rótulo;999999;Persona Externa\n"
        "QR-MISS;verify_formats;2026-08-11 11:00:00;;999999;Persona Externa\n"
        "QR-MISS;assigned_draftsman;2026-08-11 12:00:00;;999999;Persona Externa\n"
        "QR-NAME;drawing_review;2026-08-11 10:45:00;;otro;SANCHEZ MONDRAGON.manuel fernando\n"
        "QR-FINAL;2.7-Dibujo finalizado;2026-08-11 10:30:00;;CO1049796769;Sanchez Mondragon.Manuel Fernando\n"
        "QR-EXCLUDED;drawing_review;2026-08-11 10:15:00;;999999;Persona Externa\n"
        "QR-OLD;drawing_review;2026-08-10 17:00:00;;CO1049796769;Sanchez Mondragon.Manuel Fernando\n"
        "QR-OTHER;assigned_draftsman;2026-08-11 13:00:00;;999999;Persona Externa\n"
        "QR-NEXT;2.6.1-Dibujo en revision;2026-08-12 07:00:00;;999999;Persona Externa\n"
    )
    _, _, coverage, cases = read_log_validation_data(
        csv,
        {"QR-IN"},
        date(2026, 8, 11),
        chunksize=2,
    )
    assert coverage.last_date == date(2026, 8, 12)
    assert [case.qr for case in cases] == ["QR-MISS", "QR-NAME", "QR-FINAL"]
    assert cases[0].state == "verify_formats"
    assert cases[0].occurred_at == datetime(2026, 8, 11, 11, 0)
    assert cases[0].comment == "Corregir rótulo"
    assert cases[1].actor == "SANCHEZ MONDRAGON.manuel fernando (otro)"
    assert cases[2].flow == "Flujo Viejo"
    assert "CO1049796769" in cases[2].actor


def test_closed_state_professionals_match_asignador_ids_and_pms_name_formats():
    assert is_closed_state_professional("1049796769")
    assert is_closed_state_professional("co1049796769")
    assert is_closed_state_professional("Sanchez mondragon.manuel fernando")
    assert is_closed_state_professional("MANUEL FERNANDO SÁNCHEZ MONDRAGÓN")
    assert is_closed_state_professional("otro", "Silva Silva.Johan Estiven")
    assert is_closed_state_professional("Leal Cardenas.Maria")
    assert not is_closed_state_professional("999999", "Persona Externa")


def test_same_professional_matches_reversed_names_accents_and_ids():
    workbook_name = "Manuel Fernando Sánchez Mondragón"
    assert same_professional(
        workbook_name,
        "CO1049796769",
        "SANCHEZ MONDRAGON.manuel fernando",
    )
    assert not same_professional(
        workbook_name,
        "CO1033708860",
        "Silva Silva.Johan Estiven",
    )
    assert same_professional(
        "Juan Sebastian Fernandez",
        "CO1013670569",
        "Sebastian Fernandez.Juan",
    )
    assert same_professional(
        "Juan Sebastian Fernandez Buitrago",
        "CO1013670569",
        "Sebastian Fernandez.Juan",
    )
    assert not is_closed_state_professional(
        "CO1013670569",
        "Sebastian Fernandez.Juan",
    )


def test_new_flow_timing_uses_the_new_cycle_and_the_same_professional():
    csv = StringIO(
        "request_code;initial_state;final_state;date_created;comments;username;user_created_last_name\n"
        "QR-CYCLE;order_creation;drawing_request;2026-09-11 07:30:00;;;Sistema\n"
        "QR-CYCLE;drawing_request;assigned_draftsman;2026-09-11 08:00:00;;;Sistema\n"
        "QR-CYCLE;drawing_completed;drawing_review;2026-09-11 10:00:00.001;;;Territorio\n"
        "QR-CYCLE;assigned_draftsman;drawing_review;2026-09-11 10:00:00.010;;CO1049796769;Sanchez Mondragon.Manuel Fernando\n"
        "QR-CYCLE;quality_review;assigned_draftsman;2026-09-11 14:00:00;;;Sistema\n"
        "QR-CYCLE;assigned_draftsman;verify_formats;2026-09-11 17:15:00;Ajustar rótulo;CO1049796769;Sanchez Mondragon.Manuel Fernando\n"
    )
    on_day, as_of, _, _, comments, timelines = read_log_validation_data(
        csv,
        {"QR-CYCLE"},
        date(2026, 9, 11),
        include_rejection_comments=True,
        include_timelines=True,
    )
    rows = [
        {
            "sheet": "Flujo Nuevo",
            "excel_row": 2,
            "qr": "QR-CYCLE",
            "reported_state": "verify_formats",
            "professional": "Manuel Fernando Sanchez Mondragon",
        }
    ]
    item = validate_rows(
        rows,
        on_day,
        as_of,
        rejection_comments=comments,
        timelines=timelines,
        production_date=date(2026, 9, 11),
    )[0]
    assert item.result == STATUS_MATCH
    assert item.arrival_state == "assigned_draftsman"
    assert item.arrival_datetime == datetime(2026, 9, 11, 14, 0)
    assert item.exit_datetime == datetime(2026, 9, 11, 17, 15)
    assert "Manuel Fernando" in item.pms_actor
    assert item.pms_observation == "Ajustar rótulo"


def test_new_flow_reentry_uses_previous_exit_even_if_actor_is_not_authorized_for_unreported():
    qr = "SODSGISMTU-RC894_72056-01-2026AV14"
    csv = StringIO(
        "request_code;initial_state;final_state;date_created;comments;username;user_created_last_name\n"
        f"{qr};drawing_request;assigned_draftsman;2026-08-27 10:46:58;;;Sistema\n"
        f"{qr};assigned_draftsman;verify_formats;2026-09-03 12:21:19;Primera devolución;CO1013670569;Sebastian Fernandez.Juan\n"
        f"{qr};verify_formats;awaiting_terr_validation;2026-09-03 13:15:41;;;Calidad\n"
        f"{qr};awaiting_terr_validation;awaiting_valuation;2026-09-07 15:39:52;;;Territorio\n"
        f"{qr};drawing_request;assigned_draftsman;2026-09-09 07:57:13.470309;;;Sistema\n"
        f"{qr};awaiting_valuation;assigned_draftsman;2026-09-09 07:57:13.490451;;;Sistema\n"
        f"{qr};assigned_draftsman;verify_formats;2026-09-09 12:21:36.391562;Nueva devolución;CO1013670569;Sebastian Fernandez.Juan\n"
    )
    on_day, as_of, _, _, comments, timelines = read_log_validation_data(
        csv,
        {qr},
        date(2026, 9, 9),
        include_rejection_comments=True,
        include_timelines=True,
    )
    item = validate_rows(
        [
            {
                "sheet": "Flujo Nuevo",
                "excel_row": 2,
                "qr": qr,
                "reported_state": "verify_formats",
                "professional": "Juan Sebastian Fernandez Buitrago",
            }
        ],
        on_day,
        as_of,
        rejection_comments=comments,
        timelines=timelines,
        production_date=date(2026, 9, 9),
    )[0]

    assert item.arrival_state == "assigned_draftsman"
    assert item.arrival_datetime == datetime(2026, 9, 9, 7, 57, 13, 470309)
    assert item.exit_datetime == datetime(2026, 9, 9, 12, 21, 36, 391562)
    assert item.pms_observation == "Nueva devolución"


def test_old_flow_direct_reentry_ignores_13_and_accepts_state_2_as_exit():
    csv = StringIO(
        "request_code;initial_state;final_state;date_created;comments;username;user_created_last_name\n"
        "QR-OLD-CYCLE;1.4;2.5-Enviado a dibujo;2026-09-11 07:00:00;;;Sistema\n"
        "QR-OLD-CYCLE;2.5-Enviado a dibujo;2.6-Dibujante asignado;2026-09-11 08:00:00;;;Sistema\n"
        "QR-OLD-CYCLE;2.6-Dibujante asignado;2-Valorizado;2026-09-11 09:30:00;Devuelto;CO1049796769;Sanchez Mondragon.Manuel Fernando\n"
        "QR-OLD-CYCLE;2-Valorizado;1.3-RG Gestionado EEC;2026-09-11 10:00:00;;;Calidad E1\n"
        "QR-OLD-CYCLE;1.5;2.6-Dibujante asignado;2026-09-11 13:00:00;;;Sistema\n"
        "QR-OLD-CYCLE;2.6-Dibujante asignado;2.7-Dibujo finalizado;2026-09-11 16:00:00;;CO1049796769;Sanchez Mondragon.Manuel Fernando\n"
    )
    on_day, as_of, _, _, timelines = read_log_validation_data(
        csv,
        {"QR-OLD-CYCLE"},
        date(2026, 9, 11),
        include_timelines=True,
    )
    rows = [
        {
            "sheet": "Flujo Viejo",
            "excel_row": 2,
            "qr": "QR-OLD-CYCLE",
            "reported_state": "2.7-Dibujo finalizado",
            "professional": "Manuel Fernando Sanchez Mondragon",
        }
    ]
    item = validate_rows(
        rows,
        on_day,
        as_of,
        timelines=timelines,
        production_date=date(2026, 9, 11),
    )[0]
    assert item.result == STATUS_MATCH
    assert item.arrival_state == "2.6-Dibujante asignado"
    assert item.arrival_datetime == datetime(2026, 9, 11, 13, 0)
    assert item.exit_datetime == datetime(2026, 9, 11, 16, 0)


def test_output_by_another_actor_requires_state_audit():
    csv = StringIO(
        "request_code;initial_state;final_state;date_created;username;user_created_last_name\n"
        "QR-ACTOR;order_creation;drawing_request;2026-09-11 08:00:00;;Sistema\n"
        "QR-ACTOR;assigned_draftsman;drawing_review;2026-09-11 12:00:00;999999;Persona Externa\n"
    )
    on_day, as_of, _, _, timelines = read_log_validation_data(
        csv,
        {"QR-ACTOR"},
        date(2026, 9, 11),
        include_timelines=True,
    )
    item = validate_rows(
        [
            {
                "sheet": "Flujo Nuevo",
                "excel_row": 2,
                "qr": "QR-ACTOR",
                "reported_state": "drawing_review",
                "professional": "Manuel Fernando Sanchez Mondragon",
            }
        ],
        on_day,
        as_of,
        timelines=timelines,
        production_date=date(2026, 9, 11),
    )[0]
    assert item.result == STATUS_PROFESSIONAL_MISMATCH
    assert item.pms_actor == "Persona Externa (999999)"


def test_output_adds_unreported_cases_as_the_last_sheet():
    workbook = make_workbook()
    _, rows = extract_workbook_rows(workbook)
    event = LogState("drawing_review", datetime(2026, 6, 18, 10, 0))
    validations = validate_rows(
        rows,
        {"QR-N": event},
        {"QR-N": event},
    )
    cases = [
        UnreportedCase(
            qr="PREDUOBMT-303241-03-2VE28",
            state="verify_formats",
            occurred_at=datetime(2026, 6, 18, 14, 35),
            comment="Corregir formato",
            actor="Sanchez Mondragon.Manuel Fernando (CO1049796769)",
        )
    ]
    output = build_upload_workbook(
        workbook,
        validations,
        matches_only=False,
        unreported_cases=cases,
    )
    wb = load_workbook(BytesIO(output), data_only=False)
    assert wb.sheetnames[-1] == "Casos no reportados"
    ws = wb["Casos no reportados"]
    assert [cell.value for cell in ws[1]] == [
        "Flujo",
        "QR",
        "Estado PMS del día",
        "Fecha/hora PMS",
        "Usuario PMS",
        "Observación PMS",
        "Hallazgo",
    ]
    assert ws["B2"].value == "PREDUOBMT-303241-03-2VE28"
    assert ws["C2"].value == "verify_formats"
    assert ws["D2"].value == datetime(2026, 6, 18, 14, 35)
    assert ws["E2"].value == "Sanchez Mondragon.Manuel Fernando (CO1049796769)"
    assert ws["F2"].value == "Corregir formato"
    assert ws.tables["CasosNoReportados"].ref == "A1:G2"


def test_guided_state_decisions_keep_correct_and_exclude_rows():
    workbook = make_workbook()
    _, rows = extract_workbook_rows(workbook)
    event = LogState(
        "verify_formats",
        datetime(2026, 6, 18, 10, 0),
        comment="Ajustar formato del rótulo",
    )
    validations = validate_rows(rows, {"QR-N": event}, {"QR-N": event})
    by_qr = {item.qr: item for item in validations}
    decisions = {
        validation_audit_key(by_qr["QR-N"]): {
            "action": "correct",
            "state": "verify_formats",
        },
        validation_audit_key(by_qr["QR-X"]): {"action": "keep"},
        validation_audit_key(by_qr["QR-V"]): {
            "action": "exclude",
            "detail": "No corresponde al corte",
        },
    }

    output, audited, entries = apply_state_audit_decisions(
        workbook,
        validations,
        {"QR-N": event},
        decisions,
    )
    audited_by_qr = {item.qr: item for item in audited}
    assert audited_by_qr["QR-N"].result == STATUS_AUDIT_CORRECTED
    assert audited_by_qr["QR-N"].reported_state == "verify_formats"
    assert audited_by_qr["QR-N"].pms_observation == "Ajustar formato del rótulo"
    assert audited_by_qr["QR-X"].result == STATUS_AUDIT_KEEP
    assert audited_by_qr["QR-V"].result == STATUS_AUDIT_EXCLUDED
    assert len(entries) == 3

    wb = load_workbook(BytesIO(output), data_only=False)
    ws = wb["Flujo Nuevo"]
    assert ws["F2"].value == "verify_formats"
    assert ws["I2"].value is None
    assert ws["J2"].value == 1000147997
    assert ws["L2"].value == "Ajustar formato del rótulo"
    assert ws["M2"].value is None
    assert ws["N2"].value is None


def test_kept_rejection_uses_latest_comment_from_same_state_despite_later_event():
    workbook = make_workbook()
    _, rows = extract_workbook_rows(workbook)
    csv = StringIO(
        "request_code;final_state;date_created;comments\n"
        "QR-X;verify_formats;2026-06-18 09:00:00;Primera devolución\n"
        "QR-X;verify_formats;2026-06-18 10:00:00;Última devolución correcta\n"
        "QR-X;assigned_draftsman;2026-06-18 11:00:00;Respuesta posterior\n"
    )
    on_day, as_of, _, _, rejection_comments = read_log_validation_data(
        csv,
        {"QR-X"},
        date(2026, 6, 18),
        include_rejection_comments=True,
    )
    assert on_day["QR-X"].state == "assigned_draftsman"
    assert (
        rejection_comments[("QR-X", "verify_formats")].comment
        == "Última devolución correcta"
    )
    validations = validate_rows(
        rows,
        on_day,
        as_of,
        rejection_comments=rejection_comments,
    )
    target = next(item for item in validations if item.qr == "QR-X")
    assert target.result == STATUS_MISMATCH
    assert target.pms_observation == "Última devolución correcta"

    output, audited, _ = apply_state_audit_decisions(
        workbook,
        validations,
        on_day,
        {validation_audit_key(target): {"action": "keep"}},
        rejection_comments,
    )
    kept = next(item for item in audited if item.qr == "QR-X")
    assert kept.result == STATUS_AUDIT_KEEP
    assert kept.pms_observation == "Última devolución correcta"
    assert kept.observation_result == OBSERVATION_UPDATED
    wb = load_workbook(BytesIO(output), data_only=False)
    assert wb["Flujo Nuevo"]["L3"].value == "Última devolución correcta"


def test_agp_lists_require_semicolon_and_final_state_enforces_it():
    assert has_invalid_agp_separator("2000133299,2000134328")
    assert has_invalid_agp_separator("2000133299 2000134328")
    assert has_invalid_agp_separator("AGP-900,AGP-901")
    assert not has_invalid_agp_separator("2000133299;2000134328")
    assert not has_invalid_agp_separator("AGP-900")

    workbook = make_workbook()
    _, rows = extract_workbook_rows(workbook)
    validations = validate_rows(rows, {}, {})
    target = next(item for item in validations if item.qr == "QR-X")
    with pytest.raises(BitacoraError, match="varios AGP"):
        apply_state_audit_decisions(
            workbook,
            validations,
            {},
            {
                validation_audit_key(target): {
                    "action": "correct",
                    "state": "drawing_review",
                    "agp": "2000133299,2000134328",
                    "protocol": "1000000001;1000000002",
                }
            },
        )


def test_protocol_separator_rejects_alternatives_and_concatenated_codes():
    from bitacora import has_invalid_protocol_separator

    assert has_invalid_protocol_separator("1000162837,1000162838")
    assert has_invalid_protocol_separator("1000162837|1000162838")
    assert has_invalid_protocol_separator("1000162837\n1000162838")
    assert has_invalid_protocol_separator("10001628371000162838")
    assert not has_invalid_protocol_separator("1000162837;1000162838")


def test_protocols_must_begin_with_1000_in_every_token():
    assert has_invalid_protocol_prefix("9990162837")
    assert has_invalid_protocol_prefix("1000162837;2000162838")
    assert invalid_protocol_prefixes("1000162837;2000162838") == ["2000162838"]
    assert not has_invalid_protocol_prefix("1000162837;1000162838")

    pel, tratas, result = compare_protocol_quantities(
        "9990162837",
        0,
        0,
        {},
    )
    assert (pel, tratas) == (None, None)
    assert result == COUNT_INVALID_PROTOCOL_PREFIX


def test_protocol_audit_rejects_a_value_that_does_not_begin_with_1000():
    workbook = make_workbook()
    _, rows = extract_workbook_rows(workbook)
    validations = validate_rows(rows, {}, {})
    target = next(item for item in validations if item.qr == "QR-X")
    with pytest.raises(BitacoraError, match="iniciar por 1000"):
        apply_protocol_audit_decisions(
            workbook,
            validations,
            {
                "decision": {
                    "action": "set",
                    "sheet": target.sheet,
                    "excel_row": target.excel_row,
                    "reference": target.qr,
                    "protocol": "9990000001",
                }
            },
        )


def test_missing_protocol_can_be_replaced_in_all_flow_and_email_references():
    wb = load_workbook(BytesIO(make_workbook()), data_only=False)
    wb["Flujo Nuevo"]["J2"] = "1000999999;1000147997"
    wb["Flujo Viejo"]["J2"] = "1000999999"
    email = wb.create_sheet("Correo-Rotulacion")
    email.append(
        [
            "ORIGEN DE ACTUALIZACIÓN",
            "ASUNTO",
            "ESTADO DE LA OT",
            "PROTOCOLO",
            "CANTIDAD PEL",
            "CANTIDAD TRATAS",
        ]
    )
    email.append(
        [
            "Solicitud por correo",
            "Correo con protocolo errado",
            "Finalizado",
            "1000999999",
            0,
            0,
        ]
    )
    source = BytesIO()
    wb.save(source)
    wb.close()

    workbook = source.getvalue()
    _, rows = extract_workbook_rows(workbook)
    validations = validate_rows(rows, {}, {})
    output, corrected, entries = replace_protocol_references(
        workbook,
        validations,
        {"1000999999": "1000162837;1000162838"},
    )

    corrected_by_qr = {item.qr: item.protocol for item in corrected}
    assert corrected_by_qr["QR-N"] == "1000162837;1000162838;1000147997"
    assert corrected_by_qr["QR-V"] == "1000162837;1000162838"
    assert len(entries) == 3
    assert all(item["Decisión"] == "Corregir protocolo" for item in entries)

    saved = load_workbook(BytesIO(output), data_only=False)
    assert saved["Flujo Nuevo"]["J2"].value == "1000162837;1000162838;1000147997"
    assert saved["Flujo Viejo"]["J2"].value == "1000162837;1000162838"
    assert saved["Correo-Rotulacion"]["D2"].value == "1000162837;1000162838"
    saved.close()

    with pytest.raises(BitacoraError, match="iniciar por 1000"):
        replace_protocol_references(
            output,
            corrected,
            {"1000162837": "9000162837"},
        )
    with pytest.raises(BitacoraError, match="diferente"):
        replace_protocol_references(
            output,
            corrected,
            {"1000162837": "1000162837"},
        )


def test_protocol_inventory_is_an_excel_table_with_unique_protocols_and_detail():
    output = build_protocol_inventory_workbook(
        [
            {
                "Protocolo": "1000162837",
                "Hoja": "Flujo Nuevo",
                "QR o referencia": "QR-1",
                "Estado": "drawing_review",
            },
            {
                "Protocolo": "1000162837",
                "Hoja": "Correo-Rotulacion",
                "QR o referencia": "CORREO-1",
                "Estado": "Finalizado",
            },
            {
                "Protocolo": "1000162838",
                "Hoja": "Flujo Viejo",
                "QR o referencia": "QR-2",
                "Estado": "2.7-Dibujo finalizado",
            },
        ]
    )
    wb = load_workbook(BytesIO(output), data_only=False)
    assert wb.sheetnames == ["Protocolos", "Detalle"]
    ws = wb["Protocolos"]
    assert [ws["A2"].value, ws["A3"].value] == ["1000162837", "1000162838"]
    assert ws.tables["ListadoProtocolos"].ref == "A1:A3"
    detail = wb["Detalle"]
    assert detail.max_row == 4
    assert detail.tables["DetalleProtocolos"].ref == "A1:D4"
    assert UPDATE_ORIGINS[0] == "E1"
    assert "Solicitud por correo" in UPDATE_ORIGINS


def test_guided_unreported_case_can_be_added_to_selected_flow():
    workbook = make_workbook()
    _, rows = extract_workbook_rows(workbook)
    validations = validate_rows(rows, {}, {})
    case = UnreportedCase(
        qr="QR-NUEVO",
        state="drawing_review",
        occurred_at=datetime(2026, 6, 18, 15, 30),
        actor="Profesional PMS",
    )
    output, audited, entries = add_unreported_audit_rows(
        workbook,
        validations,
        [case],
        {
            case.qr: {
                "action": "add",
                "sheet": "Flujo Nuevo",
                "agp": "AGP-901",
                "protocol": "1000000002;1000000003",
                "origin": "Gestión del día",
                "user": "Profesional PMS",
                "elements": 4,
            }
        },
    )
    added = audited[-1]
    assert added.qr == "QR-NUEVO"
    assert added.result == STATUS_AUDIT_CORRECTED
    assert added.protocol == "1000000002;1000000003"
    assert entries[0]["Decisión"] == "Agregar"

    wb = load_workbook(BytesIO(output), data_only=False)
    ws = wb["Flujo Nuevo"]
    assert ws["C4"].value == "QR-NUEVO"
    assert ws["I4"].value == "AGP-901"
    assert ws["J4"].value == "1000000002;1000000003"


def test_unreported_rejection_does_not_require_agp_or_protocol():
    workbook = make_workbook()
    _, rows = extract_workbook_rows(workbook)
    validations = validate_rows(rows, {}, {})
    case = UnreportedCase(
        qr="QR-DEVUELTO",
        state="verify_formats",
        occurred_at=datetime(2026, 6, 18, 16, 0),
        comment="Última devolución del mismo estado",
        actor="Profesional PMS",
    )
    output, audited, entries = add_unreported_audit_rows(
        workbook,
        validations,
        [case],
        {
            case.qr: {
                "action": "add",
                "sheet": "Flujo Nuevo",
                "origin": "PQRs",
                "user": "Profesional PMS",
                "elements": 2,
            }
        },
    )
    added = audited[-1]
    assert added.protocol == ""
    assert added.pms_observation == "Última devolución del mismo estado"
    assert entries[0]["AGP"] == ""
    assert entries[0]["Protocolo"] == ""
    wb = load_workbook(BytesIO(output), data_only=False)
    ws = wb["Flujo Nuevo"]
    assert ws["I4"].value is None
    assert ws["J4"].value is None
    assert ws["L4"].value == "Última devolución del mismo estado"


def test_guided_protocol_decision_updates_workbook_and_validation():
    workbook = make_workbook()
    _, rows = extract_workbook_rows(workbook)
    validations = validate_rows(rows, {}, {})
    target = next(item for item in validations if item.qr == "QR-X")
    decision_key = f"{target.sheet}|{target.excel_row}|{target.qr}"
    output, audited, entries = apply_protocol_audit_decisions(
        workbook,
        validations,
        {
            decision_key: {
                "action": "set",
                "sheet": target.sheet,
                "excel_row": target.excel_row,
                "reference": target.qr,
                "protocol": "1000000004",
            }
        },
    )
    assert next(item for item in audited if item.qr == "QR-X").protocol == "1000000004"
    assert entries[0]["Decisión"] == "Protocolo diligenciado"
    wb = load_workbook(BytesIO(output), data_only=False)
    assert wb["Flujo Nuevo"]["J3"].value == "1000000004"


def test_protocol_no_apply_decision_is_resolved_in_final_reports():
    workbook = make_workbook()
    _, rows = extract_workbook_rows(workbook)
    validations = validate_rows(rows, {}, {})
    target = next(item for item in validations if item.qr == "QR-N")
    audited, emails = apply_protocol_audit_results(
        validations,
        [],
        {
            "decision": {
                "action": "no_apply",
                "sheet": target.sheet,
                "excel_row": target.excel_row,
            }
        },
        set(),
        {},
    )
    assert not emails
    assert (
        next(item for item in audited if item.qr == "QR-N").count_result
        == COUNT_AUDIT_NOT_APPLICABLE
    )


def test_output_adds_guided_audit_as_the_last_sheet():
    workbook = make_workbook()
    _, rows = extract_workbook_rows(workbook)
    validations = validate_rows(rows, {}, {})
    output = build_upload_workbook(
        workbook,
        validations,
        matches_only=False,
        audit_entries=[
            {
                "Etapa": "Estados",
                "Hoja/Flujo": "Flujo Nuevo",
                "QR/Referencia": "QR-N",
                "Decisión": "Conservar bitácora",
                "Estado inicial": "drawing_review",
                "Estado final": "drawing_review",
                "AGP": "AGP-1",
                "Protocolo": "1000147997",
                "Detalle": "Validado por el auditor",
            }
        ],
    )
    wb = load_workbook(BytesIO(output), data_only=False)
    assert wb.sheetnames[-1] == "Auditoria guiada"
    ws = wb["Auditoria guiada"]
    assert ws["A2"].value == "Estados"
    assert ws["C2"].value == "QR-N"
    assert ws.tables["AuditoriaGuiada"].ref == "A1:K2"


def test_output_adds_protocol_justification_and_image_evidence_sheet():
    workbook = make_workbook()
    _, rows = extract_workbook_rows(workbook)
    validations = validate_rows(rows, {}, {})
    one_pixel_png = base64.b64decode(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8"
        "/x8AAwMCAO+/p9sAAAAASUVORK5CYII="
    )
    output = build_upload_workbook(
        workbook,
        validations,
        matches_only=False,
        protocol_evidence=[
            {
                "protocol": "1000162999",
                "decision": "Cantidades manuales",
                "pel": 2,
                "tratas": 1,
                "justification": "AGP no permitió trasladar, validado visualmente.",
                "filename": "soporte.png",
                "image_bytes": one_pixel_png,
            }
        ],
    )

    wb = load_workbook(BytesIO(output), data_only=False)
    assert wb.sheetnames[-1] == "Soportes protocolos"
    ws = wb["Soportes protocolos"]
    assert ws["A2"].value == "1000162999"
    assert ws["C2"].value == 2
    assert ws["D2"].value == 1
    assert ws["E2"].value == "AGP no permitió trasladar, validado visualmente."
    assert ws["F2"].value == "soporte.png"
    assert ws.tables["SoportesProtocolos"].ref == "A1:G2"
    assert len(ws._images) == 1


def test_semicolon_protocols_show_partial_totals_when_some_are_missing():
    counts = {
        "1000162837": ProtocolElementCounts(
            pel_ids=frozenset({"P-1", "P-2"}),
            tratta_ids=frozenset({"T-1"}),
        ),
        "1000162838": ProtocolElementCounts(
            pel_ids=frozenset({"P-3", "P-4"}),
            tratta_ids=frozenset(),
        ),
    }
    pel, tratas, protocols, missing = resolve_protocol_counts(
        "1000162837;1000162838",
        counts,
    )
    assert (pel, tratas) == (4, 1)
    assert protocols == ["1000162837", "1000162838"]
    assert missing == []

    pel, tratas, protocols, missing = resolve_protocol_counts(
        "1000162837;1000162999",
        counts,
    )
    assert (pel, tratas) == (2, 1)
    assert protocols == ["1000162837", "1000162999"]
    assert missing == ["1000162999"]

    pel, tratas, protocols, missing = resolve_protocol_counts(
        "1000162998;1000162999",
        counts,
    )
    assert (pel, tratas) == (None, None)
    assert protocols == ["1000162998", "1000162999"]
    assert missing == ["1000162998", "1000162999"]


def test_semicolon_protocols_keep_positional_pel_and_trata_breakdown():
    counts = {
        "1000162837": ProtocolElementCounts(
            pel_ids=frozenset({"P-1"}),
            tratta_ids=frozenset({"T-1"}),
        ),
        "1000162838": ProtocolElementCounts(
            pel_ids=frozenset(),
            tratta_ids=frozenset({"T-2"}),
        ),
    }

    assert protocol_quantity_breakdown(
        "1000162837;1000162838",
        counts,
    ) == ("1;0", "1;1")
    assert protocol_quantity_breakdown("1000162837", counts) == (1, 1)
    assert protocol_quantity_breakdown(
        "1000162837;1000162999",
        counts,
    ) == ("1;?", "1;?")


def test_missing_protocol_audit_decisions_supply_counts_and_no_apply():
    resolved, no_apply = protocol_counts_with_audit_decisions(
        {},
        {
            "1000162837": {"action": "manual", "pel": 2, "tratas": 1},
            "1000162838": {"action": "no_apply"},
        },
    )

    assert resolved["1000162837"].pel == 2
    assert resolved["1000162837"].tratas == 1
    assert resolved["1000162838"].pel == 0
    assert resolved["1000162838"].tratas == 0
    assert no_apply == {"1000162838"}
    assert protocol_quantity_breakdown(
        "1000162837;1000162838",
        resolved,
    ) == ("2;0", "1;0")


def test_flow_output_writes_semicolon_counts_and_numeric_total():
    source = load_workbook(BytesIO(make_workbook()), data_only=False)
    source["Flujo Nuevo"]["J2"] = "1000162837;1000162838"
    source["Flujo Nuevo"]["M2"] = 1
    source["Flujo Nuevo"]["N2"] = 2
    prepared = BytesIO()
    source.save(prepared)
    source.close()
    workbook = prepared.getvalue()
    counts = {
        "1000162837": ProtocolElementCounts(
            pel_ids=frozenset({"P-1"}),
            tratta_ids=frozenset({"T-1"}),
        ),
        "1000162838": ProtocolElementCounts(
            pel_ids=frozenset(),
            tratta_ids=frozenset({"T-2"}),
        ),
    }
    _, rows = extract_workbook_rows(workbook)
    event = LogState("drawing_review", datetime(2026, 6, 18, 10, 0))
    validations = validate_rows(
        rows,
        {"QR-N": event},
        {"QR-N": event},
        protocol_counts=counts,
    )
    output = build_upload_workbook(workbook, validations, matches_only=True)

    wb = load_workbook(BytesIO(output), data_only=False)
    ws = wb["Flujo Nuevo"]
    headers = {cell.value: cell.column for cell in ws[1]}
    assert ws.cell(2, headers["CANTIDAD PEL"]).value == "1;0"
    assert ws.cell(2, headers["CANTIDAD TRATAS"]).value == "1;1"
    assert ws.cell(2, headers["TOTAL ELEMENTOS"]).value == 3


def test_validation_reports_incomplete_protocols_and_requires_semicolon():
    base_row = {
        "sheet": "Flujo Nuevo",
        "excel_row": 2,
        "qr": "PREDUOBMT-303241-03-2VE28",
        "reported_state": "drawing_review",
        "protocol": "1000162837;1000162999",
        "reported_pel": 4,
        "reported_tratas": 1,
        "reported_observation": "",
    }
    event = LogState("drawing_review", datetime(2026, 8, 11, 10, 0))
    counts = {
        "1000162837": ProtocolElementCounts(
            pel_ids=frozenset({"P-1", "P-2"}),
            tratta_ids=frozenset({"T-1"}),
        )
    }
    validation = validate_rows(
        [base_row],
        {base_row["qr"]: event},
        {base_row["qr"]: event},
        protocol_counts=counts,
    )[0]
    assert validation.count_result == (
        f"{COUNT_PROTOCOLS_INCOMPLETE}: 1000162999"
    )
    assert validation.export_pel == 2
    assert validation.export_tratas == 1

    invalid_separator_row = {
        **base_row,
        "protocol": "1000162837,1000162999",
    }
    validation = validate_rows(
        [invalid_separator_row],
        {base_row["qr"]: event},
        {base_row["qr"]: event},
        protocol_counts=counts,
    )[0]
    assert validation.count_result == COUNT_INVALID_PROTOCOL_SEPARATOR


def test_export_counts_and_rejection_observation_are_applied():
    workbook = make_workbook()
    _, rows = extract_workbook_rows(workbook)
    protocol_counts = read_export_protocol_counts(make_export_workbook())
    assert protocol_counts["1000147997"].pel == 6
    assert protocol_counts["1000147997"].tratas == 2

    csv = StringIO(
        "request_code;final_state;date_created;comments\n"
        "QR-N;drawing_review;2026-06-18 10:00:00;\n"
        'QR-X;verify_formats;2026-06-18 09:00:00;"Corregir&nbsp;rotulación"\n'
        "QR-X;verify_formats;2026-06-18 11:00:00;\n"
        "QR-V;2.7-Dibujo finalizado;2026-06-17 16:00:00;\n"
    )
    on_day, as_of = read_latest_log_states(
        csv,
        {row["qr"] for row in rows},
        date(2026, 6, 18),
        chunksize=2,
    )
    validations = validate_rows(
        rows,
        on_day,
        as_of,
        protocol_counts=protocol_counts,
    )
    by_qr = {item.qr: item for item in validations}
    assert by_qr["QR-N"].export_pel == 6
    assert by_qr["QR-N"].export_tratas == 2
    assert by_qr["QR-N"].count_result == COUNT_UPDATED
    assert by_qr["QR-X"].pms_observation == "Corregir rotulación"
    assert by_qr["QR-X"].observation_result == OBSERVATION_UPDATED

    output = build_upload_workbook(workbook, validations, matches_only=True)
    wb = load_workbook(BytesIO(output), data_only=False)
    ws_new = wb["Flujo Nuevo"]
    assert ws_new["B2"].value == "QR-N"
    assert ws_new["J2"].value is None
    assert ws_new["K2"].value == 6
    assert ws_new["L2"].value == 2
    assert ws_new["M2"].value == '=IF(COUNTA(K2:L2)=0,"",SUM(K2:L2))'
    assert ws_new["B3"].value == "QR-X"
    assert ws_new["J3"].value == "Corregir rotulación"

    ws_result = wb["Resultado Validacion"]
    assert ws_result["AB2"].value == 6
    assert ws_result["AC2"].value == 2
    assert ws_result["AD2"].value == COUNT_UPDATED
    assert ws_result["AE3"].value == "Corregir rotulación"
    assert ws_result["AF3"].value == OBSERVATION_UPDATED


def test_validation_and_strict_output_structure():
    workbook = make_workbook()
    _, rows = extract_workbook_rows(workbook)
    csv = StringIO(
        "request_code;final_state;date_created\n"
        "QR-N;drawing_review;2026-06-18 10:00:00\n"
        "QR-X;assigned_draftsman;2026-06-18 11:00:00\n"
        "QR-V;2.7-Dibujo finalizado;2026-06-17 16:00:00\n"
    )
    on_day, as_of = read_latest_log_states(
        csv,
        {row["qr"] for row in rows},
        date(2026, 6, 18),
        chunksize=2,
    )
    validations = validate_rows(rows, on_day, as_of)
    results = {item.qr: item.result for item in validations}
    assert results == {
        "QR-N": STATUS_MATCH,
        "QR-X": STATUS_MISMATCH,
        "QR-V": STATUS_NO_EVENT,
    }

    output = build_upload_workbook(workbook, validations, matches_only=True)
    wb = load_workbook(BytesIO(output), data_only=False)
    ws_new = wb["Flujo Nuevo"]
    assert [cell.value for cell in ws_new[1]][:15] == [
        "ORIGEN DE ACTUALIZACIÓN",
        "QR",
        "ELEMENTOS ESTIMADOS",
        "USUARIO",
        "ESTADO DE LA OT",
        "ESTADO SECUNDARIO",
        "CÓDIGO DEL INCREMENTO (AGP)",
        "PROTOCOLO",
        "ODM",
        "OBSERVACIONES",
        "CANTIDAD PEL",
        "CANTIDAD TRATAS",
        "TOTAL ELEMENTOS",
        "FECHA/HORA DE LLEGADA",
        "FECHA/HORA DE SALIDA",
    ]
    assert ws_new.max_row == 2
    assert ws_new["B2"].value == "QR-N"
    assert ws_new["M2"].value == '=IF(COUNTA(K2:L2)=0,"",SUM(K2:L2))'
    assert ws_new.tables["Nuevo"].ref == "A1:O2"

    ws_old = wb["Flujo Viejo"]
    assert ws_old.max_row == 1
    assert ws_old.tables["Viejo"].ref == "A1:P1"

    ws_result = wb["Resultado Validacion"]
    assert wb.sheetnames[-3:] == [
        "Resultado Validacion",
        "Validacion Correos",
        "Casos no reportados",
    ]
    assert [cell.value for cell in ws_result[1]] == [
        "Hoja",
        "Fila Excel",
        "QR",
        "Estado reportado",
        "Último estado PMS del día",
        "Fecha/hora PMS",
        "Último estado anterior o igual",
        "Fecha/hora anterior",
        "Resultado",
        "Profesional bitácora",
        "Profesional PMS",
        "Estado secundario reportado",
        "Estado de llegada",
        "Fecha/hora de llegada",
        "Fecha/hora de salida",
        "Fecha/hora entrega ODM",
        "Fecha/hora entrega dibujo",
        "Resultado tiempos",
        "Fuente tiempos",
        "Drawing Completed request",
        "Estado actual request",
        "Evidencia Flujo 2.0",
        "AGP",
        "ODM",
        "Protocolo",
        "Cantidad PEL reportada",
        "Cantidad TRATAS reportada",
        "Cantidad PEL EXPORT",
        "Cantidad TRATAS EXPORT",
        "Resultado cantidades",
        "Observación PMS",
        "Acción observación",
    ]
    assert ws_result.max_row == 4
    assert ws_result.tables["ResultadoValidacion"].ref == "A1:AF4"
    assert ws_result["C2"].value == "QR-N"
    assert ws_result["I2"].value == STATUS_MATCH
    assert ws_result["C3"].value == "QR-X"
    assert ws_result["I3"].value == STATUS_MISMATCH
    assert ws_result["C4"].value == "QR-V"
    assert ws_result["I4"].value == STATUS_NO_EVENT


def test_coded_template_output_uses_hseq_structure():
    workbook = make_workbook()
    _, rows = extract_workbook_rows(workbook)
    csv = StringIO(
        "request_code;final_state;date_created\n"
        "QR-N;drawing_review;2026-06-18 10:00:00\n"
        "QR-X;assigned_draftsman;2026-06-18 11:00:00\n"
        "QR-V;2.7-Dibujo finalizado;2026-06-17 16:00:00\n"
    )
    on_day, as_of = read_latest_log_states(
        csv,
        {row["qr"] for row in rows},
        date(2026, 6, 18),
        chunksize=2,
    )
    validations = validate_rows(rows, on_day, as_of)

    output = build_upload_workbook(
        workbook,
        validations,
        matches_only=True,
        template_bytes=make_template_workbook(),
    )
    wb = load_workbook(BytesIO(output), data_only=False)

    ws_new = wb["Flujo Nuevo"]
    assert ws_new["M1"].value == "F-AGP-04"
    assert ws_new["M2"].value == "Rev.02"
    assert [ws_new.cell(5, col).value for col in range(1, 13)] == [
        "ORIGEN DE ACTUALIZACIÓN",
        "QR",
        "ELEMENTOS ESTIMADOS",
        "USUARIO",
        "ESTADO DE LA OT",
        "ESTADO SECUNDARIO",
        "CÓDIGO DEL INCREMENTO (AGP)",
        "PROTOCOLO",
        "OBSERVACIONES",
        "CANTIDAD PEL",
        "CANTIDAD TRATAS",
        "TOTAL ELEMENTOS",
    ]
    assert ws_new["B6"].value == "QR-N"
    assert ws_new["I6"].value is None
    assert ws_new["L6"].value == '=IF(COUNTA(J6:K6)=0,"",SUM(J6:K6))'
    assert ws_new["M5"].value == "FECHA/HORA DE LLEGADA"
    assert ws_new["N5"].value == "FECHA/HORA DE SALIDA"
    assert ws_new.tables["Nuevo"].ref == "A5:N6"

    ws_old = wb["Flujo Viejo"]
    assert ws_old.tables["Viejo"].ref == "A5:O6"
    assert ws_old["B6"].value is None

    ws_pnc = wb["PNC"]
    assert ws_pnc.tables["Tabla4"].ref == "A5:J6"
    assert ws_pnc["C6"].value is None

    ws_result = wb["Resultado Validacion"]
    assert wb.sheetnames[-3:] == [
        "Resultado Validacion",
        "Validacion Correos",
        "Casos no reportados",
    ]
    assert ws_result["A1"].value == "Hoja"
    assert ws_result["I4"].value == STATUS_NO_EVENT

    for sheet_name in coded_header_sheets():
        ws = wb[sheet_name]
        assert ws["A1"].value is None
        assert len(ws._images) == 1
    assert len(ws_result._images) == 0


def test_coded_output_writes_typed_arrival_and_exit_datetimes():
    source = load_workbook(BytesIO(make_workbook()), data_only=False)
    source["Flujo Nuevo"]["E2"] = "Manuel Fernando Sanchez Mondragon"
    prepared = BytesIO()
    source.save(prepared)
    source.close()
    workbook = prepared.getvalue()
    _, rows = extract_workbook_rows(workbook)
    csv = StringIO(
        "request_code;initial_state;final_state;date_created;username;user_created_last_name\n"
        "QR-N;order_creation;drawing_request;2026-06-18 08:00:00;;Sistema\n"
        "QR-N;drawing_request;assigned_draftsman;2026-06-18 09:00:00;;Sistema\n"
        "QR-N;assigned_draftsman;drawing_review;2026-06-18 10:00:00;CO1049796769;Sanchez Mondragon.Manuel Fernando\n"
    )
    on_day, as_of, _, _, timelines = read_log_validation_data(
        csv,
        {row["qr"] for row in rows},
        date(2026, 6, 18),
        include_timelines=True,
    )
    validations = validate_rows(
        rows,
        on_day,
        as_of,
        timelines=timelines,
        production_date=date(2026, 6, 18),
    )
    output = build_upload_workbook(
        workbook,
        validations,
        matches_only=True,
        template_bytes=make_template_workbook(),
    )
    wb = load_workbook(BytesIO(output), data_only=False)
    ws = wb["Flujo Nuevo"]
    assert ws["M6"].value == datetime(2026, 6, 18, 8, 0)
    assert ws["N6"].value == datetime(2026, 6, 18, 10, 0)
    assert ws["M6"].number_format == "dd/mm/yyyy hh:mm:ss"
    assert ws["N6"].number_format == "dd/mm/yyyy hh:mm:ss"
    result = wb["Resultado Validacion"]
    assert result["N2"].value == datetime(2026, 6, 18, 8, 0)
    assert result["O2"].value == datetime(2026, 6, 18, 10, 0)

def test_operational_noise_is_removed_from_all_five_sheets_without_touching_formulas():
    source_wb = load_workbook(BytesIO(make_workbook()), data_only=False)
    ws_new = source_wb["Flujo Nuevo"]
    ws_new["E2"] = "Profesional S/Norte"
    ws_new["G2"] = "S/N"
    ws_new["I2"] = 0
    ws_new["M2"] = 0
    ws_new["N2"] = 2

    ws_old = source_wb["Flujo Viejo"]
    ws_old["L2"] = "S/N"
    ws_old["M2"] = "N/A"
    ws_old["N2"] = 0
    ws_old["O2"] = 2

    ws_odm = source_wb.create_sheet("Odm")
    ws_odm.append(
        [
            "ORIGEN DE ACTUALIZACIÓN",
            "QR",
            "ODM",
            "ELEMENTOS ESTIMADOS",
            "USUARIO",
            "ESTADO DE LA OT",
            "OBSERVACIONES",
        ]
    )
    ws_odm.append(
        ["N/A", "QR-ODM", "=1+1", 0, "S/N", "drawing_review", "Texto con 0 válido"]
    )

    ws_email = source_wb.create_sheet("Correo-Rotulacion")
    ws_email.append(
        [
            "ORIGEN DE ACTUALIZACIÓN",
            "ASUNTO",
            "ELEMENTOS ESTIMADOS",
            "USUARIO",
            "ESTADO DE LA OT",
            "ESTADO SECUNDARIO",
            "CÓDIGO DEL INCREMENTO (AGP)",
            "PROTOCOLO",
            "NÚMERO DE TICKET BLOQUEO (INC)",
            "OBSERVACIONES",
            "CANTIDAD PEL",
            "CANTIDAD TRATAS",
            "TOTAL ELEMENTOS",
        ]
    )

    ws_pnc = source_wb.create_sheet("PNC")
    ws_pnc.append(
        [
            "FECHA DE LLEGADA",
            "ORIGEN DE ACTUALIZACIÓN",
            "QR/ASUNTO",
            "USUARIO QUE DEVUELVE",
            "MOTIVO DE DEVOLUCIÓN",
            "RESULTADO REVISIÓN",
            "REVISADO POR",
            "CONCEPTO DE REVISIÓN",
            "DESTINO/ACCIÓN POSTERIOR",
            "FECHA DE REVISIÓN",
        ]
    )
    ws_pnc.append(
        [
            date(2026, 8, 10),
            "PNC",
            "QR-PNC-1",
            "N/A",
            "No aplica",
            "S/N",
            "SM",
            "Concepto con 0 válido",
            0,
            date(2026, 8, 11),
        ]
    )
    ws_pnc.append(
        [
            date(2026, 8, 12),
            "Correo",
            "ASUNTO-PNC-2",
            "Usuario válido",
            "Motivo válido",
            "Resultado válido",
            "Revisor válido",
            "Concepto válido",
            "Cierre",
            date(2026, 8, 13),
        ]
    )
    ws_email.append(
        [
            "S/N",
            "Correo 0 válido",
            0,
            "N/A",
            "correo_enviado",
            0,
            2000140847,
            "S/N",
            0,
            "N/A",
            0,
            0,
            0,
        ]
    )

    source = BytesIO()
    source_wb.save(source)
    source_wb.close()
    workbook = source.getvalue()
    _, rows = extract_workbook_rows(workbook)
    events = {
        "QR-N": LogState("drawing_review", datetime(2026, 6, 18, 10, 0)),
        "QR-X": LogState("verify_formats", datetime(2026, 6, 18, 11, 0)),
        "QR-V": LogState("2.7-Dibujo finalizado", datetime(2026, 6, 18, 12, 0)),
    }
    validations = validate_rows(rows, events, events)
    output = build_upload_workbook(
        workbook,
        validations,
        matches_only=False,
        template_bytes=make_template_workbook(),
    )
    wb = load_workbook(BytesIO(output), data_only=False)

    ws_new = wb["Flujo Nuevo"]
    assert ws_new["D6"].value == "Profesional S/Norte"
    assert ws_new["F6"].value is None
    assert ws_new["G6"].value is None
    assert ws_new["I6"].value is None
    assert ws_new["J6"].value is None
    assert ws_new["K6"].value == 2

    ws_old = wb["Flujo Viejo"]
    assert ws_old["I6"].value is None
    assert ws_old["J6"].value is None
    assert ws_old["K6"].value is None
    assert ws_old["L6"].value == 2

    ws_odm = wb["Odm"]
    assert ws_odm["A6"].value is None
    assert ws_odm["C6"].value == "=1+1"
    assert ws_odm["D6"].value is None
    assert ws_odm["E6"].value is None
    assert ws_odm["G6"].value == "Texto con 0 válido"

    ws_email = wb["Correo-Rotulacion"]
    assert ws_email["A6"].value is None
    assert ws_email["B6"].value == "Correo 0 válido"
    assert ws_email["C6"].value is None
    assert ws_email["D6"].value is None
    assert ws_email["F6"].value is None
    assert ws_email["G6"].value == 2000140847
    assert ws_email["H6"].value is None
    assert ws_email["I6"].value is None
    assert ws_email["J6"].value is None
    assert ws_email["K6"].value is None
    assert ws_email["L6"].value is None
    assert ws_email["M6"].value == '=IF(COUNTA(K6:L6)=0,"",SUM(K6:L6))'

    ws_pnc = wb["PNC"]
    assert ws_pnc.tables["Tabla4"].ref == "A5:J7"
    assert ws_pnc["A6"].value == datetime(2026, 8, 10)
    assert ws_pnc["B6"].value == "PNC"
    assert ws_pnc["C6"].value == "QR-PNC-1"
    assert ws_pnc["D6"].value is None
    assert ws_pnc["E6"].value is None
    assert ws_pnc["F6"].value is None
    assert ws_pnc["G6"].value is None
    assert ws_pnc["H6"].value == "Concepto con 0 válido"
    assert ws_pnc["I6"].value is None
    assert ws_pnc["J6"].value == datetime(2026, 8, 11)
    assert ws_pnc["B7"].value == "Correo"
    assert ws_pnc["C7"].value == "ASUNTO-PNC-2"


def test_operational_cleaning_ignores_thousands_of_trailing_formatted_rows():
    wb = Workbook()
    ws = wb.active
    ws.title = "Odm"
    ws.append(
        [
            "ORIGEN DE ACTUALIZACIÓN",
            "QR",
            "ODM",
            "ELEMENTOS ESTIMADOS",
            "USUARIO",
            "ESTADO DE LA OT",
            "OBSERVACIONES",
        ]
    )
    ws.append(["N/A", "QR-REAL", None, 0, "S/N", "drawing_review", "N/A"])
    ws.cell(22_606, 29).fill = PatternFill("solid", fgColor="FFFFFF")
    cells_before = len(ws._cells)

    _clean_operational_sheets(wb)

    assert ws["A2"].value is None
    assert ws["B2"].value == "QR-REAL"
    assert ws["D2"].value is None
    assert ws["E2"].value is None
    assert ws["G2"].value is None
    assert len(ws._cells) < 100


def test_email_quantities_only_update_solicitud_por_correo_rows():
    source_wb = load_workbook(BytesIO(make_workbook()), data_only=False)
    ws = source_wb.create_sheet("Correo-Rotulacion")
    ws.append(
        [
            "ORIGEN DE ACTUALIZACIÓN",
            "ASUNTO",
            "ELEMENTOS ESTIMADOS",
            "USUARIO",
            "ESTADO DE LA OT",
            "ESTADO SECUNDARIO",
            "CÓDIGO DEL INCREMENTO (AGP)",
            "PROTOCOLO",
            "NÚMERO DE TICKET BLOQUEO (INC)",
            "OBSERVACIONES",
            "CANTIDAD PEL",
            "CANTIDAD TRATAS",
            "TOTAL ELEMENTOS",
        ]
    )
    ws.append(
        [
            "Solicitud por correo",
            "REDIBUJADORA - SOLICITUD 1",
            1,
            "Ana",
            "Finalizado",
            None,
            None,
            "1000162837",
            None,
            None,
            1,
            0,
            None,
        ]
    )
    ws.append(
        [
            " solicitud POR correo ",
            "REDIBUJADORA - SOLICITUD 2",
            1,
            "Ana",
            " finalizado ",
            None,
            None,
            "1000162837;1000162838",
            None,
            None,
            4,
            1,
            None,
        ]
    )
    ws.append(
        [
            "Solicitud por correo",
            "REDIBUJADORA - TODAVÍA EN PROCESO",
            1,
            "Ana",
            "En proceso",
            None,
            None,
            "1000162837",
            None,
            None,
            6,
            6,
            None,
        ]
    )
    ws.append(
        [
            "Rotulación",
            "FACTURA - NO APLICA",
            1,
            "Ana",
            "Finalizado",
            None,
            None,
            "1000162837",
            None,
            None,
            9,
            9,
            None,
        ]
    )
    ws.append(
        [
            "Solicitud por correo",
            "REDIBUJADORA - PROTOCOLO INCOMPLETO",
            1,
            "Ana",
            "Finalizado",
            None,
            None,
            "1000162837;1000162999",
            None,
            None,
            7,
            7,
            None,
        ]
    )
    ws.append(
        [
            "Solicitud por correo",
            "REDIBUJADORA - SEPARADOR INVÁLIDO",
            1,
            "Ana",
            "Finalizado",
            None,
            None,
            "1000162837,1000162838",
            None,
            None,
            8,
            8,
            None,
        ]
    )

    source = BytesIO()
    source_wb.save(source)
    source_wb.close()
    workbook = source.getvalue()
    protocol_counts = {
        "1000162837": ProtocolElementCounts(
            pel_ids=frozenset({"P-1", "P-2"}),
            tratta_ids=frozenset({"T-1"}),
        ),
        "1000162838": ProtocolElementCounts(
            pel_ids=frozenset({"P-3", "P-4"}),
            tratta_ids=frozenset(),
        ),
    }
    email_results = evaluate_email_quantities(workbook, protocol_counts)

    assert [item.excel_row for item in email_results] == [2, 3, 6, 7]
    assert email_results[0].count_result == COUNT_UPDATED
    assert (email_results[0].export_pel, email_results[0].export_tratas) == (2, 1)
    assert email_results[1].count_result == COUNT_CONFIRMED
    assert (email_results[1].export_pel, email_results[1].export_tratas) == (4, 1)
    assert email_results[2].count_result == (
        f"{COUNT_PROTOCOLS_INCOMPLETE}: 1000162999"
    )
    assert (email_results[2].export_pel, email_results[2].export_tratas) == (2, 1)
    assert email_results[3].count_result == COUNT_INVALID_PROTOCOL_SEPARATOR

    _, rows = extract_workbook_rows(workbook)
    events = {
        "QR-N": LogState("drawing_review", datetime(2026, 6, 18, 10, 0)),
        "QR-X": LogState("verify_formats", datetime(2026, 6, 18, 11, 0)),
        "QR-V": LogState("2.7-Dibujo finalizado", datetime(2026, 6, 18, 12, 0)),
    }
    validations = validate_rows(rows, events, events)
    output = build_upload_workbook(
        workbook,
        validations,
        matches_only=False,
        template_bytes=make_template_workbook(),
        email_quantity_results=email_results,
    )
    wb = load_workbook(BytesIO(output), data_only=False)

    ws = wb["Correo-Rotulacion"]
    assert (ws["K6"].value, ws["L6"].value) == (2, 1)
    assert (ws["K7"].value, ws["L7"].value) == ("2;2", "1;0")
    assert ws["M7"].value == 5
    assert ws["E8"].value == "En proceso"
    assert (ws["K8"].value, ws["L8"].value) == (6, 6)
    assert ws["A9"].value == "Rotulación"
    assert (ws["K9"].value, ws["L9"].value) == (9, 9)
    assert (ws["K10"].value, ws["L10"].value) == (7, 7)
    assert (ws["K11"].value, ws["L11"].value) == (8, 8)

    ws_result = wb["Validacion Correos"]
    assert ws_result.max_row == 5
    assert ws_result.tables["ValidacionCorreos"].ref == "A1:I5"
    assert [ws_result.cell(row, 1).value for row in range(2, 6)] == [2, 3, 6, 7]
    assert ws_result["I2"].value == COUNT_UPDATED
    assert ws_result["I3"].value == COUNT_CONFIRMED
    assert ws_result["I4"].value == (
        f"{COUNT_PROTOCOLS_INCOMPLETE}: 1000162999"
    )
    assert (ws_result["G4"].value, ws_result["H4"].value) == (2, 1)
    assert ws_result["I5"].value == COUNT_INVALID_PROTOCOL_SEPARATOR


def test_request_snapshot_reader_parses_flow20_flags_and_metadata():
    request = StringIO(
        "request_code;contract_drawing_completed;status;new_fulfillment_process;"
        "date_created;odm;agp_project;agp_increment;protocol\n"
        "QR-20;true;3-ODM Asignadas;true;2026-10-29 11:00:00;ODM-20;AGP-20;;1000147997\n"
        "QR-OTHER;false;2.6-Dibujante asignado;true;2026-10-29 11:00:00;;;;\n"
    )

    snapshots = read_request_snapshots(request, {"QR-20"}, chunksize=1)

    assert set(snapshots) == {"QR-20"}
    snapshot = snapshots["QR-20"]
    assert snapshot.drawing_completed is True
    assert snapshot.new_fulfillment_process is True
    assert snapshot.status == FLOW20_ODM_STATE
    assert snapshot.odm == "ODM-20"
    assert snapshot.agp == "AGP-20"
    assert snapshot.protocol == "1000147997"


def test_flow20_simultaneous_odm_and_drawing_are_validated_and_exported():
    workbook = make_flow20_workbook()
    _, rows = extract_workbook_rows(workbook)
    log = StringIO(
        "request_code;initial_state;final_state;date_created;username;user_created_last_name;comments\n"
        "QR-20;0-Por Valorizar;2-Valorizado;2026-10-29 07:00:00;;Sistema;\n"
        "QR-20;2-Valorizado;2.5-Enviado a dibujo;2026-10-29 08:00:00;;Sistema;\n"
        "QR-20;2.5-Enviado a dibujo;2.6-Dibujante asignado;2026-10-29 09:00:00;;Sistema;Drawing Assignment\n"
        "QR-20;2.6-Dibujante asignado;3-ODM Asignadas;2026-10-29 10:00:00;"
        "CO1049796769;Sanchez Mondragon.Manuel Fernando;OK\n"
    )
    on_day, as_of, _, _, timelines = read_log_validation_data(
        log,
        {"QR-20"},
        date(2026, 10, 29),
        include_timelines=True,
    )
    snapshot = RequestSnapshot(
        qr="QR-20",
        drawing_completed=True,
        status=FLOW20_ODM_STATE,
        new_fulfillment_process=True,
    )
    validations = validate_rows(
        rows,
        on_day,
        as_of,
        timelines=timelines,
        production_date=date(2026, 10, 29),
        request_snapshots={"QR-20": snapshot},
    )
    item = next(row for row in validations if row.sheet == SHEET_FLOW20)

    assert item.result == STATUS_MATCH
    assert item.arrival_datetime == datetime(2026, 10, 29, 8, 0)
    assert item.odm_datetime == datetime(2026, 10, 29, 10, 0)
    assert item.drawing_datetime == datetime(2026, 10, 29, 10, 0)

    counted = apply_protocol_counts_to_validations(
        validations,
        {
            "1000147997": ProtocolElementCounts(
                pel_ids=frozenset({"P-1", "P-2"}),
                tratta_ids=frozenset({"T-1"}),
            )
        },
    )
    output = build_upload_workbook(
        workbook,
        counted,
        matches_only=True,
        template_bytes=make_template_workbook(),
    )
    wb = load_workbook(BytesIO(output), data_only=False)
    ws = wb[SHEET_FLOW20]
    assert ws.tables["Flujo20"].ref == "A5:Q6"
    assert ws["F6"].value == FLOW20_SECONDARY_COMPLETED
    assert (ws["L6"].value, ws["M6"].value) == (2, 1)
    assert ws["N6"].value == '=IF(COUNTA(L6:M6)=0,"",SUM(L6:M6))'
    assert ws["O6"].value == datetime(2026, 10, 29, 8, 0)
    assert ws["P6"].value == datetime(2026, 10, 29, 10, 0)
    assert ws["Q6"].value == datetime(2026, 10, 29, 10, 0)
    assert [cell.value for cell in ws[5]][-3:] == [
        "FECHA/HORA DE LLEGADA",
        ODM_DATETIME_HEADER,
        DRAWING_DATETIME_HEADER,
    ]
    wb.close()


def test_flow20_staggered_drawing_uses_same_professional_repeat_only():
    workbook = make_flow20_workbook()
    _, rows = extract_workbook_rows(workbook)
    base = (
        "request_code;initial_state;final_state;date_created;username;user_created_last_name;comments\n"
        "QR-20;0-Por Valorizar;2-Valorizado;2026-10-28 07:00:00;;Sistema;\n"
        "QR-20;2-Valorizado;2.5-Enviado a dibujo;2026-10-28 08:00:00;;Sistema;\n"
        "QR-20;2.5-Enviado a dibujo;2.6-Dibujante asignado;2026-10-28 09:00:00;;Sistema;Drawing Assignment\n"
        "QR-20;2.6-Dibujante asignado;3-ODM Asignadas;2026-10-28 10:00:00;"
        "CO1049796769;Sanchez Mondragon.Manuel Fernando;OK\n"
    )
    valid_repeat = (
        "QR-20;3-ODM Asignadas;3-ODM Asignadas;2026-10-29 11:00:00;"
        "CO1049796769;Sanchez Mondragon.Manuel Fernando;URL dibujo\n"
    )
    on_day, as_of, _, _, timelines = read_log_validation_data(
        StringIO(base + valid_repeat),
        {"QR-20"},
        date(2026, 10, 29),
        include_timelines=True,
    )
    snapshot = RequestSnapshot(
        qr="QR-20",
        drawing_completed=True,
        new_fulfillment_process=True,
    )
    item = next(
        row
        for row in validate_rows(
            rows,
            on_day,
            as_of,
            timelines=timelines,
            production_date=date(2026, 10, 29),
            request_snapshots={"QR-20": snapshot},
        )
        if row.sheet == SHEET_FLOW20
    )
    assert item.result == STATUS_MATCH
    assert item.odm_datetime == datetime(2026, 10, 28, 10, 0)
    assert item.drawing_datetime == datetime(2026, 10, 29, 11, 0)

    technical_repeat = (
        "QR-20;3-ODM Asignadas;3-ODM Asignadas;2026-10-29 11:00:00;"
        "CO999;Otro.Usuario;Drawing Reassignment\n"
    )
    on_day, as_of, _, _, timelines = read_log_validation_data(
        StringIO(base + technical_repeat),
        {"QR-20"},
        date(2026, 10, 29),
        include_timelines=True,
    )
    item = next(
        row
        for row in validate_rows(
            rows,
            on_day,
            as_of,
            timelines=timelines,
            production_date=date(2026, 10, 29),
            request_snapshots={"QR-20": snapshot},
        )
        if row.sheet == SHEET_FLOW20
    )
    assert item.result == STATUS_FLOW20_AMBIGUOUS
    assert item.drawing_datetime is None


def test_flow20_direct_valorized_to_odm_uses_valorized_as_arrival():
    workbook = make_flow20_workbook()
    _, rows = extract_workbook_rows(workbook)
    log = StringIO(
        "request_code;initial_state;final_state;date_created;username;user_created_last_name;comments\n"
        "QR-20;0-Por Valorizar;2-Valorizado;2026-10-29 08:00:00;;Sistema;\n"
        "QR-20;2-Valorizado;3-ODM Asignadas;2026-10-29 10:00:00;"
        "CO1049796769;Sanchez Mondragon.Manuel Fernando;OK\n"
    )
    on_day, as_of, _, _, timelines = read_log_validation_data(
        log,
        {"QR-20"},
        date(2026, 10, 29),
        include_timelines=True,
    )
    item = next(
        row
        for row in validate_rows(
            rows,
            on_day,
            as_of,
            timelines=timelines,
            production_date=date(2026, 10, 29),
            request_snapshots={
                "QR-20": RequestSnapshot(
                    qr="QR-20",
                    drawing_completed=True,
                    new_fulfillment_process=True,
                )
            },
        )
        if row.sheet == SHEET_FLOW20
    )
    assert item.result == STATUS_MATCH
    assert item.arrival_state == "2-Valorizado"
    assert item.arrival_datetime == datetime(2026, 10, 29, 8, 0)


def test_flow20_ambiguous_times_can_be_confirmed_once_in_state_audit():
    workbook = make_flow20_workbook()
    _, rows = extract_workbook_rows(workbook)
    log_text = (
        "request_code;initial_state;final_state;date_created;username;user_created_last_name;comments\n"
        "QR-20;0-Por Valorizar;2-Valorizado;2026-10-28 07:00:00;;Sistema;\n"
        "QR-20;2-Valorizado;2.5-Enviado a dibujo;2026-10-28 08:00:00;;Sistema;\n"
        "QR-20;2.5-Enviado a dibujo;2.6-Dibujante asignado;2026-10-28 09:00:00;;Sistema;\n"
        "QR-20;2.6-Dibujante asignado;3-ODM Asignadas;2026-10-28 10:00:00;"
        "CO1049796769;Sanchez Mondragon.Manuel Fernando;OK\n"
    )
    on_day, as_of, _, _, timelines = read_log_validation_data(
        StringIO(log_text),
        {"QR-20"},
        date(2026, 10, 29),
        include_timelines=True,
    )
    validations = validate_rows(
        rows,
        on_day,
        as_of,
        timelines=timelines,
        production_date=date(2026, 10, 29),
        request_snapshots={
            "QR-20": RequestSnapshot(
                qr="QR-20",
                drawing_completed=True,
                new_fulfillment_process=True,
            )
        },
    )
    item = next(row for row in validations if row.sheet == SHEET_FLOW20)
    assert item.result == STATUS_FLOW20_AMBIGUOUS
    decision = {
        validation_audit_key(item): {
            "action": "keep",
            "secondary_state": FLOW20_SECONDARY_COMPLETED,
            "arrival_datetime": "28/10/2026 08:00:00",
            "odm_datetime": "28/10/2026 10:00:00",
            "drawing_datetime": "29/10/2026 11:00:00",
        }
    }

    _, updated, audit = apply_state_audit_decisions(
        workbook,
        validations,
        on_day,
        decision,
        timelines=timelines,
        production_date=date(2026, 10, 29),
    )
    updated_item = next(row for row in updated if row.sheet == SHEET_FLOW20)
    assert updated_item.result == STATUS_AUDIT_KEEP
    assert updated_item.arrival_datetime == datetime(2026, 10, 28, 8, 0)
    assert updated_item.odm_datetime == datetime(2026, 10, 28, 10, 0)
    assert updated_item.drawing_datetime == datetime(2026, 10, 29, 11, 0)
    assert audit[0]["Fecha/hora entrega ODM"] == "28/10/2026 10:00:00"


def test_flow20_pending_accepts_a_drawing_repeat_after_the_cutoff():
    workbook = make_flow20_workbook(secondary=FLOW20_SECONDARY_PENDING)
    _, rows = extract_workbook_rows(workbook)
    log_text = (
        "request_code;initial_state;final_state;date_created;username;user_created_last_name;comments\n"
        "QR-20;0-Por Valorizar;2-Valorizado;2026-10-29 07:00:00;;Sistema;\n"
        "QR-20;2-Valorizado;2.5-Enviado a dibujo;2026-10-29 08:00:00;;Sistema;\n"
        "QR-20;2.5-Enviado a dibujo;2.6-Dibujante asignado;2026-10-29 09:00:00;;Sistema;\n"
        "QR-20;2.6-Dibujante asignado;3-ODM Asignadas;2026-10-29 10:00:00;"
        "CO1049796769;Sanchez Mondragon.Manuel Fernando;OK\n"
        "QR-20;3-ODM Asignadas;3-ODM Asignadas;2026-10-30 11:00:00;"
        "CO1049796769;Sanchez Mondragon.Manuel Fernando;URL dibujo\n"
    )
    on_day, as_of, _, _, timelines = read_log_validation_data(
        StringIO(log_text),
        {"QR-20"},
        date(2026, 10, 29),
        include_timelines=True,
    )
    timelines = read_full_log_timelines(StringIO(log_text), {"QR-20"})
    snapshot = RequestSnapshot(
        qr="QR-20",
        drawing_completed=True,
        new_fulfillment_process=True,
    )
    item = next(
        row
        for row in validate_rows(
            rows,
            on_day,
            as_of,
            timelines=timelines,
            production_date=date(2026, 10, 29),
            request_snapshots={"QR-20": snapshot},
        )
        if row.sheet == SHEET_FLOW20
    )
    assert item.result == STATUS_MATCH
    assert item.odm_datetime == datetime(2026, 10, 29, 10, 0)
    assert item.drawing_datetime is None
