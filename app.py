from __future__ import annotations

from collections import Counter
from hashlib import sha256
from html import escape
from typing import Any

import pandas as pd
import streamlit as st

from bitacora import (
    DRAWING_OUTPUT_STATES,
    FINAL_ELEMENT_STATES,
    FLOW20_ODM_STATE,
    FLOW20_DUPLICATE_RESULTS,
    FLOW20_SECONDARY_COMPLETED,
    FLOW20_SECONDARY_PENDING,
    FLOW20_STATES,
    LEGACY_STATES,
    NEW_STATES,
    REJECTION_STATES,
    SHEET_FLOW20,
    SHEET_NEW,
    SHEET_OLD,
    STATUS_AUDIT_EXCLUDED,
    STATUS_MATCH,
    UPDATE_ORIGINS,
    UPLOAD_APPROVED_RESULTS,
    BitacoraError,
    add_unreported_audit_rows,
    apply_qr_alias_corrections,
    apply_drawing_timings,
    apply_protocol_audit_decisions,
    apply_protocol_audit_results,
    apply_protocol_counts_to_validations,
    apply_state_audit_decisions,
    build_protocol_inventory_workbook,
    build_upload_workbook,
    default_template_bytes,
    email_quantity_frame,
    evaluate_email_quantities,
    extract_workbook_rows,
    flow20_duplicate_audit_entries,
    has_invalid_agp_separator,
    has_invalid_protocol_prefix,
    has_invalid_protocol_separator,
    invalid_protocol_prefixes,
    output_filename,
    parse_production_date,
    parse_audit_datetime,
    protocol_counts_with_audit_decisions,
    read_export_protocol_counts,
    read_full_log_timelines,
    read_log_validation_data,
    read_request_snapshots,
    require_log_date_coverage,
    split_protocols,
    replace_protocol_references,
    validate_rows,
    validation_audit_key,
    validation_frame,
    validation_requires_element_counts,
)


st.set_page_config(
    page_title="Auditoría guiada de bitácoras PMS",
    page_icon="✅",
    layout="wide",
)

st.markdown(
    """
    <style>
    .block-container {padding-top: 2rem; padding-bottom: 4rem;}
    .summary-panel {
        margin: 1rem 0 1.15rem; padding: 1.15rem;
        border: 1px solid #e2e8f0; border-radius: 18px;
        background: linear-gradient(145deg, #ffffff 0%, #f8fafc 100%);
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.06);
    }
    .summary-panel__eyebrow {
        margin: 0 0 .18rem; color: #0f766e; font-size: .72rem;
        font-weight: 800; letter-spacing: .09em; text-transform: uppercase;
    }
    .summary-panel__title {margin: 0; color: #0f172a; font-size: 1.1rem; font-weight: 750;}
    .summary-panel__subtitle {margin: .28rem 0 1rem; color: #64748b; font-size: .84rem;}
    .summary-grid {display: grid; grid-template-columns: repeat(var(--card-columns), minmax(0, 1fr)); gap: .8rem;}
    .summary-card {
        position: relative; min-height: 118px; padding: .95rem 1rem .9rem;
        overflow: hidden; border: 1px solid var(--card-border); border-radius: 14px;
        background: var(--card-bg);
    }
    .summary-card::before {position:absolute; inset:0 auto 0 0; width:4px; background:var(--card-accent); content:"";}
    .summary-card--success {--card-accent:#16a34a;--card-border:#bbf7d0;--card-bg:#f0fdf4;}
    .summary-card--danger {--card-accent:#dc2626;--card-border:#fecaca;--card-bg:#fef2f2;}
    .summary-card--warning {--card-accent:#d97706;--card-border:#fde68a;--card-bg:#fffbeb;}
    .summary-card--neutral {--card-accent:#64748b;--card-border:#e2e8f0;--card-bg:#f8fafc;}
    .summary-card--info {--card-accent:#2563eb;--card-border:#bfdbfe;--card-bg:#eff6ff;}
    .summary-card--teal {--card-accent:#0d9488;--card-border:#99f6e4;--card-bg:#f0fdfa;}
    .summary-card__label {margin:0;color:#334155;font-size:.78rem;font-weight:700;}
    .summary-card__value {margin:.35rem 0 .16rem;color:#0f172a;font-size:clamp(2rem,3vw,2.65rem);font-weight:780;line-height:1;}
    .summary-card__helper {margin:0;color:#64748b;font-size:.74rem;line-height:1.35;}
    .stepper {display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:.55rem;margin:1rem 0 1.35rem;}
    .step {padding:.75rem .8rem;border-radius:12px;border:1px solid #e2e8f0;background:#f8fafc;color:#64748b;font-size:.82rem;font-weight:700;}
    .step--active {border-color:#0d9488;background:#f0fdfa;color:#0f766e;box-shadow:0 0 0 2px rgba(13,148,136,.1);}
    .step--done {border-color:#bbf7d0;background:#f0fdf4;color:#15803d;}
    .case-card {padding:1.1rem 1.2rem;border:1px solid #cbd5e1;border-radius:16px;background:#fff;margin:.5rem 0 1rem;}
    .case-kicker {color:#0f766e;font-weight:800;font-size:.75rem;text-transform:uppercase;letter-spacing:.06em;}
    .case-title {color:#0f172a;font-size:1.15rem;font-weight:800;margin:.25rem 0 .45rem;}
    .case-meta {color:#475569;font-size:.85rem;line-height:1.55;}
    @media(max-width:900px){.summary-grid,.stepper{grid-template-columns:repeat(2,minmax(0,1fr));}}
    </style>
    """,
    unsafe_allow_html=True,
)


def render_summary_panel(
    *,
    eyebrow: str,
    title: str,
    subtitle: str,
    cards: list[tuple[str, int, str, str]],
) -> None:
    card_html = "".join(
        (
            f'<div class="summary-card summary-card--{escape(tone)}">'
            f'<p class="summary-card__label">{escape(label)}</p>'
            f'<p class="summary-card__value">{value:,}</p>'
            f'<p class="summary-card__helper">{escape(helper)}</p></div>'
        )
        for label, value, helper, tone in cards
    )
    st.markdown(
        '<section class="summary-panel">'
        f'<p class="summary-panel__eyebrow">{escape(eyebrow)}</p>'
        f'<h3 class="summary-panel__title">{escape(title)}</h3>'
        f'<p class="summary-panel__subtitle">{escape(subtitle)}</p>'
        f'<div class="summary-grid" style="--card-columns:{len(cards)}">{card_html}</div>'
        "</section>",
        unsafe_allow_html=True,
    )


def render_stepper(current: int) -> None:
    labels = (
        "1 · Estados",
        "2 · No reportados",
        "3 · Protocolos y cantidades",
        "4 · Excel final",
    )
    items = []
    for number, label in enumerate(labels, start=1):
        status = "step--done" if number < current else "step--active" if number == current else ""
        prefix = "✓ " if number < current else ""
        items.append(f'<div class="step {status}">{prefix}{escape(label)}</div>')
    st.markdown(f'<div class="stepper">{"".join(items)}</div>', unsafe_allow_html=True)


def case_card(title: str, kicker: str, lines: list[str]) -> None:
    details = "<br>".join(escape(line) for line in lines if line)
    st.markdown(
        '<div class="case-card">'
        f'<div class="case-kicker">{escape(kicker)}</div>'
        f'<div class="case-title">{escape(title or "Sin referencia")}</div>'
        f'<div class="case-meta">{details}</div></div>',
        unsafe_allow_html=True,
    )


def reset_audit() -> None:
    prefixes = ("audit_", "state_form_", "unreported_form_", "protocol_form_")
    for key in list(st.session_state):
        if key.startswith(prefixes):
            del st.session_state[key]


def current_context() -> dict[str, Any] | None:
    value = st.session_state.get("audit_context")
    return value if isinstance(value, dict) else None


def protocol_issue_key(sheet: str, row: int, reference: str) -> str:
    return f"{sheet}|{row}|{reference}"


def protocol_references(workbook_bytes: bytes, validations: list[Any]) -> list[dict[str, Any]]:
    refs: list[dict[str, Any]] = []
    for item in validations:
        if item.result not in UPLOAD_APPROVED_RESULTS or not validation_requires_element_counts(item):
            continue
        refs.append(
            {
                "key": protocol_issue_key(item.sheet, item.excel_row, item.qr),
                "sheet": item.sheet,
                "excel_row": item.excel_row,
                "reference": item.qr,
                "state": item.reported_state,
                "protocol": item.protocol,
            }
        )
    for item in evaluate_email_quantities(workbook_bytes, {}):
        refs.append(
            {
                "key": protocol_issue_key("Correo-Rotulacion", item.excel_row, item.subject),
                "sheet": "Correo-Rotulacion",
                "excel_row": item.excel_row,
                "reference": item.subject,
                "state": "Finalizado",
                "protocol": item.protocol,
            }
        )
    return refs


def protocol_inventory_frame(refs: list[dict[str, Any]]) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    for ref in refs:
        for protocol in split_protocols(ref["protocol"]):
            rows.append(
                {
                    "Protocolo": protocol,
                    "Hoja": ref["sheet"],
                    "QR o referencia": ref["reference"],
                    "Estado": ref["state"],
                }
            )
    if not rows:
        return pd.DataFrame(columns=["Protocolo", "Hoja", "QR o referencia", "Estado"])
    return pd.DataFrame(rows).drop_duplicates().sort_values(["Protocolo", "Hoja"])


def audit_state_stage(context: dict[str, Any]) -> None:
    validations = context["validations"]
    qr_aliases = context.get("qr_aliases", {})
    if qr_aliases:
        corrections = ", ".join(
            f"{original} → {corrected}"
            for original, corrected in sorted(qr_aliases.items())
        )
        st.info(
            "Se corrigieron QR truncados porque PMS tenía una única coincidencia con "
            f"un carácter final adicional: {corrections}. La corrección quedará en la auditoría."
        )
    duplicates = [
        item for item in validations if item.result in FLOW20_DUPLICATE_RESULTS
    ]
    issues = [
        item
        for item in validations
        if item.result != STATUS_MATCH
        and item.result not in FLOW20_DUPLICATE_RESULTS
    ]
    matches = sum(item.result == STATUS_MATCH for item in validations)
    decisions = context.setdefault("state_decisions", {})
    completed = sum(validation_audit_key(item) in decisions for item in issues)
    render_summary_panel(
        eyebrow="Etapa 1",
        title="Auditoría de estados",
        subtitle=(
            "Cada diferencia se decide una sola vez. Los hitos de Flujo 2.0 "
            "que pertenecen a fechas anteriores se excluyen automáticamente."
        ),
        cards=[
            ("Coinciden", matches, "No requieren intervención", "success"),
            (
                "Duplicados históricos",
                len(duplicates),
                "Excluidos sin volver a auditarlos",
                "neutral",
            ),
            ("Por revisar", len(issues) - completed, "Decisiones pendientes", "warning"),
            ("Revisados", completed, "Decisiones guardadas", "teal"),
        ],
    )
    if duplicates:
        with st.expander(
            f"Hitos de Flujo 2.0 excluidos por redundancia ({len(duplicates)})"
        ):
            st.dataframe(
                pd.DataFrame(
                    [
                        {
                            "QR": item.qr,
                            "Estado secundario": item.reported_secondary_state,
                            "Resultado": item.result,
                            "Fecha entrega ODM": item.odm_datetime,
                            "Fecha entrega dibujo": item.drawing_datetime,
                        }
                        for item in duplicates
                    ]
                ),
                use_container_width=True,
                hide_index=True,
            )
    if not issues:
        st.success(
            "No hay diferencias que requieran decisión manual. "
            "Los duplicados históricos ya quedaron excluidos y trazados."
            if duplicates
            else "No hay diferencias de estado. Puedes continuar."
        )
    else:
        cursor = min(st.session_state.get("audit_state_cursor", 0), len(issues) - 1)
        st.session_state.audit_state_cursor = cursor
        item = issues[cursor]
        key = validation_audit_key(item)
        saved = decisions.get(key, {})
        case_card(
            item.qr,
            f"Caso {cursor + 1} de {len(issues)} · {item.sheet} · fila {item.excel_row}",
            [
                f"Bitácora: {item.reported_state or 'vacío'}",
                (
                    f"Estado secundario: {item.reported_secondary_state or 'sin dato'}"
                    if item.sheet == SHEET_FLOW20
                    else ""
                ),
                f"PMS del día: {item.pms_state or 'sin movimiento'}",
                f"Profesional bitácora: {item.professional or 'sin dato'}",
                f"Profesional del evento PMS: {item.pms_actor or 'sin dato'}",
                f"Último anterior: {item.previous_state or 'sin dato'}",
                f"Resultado inicial: {item.result}",
                f"Llegada propuesta: {item.arrival_datetime:%d/%m/%Y %H:%M:%S}" if item.arrival_datetime else "Llegada propuesta: sin dato",
                f"Salida propuesta: {item.exit_datetime:%d/%m/%Y %H:%M:%S}" if item.exit_datetime else "Salida propuesta: sin dato",
                (
                    f"Entrega ODM propuesta: {item.odm_datetime:%d/%m/%Y %H:%M:%S}"
                    if item.sheet == SHEET_FLOW20 and item.odm_datetime
                    else "Entrega ODM propuesta: sin dato"
                    if item.sheet == SHEET_FLOW20
                    else ""
                ),
                (
                    f"Entrega dibujo propuesta: {item.drawing_datetime:%d/%m/%Y %H:%M:%S}"
                    if item.sheet == SHEET_FLOW20 and item.drawing_datetime
                    else "Entrega dibujo propuesta: pendiente/sin dato"
                    if item.sheet == SHEET_FLOW20
                    else ""
                ),
                (
                    "Drawing Completed (request): "
                    + (
                        "Sí"
                        if item.request_drawing_completed is True
                        else "No"
                        if item.request_drawing_completed is False
                        else "sin dato"
                    )
                    if item.sheet == SHEET_FLOW20
                    else ""
                ),
                (
                    f"Evidencia: {item.flow20_evidence or 'sin dato'}"
                    if item.sheet == SHEET_FLOW20
                    else ""
                ),
                f"Devolución del mismo estado: {item.pms_observation or 'sin observación'}",
            ],
        )
        options = ["Conservar bitácora", "Corregir", "Excluir"]
        action_index = {"keep": 0, "correct": 1, "exclude": 2}.get(saved.get("action"), 0)
        choice = st.radio(
            "Decisión",
            options,
            index=action_index,
            horizontal=True,
            key=f"state_choice_{key}",
        )
        target_state = item.reported_state
        secondary_state = item.reported_secondary_state
        agp = ""
        protocol = ""
        odm = ""
        if choice == "Corregir":
            state_options = sorted(
                NEW_STATES
                if item.sheet == SHEET_NEW
                else FLOW20_STATES
                if item.sheet == SHEET_FLOW20
                else LEGACY_STATES
            )
            default_state = saved.get("state") or item.pms_state or item.reported_state
            state_index = state_options.index(default_state) if default_state in state_options else 0
            target_state = st.selectbox(
                "Estado correcto",
                state_options,
                index=state_index,
                key=f"state_target_{key}",
            )
        is_flow20_odm = item.sheet == SHEET_FLOW20 and target_state == FLOW20_ODM_STATE
        if is_flow20_odm:
            secondary_options = [
                FLOW20_SECONDARY_PENDING,
                FLOW20_SECONDARY_COMPLETED,
            ]
            default_secondary = (
                saved.get("secondary_state")
                or item.reported_secondary_state
                or FLOW20_SECONDARY_PENDING
            )
            secondary_index = (
                secondary_options.index(default_secondary)
                if default_secondary in secondary_options
                else 0
            )
            secondary_state = st.selectbox(
                "Estado secundario",
                secondary_options,
                index=secondary_index,
                help=(
                    "Indica si se entregó únicamente la ODM o si también se "
                    "finalizó el dibujo."
                ),
                key=f"state_secondary_{key}",
            )
        needs_element_data = (
            choice == "Corregir"
            and (target_state in FINAL_ELEMENT_STATES or is_flow20_odm)
        )
        if needs_element_data:
            columns = st.columns(3 if is_flow20_odm else 2)
            agp = columns[0].text_input(
                "Código(s) AGP, separados por ;",
                value=str(saved.get("agp") or item.agp or ""),
                help="Obligatorio. Si hay varios, usa ;.",
                key=f"state_agp_{key}",
            )
            protocol = columns[1].text_input(
                "Protocolo(s), separados por ;",
                value=str(saved.get("protocol") or item.protocol or ""),
                help="Obligatorio. Si hay varios, usa ;.",
                key=f"state_protocol_{key}",
            )
            if is_flow20_odm:
                odm = columns[2].text_input(
                    "ODM",
                    value=str(saved.get("odm") or item.odm or ""),
                    help="La entrega en estado 3 debe llevar una ODM.",
                    key=f"state_odm_{key}",
                )
        elif choice == "Corregir" and target_state in REJECTION_STATES:
            st.info(
                "Es una devolución: no se solicita AGP ni protocolo. "
                "La observación se recuperará automáticamente del PMS."
            )
        needs_flow20_timing = choice != "Excluir" and is_flow20_odm
        needs_timing_data = (
            choice != "Excluir"
            and not needs_flow20_timing
            and target_state in DRAWING_OUTPUT_STATES
        )
        arrival_datetime = ""
        exit_datetime = ""
        odm_datetime = ""
        drawing_datetime = ""
        if needs_flow20_timing:
            timing_columns = st.columns(3)
            arrival_default = (
                str(saved.get("arrival_datetime") or "")
                or (
                    item.arrival_datetime.strftime("%d/%m/%Y %H:%M:%S")
                    if item.arrival_datetime
                    else ""
                )
            )
            odm_default = (
                str(saved.get("odm_datetime") or "")
                or (
                    item.odm_datetime.strftime("%d/%m/%Y %H:%M:%S")
                    if item.odm_datetime
                    else ""
                )
            )
            drawing_default = (
                str(saved.get("drawing_datetime") or "")
                or (
                    item.drawing_datetime.strftime("%d/%m/%Y %H:%M:%S")
                    if item.drawing_datetime
                    else ""
                )
            )
            if secondary_state == FLOW20_SECONDARY_PENDING:
                drawing_default = ""
            arrival_datetime = timing_columns[0].text_input(
                "Fecha/hora de llegada a dibujo",
                value=arrival_default,
                placeholder="DD/MM/AAAA HH:MM:SS",
                key=f"state_arrival_{key}",
            )
            odm_datetime = timing_columns[1].text_input(
                "Fecha/hora entrega ODM",
                value=odm_default,
                placeholder="DD/MM/AAAA HH:MM:SS",
                key=f"state_odm_datetime_{key}",
            )
            drawing_datetime = timing_columns[2].text_input(
                "Fecha/hora entrega dibujo",
                value=drawing_default,
                placeholder=(
                    "DD/MM/AAAA HH:MM:SS"
                    if secondary_state == FLOW20_SECONDARY_COMPLETED
                    else "Déjalo vacío: dibujo pendiente"
                ),
                disabled=secondary_state == FLOW20_SECONDARY_PENDING,
                key=f"state_drawing_datetime_{key}_{secondary_state}",
            )
            st.caption(
                "Flujo 2.0 conserva por separado la llegada, la entrega de la ODM "
                "y la entrega del dibujo."
            )
        if needs_timing_data:
            timing_columns = st.columns(2)
            arrival_default = (
                str(saved.get("arrival_datetime") or "")
                or (
                    item.arrival_datetime.strftime("%d/%m/%Y %H:%M:%S")
                    if item.arrival_datetime
                    else ""
                )
            )
            exit_default = (
                str(saved.get("exit_datetime") or "")
                or (
                    item.exit_datetime.strftime("%d/%m/%Y %H:%M:%S")
                    if item.exit_datetime
                    else ""
                )
            )
            arrival_datetime = timing_columns[0].text_input(
                "Fecha/hora de llegada",
                value=arrival_default,
                placeholder="DD/MM/AAAA HH:MM:SS",
                key=f"state_arrival_{key}",
            )
            exit_datetime = timing_columns[1].text_input(
                "Fecha/hora de salida",
                value=exit_default,
                placeholder="DD/MM/AAAA HH:MM:SS",
                key=f"state_exit_{key}",
            )
            st.caption(
                "Confirma las horas propuestas por PMS o corrígelas. Ambas quedan registradas en la auditoría."
            )
        detail = st.text_area(
            "Motivo o nota de auditoría",
            value=str(saved.get("detail") or ""),
            help="Útil al conservar o excluir; opcional al corregir.",
            key=f"state_detail_{key}",
        )
        save = st.button(
            "Guardar decisión",
            type="primary",
            use_container_width=True,
            key=f"state_save_{key}",
        )
        if save:
            timing_error = ""
            if needs_flow20_timing:
                try:
                    parsed_arrival = parse_audit_datetime(arrival_datetime)
                    parsed_odm = parse_audit_datetime(odm_datetime)
                    parsed_drawing = parse_audit_datetime(drawing_datetime)
                    if not parsed_arrival or not parsed_odm:
                        timing_error = (
                            "Debes confirmar la llegada y la fecha/hora de entrega de la ODM."
                        )
                    elif parsed_arrival > parsed_odm:
                        timing_error = (
                            "La llegada no puede ser posterior a la entrega de la ODM."
                        )
                    elif secondary_state == FLOW20_SECONDARY_COMPLETED:
                        if not parsed_drawing:
                            timing_error = (
                                "Si el dibujo quedó finalizado debes confirmar su fecha/hora."
                            )
                        elif parsed_drawing < parsed_odm:
                            timing_error = (
                                "La entrega del dibujo no puede ser anterior a la ODM."
                            )
                    elif parsed_drawing:
                        timing_error = (
                            "Un caso pendiente de dibujo no debe tener fecha de entrega del dibujo."
                        )
                except BitacoraError as exc:
                    timing_error = str(exc)
            elif needs_timing_data:
                try:
                    parsed_arrival = parse_audit_datetime(arrival_datetime)
                    parsed_exit = parse_audit_datetime(exit_datetime)
                    if not parsed_arrival or not parsed_exit:
                        timing_error = "Debes confirmar la fecha/hora de llegada y de salida."
                    elif parsed_arrival > parsed_exit:
                        timing_error = "La llegada no puede ser posterior a la salida."
                except BitacoraError as exc:
                    timing_error = str(exc)
            if timing_error:
                st.error(timing_error)
            elif needs_element_data and (
                not agp.strip()
                or not protocol.strip()
                or is_flow20_odm and not odm.strip()
            ):
                st.error(
                    "Para corregir este caso debes indicar AGP, protocolo"
                    + (" y ODM." if is_flow20_odm else ".")
                )
            elif needs_element_data and has_invalid_agp_separator(agp):
                st.error("Cuando haya varios AGP, sepáralos únicamente con ;.")
            elif needs_element_data and has_invalid_protocol_separator(protocol):
                st.error("Cuando haya varios protocolos, sepáralos únicamente con ;.")
            elif needs_element_data and has_invalid_protocol_prefix(protocol):
                st.error("Todos los protocolos deben iniciar por 1000.")
            else:
                action = {"Conservar bitácora": "keep", "Corregir": "correct", "Excluir": "exclude"}[choice]
                decisions[key] = {
                    "action": action,
                    "state": target_state,
                    "secondary_state": secondary_state,
                    "agp": agp.strip(),
                    "protocol": protocol.strip(),
                    "odm": odm.strip(),
                    "arrival_datetime": arrival_datetime.strip(),
                    "exit_datetime": exit_datetime.strip(),
                    "odm_datetime": odm_datetime.strip(),
                    "drawing_datetime": drawing_datetime.strip(),
                    "detail": detail.strip(),
                }
                context["state_decisions"] = decisions
                st.session_state.audit_state_cursor = min(cursor + 1, len(issues) - 1)
                st.rerun()

        navigation = st.columns([1, 2, 1])
        if navigation[0].button("← Anterior", disabled=cursor == 0, use_container_width=True):
            st.session_state.audit_state_cursor = cursor - 1
            st.rerun()
        navigation[1].progress(completed / len(issues), text=f"{completed} de {len(issues)} decisiones guardadas")
        if navigation[2].button("Siguiente →", disabled=cursor == len(issues) - 1, use_container_width=True):
            st.session_state.audit_state_cursor = cursor + 1
            st.rerun()

    if st.button(
        "Cerrar estados y continuar",
        type="primary",
        disabled=completed != len(issues),
        use_container_width=True,
    ):
        working_bytes, working_validations, entries = apply_state_audit_decisions(
            context["workbook_bytes"],
            validations,
            context["latest_on_day"],
            decisions,
            context.get("rejection_comments", {}),
            context.get("timelines", {}),
            context["production_date"],
        )
        context["working_bytes"] = working_bytes
        context["working_validations"] = apply_drawing_timings(
            working_validations,
            context.get("timelines", {}),
            context["production_date"],
        )
        context["audit_entries"] = [
            *context.get("audit_entries", []),
            *entries,
        ]
        st.session_state.audit_step = 2
        st.rerun()


def audit_unreported_stage(context: dict[str, Any]) -> None:
    cases = context["unreported_cases"]
    decisions = context.setdefault("unreported_decisions", {})
    completed = sum(case.qr in decisions for case in cases)
    render_summary_panel(
        eyebrow="Etapa 2",
        title="Casos no reportados",
        subtitle="Movimientos del día que no estaban en la bitácora. Decide si se incorporan o se descartan.",
        cards=[
            ("Detectados", len(cases), "Hallazgos del log", "info"),
            ("Por revisar", len(cases) - completed, "Aún sin decisión", "warning"),
            ("Revisados", completed, "Decisiones guardadas", "teal"),
        ],
    )
    if not cases:
        st.success("No se encontraron casos no reportados.")
    else:
        cursor = min(st.session_state.get("audit_unreported_cursor", 0), len(cases) - 1)
        st.session_state.audit_unreported_cursor = cursor
        case = cases[cursor]
        saved = decisions.get(case.qr, {})
        case_card(
            case.qr,
            f"Caso {cursor + 1} de {len(cases)} · sugerido para {case.flow}",
            [
                f"Estado PMS: {case.state}",
                f"Fecha/hora: {case.occurred_at:%d/%m/%Y %H:%M:%S}",
                f"Profesional PMS: {case.actor or 'sin dato'}",
                f"Devolución: {case.comment or 'sin observación'}",
            ],
        )
        choice = st.radio(
            "Decisión",
            ["Agregar a la bitácora", "No agregar"],
            index=1 if saved.get("action") == "ignore" else 0,
            horizontal=True,
            key=f"unreported_choice_{case.qr}",
        )
        adding = choice == "Agregar a la bitácora"
        sheet = case.flow
        agp = ""
        protocol = ""
        origin = ""
        user = ""
        elements = 0
        if adding:
            sheet = st.selectbox(
                "Hoja destino",
                [SHEET_NEW, SHEET_OLD],
                index=0 if (saved.get("sheet") or case.flow) == SHEET_NEW else 1,
                key=f"unreported_sheet_{case.qr}",
            )
            needs_element_data = case.state in FINAL_ELEMENT_STATES
            if needs_element_data:
                left, right = st.columns(2)
                agp = left.text_input(
                    "Código(s) AGP, separados por ;",
                    value=str(saved.get("agp") or ""),
                    key=f"unreported_agp_{case.qr}",
                )
                protocol = right.text_input(
                    "Protocolo(s), separados por ;",
                    value=str(saved.get("protocol") or ""),
                    key=f"unreported_protocol_{case.qr}",
                )
            else:
                st.info(
                    "Es una devolución: no se solicita AGP ni protocolo. "
                    "Se incorporará la observación más reciente del mismo estado."
                )
            left, right = st.columns(2)
            saved_origin = str(saved.get("origin") or "")
            origin_index = (
                UPDATE_ORIGINS.index(saved_origin)
                if saved_origin in UPDATE_ORIGINS
                else None
            )
            origin_choice = left.selectbox(
                "Origen de actualización",
                options=UPDATE_ORIGINS,
                index=origin_index,
                placeholder="Selecciona un origen",
                key=f"unreported_origin_{case.qr}",
            )
            origin = origin_choice or ""
            user = right.text_input(
                "Profesional",
                value=str(saved.get("user") or case.actor or ""),
                key=f"unreported_user_{case.qr}",
            )
            elements = left.number_input(
                "Elementos estimados (opcional)",
                min_value=0,
                value=int(saved.get("elements") or 0),
                key=f"unreported_elements_{case.qr}",
            )
        else:
            needs_element_data = False
        detail = st.text_area(
            "Motivo o nota",
            value=str(saved.get("detail") or ""),
            key=f"unreported_detail_{case.qr}",
        )
        save = st.button(
            "Guardar decisión",
            type="primary",
            use_container_width=True,
            key=f"unreported_save_{case.qr}",
        )
        if save:
            if adding and not origin:
                st.error("Selecciona el origen de actualización del caso.")
            elif needs_element_data and (not agp.strip() or not protocol.strip()):
                st.error("Para agregar el caso debes indicar AGP y protocolo. P/TRATA se obtendrán después.")
            elif needs_element_data and has_invalid_agp_separator(agp):
                st.error("Cuando haya varios AGP, sepáralos únicamente con ;.")
            elif needs_element_data and has_invalid_protocol_separator(protocol):
                st.error("Cuando haya varios protocolos, sepáralos únicamente con ;.")
            elif needs_element_data and has_invalid_protocol_prefix(protocol):
                st.error("Todos los protocolos deben iniciar por 1000.")
            else:
                decisions[case.qr] = {
                    "action": "add" if adding else "ignore",
                    "sheet": sheet,
                    "agp": agp.strip(),
                    "protocol": protocol.strip(),
                    "origin": origin.strip(),
                    "user": user.strip(),
                    "elements": elements or None,
                    "detail": detail.strip(),
                }
                context["unreported_decisions"] = decisions
                st.session_state.audit_unreported_cursor = min(cursor + 1, len(cases) - 1)
                st.rerun()

        navigation = st.columns([1, 2, 1])
        if navigation[0].button("← Anterior", key="unreported_previous", disabled=cursor == 0, use_container_width=True):
            st.session_state.audit_unreported_cursor = cursor - 1
            st.rerun()
        navigation[1].progress(completed / len(cases), text=f"{completed} de {len(cases)} decisiones guardadas")
        if navigation[2].button("Siguiente →", key="unreported_next", disabled=cursor == len(cases) - 1, use_container_width=True):
            st.session_state.audit_unreported_cursor = cursor + 1
            st.rerun()

    if st.button(
        "Cerrar casos no reportados y continuar",
        type="primary",
        disabled=completed != len(cases),
        use_container_width=True,
    ):
        working_bytes, working_validations, entries = add_unreported_audit_rows(
            context["working_bytes"],
            context["working_validations"],
            cases,
            decisions,
        )
        context["working_bytes"] = working_bytes
        context["working_validations"] = apply_drawing_timings(
            working_validations,
            context.get("timelines", {}),
            context["production_date"],
        )
        context["audit_entries"].extend(entries)
        st.session_state.audit_step = 3
        st.rerun()


def render_protocol_prevalidation(context: dict[str, Any], refs: list[dict[str, Any]]) -> bool:
    issues = [
        ref for ref in refs
        if (
            not split_protocols(ref["protocol"])
            or has_invalid_protocol_separator(ref["protocol"])
            or has_invalid_protocol_prefix(ref["protocol"])
        )
    ]
    decisions = context.setdefault("protocol_decisions", {})
    completed = sum(ref["key"] in decisions for ref in issues)
    if not issues:
        return True

    st.subheader("3A. Completar protocolos antes de ir a SAP")
    st.caption(
        "Estos casos están vacíos, usan un separador distinto de punto y coma "
        "o contienen protocolos que no comienzan por 1000."
    )
    cursor = min(st.session_state.get("audit_protocol_cursor", 0), len(issues) - 1)
    st.session_state.audit_protocol_cursor = cursor
    ref = issues[cursor]
    saved = decisions.get(ref["key"], {})
    case_card(
        ref["reference"],
        f"Referencia {cursor + 1} de {len(issues)} · {ref['sheet']} · fila {ref['excel_row']}",
        [f"Estado: {ref['state']}", f"Valor actual: {ref['protocol'] or 'vacío'}"],
    )
    with st.form(f"protocol_form_{ref['key']}"):
        choice = st.radio(
            "Decisión",
            ["Ingresar/corregir protocolo", "No aplica"],
            index=1 if saved.get("action") == "no_apply" else 0,
            horizontal=True,
        )
        setting = choice == "Ingresar/corregir protocolo"
        protocol = st.text_input(
            "Protocolo(s), separados obligatoriamente por ;",
            value=str(saved.get("protocol") or ref["protocol"] or ""),
        )
        detail = st.text_area("Justificación", value=str(saved.get("detail") or ""))
        save = st.form_submit_button("Guardar decisión", type="primary", use_container_width=True)
    if save:
        if setting and (not split_protocols(protocol) or has_invalid_protocol_separator(protocol)):
            st.error("Ingresa al menos un protocolo y usa ; cuando haya varios.")
        elif setting and has_invalid_protocol_prefix(protocol):
            invalid = ", ".join(invalid_protocol_prefixes(protocol))
            st.error(f"Todos los protocolos deben iniciar por 1000. Revisa: {invalid}.")
        else:
            decisions[ref["key"]] = {
                "action": "set" if setting else "no_apply",
                "sheet": ref["sheet"],
                "excel_row": ref["excel_row"],
                "reference": ref["reference"],
                "current_protocol": ref["protocol"],
                "protocol": protocol.strip(),
                "detail": detail.strip(),
            }
            context["protocol_decisions"] = decisions
            st.session_state.audit_protocol_cursor = min(cursor + 1, len(issues) - 1)
            st.rerun()
    st.progress(completed / len(issues), text=f"{completed} de {len(issues)} referencias revisadas")
    if completed == len(issues):
        if st.button("Aplicar protocolos y preparar listado", type="primary", use_container_width=True):
            workbook_bytes, validations, entries = apply_protocol_audit_decisions(
                context["working_bytes"], context["working_validations"], decisions
            )
            context["working_bytes"] = workbook_bytes
            context["working_validations"] = validations
            context["audit_entries"].extend(entries)
            context["protocol_prevalidation_done"] = True
            st.rerun()
    return False


def audit_protocol_stage(context: dict[str, Any], export_file: Any) -> None:
    refs = protocol_references(context["working_bytes"], context["working_validations"])
    if not context.get("protocol_prevalidation_done"):
        ready = render_protocol_prevalidation(context, refs)
        if not ready:
            return
        context["protocol_prevalidation_done"] = True
        refs = protocol_references(context["working_bytes"], context["working_validations"])

    no_apply_reference_keys = {
        key
        for key, decision in context.get("protocol_decisions", {}).items()
        if decision.get("action") == "no_apply"
    }
    refs = [ref for ref in refs if ref["key"] not in no_apply_reference_keys]

    inventory = protocol_inventory_frame(refs)
    protocols = sorted(set(inventory["Protocolo"].tolist())) if not inventory.empty else []
    missing_decisions = context.get("missing_protocol_decisions", {})
    unresolved_count = len(
        [
            protocol
            for protocol in context.get("missing_protocols", [])
            if protocol not in missing_decisions
        ]
    )
    render_summary_panel(
        eyebrow="Etapa 3",
        title="Protocolos y cantidades P/TRATA",
        subtitle="Descarga una sola lista, obtén el EXPORT de SAP y vuelve aquí para completar todos los conteos.",
        cards=[
            ("Referencias", len(refs), "Flujos y correos finalizados", "info"),
            ("Protocolos únicos", len(protocols), "Listado para la extracción", "teal"),
            ("Sin resolver", unresolved_count, "Faltantes del último EXPORT", "warning"),
        ],
    )
    st.subheader("3B. Descargar listado para SAP")
    st.dataframe(inventory, use_container_width=True, hide_index=True)
    st.download_button(
        "Descargar listado de protocolos (Excel)",
        data=build_protocol_inventory_workbook(inventory.to_dict("records")),
        file_name=f"Protocolos {context['production_date']:%d-%m-%Y}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
        disabled=inventory.empty,
    )

    st.subheader("3C. Cargar y procesar el EXPORT")
    if export_file is None:
        if "protocol_counts" not in context:
            st.info("Ahora carga el EXPORT de SAP en el panel izquierdo. Tu auditoría queda guardada mientras lo descargas.")
            return
        st.caption("Se conserva el último EXPORT procesado. Puedes cargar otro archivo para reemplazarlo.")
    else:
        export_bytes = export_file.getvalue()
        export_id = f"{export_file.name}:{sha256(export_bytes).hexdigest()}"
        if context.get("processed_export_id") == export_id:
            st.success(f"EXPORT procesado: {export_file.name}")
        else:
            counts = read_export_protocol_counts(export_bytes)
            missing = sorted(protocol for protocol in protocols if protocol not in counts)
            decisions = {
                protocol: decision
                for protocol, decision in context.get("missing_protocol_decisions", {}).items()
                if protocol in missing
            }
            resolved_counts, no_apply = protocol_counts_with_audit_decisions(
                counts,
                decisions,
            )
            validations = apply_protocol_counts_to_validations(
                context["working_validations"],
                resolved_counts,
            )
            email_results = evaluate_email_quantities(
                context["working_bytes"],
                resolved_counts,
            )
            validations, email_results = apply_protocol_audit_results(
                validations,
                email_results,
                context.get("protocol_decisions", {}),
                no_apply,
                resolved_counts,
            )
            context["protocol_counts"] = counts
            context["resolved_protocol_counts"] = resolved_counts
            context["working_validations"] = validations
            context["email_quantity_results"] = email_results
            context["missing_protocols"] = missing
            context["missing_protocol_decisions"] = decisions
            context["protocols_no_apply"] = no_apply
            context["processed_export_id"] = export_id
            context.pop("final_output", None)
            st.rerun()

    if "protocol_counts" not in context:
        return

    missing = context.get("missing_protocols", [])
    decisions = context.setdefault("missing_protocol_decisions", {})
    unresolved = [protocol for protocol in missing if protocol not in decisions]
    correction_message = context.pop("protocol_correction_message", "")
    if correction_message:
        st.success(correction_message)
    render_summary_panel(
        eyebrow="Resultado del EXPORT",
        title="Cruce de cantidades",
        subtitle=(
            "Cada protocolo conserva su PEL y TRATA en la misma posición; "
            "el total de elementos se calcula con la suma completa."
        ),
        cards=[
            ("Encontrados", len(protocols) - len(missing), "Presentes en el EXPORT", "success"),
            ("Faltantes", len(missing), "No aparecen en el EXPORT", "warning"),
            ("Resueltos", len(missing) - len(unresolved), "Con decisión de auditoría", "teal"),
            ("Pendientes", len(unresolved), "Requieren tu decisión", "danger" if unresolved else "success"),
        ],
    )
    found_protocols = [protocol for protocol in protocols if protocol not in missing]
    if found_protocols:
        with st.expander(f"Protocolos encontrados ({len(found_protocols)})"):
            st.dataframe(
                pd.DataFrame({"Protocolo encontrado": found_protocols}),
                hide_index=True,
                use_container_width=True,
            )
    if missing:
        st.warning(
            f"El EXPORT no contiene {len(missing)} protocolo(s). Revísalos uno por uno: "
            "puedes corregir el código, justificar cantidades manuales o indicar que no aplica."
        )
        cursor = min(
            int(st.session_state.get("audit_missing_protocol_cursor", 0)),
            len(missing) - 1,
        )
        protocol = missing[cursor]
        saved = decisions.get(protocol, {})
        case_card(
            protocol,
            f"Protocolo {cursor + 1} de {len(missing)}",
            [
                "No se encontró en el EXPORT cargado.",
                "El valor inferior se puede seleccionar y copiar para buscarlo en SAP o AGP.",
            ],
        )
        st.text_input(
            "Protocolo para copiar o buscar",
            value=protocol,
            key=f"audit_copy_protocol_{protocol}",
        )
        affected = inventory[inventory["Protocolo"] == protocol]
        if not affected.empty:
            st.caption("Referencias afectadas")
            st.dataframe(
                affected[["Hoja", "QR o referencia", "Estado"]],
                hide_index=True,
                use_container_width=True,
            )

        action_labels = (
            "Corregir protocolo",
            "Registrar cantidades manualmente",
            "No aplica",
        )
        action_by_label = {
            "Corregir protocolo": "correct",
            "Registrar cantidades manualmente": "manual",
            "No aplica": "no_apply",
        }
        saved_label = {
            "correct": "Corregir protocolo",
            "manual": "Registrar cantidades manualmente",
            "no_apply": "No aplica",
        }.get(saved.get("action"), action_labels[0])
        action_label = st.radio(
            "¿Qué ocurrió con este protocolo?",
            action_labels,
            index=action_labels.index(saved_label),
            horizontal=True,
            key=f"audit_missing_protocol_action_{protocol}",
        )
        action = action_by_label[action_label]
        with st.form(f"audit_missing_protocol_form_{protocol}"):
            corrected_protocol = ""
            pel = int(saved.get("pel", 0) or 0)
            tratas = int(saved.get("tratas", 0) or 0)
            justification = str(saved.get("justification", "") or "")
            evidence_file = None
            if action == "correct":
                corrected_protocol = st.text_input(
                    "Protocolo correcto",
                    value=str(saved.get("corrected_protocol", "") or ""),
                    placeholder="1000... (si son varios, sepáralos con ;)",
                )
                st.caption("La corrección se aplicará a todas las referencias donde aparezca el código actual.")
            else:
                if action == "manual":
                    left, right = st.columns(2)
                    with left:
                        pel = st.number_input(
                            "Cantidad PEL comprobada",
                            min_value=0,
                            value=pel,
                            step=1,
                        )
                    with right:
                        tratas = st.number_input(
                            "Cantidad TRATAS comprobada",
                            min_value=0,
                            value=tratas,
                            step=1,
                        )
                justification = st.text_area(
                    "Justificación",
                    value=justification,
                    placeholder=(
                        "Explica por qué se usan cantidades manuales o por qué el protocolo no aplica."
                    ),
                )
                evidence_file = st.file_uploader(
                    "Imagen de soporte" + (" (obligatoria)" if action == "manual" else " (opcional)"),
                    type=["png", "jpg", "jpeg"],
                    key=f"audit_protocol_evidence_{protocol}_{action}",
                )
                if saved.get("image_bytes") and evidence_file is None:
                    st.caption(f"Soporte guardado: {saved.get('filename', 'imagen')}")
            save_decision = st.form_submit_button(
                "Guardar decisión y continuar",
                type="primary",
                use_container_width=True,
            )

        if save_decision:
            if action == "correct":
                if not split_protocols(corrected_protocol):
                    st.error("Indica el protocolo correcto.")
                elif has_invalid_protocol_separator(corrected_protocol):
                    st.error("Si son varios protocolos, sepáralos únicamente con ;.")
                elif has_invalid_protocol_prefix(corrected_protocol):
                    invalid = ", ".join(invalid_protocol_prefixes(corrected_protocol))
                    st.error(f"Todos los protocolos deben iniciar por 1000. Revisa: {invalid}.")
                elif split_protocols(corrected_protocol) == [protocol]:
                    st.error("El protocolo correcto debe ser diferente del valor actual.")
                else:
                    workbook_bytes, corrected_validations, entries = replace_protocol_references(
                        context["working_bytes"],
                        context["working_validations"],
                        {protocol: corrected_protocol},
                    )
                    corrected_refs = protocol_references(workbook_bytes, corrected_validations)
                    corrected_refs = [
                        ref for ref in corrected_refs if ref["key"] not in no_apply_reference_keys
                    ]
                    corrected_inventory = protocol_inventory_frame(corrected_refs)
                    corrected_protocols = (
                        set(corrected_inventory["Protocolo"].tolist())
                        if not corrected_inventory.empty
                        else set()
                    )
                    counts = context["protocol_counts"]
                    new_missing = sorted(
                        item for item in corrected_protocols if item not in counts
                    )
                    decisions.pop(protocol, None)
                    decisions = {
                        item: decision
                        for item, decision in decisions.items()
                        if item in new_missing
                    }
                    resolved_counts, no_apply = protocol_counts_with_audit_decisions(
                        counts,
                        decisions,
                    )
                    corrected_validations = apply_protocol_counts_to_validations(
                        corrected_validations,
                        resolved_counts,
                    )
                    corrected_emails = evaluate_email_quantities(
                        workbook_bytes,
                        resolved_counts,
                    )
                    corrected_validations, corrected_emails = apply_protocol_audit_results(
                        corrected_validations,
                        corrected_emails,
                        context.get("protocol_decisions", {}),
                        no_apply,
                        resolved_counts,
                    )
                    context["working_bytes"] = workbook_bytes
                    context["working_validations"] = corrected_validations
                    context["email_quantity_results"] = corrected_emails
                    context["missing_protocols"] = new_missing
                    context["missing_protocol_decisions"] = decisions
                    context["resolved_protocol_counts"] = resolved_counts
                    context["protocols_no_apply"] = no_apply
                    context["audit_entries"].extend(entries)
                    replacements = context.setdefault("protocol_replacements", {})
                    replacements[protocol] = corrected_protocol.strip()
                    context["protocol_correction_message"] = (
                        f"Se reemplazó {protocol} por {corrected_protocol.strip()} "
                        f"en {len(entries)} referencia(s)."
                    )
                    context.pop("final_output", None)
                    st.session_state.audit_missing_protocol_cursor = min(
                        cursor,
                        max(len(new_missing) - 1, 0),
                    )
                    st.rerun()
            elif not justification.strip():
                st.error("Escribe una justificación para guardar esta decisión.")
            elif action == "manual" and not (
                evidence_file is not None or saved.get("image_bytes")
            ):
                st.error("Adjunta una imagen que respalde las cantidades manuales.")
            else:
                image_bytes = (
                    evidence_file.getvalue()
                    if evidence_file is not None
                    else saved.get("image_bytes")
                )
                filename = (
                    evidence_file.name
                    if evidence_file is not None
                    else saved.get("filename", "")
                )
                decisions[protocol] = {
                    "action": action,
                    "pel": int(pel) if action == "manual" else 0,
                    "tratas": int(tratas) if action == "manual" else 0,
                    "justification": justification.strip(),
                    "filename": filename,
                    "image_bytes": image_bytes,
                }
                resolved_counts, no_apply = protocol_counts_with_audit_decisions(
                    context["protocol_counts"],
                    decisions,
                )
                validations = apply_protocol_counts_to_validations(
                    context["working_validations"],
                    resolved_counts,
                )
                email_results = evaluate_email_quantities(
                    context["working_bytes"],
                    resolved_counts,
                )
                validations, email_results = apply_protocol_audit_results(
                    validations,
                    email_results,
                    context.get("protocol_decisions", {}),
                    no_apply,
                    resolved_counts,
                )
                context["working_validations"] = validations
                context["email_quantity_results"] = email_results
                context["missing_protocol_decisions"] = decisions
                context["resolved_protocol_counts"] = resolved_counts
                context["protocols_no_apply"] = no_apply
                context.pop("final_output", None)
                st.session_state.audit_missing_protocol_cursor = min(
                    cursor + 1,
                    len(missing) - 1,
                )
                st.rerun()

        nav_left, nav_progress, nav_right = st.columns([1, 3, 1])
        with nav_left:
            if st.button("← Anterior", disabled=cursor == 0, use_container_width=True):
                st.session_state.audit_missing_protocol_cursor = cursor - 1
                st.rerun()
        with nav_progress:
            st.progress(
                (len(missing) - len(unresolved)) / len(missing),
                text=f"{len(missing) - len(unresolved)} de {len(missing)} protocolos resueltos",
            )
        with nav_right:
            if st.button(
                "Siguiente →",
                disabled=cursor >= len(missing) - 1,
                use_container_width=True,
            ):
                st.session_state.audit_missing_protocol_cursor = cursor + 1
                st.rerun()

        if decisions:
            decision_rows = []
            for item in missing:
                decision = decisions.get(item)
                if not decision:
                    continue
                decision_rows.append(
                    {
                        "Protocolo": item,
                        "Decisión": (
                            "Cantidades manuales"
                            if decision.get("action") == "manual"
                            else "No aplica"
                        ),
                        "PEL": decision.get("pel", 0),
                        "TRATAS": decision.get("tratas", 0),
                        "Justificación": decision.get("justification", ""),
                        "Soporte": decision.get("filename", ""),
                    }
                )
            with st.expander(f"Decisiones guardadas ({len(decision_rows)})"):
                st.dataframe(pd.DataFrame(decision_rows), hide_index=True, use_container_width=True)
    else:
        st.success("Todos los protocolos del listado aparecen en el EXPORT.")

    if st.button(
        "Cerrar cantidades y generar Excel final",
        type="primary",
        disabled=bool(unresolved),
        use_container_width=True,
    ):
        protocol_audit_entries = []
        protocol_evidence = []
        for protocol, decision in sorted(decisions.items()):
            is_manual = decision.get("action") == "manual"
            protocol_audit_entries.append(
                {
                    "Etapa": "Protocolos",
                    "Hoja/Flujo": "Varias referencias",
                    "QR/Referencia": protocol,
                    "Decisión": "Cantidades manuales" if is_manual else "No aplica",
                    "Estado inicial": "No encontrado en EXPORT",
                    "Estado final": (
                        f"PEL {decision.get('pel', 0)}; TRATAS {decision.get('tratas', 0)}"
                        if is_manual
                        else "Justificado como No aplica"
                    ),
                    "AGP": "",
                    "Protocolo": protocol,
                    "Detalle": decision.get("justification", ""),
                }
            )
            protocol_evidence.append(
                {
                    "protocol": protocol,
                    "decision": "Cantidades manuales" if is_manual else "No aplica",
                    "pel": decision.get("pel", 0),
                    "tratas": decision.get("tratas", 0),
                    "justification": decision.get("justification", ""),
                    "filename": decision.get("filename", ""),
                    "image_bytes": decision.get("image_bytes"),
                }
            )
        context["audit_entries"].extend(protocol_audit_entries)
        context["protocol_evidence"] = protocol_evidence
        template_bytes = default_template_bytes()
        if template_bytes is None:
            raise BitacoraError(
                "No encontré la plantilla fija F-AGP-04 incluida con la aplicación."
            )
        context["final_output"] = build_upload_workbook(
            context["working_bytes"],
            context["working_validations"],
            matches_only=True,
            template_bytes=template_bytes,
            email_quantity_results=context["email_quantity_results"],
            unreported_cases=[],
            audit_entries=context["audit_entries"],
            protocol_evidence=protocol_evidence,
        )
        st.session_state.audit_step = 4
        st.rerun()


def audit_final_stage(context: dict[str, Any]) -> None:
    validations = context["working_validations"]
    email_results = context["email_quantity_results"]
    counts = Counter(item.result for item in validations)
    approved = sum(counts[result] for result in UPLOAD_APPROVED_RESULTS)
    historical_duplicates = sum(counts[result] for result in FLOW20_DUPLICATE_RESULTS)
    excluded = counts[STATUS_AUDIT_EXCLUDED] + historical_duplicates
    render_summary_panel(
        eyebrow="Etapa 4",
        title="Auditoría terminada",
        subtitle="El Excel conserva la estructura codificada F-AGP-04 y agrega la trazabilidad de tus decisiones.",
        cards=[
            ("Filas aprobadas", approved, "Incluidas en los flujos", "success"),
            ("Filas excluidas", excluded, "No pasan al cargue", "neutral"),
            (
                "Duplicados históricos",
                historical_duplicates,
                "Hitos ya reportables en fechas anteriores",
                "info",
            ),
            ("Decisiones", len(context["audit_entries"]), "Registradas en Auditoria guiada", "teal"),
        ],
    )
    st.success("Listo: los estados, casos no reportados y protocolos quedaron resueltos en orden.")
    output = context["final_output"]
    frame = validation_frame(validations)
    tabs = st.tabs(["Resultado final", "Cantidades", "Auditoría realizada"])
    with tabs[0]:
        st.dataframe(frame, use_container_width=True, hide_index=True)
    with tabs[1]:
        quantity_columns = [
            "Hoja", "Fila Excel", "QR", "Estado reportado", "Protocolo",
            "Cantidad PEL reportada", "Cantidad TRATAS reportada",
            "Cantidad PEL EXPORT", "Cantidad TRATAS EXPORT", "Resultado cantidades",
        ]
        quantity_keys = {
            (item.sheet, item.excel_row, item.qr)
            for item in validations
            if item.result in UPLOAD_APPROVED_RESULTS
            and validation_requires_element_counts(item)
        }
        flow_mask = frame.apply(
            lambda row: (
                row["Hoja"],
                int(row["Fila Excel"]),
                row["QR"],
            ) in quantity_keys,
            axis=1,
        )
        flow = frame.loc[flow_mask, quantity_columns]
        email = email_quantity_frame(email_results)[quantity_columns]
        st.dataframe(pd.concat([flow, email], ignore_index=True), use_container_width=True, hide_index=True)
    with tabs[2]:
        st.dataframe(pd.DataFrame(context["audit_entries"]), use_container_width=True, hide_index=True)
    downloads = st.columns(2)
    downloads[0].download_button(
        "Descargar Excel auditado y listo para cargar",
        data=output,
        file_name=output_filename(context["production_date"], matches_only=True),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
    downloads[1].download_button(
        "Descargar trazabilidad (CSV)",
        data=pd.DataFrame(context["audit_entries"]).to_csv(index=False).encode("utf-8-sig"),
        file_name=f"Auditoria {context['production_date']:%d-%m-%Y}.csv",
        mime="text/csv",
        use_container_width=True,
    )


st.title("Auditoría guiada de bitácoras PMS")
st.caption(
    "Revisa cada hallazgo una sola vez: primero estados, luego casos no reportados, "
    "después protocolos/P-TRATA y finalmente el Excel codificado listo para cargar."
)

context = current_context()
current_step = int(st.session_state.get("audit_step", 1)) if context else 0
with st.sidebar:
    st.header("Expediente de auditoría")
    workbook_file = st.file_uploader(
        "1. Bitácora del coordinador",
        type=["xlsx"],
        disabled=context is not None,
        help="El nombre debe contener la fecha DD-MM-AAAA.",
    )
    log_file = st.file_uploader(
        "2. Log del PMS",
        type=["csv"],
        disabled=context is not None,
        help="Debe cubrir la fecha indicada en el nombre de la bitácora.",
    )
    request_file = st.file_uploader(
        "3. Exported Request del PMS",
        type=["csv"],
        disabled=context is not None,
        help=(
            "Se exige cuando la bitácora contiene Flujo 2.0. Permite validar "
            "la marca Contract Drawing Completed."
        ),
    )
    export_file = st.file_uploader(
        "4. EXPORT de SAP",
        type=["xlsx"],
        disabled=current_step != 3,
        help="Se habilita únicamente durante la etapa 3, después de depurar los protocolos.",
    )
    if current_step < 3:
        st.caption("🔒 Se habilitará al llegar a Protocolos y cantidades.")
    elif current_step == 3:
        st.caption("🔓 Carga aquí el EXPORT obtenido con el listado depurado.")
    else:
        st.caption("🔒 Etapa de cantidades cerrada.")
    start = st.button(
        "Iniciar auditoría",
        type="primary",
        use_container_width=True,
        disabled=context is not None or not (workbook_file and log_file),
    )
    if context is not None:
        st.success(f"Corte: {context['production_date']:%d/%m/%Y}")
        if st.button("Reiniciar con otros archivos", use_container_width=True):
            reset_audit()
            st.rerun()
    st.caption("La plantilla HSEQ/SIG F-AGP-04 viene incluida automáticamente con la aplicación.")

if start:
    try:
        production_date = parse_production_date(workbook_file.name)
        workbook_bytes, rows = extract_workbook_rows(workbook_file.getvalue())
        flow20_rows = [row for row in rows if row["sheet"] == SHEET_FLOW20]
        if flow20_rows and request_file is None:
            raise BitacoraError(
                "Esta bitácora contiene Flujo 2.0. Carga también el archivo "
                "exported-request.csv para validar Drawing Completed."
            )
        qrs = {row["qr"] for row in rows if row["qr"]}
        allowed_states_by_qr: dict[str, set[str]] = {}
        for row in rows:
            if not row["qr"]:
                continue
            states = (
                NEW_STATES
                if row["sheet"] == SHEET_NEW
                else FLOW20_STATES
                if row["sheet"] == SHEET_FLOW20
                else LEGACY_STATES
            )
            allowed_states_by_qr.setdefault(row["qr"], set()).update(states)
        log_bytes = log_file.getvalue()
        (
            latest_on_day,
            latest_as_of,
            coverage,
            unreported_cases,
            rejection_comments,
            timelines,
            qr_aliases,
        ) = read_log_validation_data(
            log_bytes,
            qrs,
            production_date,
            allowed_states_by_qr=allowed_states_by_qr,
            include_rejection_comments=True,
            include_timelines=True,
            include_qr_aliases=True,
        )
        require_log_date_coverage(coverage, production_date)
        workbook_bytes, rows, qr_audit_entries = apply_qr_alias_corrections(
            workbook_bytes,
            rows,
            qr_aliases,
        )
        flow20_qrs = {
            row["qr"]
            for row in rows
            if row["sheet"] == SHEET_FLOW20 and row["qr"]
        }
        request_snapshots = (
            read_request_snapshots(request_file.getvalue(), flow20_qrs)
            if flow20_qrs and request_file is not None
            else {}
        )
        if flow20_qrs:
            full_flow20_timelines = read_full_log_timelines(log_bytes, flow20_qrs)
            for qr in flow20_qrs:
                timelines[qr] = full_flow20_timelines.get(qr, [])
        validations = validate_rows(
            rows,
            latest_on_day,
            latest_as_of,
            rejection_comments=rejection_comments,
            timelines=timelines,
            production_date=production_date,
            request_snapshots=request_snapshots,
        )
        duplicate_audit_entries = flow20_duplicate_audit_entries(validations)
        st.session_state.audit_context = {
            "production_date": production_date,
            "workbook_name": workbook_file.name,
            "workbook_bytes": workbook_bytes,
            "validations": validations,
            "latest_on_day": latest_on_day,
            "latest_as_of": latest_as_of,
            "rejection_comments": rejection_comments,
            "timelines": timelines,
            "request_snapshots": request_snapshots,
            "qr_aliases": qr_aliases,
            "coverage": coverage,
            "unreported_cases": unreported_cases,
            "state_decisions": {},
            "unreported_decisions": {},
            "protocol_decisions": {},
            "audit_entries": [
                *qr_audit_entries,
                *duplicate_audit_entries,
            ],
        }
        st.session_state.audit_step = 1
        st.rerun()
    except BitacoraError as exc:
        st.error(str(exc))
    except Exception as exc:
        st.exception(exc)

context = current_context()
if context is None:
    st.info(
        "Carga la bitácora y el log para comenzar. Si contiene Flujo 2.0, agrega "
        "también exported-request. El EXPORT de SAP se solicitará en la etapa 3."
    )
    st.stop()

step = int(st.session_state.get("audit_step", 1))
render_stepper(step)
try:
    if step == 1:
        audit_state_stage(context)
    elif step == 2:
        audit_unreported_stage(context)
    elif step == 3:
        audit_protocol_stage(context, export_file)
    else:
        audit_final_stage(context)
except BitacoraError as exc:
    st.error(str(exc))
except Exception as exc:
    st.exception(exc)
