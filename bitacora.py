from __future__ import annotations

import math
import re
import unicodedata
import zipfile
from html import unescape
from copy import copy
from dataclasses import dataclass, replace
from datetime import date, datetime, timedelta
from io import BytesIO
from pathlib import Path
from typing import Any, BinaryIO, Iterable

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo


SHEET_NEW = "Flujo Nuevo"
SHEET_OLD = "Flujo Viejo"
SHEET_FLOW20 = "Flujo 2.0"
SHEET_PNC = "PNC"
SHEET_RESULT = "Resultado Validacion"
SHEET_EMAIL_RESULT = "Validacion Correos"
SHEET_UNREPORTED = "Casos no reportados"
SHEET_GUIDED_AUDIT = "Auditoria guiada"
SHEET_PROTOCOL_EVIDENCE = "Soportes protocolos"
TEMPLATE_HEADER_ROW = 5
BUNDLED_TEMPLATE_PATH = (
    Path(__file__).resolve().parent
    / "assets"
    / "F-AGP-04 Formato Bitacora Seguimiento Gestion.xlsx"
)
LEGACY_TEMPLATE_PATH = Path(
    r"G:\.shortcut-targets-by-id\1CCgtZ-LFgl_wQdmd3LuIRjEtOZJYvRlI"
    r"\20_ENEL_AGP\AGP #2 - JA10161530\08_HSEQ\HSE 2026"
    r"\1. Gestión documental\1.2 Formatos de proyecto"
    r"\F-AGP-04 Formato Bitacora Seguimiento Gestión.xlsx"
)
FALLBACK_LOGO_PATH = (
    Path(__file__).resolve().parent
    / "tmp"
    / "inspection"
    / "logo_extracted.png"
)
CODED_HEADER_SHEETS = (
    SHEET_NEW,
    SHEET_OLD,
    SHEET_FLOW20,
    "Odm",
    "Correo-Rotulacion",
    SHEET_PNC,
    "Control de Cambios",
)
OPERATIONAL_SHEET_KEYS = {
    SHEET_NEW: "QR",
    SHEET_OLD: "QR",
    SHEET_FLOW20: "QR",
    "Odm": "QR",
    "Correo-Rotulacion": "ASUNTO",
    SHEET_PNC: "QR/ASUNTO",
}
LOGO_WIDTH_PX = 150
LOGO_HEIGHT_PX = 80
ARRIVAL_DATETIME_HEADER = "FECHA/HORA DE LLEGADA"
EXIT_DATETIME_HEADER = "FECHA/HORA DE SALIDA"
TIMING_HEADERS = (ARRIVAL_DATETIME_HEADER, EXIT_DATETIME_HEADER)
ODM_DATETIME_HEADER = "FECHA/HORA ENTREGA ODM"
DRAWING_DATETIME_HEADER = "FECHA/HORA ENTREGA DIBUJO"
FLOW20_TIMING_HEADERS = (
    ARRIVAL_DATETIME_HEADER,
    ODM_DATETIME_HEADER,
    DRAWING_DATETIME_HEADER,
)
EXCEL_DATETIME_FORMAT = "dd/mm/yyyy hh:mm:ss"

LEGACY_STATES = {
    "2-Valorizado",
    "2.5-Enviado a dibujo",
    "2.6-Dibujante asignado",
    "2.6.1-Dibujo en revision",
    "2.7-Dibujo finalizado",
}
NEW_STATES = {
    "drawing_request",
    "assigned_draftsman",
    "verify_formats",
    "drawing_review",
}
FLOW20_ODM_STATE = "3-ODM Asignadas"
FLOW20_STATES = {
    "2-Valorizado",
    "2.5-Enviado a dibujo",
    "2.6-Dibujante asignado",
    FLOW20_ODM_STATE,
}
FLOW20_SECONDARY_PENDING = "ODM_Creada-Pendiente_Dibujo"
FLOW20_SECONDARY_COMPLETED = "ODM_Creada-Dibujo_Finalizado"
FLOW20_SECONDARY_STATES = {
    FLOW20_SECONDARY_PENDING,
    FLOW20_SECONDARY_COMPLETED,
}
FLOW20_ODM_ENTRY_INITIAL_STATES = {
    "2-Valorizado",
    "2.6-Dibujante asignado",
}
ALL_DRAWING_STATES = LEGACY_STATES | NEW_STATES | {FLOW20_ODM_STATE}
FINAL_ELEMENT_STATES = {"2.7-Dibujo finalizado", "drawing_review"}
REJECTION_STATES = {"2-Valorizado", "2.6.1-Dibujo en revision", "verify_formats"}
LEGACY_ARRIVAL_STATES = {"2.5-Enviado a dibujo", "2.6-Dibujante asignado"}
NEW_ARRIVAL_STATES = {"drawing_request", "assigned_draftsman"}
LEGACY_OUTPUT_STATES = {
    "2-Valorizado",
    "2.6.1-Dibujo en revision",
    "2.7-Dibujo finalizado",
}
NEW_OUTPUT_STATES = {"verify_formats", "drawing_review"}
DRAWING_OUTPUT_STATES = LEGACY_OUTPUT_STATES | NEW_OUTPUT_STATES
UNREPORTED_TARGET_STATES = {
    "2.6.1-Dibujo en revision",
    "2.7-Dibujo finalizado",
    "verify_formats",
    "drawing_review",
}
CLOSED_UNREPORTED_STATES = {"2.7-Dibujo finalizado", "drawing_review"}
EMAIL_QUANTITY_ORIGIN = "SOLICITUD POR CORREO"
EMAIL_QUANTITY_STATE = "FINALIZADO"

UPDATE_ORIGINS = (
    "E1",
    "Proyectos",
    "Convenios: Concesiones",
    "Telecontrol",
    "Indisponibles",
    "Equipos Nuevos",
    "Conexiones",
    "Mantenimiento Programado",
    "PQRs",
    "Boletines Pagos",
    "C100",
    "SMEG",
    "IPPC",
    "Solicitud por correo",
    "Rotulación",
    "Otro",
    "Concesiones",
)

CLOSED_STATE_PROFESSIONALS = (
    ("4264437", "Adrian Ignacio Sanchez Vacca"),
    ("1233691570", "Daniela Forero Fierro"),
    ("1053845070", "David Andres Luna Yela"),
    ("1023014624", "Jenny Lorena Porras Cano"),
    ("1033708860", "Johan Estiven Silva Silva"),
    ("1022433323", "Jorge Luis Urrego Castañeda"),
    ("79170578", "Jose Pablo Siachoque Rodriguez"),
    ("1012442896", "Kevin Alejandro Torres Perdomo"),
    ("1000988576", "Luisa Fernanda Guatavita Perez"),
    ("1049796769", "Manuel Fernando Sanchez Mondragon"),
    ("1079391368", "Maria Alejandra Sierra Almario"),
    ("1108455462", "Maria Camila Leal Cardenas"),
    ("1077976852", "Stefany Colorado Fresneda"),
    ("1233691463", "Valery Steffany Cagua Marroquin"),
    ("1053829105", "William Fernando Uribe Bayona"),
    ("1019108406", "Yesika Arevalo Pinzon"),
    ("1030627620", "Yiceth Loreyna Ayala Guio"),
    ("1142514909", "Andres Felipe Perez Naranjo"),
)
CLOSED_STATE_PROFESSIONAL_NAME_ALIASES = (
    "Johan Stiven Silva Silva",
    "Maria Alejandra Sierra",
    "Maria Leal Cardenas",
    "Yesica Arevalo Pinzon",
)

STATUS_MATCH = "Coincide"
STATUS_MISMATCH = "No coincide"
STATUS_NO_EVENT = "Sin movimiento en la fecha"
STATUS_NO_QR = "Sin QR"
STATUS_PROFESSIONAL_MISMATCH = "Revisión obligatoria: profesional PMS diferente"
STATUS_FLOW20_AMBIGUOUS = "Revisión obligatoria: evidencia temporal insuficiente"
STATUS_FLOW20_DUPLICATE_ODM = "Duplicado histórico: ODM entregada en fecha anterior"
STATUS_FLOW20_DUPLICATE_DRAWING = "Duplicado histórico: dibujo entregado en fecha anterior"
FLOW20_DUPLICATE_RESULTS = {
    STATUS_FLOW20_DUPLICATE_ODM,
    STATUS_FLOW20_DUPLICATE_DRAWING,
}
STATUS_AUDIT_KEEP = "Aprobado: conservar bitácora"
STATUS_AUDIT_CORRECTED = "Corregido en auditoría"
STATUS_AUDIT_EXCLUDED = "Excluido en auditoría"
UPLOAD_APPROVED_RESULTS = {
    STATUS_MATCH,
    STATUS_AUDIT_KEEP,
    STATUS_AUDIT_CORRECTED,
}

COUNT_NOT_EVALUATED = "No evaluado"
COUNT_NOT_APPLICABLE = "No aplica por estado"
COUNT_AUDIT_NOT_APPLICABLE = "No aplica por auditoría"
COUNT_CONFIRMED = "Coincide con EXPORT"
COUNT_UPDATED = "Actualizada desde EXPORT"
COUNT_NO_PROTOCOL = "Sin protocolo"
COUNT_PROTOCOL_NOT_FOUND = "Protocolo no encontrado en EXPORT"
COUNT_PROTOCOLS_INCOMPLETE = "Protocolos incompletos en EXPORT"
COUNT_INVALID_PROTOCOL_SEPARATOR = "Separador inválido entre protocolos: usa ;"
COUNT_INVALID_PROTOCOL_PREFIX = "Protocolo inválido: debe iniciar por 1000"
PROTOCOL_REQUIRED_PREFIX = "1000"

OBSERVATION_UPDATED = "Actualizada desde PMS"
OBSERVATION_CLEANED = "Marcador eliminado"
OBSERVATION_NOT_APPLICABLE = "No aplica"
OBSERVATION_NOT_FOUND = "Sin devolución diligenciada en PMS"
OBSERVATION_STATE_MISMATCH = "No aplicada: el estado no coincide"

TIMING_CAPTURED = "Ciclo capturado desde PMS"
TIMING_NO_ARRIVAL = "Salida encontrada; llegada no localizada"
TIMING_NO_EXIT = "Salida del profesional no localizada"
TIMING_NOT_APPLICABLE = "No aplica: estado sin salida"
TIMING_MANUAL = "Fechas confirmadas en auditoría"
TIMING_FLOW20_PENDING = "ODM capturada; dibujo pendiente"
TIMING_FLOW20_COMPLETED = "ODM y dibujo capturados"
TIMING_FLOW20_AMBIGUOUS = "ODM capturada; fecha de dibujo no demostrable"

LOG_REQUIRED_COLUMNS = {"request_code", "final_state", "date_created"}
LOG_OPTIONAL_COLUMNS = {
    "initial_state",
    "comments",
    "username",
    "user_created_last_name",
}
REQUEST_REQUIRED_COLUMNS = {"request_code", "contract_drawing_completed"}
REQUEST_OPTIONAL_COLUMNS = {
    "status",
    "new_fulfillment_process",
    "date_created",
    "odm",
    "agp_project",
    "agp_increment",
    "protocol",
}
EXPORT_REQUIRED_COLUMNS = {
    "PROTOCOLLO AP",
    "GIS GUID",
    "TIPO PEL O TRATT",
}
REMOVED_UPLOAD_COLUMNS = {
    "FECHA DE LLEGADA",
    "FECHA ENTREGADO",
    "CONTROL ANS",
    "VALIDA",
    "CONFIRMA",
}


class BitacoraError(ValueError):
    """Error de entrada que se puede mostrar directamente al usuario."""


@dataclass(frozen=True)
class LogState:
    state: str
    occurred_at: datetime
    comment: str = ""
    actor: str = ""


@dataclass(frozen=True)
class DrawingLogEvent:
    qr: str
    initial_state: str
    state: str
    occurred_at: datetime
    comment: str = ""
    username: str = ""
    user_name: str = ""

    @property
    def actor(self) -> str:
        return format_pms_actor(self.username, self.user_name)


@dataclass(frozen=True)
class LogCoverage:
    first_datetime: datetime | None
    last_datetime: datetime | None

    @property
    def first_date(self) -> date | None:
        return self.first_datetime.date() if self.first_datetime else None

    @property
    def last_date(self) -> date | None:
        return self.last_datetime.date() if self.last_datetime else None


@dataclass(frozen=True)
class RequestSnapshot:
    qr: str
    drawing_completed: bool | None
    status: str = ""
    new_fulfillment_process: bool | None = None
    captured_at: datetime | None = None
    odm: str = ""
    agp: str = ""
    protocol: str = ""


@dataclass(frozen=True)
class UnreportedCase:
    qr: str
    state: str
    occurred_at: datetime
    comment: str = ""
    actor: str = ""

    @property
    def flow(self) -> str:
        return SHEET_OLD if self.state in LEGACY_STATES else SHEET_NEW

    def as_dict(self) -> dict[str, Any]:
        return {
            "Flujo": self.flow,
            "QR": self.qr,
            "Estado PMS del día": self.state,
            "Fecha/hora PMS": self.occurred_at,
            "Usuario PMS": self.actor,
            "Observación PMS": self.comment,
            "Hallazgo": "QR con movimiento objetivo que no aparece en la bitácora",
        }


@dataclass(frozen=True)
class ProtocolElementCounts:
    pel_ids: frozenset[str]
    tratta_ids: frozenset[str]

    @property
    def pel(self) -> int:
        return len(self.pel_ids)

    @property
    def tratas(self) -> int:
        return len(self.tratta_ids)


@dataclass(frozen=True)
class EmailQuantityResult:
    excel_row: int
    origin: str
    subject: str
    protocol: str
    reported_pel: Any = None
    reported_tratas: Any = None
    export_pel: int | None = None
    export_tratas: int | None = None
    export_pel_breakdown: Any = None
    export_tratas_breakdown: Any = None
    count_result: str = COUNT_NOT_EVALUATED

    def as_dict(self) -> dict[str, Any]:
        return {
            "Hoja": "Correo-Rotulacion",
            "Fila Excel": self.excel_row,
            "QR": self.subject,
            "Estado reportado": self.origin,
            "Protocolo": self.protocol,
            "Cantidad PEL reportada": self.reported_pel,
            "Cantidad TRATAS reportada": self.reported_tratas,
            "Cantidad PEL EXPORT": self.export_pel,
            "Cantidad TRATAS EXPORT": self.export_tratas,
            "Resultado cantidades": self.count_result,
        }


@dataclass(frozen=True)
class ValidationRow:
    sheet: str
    excel_row: int
    qr: str
    reported_state: str
    pms_state: str
    pms_datetime: datetime | None
    previous_state: str
    previous_datetime: datetime | None
    result: str
    professional: str = ""
    pms_actor: str = ""
    reported_secondary_state: str = ""
    agp: str = ""
    odm: str = ""
    protocol: str = ""
    reported_pel: Any = None
    reported_tratas: Any = None
    export_pel: int | None = None
    export_tratas: int | None = None
    export_pel_breakdown: Any = None
    export_tratas_breakdown: Any = None
    count_result: str = COUNT_NOT_EVALUATED
    pms_observation: str = ""
    observation_result: str = OBSERVATION_NOT_APPLICABLE
    arrival_state: str = ""
    arrival_datetime: datetime | None = None
    exit_datetime: datetime | None = None
    timing_result: str = TIMING_NOT_APPLICABLE
    timing_source: str = ""
    request_drawing_completed: bool | None = None
    request_status: str = ""
    odm_datetime: datetime | None = None
    drawing_datetime: datetime | None = None
    flow20_evidence: str = ""

    def as_dict(self) -> dict[str, Any]:
        return {
            "Hoja": self.sheet,
            "Fila Excel": self.excel_row,
            "QR": self.qr,
            "Estado reportado": self.reported_state,
            "Último estado PMS del día": self.pms_state,
            "Fecha/hora PMS": self.pms_datetime,
            "Último estado anterior o igual": self.previous_state,
            "Fecha/hora anterior": self.previous_datetime,
            "Resultado": self.result,
            "Profesional bitácora": self.professional,
            "Profesional PMS": self.pms_actor,
            "Estado secundario reportado": self.reported_secondary_state,
            "Estado de llegada": self.arrival_state,
            "Fecha/hora de llegada": self.arrival_datetime,
            "Fecha/hora de salida": self.exit_datetime,
            "Fecha/hora entrega ODM": self.odm_datetime,
            "Fecha/hora entrega dibujo": self.drawing_datetime,
            "Resultado tiempos": self.timing_result,
            "Fuente tiempos": self.timing_source,
            "Drawing Completed request": self.request_drawing_completed,
            "Estado actual request": self.request_status,
            "Evidencia Flujo 2.0": self.flow20_evidence,
            "AGP": self.agp,
            "ODM": self.odm,
            "Protocolo": self.protocol,
            "Cantidad PEL reportada": self.reported_pel,
            "Cantidad TRATAS reportada": self.reported_tratas,
            "Cantidad PEL EXPORT": self.export_pel,
            "Cantidad TRATAS EXPORT": self.export_tratas,
            "Resultado cantidades": self.count_result,
            "Observación PMS": self.pms_observation,
            "Acción observación": self.observation_result,
        }


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    text = str(value).strip()
    return "" if text.lower() == "nan" else text


def normalize_header(value: Any) -> str:
    text = unicodedata.normalize("NFKD", normalize_text(value))
    text = "".join(char for char in text if not unicodedata.combining(char))
    return " ".join(text.upper().replace("_", " ").split())


def normalize_identifier(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if math.isnan(value):
            return ""
        if value.is_integer():
            return str(int(value))
    text = normalize_text(value)
    if re.fullmatch(r"\d+\.0+", text):
        return text.split(".", 1)[0]
    return text


def parse_optional_bool(value: Any) -> bool | None:
    key = normalize_text(value).strip().lower()
    if key in {"true", "1", "yes", "si", "sí"}:
        return True
    if key in {"false", "0", "no"}:
        return False
    return None


def _pms_name_tokens(value: Any) -> tuple[str, ...]:
    return tuple(re.findall(r"[A-Z]+", normalize_header(value)))


def _pms_name_signature(value: Any) -> tuple[str, ...]:
    return tuple(sorted(_pms_name_tokens(value)))


def _professional_name_signatures(names: Iterable[str]) -> frozenset[tuple[str, ...]]:
    signatures: set[tuple[str, ...]] = set()
    for name in names:
        tokens = _pms_name_tokens(name)
        if not tokens:
            continue
        signatures.add(tuple(sorted(tokens)))
        if len(tokens) >= 4:
            signatures.add(tuple(sorted((tokens[0], tokens[-2], tokens[-1]))))
    return frozenset(signatures)


_CLOSED_STATE_PROFESSIONAL_IDS = frozenset(
    identifier for identifier, _ in CLOSED_STATE_PROFESSIONALS
)
_CLOSED_STATE_PROFESSIONAL_SIGNATURES = _professional_name_signatures(
    [name for _, name in CLOSED_STATE_PROFESSIONALS]
    + list(CLOSED_STATE_PROFESSIONAL_NAME_ALIASES)
)


def _build_professional_identity_index() -> dict[tuple[str, ...], str]:
    index: dict[tuple[str, ...], str] = {}
    for identifier, name in CLOSED_STATE_PROFESSIONALS:
        for signature in _professional_name_signatures([name]):
            index[signature] = identifier
    alias_targets = {
        "Johan Stiven Silva Silva": "1033708860",
        "Maria Alejandra Sierra": "1079391368",
        "Maria Leal Cardenas": "1108455462",
        "Yesica Arevalo Pinzon": "1019108406",
    }
    for alias, identifier in alias_targets.items():
        for signature in _professional_name_signatures([alias]):
            index[signature] = identifier
    return index


_PROFESSIONAL_ID_BY_SIGNATURE = _build_professional_identity_index()


def professional_identity(value: Any) -> str:
    """Resolve a PMS/user display value to the canonical professional ID."""
    normalized = normalize_header(value)
    identifier_match = re.search(r"(?:^|[^A-Z0-9])(?:CO)?(\d{6,})(?:$|[^A-Z0-9])", normalized)
    if identifier_match and identifier_match.group(1) in _CLOSED_STATE_PROFESSIONAL_IDS:
        return identifier_match.group(1)
    signature = _pms_name_signature(value)
    return _PROFESSIONAL_ID_BY_SIGNATURE.get(signature, "") if signature else ""


def same_professional(workbook_name: Any, *pms_values: Any) -> bool:
    """Compare any workbook/PMS professional, keeping the whitelist separate."""
    workbook_identity = professional_identity(workbook_name)
    workbook_signature = _pms_name_signature(workbook_name)
    workbook_tokens = frozenset(_pms_name_tokens(workbook_name))
    if not workbook_identity and not workbook_signature:
        return False
    for value in pms_values:
        if not normalize_text(value):
            continue
        pms_identity = professional_identity(value)
        if workbook_identity and pms_identity:
            if pms_identity == workbook_identity:
                return True
            continue
        pms_signature = _pms_name_signature(value)
        if workbook_signature and pms_signature == workbook_signature:
            return True
        pms_tokens = frozenset(_pms_name_tokens(value))
        if (
            len(workbook_tokens) >= 3
            and len(pms_tokens) >= 3
            and (
                workbook_tokens.issubset(pms_tokens)
                or pms_tokens.issubset(workbook_tokens)
            )
        ):
            return True
    return False


def is_closed_state_professional(*values: Any) -> bool:
    """Match PMS users by Asignador ID or by a name independent of order/accents."""
    for value in values:
        if professional_identity(value):
            return True
    return False


def format_pms_actor(username: Any, user_name: Any) -> str:
    username_text = normalize_text(username)
    name_text = normalize_text(user_name)
    if name_text and username_text and normalize_header(name_text) != normalize_header(
        username_text
    ):
        return f"{name_text} ({username_text})"
    return name_text or username_text


def clean_placeholder_observation(value: Any) -> str:
    text = normalize_text(value)
    normalized = unicodedata.normalize("NFKD", text)
    normalized = "".join(char for char in normalized if not unicodedata.combining(char))
    key = re.sub(r"[^A-Z0-9]+", "", normalized.upper())
    if key in {"NA", "NOAPLICA", "SM"}:
        return ""
    return text


def clean_operational_noise(value: Any) -> Any:
    """Remove standalone upload placeholders without touching formulas or real text."""
    if isinstance(value, str) and value.startswith("="):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool) and value == 0:
        return None

    text = normalize_text(value)
    normalized = unicodedata.normalize("NFKD", text)
    normalized = "".join(char for char in normalized if not unicodedata.combining(char))
    key = re.sub(r"[^A-Z0-9]+", "", normalized.upper())
    if key in {"NA", "SN", "0"}:
        return None
    return value


def clean_pms_comment(value: Any) -> str:
    text = normalize_text(value)
    if not text:
        return ""
    text = re.sub(r"(?i)<br\s*/?>|</p\s*>|</div\s*>|</li\s*>", "\n", text)
    text = re.sub(r"<[^>]+>", "", text)
    text = unescape(text).replace("\xa0", " ")
    lines = [" ".join(line.split()) for line in text.splitlines()]
    cleaned = "\n".join(line for line in lines if line)
    return clean_placeholder_observation(cleaned)


def _numeric_count(value: Any) -> int | None:
    text = normalize_text(value)
    if not text:
        return None
    try:
        number = float(text.replace(",", "."))
    except ValueError:
        return None
    return int(number) if number.is_integer() else None


def split_protocols(value: Any) -> list[str]:
    text = normalize_identifier(value)
    if not text:
        return []
    parts = [normalize_identifier(item) for item in text.split(";")]
    return list(dict.fromkeys(item for item in parts if item))


def has_invalid_protocol_separator(value: Any) -> bool:
    """Detect multiple numeric protocols that were not separated with semicolons."""
    text = normalize_identifier(value)
    if not text:
        return False
    if re.search(r"[,|\r\n]", text):
        return True
    numeric_protocols = list(
        dict.fromkeys(re.findall(r"(?<!\d)\d{7,}(?!\d)", text))
    )
    concatenated_numeric = any(
        len(match) > 12 for match in re.findall(r"(?<!\d)\d+(?!\d)", text)
    )
    return concatenated_numeric or (
        len(numeric_protocols) > 1 and ";" not in text
    )


def invalid_protocol_prefixes(value: Any) -> list[str]:
    """Return protocol tokens that do not use the required PMS/SAP prefix."""
    return [
        protocol
        for protocol in split_protocols(value)
        if not protocol.startswith(PROTOCOL_REQUIRED_PREFIX)
    ]


def has_invalid_protocol_prefix(value: Any) -> bool:
    return bool(invalid_protocol_prefixes(value))


def has_invalid_agp_separator(value: Any) -> bool:
    """Detect several AGP identifiers written without the required semicolon."""
    text = normalize_identifier(value)
    if not text:
        return False
    if re.search(r"[,|\r\n]", text):
        return True
    if ";" in text:
        return False
    numeric_agps = list(dict.fromkeys(re.findall(r"(?<!\d)\d{7,}(?!\d)", text)))
    labelled_agps = re.findall(r"(?i)\bAGP(?:[-_ ]?[A-Z0-9]+)", text)
    return len(numeric_agps) > 1 or len(labelled_agps) > 1


def canonicalize_state(value: Any) -> str:
    raw = normalize_text(value)
    normalized = unicodedata.normalize("NFKD", raw)
    normalized = "".join(char for char in normalized if not unicodedata.combining(char))
    key = " ".join(normalized.replace("_", " ").split()).lower()
    aliases = {
        "2": "2-Valorizado",
        "2 valorizado": "2-Valorizado",
        "2-valorizado": "2-Valorizado",
        "2.5": "2.5-Enviado a dibujo",
        "2.5 enviado a dibujo": "2.5-Enviado a dibujo",
        "2.5-enviado a dibujo": "2.5-Enviado a dibujo",
        "2.6": "2.6-Dibujante asignado",
        "2.6 dibujante asignado": "2.6-Dibujante asignado",
        "2.6-dibujante asignado": "2.6-Dibujante asignado",
        "2.6.1": "2.6.1-Dibujo en revision",
        "2.6.1 dibujo en revision": "2.6.1-Dibujo en revision",
        "2.6.1-dibujo en revision": "2.6.1-Dibujo en revision",
        "2.7": "2.7-Dibujo finalizado",
        "2.7 dibujo finalizado": "2.7-Dibujo finalizado",
        "2.7-dibujo finalizado": "2.7-Dibujo finalizado",
        "drawing request": "drawing_request",
        "assigned draftsman": "assigned_draftsman",
        "verify formats": "verify_formats",
        "drawing review": "drawing_review",
        "3": FLOW20_ODM_STATE,
        "3 odm asignadas": FLOW20_ODM_STATE,
        "3-odm asignadas": FLOW20_ODM_STATE,
    }
    return aliases.get(key, raw)


def canonicalize_flow20_secondary(value: Any) -> str:
    normalized = normalize_header(value)
    compact = re.sub(r"[^A-Z0-9]+", "", normalized)
    aliases = {
        "ODMCREADAPENDIENTEDIBUJO": FLOW20_SECONDARY_PENDING,
        "ODMCREADADIBUJOFINALIZADO": FLOW20_SECONDARY_COMPLETED,
    }
    return aliases.get(compact, normalize_text(value))


def validation_requires_element_counts(item: ValidationRow) -> bool:
    return item.reported_state in FINAL_ELEMENT_STATES or (
        item.sheet == SHEET_FLOW20
        and item.reported_state == FLOW20_ODM_STATE
        and item.reported_secondary_state == FLOW20_SECONDARY_COMPLETED
    )


def parse_production_date(filename: str) -> date:
    matches = re.findall(r"(?<!\d)(\d{2})[-_/](\d{2})[-_/](\d{4})(?!\d)", filename)
    if not matches:
        raise BitacoraError(
            "No encontré una fecha DD-MM-AAAA en el nombre de la bitácora. "
            "Ejemplo esperado: Produccion 18-06-2026.xlsx."
        )
    day, month, year = matches[-1]
    try:
        return date(int(year), int(month), int(day))
    except ValueError as exc:
        raise BitacoraError(f"La fecha del nombre no es válida: {day}-{month}-{year}.") from exc


def _rewind(source: str | Path | BinaryIO | bytes) -> str | Path | BinaryIO:
    if isinstance(source, bytes):
        return BytesIO(source)
    if hasattr(source, "seek"):
        source.seek(0)
    return source


def _header_map(ws, header_row: int = 1) -> dict[str, int]:
    headers: dict[str, int] = {}
    for cell in ws[header_row]:
        normalized = normalize_header(cell.value)
        if normalized:
            headers[normalized] = cell.column
    return headers


def _detect_header_row(
    ws,
    required_headers: set[str],
    *,
    max_scan_rows: int = 20,
) -> int:
    required = {normalize_header(item) for item in required_headers}
    for row_number in range(1, min(ws.max_row, max_scan_rows) + 1):
        headers = _header_map(ws, header_row=row_number)
        if required.issubset(headers):
            return row_number
    expected = ", ".join(sorted(required))
    raise BitacoraError(
        f"No encontré una fila de encabezados con {expected} en la hoja '{ws.title}'."
    )


def extract_workbook_rows(
    workbook_source: str | Path | BinaryIO | bytes,
) -> tuple[bytes, list[dict[str, Any]]]:
    source = _rewind(workbook_source)
    if isinstance(source, (str, Path)):
        workbook_bytes = Path(source).read_bytes()
    else:
        workbook_bytes = source.read()

    try:
        wb = load_workbook(BytesIO(workbook_bytes), read_only=True, data_only=False)
    except Exception as exc:  # openpyxl emits several format-specific exceptions
        raise BitacoraError("No pude abrir la bitácora. Verifica que sea un archivo .xlsx válido.") from exc

    missing_sheets = [name for name in (SHEET_NEW, SHEET_OLD) if name not in wb.sheetnames]
    if missing_sheets:
        raise BitacoraError(f"Faltan las hojas requeridas: {', '.join(missing_sheets)}.")

    rows: list[dict[str, Any]] = []
    validation_sheets = [SHEET_NEW, SHEET_OLD]
    if SHEET_FLOW20 in wb.sheetnames:
        validation_sheets.append(SHEET_FLOW20)
    for sheet_name in validation_sheets:
        ws = wb[sheet_name]
        header_row = _detect_header_row(ws, {"QR", "ESTADO DE LA OT"})
        headers = _header_map(ws, header_row=header_row)
        missing = [name for name in ("QR", "ESTADO DE LA OT") if name not in headers]
        if missing:
            raise BitacoraError(
                f"En la hoja '{sheet_name}' faltan columnas: {', '.join(missing)}."
            )
        qr_col = headers["QR"]
        state_col = headers["ESTADO DE LA OT"]
        agp_col = headers.get("CODIGO DEL INCREMENTO (AGP)")
        protocol_col = headers.get("PROTOCOLO")
        pel_col = headers.get("CANTIDAD PEL")
        tratas_col = headers.get("CANTIDAD TRATAS")
        observation_col = headers.get("OBSERVACIONES")
        user_col = headers.get("USUARIO")
        secondary_col = headers.get("ESTADO SECUNDARIO")
        odm_col = headers.get("ODM")
        delivered_col = headers.get("FECHA ENTREGADO")
        for row_number in range(header_row + 1, ws.max_row + 1):
            qr = normalize_text(ws.cell(row_number, qr_col).value)
            state = canonicalize_state(ws.cell(row_number, state_col).value)
            if not qr and not state:
                continue
            rows.append(
                {
                    "sheet": sheet_name,
                    "excel_row": row_number,
                    "qr": qr,
                    "reported_state": state,
                    "reported_secondary_state": (
                        canonicalize_flow20_secondary(
                            ws.cell(row_number, secondary_col).value
                        )
                        if secondary_col
                        else ""
                    ),
                    "professional": (
                        normalize_text(ws.cell(row_number, user_col).value)
                        if user_col
                        else ""
                    ),
                    "agp": (
                        normalize_identifier(ws.cell(row_number, agp_col).value)
                        if agp_col
                        else ""
                    ),
                    "protocol": (
                        normalize_identifier(ws.cell(row_number, protocol_col).value)
                        if protocol_col
                        else ""
                    ),
                    "odm": (
                        normalize_identifier(ws.cell(row_number, odm_col).value)
                        if odm_col
                        else ""
                    ),
                    "reported_delivery_date": (
                        ws.cell(row_number, delivered_col).value
                        if delivered_col
                        else None
                    ),
                    "reported_pel": ws.cell(row_number, pel_col).value if pel_col else None,
                    "reported_tratas": (
                        ws.cell(row_number, tratas_col).value if tratas_col else None
                    ),
                    "reported_observation": (
                        normalize_text(ws.cell(row_number, observation_col).value)
                        if observation_col
                        else ""
                    ),
                }
            )
    wb.close()
    return workbook_bytes, rows


def apply_qr_alias_corrections(
    workbook_bytes: bytes,
    rows: list[dict[str, Any]],
    aliases: dict[str, str],
) -> tuple[bytes, list[dict[str, Any]], list[dict[str, Any]]]:
    """Replace safely inferred truncated QR values and record the correction."""
    if not aliases:
        return workbook_bytes, rows, []
    try:
        wb = load_workbook(BytesIO(workbook_bytes), data_only=False)
    except Exception as exc:
        raise BitacoraError("No pude corregir los QR truncados en la bitácora.") from exc

    sheet_qr_columns: dict[str, int] = {}
    updated_rows: list[dict[str, Any]] = []
    audit_entries: list[dict[str, Any]] = []
    try:
        for row in rows:
            original = normalize_text(row.get("qr"))
            corrected = aliases.get(original)
            if not corrected:
                updated_rows.append(row)
                continue
            sheet_name = row["sheet"]
            if sheet_name not in sheet_qr_columns:
                ws = wb[sheet_name]
                header_row = _detect_header_row(ws, {"QR", "ESTADO DE LA OT"})
                sheet_qr_columns[sheet_name] = _header_map(ws, header_row=header_row)["QR"]
            wb[sheet_name].cell(
                row["excel_row"], sheet_qr_columns[sheet_name]
            ).value = corrected
            corrected_row = dict(row)
            corrected_row["qr"] = corrected
            updated_rows.append(corrected_row)
            audit_entries.append(
                {
                    "Etapa": "Prevalidación QR",
                    "Hoja/Flujo": sheet_name,
                    "QR/Referencia": f"{original} → {corrected}",
                    "Decisión": "Corregir QR automáticamente",
                    "Estado inicial": original,
                    "Estado final": corrected,
                    "AGP": normalize_identifier(row.get("agp")),
                    "Protocolo": normalize_identifier(row.get("protocol")),
                    "Detalle": (
                        "El QR de la bitácora estaba truncado en un carácter final. "
                        "Se encontró una única referencia compatible en PMS para la fecha de corte."
                    ),
                }
            )
        output = BytesIO()
        wb.save(output)
        return output.getvalue(), updated_rows, audit_entries
    finally:
        wb.close()


def read_request_snapshots(
    request_source: str | Path | BinaryIO | bytes,
    qrs: set[str] | None = None,
    *,
    chunksize: int = 200_000,
) -> dict[str, RequestSnapshot]:
    """Read the current PMS request snapshot used by Flujo 2.0."""
    source = _rewind(request_source)
    targets = (
        {normalize_text(qr) for qr in qrs if normalize_text(qr)}
        if qrs is not None
        else None
    )
    snapshots: dict[str, RequestSnapshot] = {}
    try:
        chunks = pd.read_csv(
            source,
            sep=";",
            usecols=lambda column: column
            in REQUEST_REQUIRED_COLUMNS | REQUEST_OPTIONAL_COLUMNS,
            dtype=str,
            encoding="utf-8-sig",
            chunksize=chunksize,
            on_bad_lines="skip",
            low_memory=False,
        )
        saw_required_columns = False
        for chunk in chunks:
            if not REQUEST_REQUIRED_COLUMNS.issubset(chunk.columns):
                continue
            saw_required_columns = True
            for column in REQUEST_OPTIONAL_COLUMNS:
                if column not in chunk.columns:
                    chunk[column] = ""
            chunk = chunk.copy()
            chunk["normalized_qr"] = chunk["request_code"].map(normalize_text)
            if targets is not None:
                chunk = chunk[chunk["normalized_qr"].isin(targets)]
            if chunk.empty:
                continue
            chunk["parsed_date"] = pd.to_datetime(
                chunk["date_created"], errors="coerce", format="mixed"
            )
            for record in chunk.itertuples(index=False):
                qr = record.normalized_qr
                captured_at = (
                    record.parsed_date.to_pydatetime()
                    if not pd.isna(record.parsed_date)
                    else None
                )
                snapshot = RequestSnapshot(
                    qr=qr,
                    drawing_completed=parse_optional_bool(
                        record.contract_drawing_completed
                    ),
                    status=canonicalize_state(record.status),
                    new_fulfillment_process=parse_optional_bool(
                        record.new_fulfillment_process
                    ),
                    captured_at=captured_at,
                    odm=normalize_identifier(record.odm),
                    agp=(
                        normalize_identifier(record.agp_project)
                        or normalize_identifier(record.agp_increment)
                    ),
                    protocol=normalize_identifier(record.protocol),
                )
                previous = snapshots.get(qr)
                if (
                    previous is None
                    or previous.captured_at is None
                    or (
                        captured_at is not None
                        and captured_at >= previous.captured_at
                    )
                ):
                    snapshots[qr] = snapshot
        if not saw_required_columns:
            raise BitacoraError(
                "El exported-request no contiene request_code y "
                "contract_drawing_completed."
            )
        return snapshots
    except BitacoraError:
        raise
    except ValueError as exc:
        raise BitacoraError(
            "El exported-request no contiene request_code y "
            "contract_drawing_completed."
        ) from exc
    except Exception as exc:
        raise BitacoraError("No pude leer el archivo exported-request.csv.") from exc


def read_export_protocol_counts(
    export_source: str | Path | BinaryIO | bytes,
) -> dict[str, ProtocolElementCounts]:
    source = _rewind(export_source)
    try:
        wb = load_workbook(source, read_only=True, data_only=True)
    except Exception as exc:
        raise BitacoraError("No pude abrir el archivo EXPORT. Verifica que sea un .xlsx válido.") from exc

    try:
        sheet_name = next(
            (name for name in wb.sheetnames if normalize_header(name) == "SHEET1"),
            wb.sheetnames[0] if wb.sheetnames else None,
        )
        if sheet_name is None:
            raise BitacoraError("El archivo EXPORT no contiene hojas.")
        ws = wb[sheet_name]
        header_row = _detect_header_row(ws, EXPORT_REQUIRED_COLUMNS, max_scan_rows=10)
        headers = _header_map(ws, header_row=header_row)
        protocol_col = headers["PROTOCOLLO AP"]
        guid_col = headers["GIS GUID"]
        type_col = headers["TIPO PEL O TRATT"]

        grouped: dict[str, dict[str, set[str]]] = {}
        for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
            protocol = normalize_identifier(row[protocol_col - 1])
            element_type = normalize_text(row[type_col - 1]).upper()
            guid = normalize_identifier(row[guid_col - 1])
            if not protocol or element_type not in {"P", "T"}:
                continue
            item = grouped.setdefault(protocol, {"P": set(), "T": set()})
            if guid:
                item[element_type].add(guid)

        if not grouped:
            raise BitacoraError(
                "La hoja Sheet1 del EXPORT no contiene protocolos con tipos P o T."
            )
        return {
            protocol: ProtocolElementCounts(
                pel_ids=frozenset(values["P"]),
                tratta_ids=frozenset(values["T"]),
            )
            for protocol, values in grouped.items()
        }
    finally:
        wb.close()


def resolve_protocol_counts(
    protocol_value: Any,
    protocol_counts: dict[str, ProtocolElementCounts],
) -> tuple[int | None, int | None, list[str], list[str]]:
    protocols = split_protocols(protocol_value)
    if not protocols:
        return None, None, [], []
    missing = [protocol for protocol in protocols if protocol not in protocol_counts]
    found = [protocol for protocol in protocols if protocol in protocol_counts]
    if not found:
        return None, None, protocols, missing
    pel = sum(protocol_counts[protocol].pel for protocol in found)
    tratas = sum(protocol_counts[protocol].tratas for protocol in found)
    return pel, tratas, protocols, missing


def protocol_quantity_breakdown(
    protocol_value: Any,
    protocol_counts: dict[str, ProtocolElementCounts],
) -> tuple[Any, Any]:
    """Return PEL/TRATA counts aligned positionally with the protocol list."""
    protocols = split_protocols(protocol_value)
    if not protocols:
        return None, None
    pel_values = [
        protocol_counts[protocol].pel if protocol in protocol_counts else None
        for protocol in protocols
    ]
    tratta_values = [
        protocol_counts[protocol].tratas if protocol in protocol_counts else None
        for protocol in protocols
    ]

    def display(values: list[int | None]) -> Any:
        if len(values) == 1:
            return values[0]
        return ";".join("?" if value is None else str(value) for value in values)

    return display(pel_values), display(tratta_values)


def protocol_counts_with_audit_decisions(
    protocol_counts: dict[str, ProtocolElementCounts],
    decisions: dict[str, dict[str, Any]],
) -> tuple[dict[str, ProtocolElementCounts], set[str]]:
    """Resolve missing protocols with audited manual counts or explicit No aplica."""
    resolved = dict(protocol_counts)
    no_apply: set[str] = set()
    for protocol, decision in decisions.items():
        action = normalize_text(decision.get("action")).lower()
        if action not in {"manual", "no_apply"}:
            continue
        if action == "no_apply":
            pel = tratas = 0
            no_apply.add(protocol)
        else:
            pel = max(int(decision.get("pel") or 0), 0)
            tratas = max(int(decision.get("tratas") or 0), 0)
        resolved[protocol] = ProtocolElementCounts(
            pel_ids=frozenset(f"AUDIT:{protocol}:P:{index}" for index in range(pel)),
            tratta_ids=frozenset(
                f"AUDIT:{protocol}:T:{index}" for index in range(tratas)
            ),
        )
    return resolved, no_apply


def compare_protocol_quantities(
    protocol_value: Any,
    reported_pel: Any,
    reported_tratas: Any,
    protocol_counts: dict[str, ProtocolElementCounts],
) -> tuple[int | None, int | None, str]:
    protocol = normalize_identifier(protocol_value)
    if not split_protocols(protocol):
        return None, None, COUNT_NO_PROTOCOL
    if has_invalid_protocol_separator(protocol):
        return None, None, COUNT_INVALID_PROTOCOL_SEPARATOR
    if has_invalid_protocol_prefix(protocol):
        return None, None, COUNT_INVALID_PROTOCOL_PREFIX

    export_pel, export_tratas, protocols, missing = resolve_protocol_counts(
        protocol,
        protocol_counts,
    )
    if missing:
        result = (
            COUNT_PROTOCOL_NOT_FOUND
            if len(missing) == len(protocols)
            else f"{COUNT_PROTOCOLS_INCOMPLETE}: {', '.join(missing)}"
        )
        return export_pel, export_tratas, result
    if (
        _numeric_count(reported_pel) == export_pel
        and _numeric_count(reported_tratas) == export_tratas
    ):
        return export_pel, export_tratas, COUNT_CONFIRMED
    return export_pel, export_tratas, COUNT_UPDATED


def evaluate_email_quantities(
    workbook_source: str | Path | BinaryIO | bytes,
    protocol_counts: dict[str, ProtocolElementCounts],
) -> list[EmailQuantityResult]:
    source = _rewind(workbook_source)
    if isinstance(source, (str, Path)):
        workbook_bytes = Path(source).read_bytes()
    else:
        workbook_bytes = source.read()

    try:
        wb = load_workbook(BytesIO(workbook_bytes), read_only=True, data_only=False)
    except Exception as exc:
        raise BitacoraError("No pude leer la hoja Correo-Rotulacion.") from exc

    try:
        if "Correo-Rotulacion" not in wb.sheetnames:
            return []
        ws = wb["Correo-Rotulacion"]
        header_row = _detect_header_row(
            ws,
            {"ORIGEN DE ACTUALIZACIÓN", "ASUNTO", "ESTADO DE LA OT"},
        )
        headers = _header_map(ws, header_row=header_row)
        origin_col = headers["ORIGEN DE ACTUALIZACION"]
        subject_col = headers["ASUNTO"]
        state_col = headers["ESTADO DE LA OT"]
        protocol_col = headers.get("PROTOCOLO")
        pel_col = headers.get("CANTIDAD PEL")
        tratas_col = headers.get("CANTIDAD TRATAS")
        results: list[EmailQuantityResult] = []

        for row_number in range(header_row + 1, ws.max_row + 1):
            origin = normalize_text(ws.cell(row_number, origin_col).value)
            if normalize_header(origin) != EMAIL_QUANTITY_ORIGIN:
                continue
            state = normalize_text(ws.cell(row_number, state_col).value)
            if normalize_header(state) != EMAIL_QUANTITY_STATE:
                continue
            subject = normalize_text(ws.cell(row_number, subject_col).value)
            protocol = (
                normalize_identifier(ws.cell(row_number, protocol_col).value)
                if protocol_col
                else ""
            )
            reported_pel = ws.cell(row_number, pel_col).value if pel_col else None
            reported_tratas = (
                ws.cell(row_number, tratas_col).value if tratas_col else None
            )
            export_pel, export_tratas, count_result = compare_protocol_quantities(
                protocol,
                reported_pel,
                reported_tratas,
                protocol_counts,
            )
            export_pel_breakdown, export_tratas_breakdown = protocol_quantity_breakdown(
                protocol,
                protocol_counts,
            )
            results.append(
                EmailQuantityResult(
                    excel_row=row_number,
                    origin=origin,
                    subject=subject,
                    protocol=protocol,
                    reported_pel=reported_pel,
                    reported_tratas=reported_tratas,
                    export_pel=export_pel,
                    export_tratas=export_tratas,
                    export_pel_breakdown=export_pel_breakdown,
                    export_tratas_breakdown=export_tratas_breakdown,
                    count_result=count_result,
                )
            )
        return results
    finally:
        wb.close()


def _read_log_chunks(
    log_source: str | Path | BinaryIO | bytes,
    *,
    chunksize: int = 250_000,
) -> Iterable[pd.DataFrame]:
    source = _rewind(log_source)
    try:
        yield from pd.read_csv(
            source,
            sep=";",
            usecols=lambda column: column in LOG_REQUIRED_COLUMNS | LOG_OPTIONAL_COLUMNS,
            dtype=str,
            encoding="utf-8-sig",
            chunksize=chunksize,
            on_bad_lines="skip",
            low_memory=False,
        )
    except ValueError as exc:
        raise BitacoraError(
            "El CSV no contiene las columnas request_code, final_state y date_created."
        ) from exc
    except Exception as exc:
        raise BitacoraError("No pude leer el archivo exported-logs.csv.") from exc


def _drawing_event_from_record(record: Any) -> DrawingLogEvent:
    return DrawingLogEvent(
        qr=record.normalized_qr,
        initial_state=canonicalize_state(getattr(record, "initial_state", "")),
        state=canonicalize_state(record.final_state),
        occurred_at=record.parsed_date.to_pydatetime(),
        comment=clean_pms_comment(getattr(record, "comments", "")),
        username=normalize_text(getattr(record, "username", "")),
        user_name=normalize_text(getattr(record, "user_created_last_name", "")),
    )


def _truncated_qr_target(log_qr: Any, target_qrs: set[str]) -> str:
    """Return the workbook QR when PMS has exactly one additional final character."""
    actual = normalize_text(log_qr)
    if len(actual) < 2 or actual in target_qrs or not actual[-1].isalnum():
        return ""
    candidate = actual[:-1]
    return candidate if candidate in target_qrs else ""


def _sorted_unique_timeline(events: Iterable[DrawingLogEvent]) -> list[DrawingLogEvent]:
    unique: dict[tuple[Any, ...], DrawingLogEvent] = {}
    for event in events:
        key = (
            event.occurred_at,
            event.initial_state,
            event.state,
            normalize_header(event.username),
            normalize_header(event.user_name),
            event.comment,
        )
        unique[key] = event
    return sorted(unique.values(), key=lambda event: event.occurred_at)


def _collect_timeline_events(
    log_source: str | Path | BinaryIO | bytes,
    qrs: set[str],
    end: pd.Timestamp,
    *,
    chunksize: int,
) -> dict[str, list[DrawingLogEvent]]:
    """Read drawing-family history only for a small, already identified QR set."""
    targets = {normalize_text(qr) for qr in qrs if normalize_text(qr)}
    timelines: dict[str, list[DrawingLogEvent]] = {}
    if not targets:
        return timelines
    for chunk in _read_log_chunks(log_source, chunksize=chunksize):
        if not LOG_REQUIRED_COLUMNS.issubset(chunk.columns):
            continue
        chunk = chunk.copy()
        chunk["parsed_date"] = pd.to_datetime(
            chunk["date_created"], errors="coerce", format="mixed"
        )
        chunk["normalized_qr"] = chunk["request_code"].map(normalize_text)
        for column in LOG_OPTIONAL_COLUMNS:
            if column not in chunk.columns:
                chunk[column] = ""
        chunk = chunk[
            chunk["normalized_qr"].isin(targets)
            & chunk["parsed_date"].notna()
            & (chunk["parsed_date"] < end)
        ].copy()
        if chunk.empty:
            continue
        chunk["final_state"] = chunk["final_state"].map(canonicalize_state)
        chunk = chunk[chunk["final_state"].isin(ALL_DRAWING_STATES)]
        for record in chunk.itertuples(index=False):
            event = _drawing_event_from_record(record)
            timelines.setdefault(event.qr, []).append(event)
    return {
        qr: _sorted_unique_timeline(events)
        for qr, events in timelines.items()
    }


def read_full_log_timelines(
    log_source: str | Path | BinaryIO | bytes,
    qrs: set[str],
    *,
    chunksize: int = 250_000,
) -> dict[str, list[DrawingLogEvent]]:
    """Read the complete history for a bounded QR set, including territory states."""
    targets = {normalize_text(qr) for qr in qrs if normalize_text(qr)}
    timelines: dict[str, list[DrawingLogEvent]] = {}
    if not targets:
        return timelines
    for chunk in _read_log_chunks(log_source, chunksize=chunksize):
        if not LOG_REQUIRED_COLUMNS.issubset(chunk.columns):
            continue
        chunk = chunk.copy()
        chunk["parsed_date"] = pd.to_datetime(
            chunk["date_created"], errors="coerce", format="mixed"
        )
        chunk["normalized_qr"] = chunk["request_code"].map(normalize_text)
        for column in LOG_OPTIONAL_COLUMNS:
            if column not in chunk.columns:
                chunk[column] = ""
        chunk = chunk[
            chunk["normalized_qr"].isin(targets)
            & chunk["parsed_date"].notna()
        ].copy()
        if chunk.empty:
            continue
        chunk["final_state"] = chunk["final_state"].map(canonicalize_state)
        for record in chunk.itertuples(index=False):
            event = _drawing_event_from_record(record)
            timelines.setdefault(event.qr, []).append(event)
    return {
        qr: _sorted_unique_timeline(events)
        for qr, events in timelines.items()
    }


def read_log_validation_data(
    log_source: str | Path | BinaryIO | bytes,
    qrs: set[str],
    production_date: date,
    *,
    chunksize: int = 250_000,
    allowed_states_by_qr: dict[str, set[str]] | None = None,
    include_rejection_comments: bool = False,
    include_timelines: bool = False,
    include_qr_aliases: bool = False,
) -> Any:
    """Return validation states, log coverage, and target events omitted from the workbook."""
    target_qrs = {normalize_text(qr) for qr in qrs if normalize_text(qr)}
    latest_on_day: dict[str, LogState] = {}
    latest_as_of: dict[str, LogState] = {}
    latest_meaningful_comment: dict[tuple[str, str], LogState] = {}
    latest_unreported: dict[str, LogState] = {}
    latest_unreported_comment: dict[tuple[str, str], LogState] = {}
    timeline_events: dict[str, list[DrawingLogEvent]] = {}
    exact_seen_qrs: set[str] = set()
    first_log_datetime: pd.Timestamp | None = None
    last_log_datetime: pd.Timestamp | None = None
    start = pd.Timestamp(production_date)
    end = pd.Timestamp(production_date + timedelta(days=1))
    saw_required_columns = False

    for chunk in _read_log_chunks(log_source, chunksize=chunksize):
        if not LOG_REQUIRED_COLUMNS.issubset(chunk.columns):
            continue
        saw_required_columns = True
        chunk = chunk.copy()
        chunk["parsed_date"] = pd.to_datetime(
            chunk["date_created"], errors="coerce", format="mixed"
        )
        chunk["normalized_qr"] = chunk["request_code"].map(normalize_text)
        exact_seen_qrs.update(
            chunk.loc[chunk["normalized_qr"].isin(target_qrs), "normalized_qr"]
        )
        for column in LOG_OPTIONAL_COLUMNS:
            if column not in chunk.columns:
                chunk[column] = ""
        valid_dates = chunk["parsed_date"].dropna()
        if not valid_dates.empty:
            chunk_first = valid_dates.min()
            chunk_last = valid_dates.max()
            if first_log_datetime is None or chunk_first < first_log_datetime:
                first_log_datetime = chunk_first
            if last_log_datetime is None or chunk_last > last_log_datetime:
                last_log_datetime = chunk_last

        unreported_chunk = chunk[
            chunk["parsed_date"].notna()
            & (chunk["parsed_date"] >= start)
            & (chunk["parsed_date"] < end)
            & chunk["normalized_qr"].ne("")
            & ~chunk["normalized_qr"].isin(target_qrs)
        ].copy()
        if not unreported_chunk.empty:
            unreported_chunk["final_state"] = unreported_chunk["final_state"].map(
                canonicalize_state
            )
            unreported_chunk = unreported_chunk[
                unreported_chunk["final_state"].isin(UNREPORTED_TARGET_STATES)
            ]
            if not unreported_chunk.empty:
                approved_professional = pd.Series(
                    [
                        is_closed_state_professional(username, user_name)
                        for username, user_name in zip(
                            unreported_chunk["username"],
                            unreported_chunk["user_created_last_name"],
                        )
                    ],
                    index=unreported_chunk.index,
                )
                is_closed_state = unreported_chunk["final_state"].isin(
                    CLOSED_UNREPORTED_STATES
                )
                unreported_chunk = unreported_chunk[
                    ~is_closed_state | approved_professional
                ]
            for record in unreported_chunk.itertuples(index=False):
                qr = record.normalized_qr
                comment = clean_pms_comment(record.comments)
                actor = format_pms_actor(
                    record.username,
                    record.user_created_last_name,
                )
                event = LogState(
                    record.final_state,
                    record.parsed_date.to_pydatetime(),
                    comment,
                    actor,
                )
                if comment:
                    comment_key = (qr, event.state)
                    previous_comment = latest_unreported_comment.get(comment_key)
                    if (
                        previous_comment is None
                        or event.occurred_at > previous_comment.occurred_at
                    ):
                        latest_unreported_comment[comment_key] = event
                previous = latest_unreported.get(qr)
                if previous is None or event.occurred_at > previous.occurred_at:
                    latest_unreported[qr] = event

        if not target_qrs:
            continue
        normalized_qr = chunk["normalized_qr"]
        potential_alias = (
            normalized_qr.str.len().gt(1)
            & normalized_qr.str[-1].str.match(r"[A-Za-z0-9]", na=False)
            & normalized_qr.str[:-1].isin(target_qrs)
            & ~normalized_qr.isin(target_qrs)
        )
        chunk = chunk[normalized_qr.isin(target_qrs) | potential_alias]
        if chunk.empty:
            continue
        chunk["final_state"] = chunk["final_state"].map(canonicalize_state)
        chunk = chunk[chunk["final_state"].isin(ALL_DRAWING_STATES)]
        if allowed_states_by_qr:
            allowed_mask = [
                state
                in allowed_states_by_qr.get(
                    _truncated_qr_target(qr, target_qrs) or qr,
                    ALL_DRAWING_STATES,
                )
                for qr, state in zip(chunk["normalized_qr"], chunk["final_state"])
            ]
            chunk = chunk[allowed_mask]
        if chunk.empty:
            continue
        chunk = chunk[chunk["parsed_date"].notna() & (chunk["parsed_date"] < end)]
        if chunk.empty:
            continue

        for record in chunk.itertuples(index=False):
            qr = record.normalized_qr
            comment = clean_pms_comment(record.comments)
            actor = format_pms_actor(
                record.username,
                record.user_created_last_name,
            )
            event = LogState(
                record.final_state,
                record.parsed_date.to_pydatetime(),
                comment,
                actor,
            )
            if include_timelines:
                timeline_events.setdefault(qr, []).append(
                    _drawing_event_from_record(record)
                )
            if comment:
                comment_key = (qr, event.state)
                previous_comment = latest_meaningful_comment.get(comment_key)
                if (
                    previous_comment is None
                    or event.occurred_at > previous_comment.occurred_at
                ):
                    latest_meaningful_comment[comment_key] = event
            previous = latest_as_of.get(qr)
            if previous is None or event.occurred_at > previous.occurred_at:
                latest_as_of[qr] = event
            if record.parsed_date >= start:
                previous_day = latest_on_day.get(qr)
                if previous_day is None or event.occurred_at > previous_day.occurred_at:
                    latest_on_day[qr] = event

    if not saw_required_columns:
        raise BitacoraError(
            "El CSV no contiene las columnas request_code, final_state y date_created."
        )
    alias_candidates: dict[str, set[str]] = {}
    for actual_qr, event in latest_on_day.items():
        target_qr = _truncated_qr_target(actual_qr, target_qrs)
        if not target_qr:
            continue
        allowed_states = (
            allowed_states_by_qr.get(target_qr, ALL_DRAWING_STATES)
            if allowed_states_by_qr
            else ALL_DRAWING_STATES
        )
        if event.state in allowed_states:
            alias_candidates.setdefault(target_qr, set()).add(actual_qr)
    qr_aliases = {
        target_qr: next(iter(candidates))
        for target_qr, candidates in alias_candidates.items()
        if len(candidates) == 1 and target_qr not in exact_seen_qrs
    }
    for corrected_qr in qr_aliases.values():
        latest_unreported.pop(corrected_qr, None)

    for qr, event in list(latest_on_day.items()):
        if event.comment or event.state not in REJECTION_STATES:
            continue
        comment_event = latest_meaningful_comment.get((qr, event.state))
        if comment_event is not None and comment_event.occurred_at <= event.occurred_at:
            latest_on_day[qr] = LogState(
                event.state,
                event.occurred_at,
                comment_event.comment,
                event.actor,
            )
    for qr, event in list(latest_unreported.items()):
        if event.comment or event.state not in REJECTION_STATES:
            continue
        comment_event = latest_unreported_comment.get((qr, event.state))
        if comment_event is not None and comment_event.occurred_at <= event.occurred_at:
            latest_unreported[qr] = LogState(
                event.state,
                event.occurred_at,
                comment_event.comment,
                event.actor,
            )
    coverage = LogCoverage(
        first_datetime=(
            first_log_datetime.to_pydatetime() if first_log_datetime is not None else None
        ),
        last_datetime=(
            last_log_datetime.to_pydatetime() if last_log_datetime is not None else None
        ),
    )
    unreported_cases = [
        UnreportedCase(
            qr=qr,
            state=event.state,
            occurred_at=event.occurred_at,
            comment=event.comment,
            actor=event.actor,
        )
        for qr, event in sorted(
            latest_unreported.items(),
            key=lambda item: (item[1].occurred_at, item[0]),
            reverse=True,
        )
    ]
    if include_timelines:
        missing_timeline_qrs = set(latest_unreported) - set(timeline_events)
        if missing_timeline_qrs:
            timeline_events.update(
                _collect_timeline_events(
                    log_source,
                    missing_timeline_qrs,
                    end,
                    chunksize=chunksize,
                )
            )
        timeline_events = {
            qr: _sorted_unique_timeline(events)
            for qr, events in timeline_events.items()
        }
    result = (latest_on_day, latest_as_of, coverage, unreported_cases)
    additions: list[Any] = []
    if include_rejection_comments:
        rejection_comments = {
            key: event
            for key, event in latest_meaningful_comment.items()
            if key[1] in REJECTION_STATES
        }
        additions.append(rejection_comments)
    if include_timelines:
        additions.append(timeline_events)
    if include_qr_aliases:
        additions.append(qr_aliases)
    return (*result, *additions)


def read_log_states_with_coverage(
    log_source: str | Path | BinaryIO | bytes,
    qrs: set[str],
    production_date: date,
    *,
    chunksize: int = 250_000,
    allowed_states_by_qr: dict[str, set[str]] | None = None,
) -> tuple[dict[str, LogState], dict[str, LogState], LogCoverage]:
    """Return matching states plus the real date coverage found in the log."""
    latest_on_day, latest_as_of, coverage, _ = read_log_validation_data(
        log_source,
        qrs,
        production_date,
        chunksize=chunksize,
        allowed_states_by_qr=allowed_states_by_qr,
    )
    return latest_on_day, latest_as_of, coverage


def read_latest_log_states(
    log_source: str | Path | BinaryIO | bytes,
    qrs: set[str],
    production_date: date,
    *,
    chunksize: int = 250_000,
    allowed_states_by_qr: dict[str, set[str]] | None = None,
) -> tuple[dict[str, LogState], dict[str, LogState]]:
    """Return latest drawing state on the day and latest state up to end of day."""
    latest_on_day, latest_as_of, _ = read_log_states_with_coverage(
        log_source,
        qrs,
        production_date,
        chunksize=chunksize,
        allowed_states_by_qr=allowed_states_by_qr,
    )
    return latest_on_day, latest_as_of


def require_log_date_coverage(coverage: LogCoverage, production_date: date) -> None:
    """Stop validation when the uploaded snapshot cannot contain the requested day."""
    if coverage.first_date is None or coverage.last_date is None:
        raise BitacoraError(
            "El exported-logs no contiene fechas válidas en la columna date_created."
        )
    if coverage.first_date <= production_date <= coverage.last_date:
        return

    if production_date > coverage.last_date:
        detail = (
            f"El log termina el {coverage.last_date:%d/%m/%Y} y la bitácora es del "
            f"{production_date:%d/%m/%Y}."
        )
        recommendation = (
            "Descarga un exported-logs generado al finalizar esa jornada o en una fecha "
            "posterior y vuelve a validar."
        )
    else:
        detail = (
            f"El log comienza el {coverage.first_date:%d/%m/%Y} y la bitácora es del "
            f"{production_date:%d/%m/%Y}."
        )
        recommendation = "Carga un exported-logs que incluya esa fecha y vuelve a validar."
    raise BitacoraError(
        f"El archivo del PMS no cubre la fecha de producción. {detail} {recommendation}"
    )


def _flow_timing_states(sheet: str) -> tuple[set[str], set[str]]:
    if sheet == SHEET_OLD:
        return LEGACY_ARRIVAL_STATES, LEGACY_OUTPUT_STATES
    return NEW_ARRIVAL_STATES, NEW_OUTPUT_STATES


def _is_business_exit(event: DrawingLogEvent, sheet: str) -> bool:
    """Reject no-op/technical log rows and retain actual drawing outputs."""
    _, output_states = _flow_timing_states(sheet)
    if event.state not in output_states or event.initial_state == event.state:
        return False
    if sheet == SHEET_NEW:
        return event.initial_state in {"", "assigned_draftsman"}
    if event.state == "2-Valorizado":
        return event.initial_state == "2.6-Dibujante asignado"
    return event.initial_state in {"", "2.6-Dibujante asignado"}


def _latest_output_event(
    events: Iterable[DrawingLogEvent],
    sheet: str,
    production_date: date,
    *,
    professional: str = "",
    desired_state: str = "",
) -> DrawingLogEvent | None:
    candidates = [
        event
        for event in events
        if event.occurred_at.date() == production_date
        and _is_business_exit(event, sheet)
        and (not desired_state or event.state == desired_state)
        and (
            not professional
            or same_professional(professional, event.username, event.user_name)
        )
    ]
    return max(candidates, key=lambda event: event.occurred_at, default=None)


def _latest_output_row_even_if_transition_is_unusual(
    events: Iterable[DrawingLogEvent],
    sheet: str,
    production_date: date,
) -> DrawingLogEvent | None:
    _, output_states = _flow_timing_states(sheet)
    candidates = [
        event
        for event in events
        if event.occurred_at.date() == production_date
        and event.state in output_states
    ]
    return max(candidates, key=lambda event: event.occurred_at, default=None)


def _cycle_arrival_for_exit(
    events: Iterable[DrawingLogEvent],
    sheet: str,
    exit_event: DrawingLogEvent,
) -> DrawingLogEvent | None:
    arrival_states, _ = _flow_timing_states(sheet)
    ordered = sorted(
        (event for event in events if event.occurred_at <= exit_event.occurred_at),
        key=lambda event: event.occurred_at,
    )
    # Any real drawing exit closes the preceding cycle.  The authorized-user
    # list belongs only to the "casos no reportados" control; using it here can
    # make a later re-entry inherit an arrival from an older cycle.
    previous_exits = [
        event
        for event in ordered
        if event.occurred_at < exit_event.occurred_at
        and _is_business_exit(event, sheet)
    ]
    boundary = max(
        (event.occurred_at for event in previous_exits),
        default=datetime.min,
    )
    arrivals = [
        event
        for event in ordered
        if boundary < event.occurred_at < exit_event.occurred_at
        and event.state in arrival_states
        and event.initial_state != event.state
    ]
    return min(arrivals, key=lambda event: event.occurred_at, default=None)


def _flow20_odm_event(
    events: Iterable[DrawingLogEvent],
    cutoff_end: datetime,
) -> DrawingLogEvent | None:
    candidates = [
        event
        for event in events
        if event.occurred_at < cutoff_end
        and event.state == FLOW20_ODM_STATE
        and event.initial_state in FLOW20_ODM_ENTRY_INITIAL_STATES
        and event.initial_state != event.state
    ]
    return max(candidates, key=lambda event: event.occurred_at, default=None)


def _flow20_arrival_for_odm(
    events: Iterable[DrawingLogEvent],
    odm_event: DrawingLogEvent,
) -> DrawingLogEvent | None:
    ordered = sorted(
        (event for event in events if event.occurred_at < odm_event.occurred_at),
        key=lambda event: event.occurred_at,
    )
    previous_outputs = [
        event
        for event in ordered
        if event.state in {"2-Valorizado", FLOW20_ODM_STATE}
        and event.initial_state != event.state
    ]
    boundary = max(
        (event.occurred_at for event in previous_outputs),
        default=datetime.min,
    )
    arrivals = [
        event
        for event in ordered
        if boundary < event.occurred_at < odm_event.occurred_at
        and event.state in LEGACY_ARRIVAL_STATES
        and event.initial_state != event.state
    ]
    arrival = min(arrivals, key=lambda event: event.occurred_at, default=None)
    if arrival is not None:
        return arrival
    if odm_event.initial_state == "2-Valorizado":
        direct_arrivals = [
            event
            for event in ordered
            if event.state == "2-Valorizado"
            and event.initial_state != event.state
        ]
        return max(
            direct_arrivals,
            key=lambda event: event.occurred_at,
            default=None,
        )
    return None


def _flow20_is_technical_repeat(event: DrawingLogEvent) -> bool:
    comment = normalize_header(event.comment)
    technical_markers = (
        "DRAWING ASSIGNMENT",
        "DRAWING REASSIGNMENT",
        "REASSIGNED",
        "REASIGNACION",
        "MARCACION DE NODOS",
    )
    return any(marker in comment for marker in technical_markers)


def _flow20_drawing_event(
    events: Iterable[DrawingLogEvent],
    odm_event: DrawingLogEvent,
    professional: str,
) -> DrawingLogEvent | None:
    candidates = [
        event
        for event in events
        if event.occurred_at > odm_event.occurred_at
        and event.initial_state
        and event.initial_state == event.state
        and not _flow20_is_technical_repeat(event)
        and same_professional(professional, event.username, event.user_name)
    ]
    return min(candidates, key=lambda event: event.occurred_at, default=None)


def _validate_flow20_row(
    row: dict[str, Any],
    previous_event: LogState | None,
    request_snapshot: RequestSnapshot | None,
    events: list[DrawingLogEvent],
    production_date: date,
) -> ValidationRow:
    qr = normalize_text(row.get("qr"))
    reported = canonicalize_state(row.get("reported_state"))
    secondary = canonicalize_flow20_secondary(
        row.get("reported_secondary_state")
    )
    professional = normalize_text(row.get("professional"))
    cutoff_end = datetime.combine(
        production_date + timedelta(days=1), datetime.min.time()
    )
    odm_event = _flow20_odm_event(events, cutoff_end)
    arrival = _flow20_arrival_for_odm(events, odm_event) if odm_event else None
    drawing_event = (
        _flow20_drawing_event(events, odm_event, professional)
        if odm_event and request_snapshot is not None
        and request_snapshot.drawing_completed is True
        else None
    )
    drawing_as_of = (
        drawing_event
        if drawing_event is not None and drawing_event.occurred_at < cutoff_end
        else None
    )
    future_drawing = (
        drawing_event
        if drawing_event is not None and drawing_event.occurred_at >= cutoff_end
        else None
    )
    simultaneous = bool(
        odm_event
        and odm_event.occurred_at.date() == production_date
        and request_snapshot is not None
        and request_snapshot.drawing_completed is True
        and drawing_event is None
    )
    drawing_datetime = (
        drawing_as_of.occurred_at
        if drawing_as_of is not None
        else odm_event.occurred_at
        if simultaneous and odm_event is not None
        else None
    )
    drawing_actor = drawing_as_of or (odm_event if simultaneous else None)
    actor_event = drawing_actor or odm_event
    professional_mismatch = bool(
        actor_event
        and professional
        and not same_professional(
            professional,
            actor_event.username,
            actor_event.user_name,
        )
    )
    request_flag = (
        request_snapshot.drawing_completed if request_snapshot else None
    )
    new_flow_flag = (
        request_snapshot.new_fulfillment_process if request_snapshot else None
    )

    evidence_parts: list[str] = []
    if odm_event:
        evidence_parts.append(
            f"ODM: {odm_event.initial_state} → {odm_event.state} "
            f"({odm_event.occurred_at:%d/%m/%Y %H:%M:%S})"
        )
    else:
        evidence_parts.append("ODM no localizada")
    if drawing_as_of:
        evidence_parts.append(
            f"Dibujo: repetición {drawing_as_of.state} "
            f"({drawing_as_of.occurred_at:%d/%m/%Y %H:%M:%S})"
        )
    elif simultaneous:
        evidence_parts.append("Dibujo inferido como entrega simultánea con la ODM")
    elif future_drawing:
        evidence_parts.append(
            f"Dibujo posterior al corte: {future_drawing.occurred_at:%d/%m/%Y %H:%M:%S}"
        )
    elif request_flag is True:
        evidence_parts.append("Drawing Completed=True sin evento fechable")
    elif request_flag is False:
        evidence_parts.append("Drawing Completed=False")
    else:
        evidence_parts.append("Drawing Completed sin dato")

    result = STATUS_MISMATCH
    timing_result = TIMING_FLOW20_AMBIGUOUS
    duplicate_detail = ""
    if not qr:
        result = STATUS_NO_QR
    elif reported == FLOW20_ODM_STATE:
        if odm_event is None:
            result = STATUS_NO_EVENT
        elif secondary == FLOW20_SECONDARY_PENDING:
            if drawing_as_of is not None:
                result = STATUS_MISMATCH
            elif odm_event.occurred_at.date() < production_date:
                result = STATUS_FLOW20_DUPLICATE_ODM
                timing_result = TIMING_FLOW20_PENDING
                duplicate_detail = (
                    "La ODM ya había sido entregada el "
                    f"{odm_event.occurred_at:%d/%m/%Y}; no es un hito del corte."
                )
            elif request_snapshot is None:
                result = STATUS_NO_EVENT
            elif new_flow_flag is False:
                result = STATUS_MISMATCH
            elif professional_mismatch:
                result = STATUS_PROFESSIONAL_MISMATCH
            elif request_flag is False or future_drawing is not None:
                result = STATUS_MATCH
                timing_result = TIMING_FLOW20_PENDING
            else:
                result = STATUS_FLOW20_AMBIGUOUS
        elif secondary == FLOW20_SECONDARY_COMPLETED:
            if drawing_datetime is not None and drawing_datetime.date() < production_date:
                result = STATUS_FLOW20_DUPLICATE_DRAWING
                timing_result = TIMING_FLOW20_COMPLETED
                duplicate_detail = (
                    "El dibujo ya había sido entregado el "
                    f"{drawing_datetime:%d/%m/%Y}; no es un hito del corte."
                )
            elif request_snapshot is None:
                result = STATUS_NO_EVENT
            elif new_flow_flag is False:
                result = STATUS_MISMATCH
            elif (
                drawing_datetime is not None
                and drawing_datetime.date() == production_date
                and request_flag is True
            ):
                if professional_mismatch:
                    result = STATUS_PROFESSIONAL_MISMATCH
                else:
                    result = STATUS_MATCH
                    timing_result = TIMING_FLOW20_COMPLETED
            elif request_flag is True and drawing_datetime is None:
                result = STATUS_FLOW20_AMBIGUOUS
            else:
                result = STATUS_MISMATCH
        else:
            result = STATUS_MISMATCH
    else:
        day_events = [
            event
            for event in events
            if event.occurred_at.date() == production_date
            and event.state in FLOW20_STATES
        ]
        day_event = max(day_events, key=lambda event: event.occurred_at, default=None)
        if day_event is None:
            result = STATUS_NO_EVENT
        elif (
            professional
            and day_event.state in {"2-Valorizado", FLOW20_ODM_STATE}
            and not same_professional(
                professional, day_event.username, day_event.user_name
            )
        ):
            result = STATUS_PROFESSIONAL_MISMATCH
        elif day_event.state == reported:
            result = STATUS_MATCH
        else:
            result = STATUS_MISMATCH
        actor_event = day_event
        timing_result = TIMING_NOT_APPLICABLE

    if duplicate_detail:
        evidence_parts.append(duplicate_detail)

    pms_state = (
        FLOW20_ODM_STATE
        if odm_event and reported == FLOW20_ODM_STATE
        else actor_event.state
        if actor_event
        else ""
    )
    pms_datetime = (
        drawing_datetime
        if secondary == FLOW20_SECONDARY_COMPLETED and drawing_datetime
        else odm_event.occurred_at
        if odm_event and reported == FLOW20_ODM_STATE
        else actor_event.occurred_at
        if actor_event
        else None
    )
    return ValidationRow(
        sheet=SHEET_FLOW20,
        excel_row=int(row["excel_row"]),
        qr=qr,
        reported_state=reported,
        pms_state=pms_state,
        pms_datetime=pms_datetime,
        previous_state=previous_event.state if previous_event else "",
        previous_datetime=previous_event.occurred_at if previous_event else None,
        result=result,
        professional=professional,
        pms_actor=actor_event.actor if actor_event else "",
        reported_secondary_state=secondary,
        agp=normalize_identifier(row.get("agp")),
        odm=normalize_identifier(row.get("odm")),
        protocol=normalize_identifier(row.get("protocol")),
        reported_pel=row.get("reported_pel"),
        reported_tratas=row.get("reported_tratas"),
        count_result=COUNT_NOT_EVALUATED,
        arrival_state=arrival.state if arrival else "",
        arrival_datetime=arrival.occurred_at if arrival else None,
        exit_datetime=drawing_datetime or (odm_event.occurred_at if odm_event else None),
        timing_result=timing_result,
        timing_source="PMS" if odm_event else "",
        request_drawing_completed=request_flag,
        request_status=request_snapshot.status if request_snapshot else "",
        odm_datetime=odm_event.occurred_at if odm_event else None,
        drawing_datetime=drawing_datetime,
        flow20_evidence=" | ".join(evidence_parts),
    )


def apply_drawing_timings(
    validations: list[ValidationRow],
    timelines: dict[str, list[DrawingLogEvent]],
    production_date: date,
) -> list[ValidationRow]:
    """Attach the current drawing cycle's exact arrival and professional exit."""
    updated: list[ValidationRow] = []
    for item in validations:
        if item.sheet == SHEET_FLOW20:
            updated.append(item)
            continue
        if item.reported_state not in DRAWING_OUTPUT_STATES:
            updated.append(
                replace(
                    item,
                    arrival_state="",
                    arrival_datetime=None,
                    exit_datetime=None,
                    timing_result=TIMING_NOT_APPLICABLE,
                    timing_source="",
                )
            )
            continue
        if item.timing_source == "Auditoría":
            updated.append(item)
            continue

        events = timelines.get(item.qr, [])
        desired_state = (
            item.reported_state if item.result in UPLOAD_APPROVED_RESULTS else ""
        )
        exit_event = _latest_output_event(
            events,
            item.sheet,
            production_date,
            professional=item.professional,
            desired_state=desired_state,
        )
        if exit_event is None and item.result not in UPLOAD_APPROVED_RESULTS:
            exit_event = _latest_output_row_even_if_transition_is_unusual(
                events,
                item.sheet,
                production_date,
            )
        if exit_event is None:
            updated.append(
                replace(
                    item,
                    arrival_state="",
                    arrival_datetime=None,
                    exit_datetime=None,
                    timing_result=TIMING_NO_EXIT,
                    timing_source="",
                )
            )
            continue

        arrival = _cycle_arrival_for_exit(events, item.sheet, exit_event)
        updated.append(
            replace(
                item,
                pms_actor=exit_event.actor or item.pms_actor,
                arrival_state=arrival.state if arrival else "",
                arrival_datetime=arrival.occurred_at if arrival else None,
                exit_datetime=exit_event.occurred_at,
                timing_result=TIMING_CAPTURED if arrival else TIMING_NO_ARRIVAL,
                timing_source="PMS",
            )
        )
    return updated


def parse_audit_datetime(value: Any) -> datetime | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value
    text = normalize_text(value)
    for format_string in (
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%dT%H:%M:%S",
    ):
        try:
            return datetime.strptime(text, format_string)
        except ValueError:
            continue
    raise BitacoraError(
        "La fecha/hora debe usar el formato DD/MM/AAAA HH:MM:SS."
    )


def _manual_timing_fields(
    decision: dict[str, Any],
    target_state: str,
) -> dict[str, Any]:
    if target_state not in DRAWING_OUTPUT_STATES:
        return {
            "arrival_state": "",
            "arrival_datetime": None,
            "exit_datetime": None,
            "timing_result": TIMING_NOT_APPLICABLE,
            "timing_source": "",
        }
    arrival = parse_audit_datetime(decision.get("arrival_datetime"))
    exit_datetime = parse_audit_datetime(decision.get("exit_datetime"))
    if bool(arrival) != bool(exit_datetime):
        raise BitacoraError(
            "Para confirmar los tiempos debes diligenciar llegada y salida completas."
        )
    if arrival and exit_datetime:
        if arrival > exit_datetime:
            raise BitacoraError("La fecha/hora de llegada no puede ser posterior a la salida.")
        return {
            "arrival_datetime": arrival,
            "exit_datetime": exit_datetime,
            "timing_result": TIMING_MANUAL,
            "timing_source": "Auditoría",
        }
    return {}


def _manual_flow20_timing_fields(
    decision: dict[str, Any],
    target_state: str,
    secondary_state: str,
) -> dict[str, Any]:
    if target_state != FLOW20_ODM_STATE:
        return {
            "arrival_state": "",
            "arrival_datetime": None,
            "exit_datetime": None,
            "odm_datetime": None,
            "drawing_datetime": None,
            "timing_result": TIMING_NOT_APPLICABLE,
            "timing_source": "",
        }
    arrival = parse_audit_datetime(decision.get("arrival_datetime"))
    odm_datetime = parse_audit_datetime(decision.get("odm_datetime"))
    drawing_datetime = parse_audit_datetime(decision.get("drawing_datetime"))
    if not arrival or not odm_datetime:
        raise BitacoraError(
            "Para Flujo 2.0 debes confirmar la llegada y la entrega de la ODM."
        )
    if arrival > odm_datetime:
        raise BitacoraError(
            "La fecha/hora de llegada no puede ser posterior a la entrega de la ODM."
        )
    if secondary_state == FLOW20_SECONDARY_COMPLETED:
        if not drawing_datetime:
            raise BitacoraError(
                "Para un dibujo finalizado debes confirmar la fecha/hora del dibujo."
            )
        if drawing_datetime < odm_datetime:
            raise BitacoraError(
                "La entrega del dibujo no puede ser anterior a la entrega de la ODM."
            )
        timing_result = TIMING_FLOW20_COMPLETED
    else:
        if drawing_datetime:
            raise BitacoraError(
                "Un caso pendiente de dibujo no debe tener fecha de entrega del dibujo."
            )
        timing_result = TIMING_FLOW20_PENDING
    return {
        "arrival_datetime": arrival,
        "exit_datetime": drawing_datetime or odm_datetime,
        "odm_datetime": odm_datetime,
        "drawing_datetime": drawing_datetime,
        "timing_result": timing_result,
        "timing_source": "Auditoría",
    }


def validate_rows(
    rows: list[dict[str, Any]],
    latest_on_day: dict[str, LogState],
    latest_as_of: dict[str, LogState],
    protocol_counts: dict[str, ProtocolElementCounts] | None = None,
    rejection_comments: dict[tuple[str, str], LogState] | None = None,
    timelines: dict[str, list[DrawingLogEvent]] | None = None,
    production_date: date | None = None,
    request_snapshots: dict[str, RequestSnapshot] | None = None,
) -> list[ValidationRow]:
    validations: list[ValidationRow] = []
    for row in rows:
        qr = row["qr"]
        reported = canonicalize_state(row["reported_state"])
        day_event = latest_on_day.get(qr)
        previous_event = latest_as_of.get(qr)
        if row["sheet"] == SHEET_FLOW20 and production_date is not None:
            validations.append(
                _validate_flow20_row(
                    row,
                    previous_event,
                    (request_snapshots or {}).get(qr),
                    (timelines or {}).get(qr, []),
                    production_date,
                )
            )
            continue
        professional = normalize_text(row.get("professional"))
        professional_mismatch = False
        if (
            timelines is not None
            and production_date is not None
            and reported in DRAWING_OUTPUT_STATES
        ):
            events = timelines.get(qr, [])
            professional_event = _latest_output_event(
                events,
                row["sheet"],
                production_date,
                professional=professional,
            )
            any_event = _latest_output_event(
                events,
                row["sheet"],
                production_date,
            ) or _latest_output_row_even_if_transition_is_unusual(
                events,
                row["sheet"],
                production_date,
            )
            selected = professional_event or any_event
            if selected is not None:
                day_event = LogState(
                    selected.state,
                    selected.occurred_at,
                    selected.comment,
                    selected.actor,
                )
            else:
                day_event = None
            professional_mismatch = professional_event is None and any_event is not None
        if not qr:
            result = STATUS_NO_QR
        elif day_event is None:
            result = STATUS_NO_EVENT
        elif professional_mismatch:
            result = STATUS_PROFESSIONAL_MISMATCH
        elif canonicalize_state(day_event.state) == reported:
            result = STATUS_MATCH
        else:
            result = STATUS_MISMATCH

        protocol = normalize_identifier(row.get("protocol"))
        reported_pel = row.get("reported_pel")
        reported_tratas = row.get("reported_tratas")
        export_pel: int | None = None
        export_tratas: int | None = None
        export_pel_breakdown: Any = None
        export_tratas_breakdown: Any = None
        count_result = COUNT_NOT_EVALUATED
        if protocol_counts is not None:
            if reported not in FINAL_ELEMENT_STATES:
                count_result = COUNT_NOT_APPLICABLE
            else:
                export_pel, export_tratas, count_result = compare_protocol_quantities(
                    protocol,
                    reported_pel,
                    reported_tratas,
                    protocol_counts,
                )
                export_pel_breakdown, export_tratas_breakdown = protocol_quantity_breakdown(
                    protocol,
                    protocol_counts,
                )

        original_observation = normalize_text(row.get("reported_observation"))
        cleaned_original = clean_placeholder_observation(original_observation)
        pms_observation = ""
        if reported in REJECTION_STATES:
            matching_comment = None
            if timelines is not None and production_date is not None:
                matching_event = _latest_output_event(
                    timelines.get(qr, []),
                    row["sheet"],
                    production_date,
                    professional=professional,
                    desired_state=reported,
                )
                if matching_event is not None:
                    comment_candidates = [
                        event
                        for event in timelines.get(qr, [])
                        if event.state == reported
                        and event.comment
                        and event.occurred_at <= matching_event.occurred_at
                        and same_professional(
                            professional,
                            event.username,
                            event.user_name,
                        )
                    ]
                    comment = (
                        max(comment_candidates, key=lambda event: event.occurred_at).comment
                        if comment_candidates
                        else matching_event.comment
                    )
                    matching_comment = LogState(
                        matching_event.state,
                        matching_event.occurred_at,
                        comment,
                        matching_event.actor,
                    )
            else:
                matching_comment = (rejection_comments or {}).get((qr, reported))
            if matching_comment is not None:
                pms_observation = matching_comment.comment
            if result != STATUS_MATCH:
                observation_result = OBSERVATION_STATE_MISMATCH
            elif day_event and day_event.comment:
                pms_observation = day_event.comment
                observation_result = OBSERVATION_UPDATED
            elif pms_observation:
                observation_result = OBSERVATION_UPDATED
            else:
                observation_result = OBSERVATION_NOT_FOUND
        elif original_observation and not cleaned_original:
            observation_result = OBSERVATION_CLEANED
        else:
            observation_result = OBSERVATION_NOT_APPLICABLE

        validations.append(
            ValidationRow(
                sheet=row["sheet"],
                excel_row=row["excel_row"],
                qr=qr,
                reported_state=reported,
                pms_state=day_event.state if day_event else "",
                pms_datetime=day_event.occurred_at if day_event else None,
                previous_state=previous_event.state if previous_event else "",
                previous_datetime=previous_event.occurred_at if previous_event else None,
                result=result,
                professional=professional,
                pms_actor=day_event.actor if day_event else "",
                reported_secondary_state=canonicalize_flow20_secondary(
                    row.get("reported_secondary_state")
                ),
                agp=normalize_identifier(row.get("agp")),
                odm=normalize_identifier(row.get("odm")),
                protocol=protocol,
                reported_pel=reported_pel,
                reported_tratas=reported_tratas,
                export_pel=export_pel,
                export_tratas=export_tratas,
                export_pel_breakdown=export_pel_breakdown,
                export_tratas_breakdown=export_tratas_breakdown,
                count_result=count_result,
                pms_observation=pms_observation,
                observation_result=observation_result,
            )
        )
    if timelines is not None and production_date is not None:
        return apply_drawing_timings(validations, timelines, production_date)
    return validations


def validation_audit_key(item: ValidationRow) -> str:
    return f"{item.sheet}|{item.excel_row}|{item.qr}"


def flow20_duplicate_audit_entries(
    validations: Iterable[ValidationRow],
) -> list[dict[str, Any]]:
    """Record historical Flow 2.0 milestones excluded without manual review."""
    entries: list[dict[str, Any]] = []
    for item in validations:
        if item.result not in FLOW20_DUPLICATE_RESULTS:
            continue
        milestone_datetime = (
            item.odm_datetime
            if item.result == STATUS_FLOW20_DUPLICATE_ODM
            else item.drawing_datetime
        )
        milestone = (
            "entrega de ODM"
            if item.result == STATUS_FLOW20_DUPLICATE_ODM
            else "entrega de dibujo"
        )
        entries.append(
            {
                "Etapa": "Control de redundancia Flujo 2.0",
                "Hoja/Flujo": item.sheet,
                "QR/Referencia": item.qr,
                "Decisión": "Excluir duplicado automáticamente",
                "Estado inicial": item.reported_state,
                "Estado final": item.reported_state,
                "Estado secundario": item.reported_secondary_state,
                "AGP": item.agp,
                "Protocolo": item.protocol,
                "ODM": item.odm,
                "Fecha/hora de llegada": item.arrival_datetime,
                "Fecha/hora de salida": item.exit_datetime,
                "Fecha/hora entrega ODM": item.odm_datetime,
                "Fecha/hora entrega dibujo": item.drawing_datetime,
                "Detalle": (
                    f"Se excluyó porque la {milestone} pertenece al "
                    f"{milestone_datetime:%d/%m/%Y %H:%M:%S} y no a la "
                    "fecha de corte de esta bitácora."
                    if milestone_datetime
                    else "Se excluyó porque el hito pertenece a una fecha anterior."
                ),
            }
        )
    return entries


def apply_state_audit_decisions(
    workbook_bytes: bytes,
    validations: list[ValidationRow],
    latest_on_day: dict[str, LogState],
    decisions: dict[str, dict[str, Any]],
    rejection_comments: dict[tuple[str, str], LogState] | None = None,
    timelines: dict[str, list[DrawingLogEvent]] | None = None,
    production_date: date | None = None,
) -> tuple[bytes, list[ValidationRow], list[dict[str, Any]]]:
    """Apply the user's state decisions while preserving an auditable result."""
    try:
        wb = load_workbook(BytesIO(workbook_bytes), data_only=False)
    except Exception as exc:
        raise BitacoraError("No pude aplicar las decisiones de auditoría.") from exc

    updated: list[ValidationRow] = []
    audit_entries: list[dict[str, Any]] = []
    sheet_headers: dict[str, dict[str, int]] = {}

    def headers_for(sheet_name: str) -> dict[str, int]:
        headers = sheet_headers.get(sheet_name)
        if headers is None:
            ws = wb[sheet_name]
            header_row = _detect_header_row(ws, {"QR", "ESTADO DE LA OT"})
            headers = _header_map(ws, header_row=header_row)
            sheet_headers[sheet_name] = headers
        return headers

    def matching_rejection_comment(qr: str, state: str) -> str:
        item_professional = next(
            (
                row.professional
                for row in validations
                if row.qr == qr
            ),
            "",
        )
        if timelines is not None and production_date is not None:
            sheet_name = next(
                (row.sheet for row in validations if row.qr == qr),
                SHEET_NEW,
            )
            exit_event = _latest_output_event(
                timelines.get(qr, []),
                sheet_name,
                production_date,
                professional=item_professional,
                desired_state=state,
            )
            if exit_event is not None:
                comments = [
                    event
                    for event in timelines.get(qr, [])
                    if event.state == state
                    and event.comment
                    and event.occurred_at <= exit_event.occurred_at
                    and same_professional(
                        item_professional,
                        event.username,
                        event.user_name,
                    )
                ]
                if comments:
                    return max(comments, key=lambda event: event.occurred_at).comment
                return exit_event.comment
        event = (rejection_comments or {}).get((qr, state))
        if event is not None and event.comment:
            return event.comment
        day_event = latest_on_day.get(qr)
        if day_event is not None and day_event.state == state:
            return day_event.comment
        return ""

    for item in validations:
        decision = decisions.get(validation_audit_key(item))
        if not decision:
            updated.append(item)
            continue

        action = normalize_text(decision.get("action")).lower()
        base_entry = {
            "Etapa": "Estados",
            "Hoja/Flujo": item.sheet,
            "QR/Referencia": item.qr,
            "Estado inicial": item.reported_state,
            "Estado secundario": item.reported_secondary_state,
            "AGP": normalize_identifier(decision.get("agp")),
            "Protocolo": normalize_identifier(decision.get("protocol")),
            "ODM": normalize_identifier(decision.get("odm")),
            "Fecha/hora de llegada": normalize_text(
                decision.get("arrival_datetime")
            ),
            "Fecha/hora de salida": normalize_text(
                decision.get("exit_datetime")
            ),
            "Fecha/hora entrega ODM": normalize_text(
                decision.get("odm_datetime")
            ),
            "Fecha/hora entrega dibujo": normalize_text(
                decision.get("drawing_datetime")
            ),
        }
        if action == "keep":
            secondary_state = canonicalize_flow20_secondary(
                decision.get("secondary_state")
                or item.reported_secondary_state
            )
            timing_fields = (
                _manual_flow20_timing_fields(
                    decision,
                    item.reported_state,
                    secondary_state,
                )
                if item.sheet == SHEET_FLOW20
                else _manual_timing_fields(decision, item.reported_state)
            )
            observation = ""
            observation_result = item.observation_result
            if item.reported_state in REJECTION_STATES:
                observation = matching_rejection_comment(
                    item.qr,
                    item.reported_state,
                )
                headers = headers_for(item.sheet)
                observation_column = headers.get("OBSERVACIONES")
                if observation and observation_column:
                    wb[item.sheet].cell(
                        item.excel_row,
                        observation_column,
                    ).value = observation
                observation_result = (
                    OBSERVATION_UPDATED if observation else OBSERVATION_NOT_FOUND
                )
            updated.append(
                replace(
                    item,
                    result=STATUS_AUDIT_KEEP,
                    reported_secondary_state=secondary_state,
                    pms_observation=observation or item.pms_observation,
                    observation_result=observation_result,
                    flow20_evidence=(
                        f"{item.flow20_evidence} | Fechas confirmadas por auditoría"
                        if item.sheet == SHEET_FLOW20
                        and item.flow20_evidence
                        else "Fechas confirmadas por auditoría"
                        if item.sheet == SHEET_FLOW20
                        else item.flow20_evidence
                    ),
                    **timing_fields,
                )
            )
            audit_entries.append(
                {
                    **base_entry,
                    "Decisión": "Conservar bitácora",
                    "Estado final": item.reported_state,
                    "Estado secundario": secondary_state,
                    "Detalle": (
                        "El auditor confirmó la información reportada y se recuperó "
                        "la devolución más reciente del mismo estado."
                        if observation
                        else "El auditor confirmó la información reportada."
                    ),
                }
            )
            continue
        if action == "exclude":
            updated.append(replace(item, result=STATUS_AUDIT_EXCLUDED))
            audit_entries.append(
                {
                    **base_entry,
                    "Decisión": "Excluir",
                    "Estado final": item.reported_state,
                    "Detalle": normalize_text(decision.get("detail"))
                    or "Fila excluida por decisión del auditor.",
                }
            )
            continue
        if action != "correct":
            updated.append(item)
            continue

        target_state = canonicalize_state(
            decision.get("state") or item.pms_state or item.reported_state
        )
        secondary_state = canonicalize_flow20_secondary(
            decision.get("secondary_state")
            or item.reported_secondary_state
        )
        if item.sheet == SHEET_FLOW20 and target_state == FLOW20_ODM_STATE:
            if secondary_state not in FLOW20_SECONDARY_STATES:
                wb.close()
                raise BitacoraError(
                    "Para el estado 3-ODM Asignadas debes indicar si el dibujo "
                    "quedó pendiente o finalizado."
                )
            timing_fields = _manual_flow20_timing_fields(
                decision,
                target_state,
                secondary_state,
            )
        else:
            timing_fields = _manual_timing_fields(decision, target_state)
        ws = wb[item.sheet]
        headers = headers_for(item.sheet)

        ws.cell(item.excel_row, headers["ESTADO DE LA OT"]).value = target_state
        agp = normalize_identifier(decision.get("agp"))
        protocol = normalize_identifier(decision.get("protocol"))
        odm = normalize_identifier(decision.get("odm"))
        is_flow20_odm = item.sheet == SHEET_FLOW20 and target_state == FLOW20_ODM_STATE
        write_element_data = target_state in FINAL_ELEMENT_STATES or is_flow20_odm
        requires_quantity_validation = (
            target_state in FINAL_ELEMENT_STATES
            or is_flow20_odm
            and secondary_state == FLOW20_SECONDARY_COMPLETED
        )
        if write_element_data:
            if not agp or not protocol or (is_flow20_odm and not odm):
                wb.close()
                raise BitacoraError(
                    "Para corregir este caso debes indicar AGP, protocolo"
                    + (" y ODM." if is_flow20_odm else ".")
                )
            if has_invalid_agp_separator(agp):
                wb.close()
                raise BitacoraError(
                    "Cuando haya varios AGP debes separarlos únicamente con ;."
                )
            if has_invalid_protocol_separator(protocol):
                wb.close()
                raise BitacoraError(
                    "Cuando haya varios protocolos debes separarlos únicamente con ;."
                )
            if has_invalid_protocol_prefix(protocol):
                wb.close()
                raise BitacoraError(
                    "Todos los protocolos deben iniciar por 1000. Revisa el valor diligenciado."
                )
        if headers.get("CODIGO DEL INCREMENTO (AGP)") and write_element_data:
            ws.cell(
                item.excel_row,
                headers["CODIGO DEL INCREMENTO (AGP)"],
            ).value = agp or None
        if headers.get("PROTOCOLO") and write_element_data:
            ws.cell(item.excel_row, headers["PROTOCOLO"]).value = protocol or None
        if is_flow20_odm:
            secondary_column = headers.get("ESTADO SECUNDARIO")
            odm_column = headers.get("ODM")
            if not secondary_column or not odm_column:
                wb.close()
                raise BitacoraError(
                    "La hoja Flujo 2.0 no contiene ESTADO SECUNDARIO y ODM."
                )
            ws.cell(item.excel_row, secondary_column).value = secondary_state
            ws.cell(item.excel_row, odm_column).value = odm

        observation = ""
        if target_state in REJECTION_STATES:
            observation = matching_rejection_comment(item.qr, target_state)
        if headers.get("OBSERVACIONES") and target_state in REJECTION_STATES:
            ws.cell(item.excel_row, headers["OBSERVACIONES"]).value = observation or None

        reported_pel = item.reported_pel
        reported_tratas = item.reported_tratas
        if target_state in FINAL_ELEMENT_STATES | REJECTION_STATES or is_flow20_odm:
            if headers.get("CANTIDAD PEL"):
                ws.cell(item.excel_row, headers["CANTIDAD PEL"]).value = None
            if headers.get("CANTIDAD TRATAS"):
                ws.cell(item.excel_row, headers["CANTIDAD TRATAS"]).value = None
            reported_pel = None
            reported_tratas = None

        updated.append(
            replace(
                item,
                reported_state=target_state,
                result=STATUS_AUDIT_CORRECTED,
                reported_secondary_state=(
                    secondary_state if is_flow20_odm else ""
                ),
                agp=agp if write_element_data else item.agp,
                odm=odm if is_flow20_odm else item.odm,
                protocol=protocol if write_element_data else item.protocol,
                reported_pel=reported_pel,
                reported_tratas=reported_tratas,
                export_pel=None,
                export_tratas=None,
                count_result=(
                    COUNT_NOT_EVALUATED
                    if requires_quantity_validation
                    else COUNT_NOT_APPLICABLE
                ),
                pms_observation=observation,
                observation_result=(
                    OBSERVATION_UPDATED
                    if observation
                    else OBSERVATION_NOT_APPLICABLE
                ),
                flow20_evidence=(
                    "Estado, subestado y fechas confirmados por auditoría"
                    if item.sheet == SHEET_FLOW20
                    else item.flow20_evidence
                ),
                **timing_fields,
            )
        )
        audit_entries.append(
            {
                **base_entry,
                "Decisión": "Corregir",
                "Estado final": target_state,
                "Estado secundario": secondary_state if is_flow20_odm else "",
                "AGP": agp if write_element_data else "",
                "Protocolo": protocol if write_element_data else "",
                "ODM": odm if is_flow20_odm else "",
                "Detalle": (
                    "Devolución recuperada del PMS."
                    if observation
                    else "Estado y datos actualizados durante la auditoría."
                ),
            }
        )

    if timelines is not None and production_date is not None:
        updated = apply_drawing_timings(updated, timelines, production_date)

    output = BytesIO()
    wb.save(output)
    wb.close()
    return output.getvalue(), updated, audit_entries


def add_unreported_audit_rows(
    workbook_bytes: bytes,
    validations: list[ValidationRow],
    cases: list[UnreportedCase],
    decisions: dict[str, dict[str, Any]],
) -> tuple[bytes, list[ValidationRow], list[dict[str, Any]]]:
    """Add selected omitted cases to their operational sheet and record every choice."""
    try:
        wb = load_workbook(BytesIO(workbook_bytes), data_only=False)
    except Exception as exc:
        raise BitacoraError("No pude agregar los casos no reportados.") from exc

    updated = list(validations)
    audit_entries: list[dict[str, Any]] = []
    for case in cases:
        decision = decisions.get(case.qr)
        if not decision:
            continue
        action = normalize_text(decision.get("action")).lower()
        if action == "ignore":
            audit_entries.append(
                {
                    "Etapa": "Casos no reportados",
                    "Hoja/Flujo": case.flow,
                    "QR/Referencia": case.qr,
                    "Decisión": "No agregar",
                    "Estado inicial": "Ausente",
                    "Estado final": case.state,
                    "AGP": "",
                    "Protocolo": "",
                    "Detalle": normalize_text(decision.get("detail"))
                    or "Caso descartado por el auditor.",
                }
            )
            continue
        if action != "add":
            continue

        sheet_name = normalize_text(decision.get("sheet")) or case.flow
        if sheet_name not in {SHEET_NEW, SHEET_OLD}:
            sheet_name = case.flow
        ws = wb[sheet_name]
        header_row = _detect_header_row(ws, {"QR", "ESTADO DE LA OT"})
        headers = _header_map(ws, header_row=header_row)
        target_row = _last_data_row(ws, "QR", header_row=header_row) + 1
        agp = normalize_identifier(decision.get("agp"))
        protocol = normalize_identifier(decision.get("protocol"))
        if case.state in FINAL_ELEMENT_STATES and (not agp or not protocol):
            wb.close()
            raise BitacoraError(
                "Para agregar un caso de dibujo finalizado debes indicar AGP y protocolo."
            )
        if has_invalid_agp_separator(agp):
            wb.close()
            raise BitacoraError(
                "Cuando haya varios AGP debes separarlos únicamente con ;."
            )
        if has_invalid_protocol_separator(protocol):
            wb.close()
            raise BitacoraError(
                "Cuando haya varios protocolos debes separarlos únicamente con ;."
            )
        if protocol and has_invalid_protocol_prefix(protocol):
            wb.close()
            raise BitacoraError(
                "Todos los protocolos deben iniciar por 1000. Revisa el valor diligenciado."
            )
        user = normalize_text(decision.get("user")) or case.actor
        origin = normalize_text(decision.get("origin"))
        elements = decision.get("elements")
        values = {
            "ORIGEN DE ACTUALIZACION": origin,
            "QR": case.qr,
            "ELEMENTOS ESTIMADOS": elements,
            "USUARIO": user,
            "ESTADO DE LA OT": case.state,
            "CODIGO DEL INCREMENTO (AGP)": agp,
            "PROTOCOLO": protocol,
            "OBSERVACIONES": case.comment if case.state in REJECTION_STATES else "",
            "CANTIDAD PEL": None,
            "CANTIDAD TRATAS": None,
        }
        for header, value in values.items():
            column = headers.get(header)
            if column:
                ws.cell(target_row, column).value = value or None

        updated.append(
            ValidationRow(
                sheet=sheet_name,
                excel_row=target_row,
                qr=case.qr,
                reported_state=case.state,
                pms_state=case.state,
                pms_datetime=case.occurred_at,
                previous_state=case.state,
                previous_datetime=case.occurred_at,
                result=STATUS_AUDIT_CORRECTED,
                professional=user,
                pms_actor=case.actor,
                protocol=protocol,
                count_result=COUNT_NOT_EVALUATED,
                pms_observation=case.comment,
                observation_result=(
                    OBSERVATION_UPDATED
                    if case.comment
                    else OBSERVATION_NOT_APPLICABLE
                ),
            )
        )
        audit_entries.append(
            {
                "Etapa": "Casos no reportados",
                "Hoja/Flujo": sheet_name,
                "QR/Referencia": case.qr,
                "Decisión": "Agregar",
                "Estado inicial": "Ausente",
                "Estado final": case.state,
                "AGP": agp,
                "Protocolo": protocol,
                "Detalle": "Fila incorporada a la bitácora durante la auditoría.",
            }
        )

    output = BytesIO()
    wb.save(output)
    wb.close()
    return output.getvalue(), updated, audit_entries


def apply_protocol_counts_to_validations(
    validations: list[ValidationRow],
    protocol_counts: dict[str, ProtocolElementCounts],
) -> list[ValidationRow]:
    updated: list[ValidationRow] = []
    for item in validations:
        if item.result not in UPLOAD_APPROVED_RESULTS:
            updated.append(
                replace(
                    item,
                    export_pel=None,
                    export_tratas=None,
                    export_pel_breakdown=None,
                    export_tratas_breakdown=None,
                    count_result=COUNT_NOT_APPLICABLE,
                )
            )
            continue
        if not validation_requires_element_counts(item):
            updated.append(
                replace(
                    item,
                    export_pel=None,
                    export_tratas=None,
                    export_pel_breakdown=None,
                    export_tratas_breakdown=None,
                    count_result=COUNT_NOT_APPLICABLE,
                )
            )
            continue
        export_pel, export_tratas, count_result = compare_protocol_quantities(
            item.protocol,
            item.reported_pel,
            item.reported_tratas,
            protocol_counts,
        )
        export_pel_breakdown, export_tratas_breakdown = protocol_quantity_breakdown(
            item.protocol,
            protocol_counts,
        )
        updated.append(
            replace(
                item,
                export_pel=export_pel,
                export_tratas=export_tratas,
                export_pel_breakdown=export_pel_breakdown,
                export_tratas_breakdown=export_tratas_breakdown,
                count_result=count_result,
            )
        )
    return updated


def apply_protocol_audit_results(
    validations: list[ValidationRow],
    email_results: list[EmailQuantityResult],
    decisions: dict[str, dict[str, Any]],
    no_apply_protocols: set[str],
    protocol_counts: dict[str, ProtocolElementCounts],
) -> tuple[list[ValidationRow], list[EmailQuantityResult]]:
    """Turn justified protocol gaps into resolved audit outcomes."""
    no_apply_positions = {
        (
            normalize_text(decision.get("sheet")),
            int(decision.get("excel_row") or 0),
        )
        for decision in decisions.values()
        if normalize_text(decision.get("action")).lower() == "no_apply"
    }

    def audited_result(
        sheet_name: str,
        row_number: int,
        protocol: str,
        current_result: str,
    ) -> str:
        if (sheet_name, row_number) in no_apply_positions:
            return COUNT_AUDIT_NOT_APPLICABLE
        justified = [
            item
            for item in split_protocols(protocol)
            if item in no_apply_protocols
        ]
        if justified:
            return f"{COUNT_AUDIT_NOT_APPLICABLE}: {', '.join(justified)}"
        missing = [
            item
            for item in split_protocols(protocol)
            if item not in protocol_counts
        ]
        if missing and set(missing).issubset(no_apply_protocols):
            return f"{COUNT_AUDIT_NOT_APPLICABLE}: {', '.join(missing)}"
        return current_result

    audited_validations = [
        replace(
            item,
            count_result=audited_result(
                item.sheet,
                item.excel_row,
                item.protocol,
                item.count_result,
            ),
        )
        for item in validations
    ]
    audited_emails = [
        replace(
            item,
            count_result=audited_result(
                "Correo-Rotulacion",
                item.excel_row,
                item.protocol,
                item.count_result,
            ),
        )
        for item in email_results
    ]
    return audited_validations, audited_emails


def apply_protocol_audit_decisions(
    workbook_bytes: bytes,
    validations: list[ValidationRow],
    decisions: dict[str, dict[str, Any]],
) -> tuple[bytes, list[ValidationRow], list[dict[str, Any]]]:
    """Fill blank protocols once for flow or email rows, or record No aplica."""
    try:
        wb = load_workbook(BytesIO(workbook_bytes), data_only=False)
    except Exception as exc:
        raise BitacoraError("No pude aplicar la prevalidación de protocolos.") from exc

    updated_by_position = {
        (item.sheet, item.excel_row): item for item in validations
    }
    audit_entries: list[dict[str, Any]] = []
    header_cache: dict[str, dict[str, int]] = {}

    for decision in decisions.values():
        sheet_name = normalize_text(decision.get("sheet"))
        row_number = int(decision.get("excel_row") or 0)
        reference = normalize_text(decision.get("reference"))
        current_protocol = normalize_identifier(decision.get("current_protocol"))
        action = normalize_text(decision.get("action")).lower()
        protocol = normalize_identifier(decision.get("protocol"))
        if sheet_name not in wb.sheetnames or row_number < 1:
            continue

        if action == "set" and protocol:
            if has_invalid_protocol_separator(protocol):
                wb.close()
                raise BitacoraError(
                    "Cuando haya varios protocolos debes separarlos únicamente con ;."
                )
            if has_invalid_protocol_prefix(protocol):
                wb.close()
                raise BitacoraError(
                    "Todos los protocolos deben iniciar por 1000. Revisa el valor diligenciado."
                )
            ws = wb[sheet_name]
            headers = header_cache.get(sheet_name)
            if headers is None:
                key_header = "ASUNTO" if sheet_name == "Correo-Rotulacion" else "QR"
                header_row = _detect_header_row(ws, {key_header, "PROTOCOLO"})
                headers = _header_map(ws, header_row=header_row)
                header_cache[sheet_name] = headers
            ws.cell(row_number, headers["PROTOCOLO"]).value = protocol
            position = (sheet_name, row_number)
            if position in updated_by_position:
                updated_by_position[position] = replace(
                    updated_by_position[position],
                    protocol=protocol,
                )
            decision_label = "Protocolo diligenciado"
            detail = "Protocolo incorporado durante la prevalidación."
        elif action == "no_apply":
            decision_label = "No aplica"
            detail = normalize_text(decision.get("detail")) or (
                "El auditor confirmó que esta referencia no requiere protocolo."
            )
        else:
            continue

        audit_entries.append(
            {
                "Etapa": "Protocolos",
                "Hoja/Flujo": sheet_name,
                "QR/Referencia": reference,
                "Decisión": decision_label,
                "Estado inicial": current_protocol or "Sin protocolo",
                "Estado final": "Revisado",
                "AGP": "",
                "Protocolo": protocol,
                "Detalle": detail,
            }
        )

    output = BytesIO()
    wb.save(output)
    wb.close()
    updated = [
        updated_by_position[(item.sheet, item.excel_row)] for item in validations
    ]
    return output.getvalue(), updated, audit_entries


def replace_protocol_references(
    workbook_bytes: bytes,
    validations: list[ValidationRow],
    replacements: dict[str, str],
) -> tuple[bytes, list[ValidationRow], list[dict[str, Any]]]:
    """Replace exact protocol tokens in every eligible flow and email reference."""
    normalized_replacements: dict[str, str] = {}
    for current_value, replacement_value in replacements.items():
        current = normalize_identifier(current_value)
        replacement = normalize_identifier(replacement_value)
        if not current or not replacement:
            raise BitacoraError(
                "Indica el protocolo mal diligenciado y su valor correcto."
            )
        if has_invalid_protocol_separator(replacement):
            raise BitacoraError(
                "Cuando la corrección tenga varios protocolos debes separarlos únicamente con ;."
            )
        if has_invalid_protocol_prefix(replacement):
            raise BitacoraError(
                "Todos los protocolos corregidos deben iniciar por 1000."
            )
        if current == replacement:
            raise BitacoraError(
                "El protocolo correcto debe ser diferente del valor mal diligenciado."
            )
        normalized_replacements[current] = replacement

    def replace_tokens(value: Any) -> tuple[str, bool]:
        original = split_protocols(value)
        if not original:
            return normalize_identifier(value), False
        corrected: list[str] = []
        changed = False
        for protocol in original:
            replacement = normalized_replacements.get(protocol)
            tokens = split_protocols(replacement) if replacement is not None else [protocol]
            if replacement is not None:
                changed = True
            for token in tokens:
                if token not in corrected:
                    corrected.append(token)
        return ";".join(corrected), changed

    email_rows = evaluate_email_quantities(workbook_bytes, {})
    try:
        wb = load_workbook(BytesIO(workbook_bytes), data_only=False)
    except Exception as exc:
        raise BitacoraError("No pude corregir los protocolos en la bitácora.") from exc

    header_cache: dict[str, dict[str, int]] = {}

    def headers_for(sheet_name: str) -> dict[str, int]:
        headers = header_cache.get(sheet_name)
        if headers is None:
            ws = wb[sheet_name]
            key_header = "ASUNTO" if sheet_name == "Correo-Rotulacion" else "QR"
            header_row = _detect_header_row(ws, {key_header, "PROTOCOLO"})
            headers = _header_map(ws, header_row=header_row)
            header_cache[sheet_name] = headers
        return headers

    audit_entries: list[dict[str, Any]] = []
    updated: list[ValidationRow] = []
    for item in validations:
        corrected, changed = replace_tokens(item.protocol)
        if not changed:
            updated.append(item)
            continue
        headers = headers_for(item.sheet)
        wb[item.sheet].cell(
            item.excel_row,
            headers["PROTOCOLO"],
        ).value = corrected
        updated.append(
            replace(
                item,
                protocol=corrected,
                export_pel=None,
                export_tratas=None,
                count_result=COUNT_NOT_EVALUATED,
            )
        )
        audit_entries.append(
            {
                "Etapa": "Protocolos",
                "Hoja/Flujo": item.sheet,
                "QR/Referencia": item.qr,
                "Decisión": "Corregir protocolo",
                "Estado inicial": item.protocol,
                "Estado final": corrected,
                "AGP": "",
                "Protocolo": corrected,
                "Detalle": "Protocolo reemplazado en todas sus referencias y reenviado al cruce con SAP.",
            }
        )

    for item in email_rows:
        corrected, changed = replace_tokens(item.protocol)
        if not changed:
            continue
        headers = headers_for("Correo-Rotulacion")
        wb["Correo-Rotulacion"].cell(
            item.excel_row,
            headers["PROTOCOLO"],
        ).value = corrected
        audit_entries.append(
            {
                "Etapa": "Protocolos",
                "Hoja/Flujo": "Correo-Rotulacion",
                "QR/Referencia": item.subject,
                "Decisión": "Corregir protocolo",
                "Estado inicial": item.protocol,
                "Estado final": corrected,
                "AGP": "",
                "Protocolo": corrected,
                "Detalle": "Protocolo reemplazado en todas sus referencias y reenviado al cruce con SAP.",
            }
        )

    if not audit_entries:
        wb.close()
        raise BitacoraError(
            "El protocolo seleccionado ya no aparece en las referencias pendientes."
        )

    output = BytesIO()
    wb.save(output)
    wb.close()
    return output.getvalue(), updated, audit_entries


def validation_frame(validations: list[ValidationRow]) -> pd.DataFrame:
    return pd.DataFrame([item.as_dict() for item in validations])


def email_quantity_frame(results: list[EmailQuantityResult]) -> pd.DataFrame:
    columns = [
        "Hoja",
        "Fila Excel",
        "QR",
        "Estado reportado",
        "Protocolo",
        "Cantidad PEL reportada",
        "Cantidad TRATAS reportada",
        "Cantidad PEL EXPORT",
        "Cantidad TRATAS EXPORT",
        "Resultado cantidades",
    ]
    return pd.DataFrame([item.as_dict() for item in results], columns=columns)


def unreported_cases_frame(cases: list[UnreportedCase]) -> pd.DataFrame:
    columns = [
        "Flujo",
        "QR",
        "Estado PMS del día",
        "Fecha/hora PMS",
        "Usuario PMS",
        "Observación PMS",
        "Hallazgo",
    ]
    return pd.DataFrame([item.as_dict() for item in cases], columns=columns)


def _table_style(ws, table_name: str) -> TableStyleInfo:
    if table_name in ws.tables and ws.tables[table_name].tableStyleInfo is not None:
        return copy(ws.tables[table_name].tableStyleInfo)
    return TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )


def _copy_cell_format(source, target) -> None:
    if source.has_style:
        target._style = copy(source._style)
    if source.number_format:
        target.number_format = source.number_format
    target.font = copy(source.font)
    target.fill = copy(source.fill)
    target.border = copy(source.border)
    target.alignment = copy(source.alignment)
    target.protection = copy(source.protection)


def _rebuild_table(
    ws,
    table_name: str,
    last_column: int,
    last_row: int,
    *,
    header_row: int = 1,
) -> None:
    style = _table_style(ws, table_name)
    if table_name in ws.tables:
        del ws.tables[table_name]
    if last_row < header_row or last_column < 1:
        return
    from openpyxl.utils import get_column_letter

    table = Table(
        displayName=table_name,
        ref=f"A{header_row}:{get_column_letter(last_column)}{last_row}",
    )
    table.tableStyleInfo = style
    ws.add_table(table)


def build_protocol_inventory_workbook(
    records: Iterable[dict[str, Any]],
) -> bytes:
    """Build the SAP protocol request as an Excel table plus an audit detail."""
    detail_rows: list[dict[str, str]] = []
    for record in records:
        protocol = normalize_identifier(record.get("Protocolo"))
        if not protocol:
            continue
        if has_invalid_protocol_separator(protocol):
            raise BitacoraError(
                "El listado contiene protocolos sin separar correctamente con ;."
            )
        if has_invalid_protocol_prefix(protocol):
            raise BitacoraError(
                "El listado contiene protocolos que no inician por 1000."
            )
        detail_rows.append(
            {
                "Protocolo": protocol,
                "Hoja": normalize_text(record.get("Hoja")),
                "QR o referencia": normalize_text(record.get("QR o referencia")),
                "Estado": normalize_text(record.get("Estado")),
            }
        )
    detail_rows = list(
        {
            (
                row["Protocolo"],
                row["Hoja"],
                row["QR o referencia"],
                row["Estado"],
            ): row
            for row in detail_rows
        }.values()
    )
    detail_rows.sort(
        key=lambda row: (
            row["Protocolo"],
            row["Hoja"],
            row["QR o referencia"],
        )
    )
    unique_protocols = sorted({row["Protocolo"] for row in detail_rows})

    wb = Workbook()
    ws = wb.active
    ws.title = "Protocolos"
    ws.append(["Protocolo"])
    for protocol in unique_protocols:
        ws.append([protocol])
        ws.cell(ws.max_row, 1).number_format = "@"
        ws.cell(ws.max_row, 1).alignment = Alignment(horizontal="center")

    detail_ws = wb.create_sheet("Detalle")
    detail_headers = ["Protocolo", "Hoja", "QR o referencia", "Estado"]
    detail_ws.append(detail_headers)
    for row in detail_rows:
        detail_ws.append([row[header] for header in detail_headers])
        detail_ws.cell(detail_ws.max_row, 1).number_format = "@"
        detail_ws.cell(detail_ws.max_row, 1).alignment = Alignment(
            horizontal="center"
        )

    header_fill = PatternFill("solid", fgColor="0F766E")
    for sheet, widths in (
        (ws, {"A": 24}),
        (detail_ws, {"A": 24, "B": 22, "C": 48, "D": 28}),
    ):
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center", vertical="center")
        sheet.row_dimensions[1].height = 26
        for column, width in widths.items():
            sheet.column_dimensions[column].width = width
        sheet.freeze_panes = "A2"
        sheet.sheet_view.showGridLines = False

    if unique_protocols:
        _rebuild_table(ws, "ListadoProtocolos", 1, ws.max_row)
    if detail_rows:
        _rebuild_table(detail_ws, "DetalleProtocolos", 4, detail_ws.max_row)

    output = BytesIO()
    wb.save(output)
    wb.close()
    return output.getvalue()


def _delete_columns_by_header(
    ws,
    normalized_headers: set[str],
    *,
    header_row: int = 1,
) -> None:
    indexes = [
        cell.column
        for cell in ws[header_row]
        if normalize_header(cell.value) in normalized_headers
    ]
    for index in sorted(indexes, reverse=True):
        ws.delete_cols(index)


def _last_data_row(ws, key_header: str, *, header_row: int = 1) -> int:
    headers = _header_map(ws, header_row=header_row)
    key_col = headers.get(normalize_header(key_header), 1)
    cells = getattr(ws, "_cells", None)
    if isinstance(cells, dict):
        populated_rows = [
            row
            for (row, column), cell in cells.items()
            if column == key_col
            and row > header_row
            and normalize_text(cell.value)
        ]
        return max(populated_rows, default=header_row)
    for row in range(ws.max_row, header_row, -1):
        if normalize_text(ws.cell(row, key_col).value):
            return row
    return header_row


def _set_total_formulas(ws, last_row: int, *, header_row: int = 1) -> None:
    headers = _header_map(ws, header_row=header_row)
    pel_col = headers.get("CANTIDAD PEL")
    tratas_col = headers.get("CANTIDAD TRATAS")
    total_col = headers.get("TOTAL ELEMENTOS")
    if not (pel_col and tratas_col and total_col):
        return
    from openpyxl.utils import get_column_letter

    pel_letter = get_column_letter(pel_col)
    tratas_letter = get_column_letter(tratas_col)
    for row in range(header_row + 1, last_row + 1):
        pel_value = ws.cell(row, pel_col).value
        tratas_value = ws.cell(row, tratas_col).value
        if ";" in normalize_text(pel_value) or ";" in normalize_text(tratas_value):
            tokens = [
                token.strip()
                for value in (pel_value, tratas_value)
                for token in normalize_text(value).split(";")
                if token.strip()
            ]
            try:
                ws.cell(row, total_col).value = sum(int(token) for token in tokens)
            except ValueError:
                ws.cell(row, total_col).value = None
            continue
        ws.cell(row, total_col).value = (
            f'=IF(COUNTA({pel_letter}{row}:{tratas_letter}{row})=0,"",'
            f"SUM({pel_letter}{row}:{tratas_letter}{row}))"
        )


def _remove_unvalidated_rows(ws, validations: list[ValidationRow]) -> None:
    invalid_rows = [
        item.excel_row
        for item in validations
        if item.result not in UPLOAD_APPROVED_RESULTS
    ]
    for row_number in sorted(invalid_rows, reverse=True):
        ws.delete_rows(row_number)


def _try_detect_header_row(ws, required_headers: set[str]) -> int | None:
    try:
        return _detect_header_row(ws, required_headers)
    except BitacoraError:
        return None


def _ensure_named_timing_columns(
    ws,
    timing_headers: tuple[str, ...],
    *,
    header_row: int,
) -> dict[str, int]:
    headers = _header_map(ws, header_row=header_row)
    style_column = headers.get("TOTAL ELEMENTOS") or max(headers.values(), default=1)
    for header in timing_headers:
        normalized = normalize_header(header)
        if normalized in headers:
            continue
        target_column = max(headers.values(), default=0) + 1
        _copy_cell_format(
            ws.cell(header_row, style_column),
            ws.cell(header_row, target_column),
        )
        ws.cell(header_row, target_column).value = header
        for row_number in range(header_row + 1, max(ws.max_row, header_row + 1) + 1):
            _copy_cell_format(
                ws.cell(row_number, style_column),
                ws.cell(row_number, target_column),
            )
            ws.cell(row_number, target_column).number_format = EXCEL_DATETIME_FORMAT
        from openpyxl.utils import get_column_letter

        ws.column_dimensions[get_column_letter(target_column)].width = 22
        headers[normalized] = target_column
    return headers


def _ensure_timing_columns(ws, *, header_row: int) -> dict[str, int]:
    return _ensure_named_timing_columns(
        ws,
        TIMING_HEADERS,
        header_row=header_row,
    )


def _ensure_flow20_timing_columns(ws, *, header_row: int) -> dict[str, int]:
    return _ensure_named_timing_columns(
        ws,
        FLOW20_TIMING_HEADERS,
        header_row=header_row,
    )


def _extend_coded_header_to_last_column(ws, last_column: int) -> None:
    """Move the F-AGP-04 code box so it still closes the widened table."""
    code_cell = next(
        (
            cell
            for row in ws.iter_rows(min_row=1, max_row=4)
            for cell in row
            if normalize_header(cell.value).startswith("F-AGP-04")
        ),
        None,
    )
    if code_cell is None:
        return
    code_start = code_cell.column
    code_end = code_start
    top_title_end = code_start - 1
    bottom_title_end = code_start - 1
    code_ranges: dict[int, Any] = {}
    header_ranges = []
    for merged in list(ws.merged_cells.ranges):
        if merged.max_row > 4 or merged.min_col < 3:
            continue
        header_ranges.append(merged)
        if merged.min_col == code_start:
            code_end = max(code_end, merged.max_col)
            code_ranges[merged.min_row] = merged
        elif merged.min_row == 1:
            top_title_end = merged.max_col
        elif merged.min_row == 3:
            bottom_title_end = merged.max_col
    if code_end >= last_column or last_column < 4:
        return

    code_values = {
        row: ws.cell(row, code_start).value
        for row in range(1, 5)
    }
    title_values = {1: ws.cell(1, 3).value, 3: ws.cell(3, 3).value}
    format_sources = {
        "top_anchor": copy(ws.cell(1, 3)),
        "top_end": copy(ws.cell(2, top_title_end)),
        "bottom_anchor": copy(ws.cell(3, 3)),
        "bottom_end": copy(ws.cell(4, bottom_title_end)),
        **{
            f"code_{row}_start": copy(ws.cell(row, code_start))
            for row in range(1, 5)
        },
        **{
            f"code_{row}_end": copy(ws.cell(row, code_end))
            for row in range(1, 5)
        },
    }
    for merged in header_ranges:
        ws.unmerge_cells(str(merged))
    for row in range(1, 5):
        for column in range(3, last_column + 1):
            ws.cell(row, column).value = None

    new_code_start = last_column - 1
    _copy_cell_format(format_sources["top_anchor"], ws.cell(1, 3))
    _copy_cell_format(format_sources["top_end"], ws.cell(2, new_code_start - 1))
    _copy_cell_format(format_sources["bottom_anchor"], ws.cell(3, 3))
    _copy_cell_format(format_sources["bottom_end"], ws.cell(4, new_code_start - 1))
    ws.cell(1, 3).value = title_values[1]
    ws.cell(3, 3).value = title_values[3]
    ws.merge_cells(start_row=1, start_column=3, end_row=2, end_column=new_code_start - 1)
    ws.merge_cells(start_row=3, start_column=3, end_row=4, end_column=new_code_start - 1)

    for row in range(1, 5):
        _copy_cell_format(
            format_sources[f"code_{row}_start"],
            ws.cell(row, new_code_start),
        )
        _copy_cell_format(
            format_sources[f"code_{row}_end"],
            ws.cell(row, last_column),
        )
        ws.cell(row, new_code_start).value = code_values[row]
        ws.merge_cells(
            start_row=row,
            start_column=new_code_start,
            end_row=row,
            end_column=last_column,
        )


def _clean_operational_sheets(
    wb,
    protected_cells: set[tuple[str, int, int]] | None = None,
) -> None:
    protected = protected_cells or set()
    for sheet_name, key_header in OPERATIONAL_SHEET_KEYS.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        header_row = _try_detect_header_row(ws, {key_header})
        if header_row is None:
            continue
        headers = _header_map(ws, header_row=header_row)
        observation_col = headers.get("OBSERVACIONES")
        last_data_row = _last_data_row(ws, key_header, header_row=header_row)
        last_data_column = max(headers.values(), default=ws.max_column)
        for row_number in range(header_row + 1, last_data_row + 1):
            for column in range(1, last_data_column + 1):
                cell = ws.cell(row_number, column)
                if (sheet_name, row_number, column) in protected:
                    continue
                cleaned = clean_operational_noise(cell.value)
                if sheet_name == SHEET_PNC and isinstance(cleaned, str):
                    cleaned = clean_placeholder_observation(cleaned) or None
                if cleaned != cell.value:
                    cell.value = cleaned
            if observation_col:
                cell = ws.cell(row_number, observation_col)
                cleaned = clean_placeholder_observation(cell.value)
                if cleaned != normalize_text(cell.value):
                    cell.value = cleaned or None


def _apply_cleaning_and_enrichment(
    wb,
    validations: list[ValidationRow],
    email_quantities: list[EmailQuantityResult],
) -> None:
    protected_cells: set[tuple[str, int, int]] = set()
    by_sheet: dict[str, list[ValidationRow]] = {}
    for item in validations:
        by_sheet.setdefault(item.sheet, []).append(item)

    for sheet_name, items in by_sheet.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        header_row = _try_detect_header_row(
            ws,
            {"QR", "ESTADO DE LA OT", "OBSERVACIONES"},
        )
        if header_row is None:
            continue
        headers = (
            _ensure_flow20_timing_columns(ws, header_row=header_row)
            if sheet_name == SHEET_FLOW20
            else _ensure_timing_columns(ws, header_row=header_row)
        )
        pel_col = headers.get("CANTIDAD PEL")
        tratas_col = headers.get("CANTIDAD TRATAS")
        total_col = headers.get("TOTAL ELEMENTOS")
        observation_col = headers.get("OBSERVACIONES")
        arrival_col = headers.get(normalize_header(ARRIVAL_DATETIME_HEADER))
        exit_col = headers.get(normalize_header(EXIT_DATETIME_HEADER))
        odm_datetime_col = headers.get(normalize_header(ODM_DATETIME_HEADER))
        drawing_datetime_col = headers.get(
            normalize_header(DRAWING_DATETIME_HEADER)
        )
        for item in items:
            if (
                (
                    item.count_result in {COUNT_CONFIRMED, COUNT_UPDATED}
                    or str(item.count_result).startswith(COUNT_AUDIT_NOT_APPLICABLE)
                )
                and item.export_pel is not None
                and item.export_tratas is not None
            ):
                if pel_col:
                    ws.cell(item.excel_row, pel_col).value = (
                        item.export_pel_breakdown
                        if item.export_pel_breakdown is not None
                        else item.export_pel
                    )
                    protected_cells.add((sheet_name, item.excel_row, pel_col))
                if tratas_col:
                    ws.cell(item.excel_row, tratas_col).value = (
                        item.export_tratas_breakdown
                        if item.export_tratas_breakdown is not None
                        else item.export_tratas
                    )
                    protected_cells.add((sheet_name, item.excel_row, tratas_col))
                if total_col:
                    ws.cell(item.excel_row, total_col).value = (
                        item.export_pel + item.export_tratas
                    )
                    protected_cells.add((sheet_name, item.excel_row, total_col))
            if observation_col and item.pms_observation:
                ws.cell(item.excel_row, observation_col).value = item.pms_observation
            if arrival_col:
                arrival_cell = ws.cell(item.excel_row, arrival_col)
                arrival_cell.value = item.arrival_datetime
                arrival_cell.number_format = EXCEL_DATETIME_FORMAT
            if exit_col:
                exit_cell = ws.cell(item.excel_row, exit_col)
                exit_cell.value = item.exit_datetime
                exit_cell.number_format = EXCEL_DATETIME_FORMAT
            if odm_datetime_col:
                odm_cell = ws.cell(item.excel_row, odm_datetime_col)
                odm_cell.value = item.odm_datetime
                odm_cell.number_format = EXCEL_DATETIME_FORMAT
            if drawing_datetime_col:
                drawing_cell = ws.cell(item.excel_row, drawing_datetime_col)
                drawing_cell.value = item.drawing_datetime
                drawing_cell.number_format = EXCEL_DATETIME_FORMAT

    if email_quantities and "Correo-Rotulacion" in wb.sheetnames:
        ws = wb["Correo-Rotulacion"]
        header_row = _try_detect_header_row(ws, {"ORIGEN DE ACTUALIZACIÓN", "ASUNTO"})
        if header_row is not None:
            headers = _header_map(ws, header_row=header_row)
            pel_col = headers.get("CANTIDAD PEL")
            tratas_col = headers.get("CANTIDAD TRATAS")
            total_col = headers.get("TOTAL ELEMENTOS")
            for item in email_quantities:
                if (
                    (
                        item.count_result in {COUNT_CONFIRMED, COUNT_UPDATED}
                        or str(item.count_result).startswith(COUNT_AUDIT_NOT_APPLICABLE)
                    )
                    and item.export_pel is not None
                    and item.export_tratas is not None
                ):
                    if pel_col:
                        ws.cell(item.excel_row, pel_col).value = (
                            item.export_pel_breakdown
                            if item.export_pel_breakdown is not None
                            else item.export_pel
                        )
                        protected_cells.add(("Correo-Rotulacion", item.excel_row, pel_col))
                    if tratas_col:
                        ws.cell(item.excel_row, tratas_col).value = (
                            item.export_tratas_breakdown
                            if item.export_tratas_breakdown is not None
                            else item.export_tratas
                        )
                        protected_cells.add(("Correo-Rotulacion", item.excel_row, tratas_col))
                    if total_col:
                        ws.cell(item.excel_row, total_col).value = (
                            item.export_pel + item.export_tratas
                        )
                        protected_cells.add(("Correo-Rotulacion", item.excel_row, total_col))

    _clean_operational_sheets(wb, protected_cells)


def _add_validation_result_sheet(wb, validations: list[ValidationRow]) -> None:
    if SHEET_RESULT in wb.sheetnames:
        del wb[SHEET_RESULT]

    ws = wb.create_sheet(SHEET_RESULT)
    headers = [
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
    ws.append(headers)

    for item in validations:
        values = item.as_dict()
        ws.append([values[header] for header in headers])

    header_fill = PatternFill("solid", fgColor="1F4E78")
    match_fill = PatternFill("solid", fgColor="D9EAD3")
    mismatch_fill = PatternFill("solid", fgColor="F4CCCC")
    neutral_fill = PatternFill("solid", fgColor="E7E6E6")
    warning_fill = PatternFill("solid", fgColor="FFF2CC")

    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )
    ws.row_dimensions[1].height = 36

    header_columns = {
        normalize_header(cell.value): cell.column
        for cell in ws[1]
        if normalize_text(cell.value)
    }

    for row in range(2, ws.max_row + 1):
        result_cell = ws.cell(row, header_columns["RESULTADO"])
        if result_cell.value in UPLOAD_APPROVED_RESULTS:
            result_cell.fill = match_fill
        elif result_cell.value == STATUS_MISMATCH:
            result_cell.fill = mismatch_fill
        elif result_cell.value == STATUS_AUDIT_EXCLUDED or result_cell.value in FLOW20_DUPLICATE_RESULTS:
            result_cell.fill = neutral_fill
        else:
            result_cell.fill = warning_fill
        count_cell = ws.cell(row, header_columns["RESULTADO CANTIDADES"])
        if count_cell.value == COUNT_CONFIRMED:
            count_cell.fill = match_fill
        elif count_cell.value == COUNT_UPDATED:
            count_cell.fill = warning_fill
        elif count_cell.value in {
            COUNT_NO_PROTOCOL,
            COUNT_PROTOCOL_NOT_FOUND,
            COUNT_INVALID_PROTOCOL_SEPARATOR,
            COUNT_INVALID_PROTOCOL_PREFIX,
        } or str(count_cell.value).startswith(COUNT_PROTOCOLS_INCOMPLETE):
            count_cell.fill = mismatch_fill
        else:
            count_cell.fill = neutral_fill

        observation_cell = ws.cell(row, header_columns["ACCION OBSERVACION"])
        if observation_cell.value == OBSERVATION_UPDATED:
            observation_cell.fill = match_fill
        elif observation_cell.value == OBSERVATION_STATE_MISMATCH:
            observation_cell.fill = mismatch_fill
        elif observation_cell.value == OBSERVATION_NOT_FOUND:
            observation_cell.fill = warning_fill
        else:
            observation_cell.fill = neutral_fill
        for column in (
            header_columns["FECHA/HORA PMS"],
            header_columns["FECHA/HORA ANTERIOR"],
            header_columns["FECHA/HORA DE LLEGADA"],
            header_columns["FECHA/HORA DE SALIDA"],
            header_columns["FECHA/HORA ENTREGA ODM"],
            header_columns["FECHA/HORA ENTREGA DIBUJO"],
        ):
            ws.cell(row, column).number_format = "yyyy-mm-dd hh:mm:ss"
        for column in (
            header_columns["OBSERVACION PMS"],
            header_columns["ACCION OBSERVACION"],
        ):
            ws.cell(row, column).alignment = Alignment(vertical="top", wrap_text=True)
        if ws.cell(row, header_columns["OBSERVACION PMS"]).value:
            ws.row_dimensions[row].height = 72

    widths = {
        "A": 16,
        "B": 10,
        "C": 34,
        "D": 28,
        "E": 28,
        "F": 22,
        "G": 30,
        "H": 22,
        "I": 22,
        "J": 24,
        "K": 20,
        "L": 22,
        "M": 20,
        "N": 22,
        "O": 22,
        "P": 22,
        "Q": 22,
        "R": 30,
        "S": 18,
        "T": 20,
        "U": 28,
        "V": 60,
        "W": 34,
        "X": 18,
        "Y": 34,
        "Z": 18,
        "AA": 18,
        "AB": 18,
        "AC": 18,
        "AD": 36,
        "AE": 60,
        "AF": 30,
    }
    for column, width in widths.items():
        ws.column_dimensions[column].width = width

    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    _rebuild_table(ws, "ResultadoValidacion", len(headers), ws.max_row)


def _add_email_quantity_result_sheet(
    wb,
    results: list[EmailQuantityResult],
) -> None:
    if SHEET_EMAIL_RESULT in wb.sheetnames:
        del wb[SHEET_EMAIL_RESULT]

    ws = wb.create_sheet(SHEET_EMAIL_RESULT)
    headers = [
        "Fila Excel",
        "Origen de actualización",
        "Asunto",
        "Protocolo",
        "Cantidad PEL reportada",
        "Cantidad TRATAS reportada",
        "Cantidad PEL EXPORT",
        "Cantidad TRATAS EXPORT",
        "Resultado cantidades",
    ]
    ws.append(headers)
    for item in results:
        ws.append(
            [
                item.excel_row,
                item.origin,
                item.subject,
                item.protocol,
                item.reported_pel,
                item.reported_tratas,
                item.export_pel,
                item.export_tratas,
                item.count_result,
            ]
        )

    header_fill = PatternFill("solid", fgColor="1F4E78")
    match_fill = PatternFill("solid", fgColor="D9EAD3")
    warning_fill = PatternFill("solid", fgColor="FFF2CC")
    mismatch_fill = PatternFill("solid", fgColor="F4CCCC")
    neutral_fill = PatternFill("solid", fgColor="E7E6E6")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )
    ws.row_dimensions[1].height = 36

    for row in range(2, ws.max_row + 1):
        result_cell = ws.cell(row, 9)
        if result_cell.value == COUNT_CONFIRMED:
            result_cell.fill = match_fill
        elif result_cell.value == COUNT_UPDATED:
            result_cell.fill = warning_fill
        elif str(result_cell.value).startswith(COUNT_AUDIT_NOT_APPLICABLE):
            result_cell.fill = neutral_fill
        else:
            result_cell.fill = mismatch_fill
        ws.cell(row, 3).alignment = Alignment(vertical="top", wrap_text=True)
        ws.cell(row, 9).alignment = Alignment(vertical="top", wrap_text=True)
        ws.row_dimensions[row].height = 36

    widths = {
        "A": 12,
        "B": 24,
        "C": 65,
        "D": 30,
        "E": 22,
        "F": 24,
        "G": 20,
        "H": 22,
        "I": 42,
    }
    for column, width in widths.items():
        ws.column_dimensions[column].width = width

    if results:
        _rebuild_table(ws, "ValidacionCorreos", len(headers), ws.max_row)
    else:
        ws.merge_cells("A3:I3")
        ws["A3"] = (
            "No se encontraron filas con origen Solicitud por correo "
            "y estado Finalizado."
        )
        ws["A3"].font = Font(italic=True, color="548235")
        ws["A3"].alignment = Alignment(horizontal="center")

    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "1F4E78"


def _add_unreported_cases_sheet(wb, cases: list[UnreportedCase]) -> None:
    if SHEET_UNREPORTED in wb.sheetnames:
        del wb[SHEET_UNREPORTED]

    ws = wb.create_sheet(SHEET_UNREPORTED)
    headers = [
        "Flujo",
        "QR",
        "Estado PMS del día",
        "Fecha/hora PMS",
        "Usuario PMS",
        "Observación PMS",
        "Hallazgo",
    ]
    ws.append(headers)
    for case in cases:
        values = case.as_dict()
        ws.append([values[header] for header in headers])

    header_fill = PatternFill("solid", fgColor="C65911")
    finding_fill = PatternFill("solid", fgColor="FCE4D6")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )
    ws.row_dimensions[1].height = 34

    for row in range(2, ws.max_row + 1):
        for column in range(1, 5):
            ws.cell(row, column).alignment = Alignment(vertical="center")
        ws.cell(row, 4).number_format = "yyyy-mm-dd hh:mm:ss"
        ws.cell(row, 5).alignment = Alignment(vertical="top", wrap_text=True)
        ws.cell(row, 6).alignment = Alignment(vertical="top", wrap_text=True)
        ws.cell(row, 7).fill = finding_fill
        ws.cell(row, 7).alignment = Alignment(vertical="top", wrap_text=True)
        if ws.cell(row, 5).value or ws.cell(row, 6).value:
            ws.row_dimensions[row].height = 42
        else:
            ws.row_dimensions[row].height = 26

    widths = {
        "A": 16,
        "B": 38,
        "C": 30,
        "D": 22,
        "E": 42,
        "F": 60,
        "G": 52,
    }
    for column, width in widths.items():
        ws.column_dimensions[column].width = width

    if cases:
        _rebuild_table(ws, "CasosNoReportados", len(headers), ws.max_row)
    else:
        ws.merge_cells("A3:G3")
        ws["A3"] = "No se encontraron movimientos objetivo omitidos en esta fecha."
        ws["A3"].font = Font(italic=True, color="548235")
        ws["A3"].alignment = Alignment(horizontal="center")

    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "C65911"


def _add_guided_audit_sheet(
    wb,
    entries: list[dict[str, Any]],
) -> None:
    if SHEET_GUIDED_AUDIT in wb.sheetnames:
        del wb[SHEET_GUIDED_AUDIT]

    headers = [
        "Etapa",
        "Hoja/Flujo",
        "QR/Referencia",
        "Decisión",
        "Estado inicial",
        "Estado final",
        "Estado secundario",
        "AGP",
        "Protocolo",
        "ODM",
        "Fecha/hora de llegada",
        "Fecha/hora de salida",
        "Fecha/hora entrega ODM",
        "Fecha/hora entrega dibujo",
        "Detalle",
    ]
    ws = wb.create_sheet(SHEET_GUIDED_AUDIT)
    ws.append(headers)
    for entry in entries:
        ws.append([entry.get(header, "") for header in headers])

    header_fill = PatternFill("solid", fgColor="0F766E")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )
    ws.row_dimensions[1].height = 34

    for row in range(2, ws.max_row + 1):
        for column in range(1, len(headers) + 1):
            ws.cell(row, column).alignment = Alignment(vertical="top", wrap_text=True)
        for column in (11, 12, 13, 14):
            if ws.cell(row, column).value:
                parsed = parse_audit_datetime(ws.cell(row, column).value)
                ws.cell(row, column).value = parsed
                ws.cell(row, column).number_format = EXCEL_DATETIME_FORMAT
        decision = normalize_header(ws.cell(row, 4).value)
        if decision in {"AGREGAR", "CORREGIR", "CONSERVAR BITACORA"}:
            ws.cell(row, 4).fill = PatternFill("solid", fgColor="D9EAD3")
        elif decision in {"EXCLUIR", "NO AGREGAR"} or decision.startswith(
            "EXCLUIR DUPLICADO"
        ):
            ws.cell(row, 4).fill = PatternFill("solid", fgColor="E7E6E6")
        elif decision == "NO APLICA":
            ws.cell(row, 4).fill = PatternFill("solid", fgColor="FFF2CC")
        ws.row_dimensions[row].height = 34

    widths = {
        "A": 24,
        "B": 18,
        "C": 38,
        "D": 24,
        "E": 28,
        "F": 28,
        "G": 34,
        "H": 22,
        "I": 34,
        "J": 22,
        "K": 22,
        "L": 22,
        "M": 22,
        "N": 22,
        "O": 62,
    }
    for column, width in widths.items():
        ws.column_dimensions[column].width = width

    if entries:
        _rebuild_table(ws, "AuditoriaGuiada", len(headers), ws.max_row)
    else:
        ws.merge_cells("A3:O3")
        ws["A3"] = "La auditoría guiada no registró decisiones manuales."
        ws["A3"].font = Font(italic=True, color="548235")
        ws["A3"].alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "0F766E"


def _add_protocol_evidence_sheet(
    wb,
    evidence_entries: list[dict[str, Any]],
) -> None:
    if not evidence_entries:
        return
    if SHEET_PROTOCOL_EVIDENCE in wb.sheetnames:
        del wb[SHEET_PROTOCOL_EVIDENCE]

    headers = [
        "Protocolo",
        "Decisión",
        "Cantidad PEL",
        "Cantidad TRATAS",
        "Justificación",
        "Archivo",
        "Soporte visual",
    ]
    ws = wb.create_sheet(SHEET_PROTOCOL_EVIDENCE)
    ws.append(headers)
    for entry in evidence_entries:
        ws.append(
            [
                normalize_identifier(entry.get("protocol")),
                normalize_text(entry.get("decision")),
                entry.get("pel"),
                entry.get("tratas"),
                normalize_text(entry.get("justification")),
                normalize_text(entry.get("filename")),
                "",
            ]
        )
        row_number = ws.max_row
        image_bytes = entry.get("image_bytes")
        if image_bytes:
            try:
                image = OpenpyxlImage(BytesIO(image_bytes))
                scale = min(520 / image.width, 260 / image.height, 1)
                image.width = int(image.width * scale)
                image.height = int(image.height * scale)
                ws.add_image(image, f"G{row_number}")
                ws.row_dimensions[row_number].height = max(48, image.height * 0.75 + 8)
            except Exception:
                ws.cell(row_number, 7).value = "No fue posible incrustar la imagen."

    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 34
    for row in range(2, ws.max_row + 1):
        for column in range(1, 7):
            ws.cell(row, column).alignment = Alignment(vertical="top", wrap_text=True)
    widths = {
        "A": 24,
        "B": 24,
        "C": 16,
        "D": 18,
        "E": 60,
        "F": 32,
        "G": 76,
    }
    for column, width in widths.items():
        ws.column_dimensions[column].width = width
    _rebuild_table(ws, "SoportesProtocolos", len(headers), ws.max_row)
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "D97706"


def default_template_bytes() -> bytes | None:
    for path in (BUNDLED_TEMPLATE_PATH, LEGACY_TEMPLATE_PATH):
        try:
            if path.exists():
                return path.read_bytes()
        except OSError:
            continue
    return None


def _extract_template_logo_bytes(template_bytes: bytes) -> bytes | None:
    try:
        with zipfile.ZipFile(BytesIO(template_bytes)) as archive:
            media_files = sorted(
                name
                for name in archive.namelist()
                if name.startswith("xl/media/")
                and name.lower().endswith((".png", ".jpg", ".jpeg"))
            )
            if not media_files:
                return None
            png_files = [name for name in media_files if name.lower().endswith(".png")]
            return archive.read(png_files[0] if png_files else media_files[0])
    except Exception:
        return None


def _add_oca_logos_from_template(wb, template_bytes: bytes) -> None:
    logo_bytes = _extract_template_logo_bytes(template_bytes)
    if not logo_bytes:
        try:
            logo_bytes = FALLBACK_LOGO_PATH.read_bytes()
        except OSError:
            logo_bytes = None
    if not logo_bytes:
        return

    for sheet_name in CODED_HEADER_SHEETS:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        ws._images = []
        ws["A1"].value = None
        image = OpenpyxlImage(BytesIO(logo_bytes))
        image.width = LOGO_WIDTH_PX
        image.height = LOGO_HEIGHT_PX
        ws.add_image(image, "A1")


def _sheet_rows_by_header(
    ws,
    *,
    header_row: int = 1,
    last_row: int | None = None,
) -> list[dict[str, Any]]:
    headers = _header_map(ws, header_row=header_row)
    rows: list[dict[str, Any]] = []
    data_last_row = ws.max_row if last_row is None else last_row
    for row_number in range(header_row + 1, data_last_row + 1):
        record: dict[str, Any] = {}
        has_value = False
        for header, column in headers.items():
            value = ws.cell(row_number, column).value
            record[header] = value
            if normalize_text(value):
                has_value = True
        if has_value:
            rows.append(record)
    return rows


def _resize_template_data_area(ws, data_rows: int, *, header_row: int) -> int:
    desired_rows = max(data_rows, 1)
    current_rows = max(ws.max_row - header_row, 0)
    if current_rows < desired_rows:
        ws.insert_rows(header_row + current_rows + 1, desired_rows - current_rows)
    elif current_rows > desired_rows:
        ws.delete_rows(header_row + desired_rows + 1, current_rows - desired_rows)
    return header_row + desired_rows


def _clear_template_values(ws, *, header_row: int, last_row: int, last_column: int) -> None:
    for row in range(header_row + 1, last_row + 1):
        for column in range(1, last_column + 1):
            ws.cell(row, column).value = None


def _copy_rows_into_template_sheet(
    source_ws,
    target_ws,
    *,
    table_name: str,
    key_header: str,
) -> None:
    if source_ws is not None:
        source_header_row = _detect_header_row(source_ws, {key_header})
        source_last_row = _last_data_row(
            source_ws,
            key_header,
            header_row=source_header_row,
        )
        source_rows = _sheet_rows_by_header(
            source_ws,
            header_row=source_header_row,
            last_row=source_last_row,
        )
    else:
        source_rows = []
    if target_ws.title in {SHEET_NEW, SHEET_OLD, SHEET_FLOW20}:
        target_headers = (
            _ensure_flow20_timing_columns(
                target_ws,
                header_row=TEMPLATE_HEADER_ROW,
            )
            if target_ws.title == SHEET_FLOW20
            else _ensure_timing_columns(
                target_ws,
                header_row=TEMPLATE_HEADER_ROW,
            )
        )
        _extend_coded_header_to_last_column(
            target_ws,
            max(target_headers.values()),
        )
    else:
        target_headers = _header_map(target_ws, header_row=TEMPLATE_HEADER_ROW)
    if not target_headers:
        raise BitacoraError(f"La hoja '{target_ws.title}' de la plantilla no tiene encabezados.")

    last_column = max(target_headers.values())
    last_row = _resize_template_data_area(
        target_ws,
        len(source_rows),
        header_row=TEMPLATE_HEADER_ROW,
    )

    style_row = TEMPLATE_HEADER_ROW + 1
    for row in range(TEMPLATE_HEADER_ROW + 1, last_row + 1):
        target_ws.row_dimensions[row].height = target_ws.row_dimensions[style_row].height
        for column in range(1, last_column + 1):
            _copy_cell_format(target_ws.cell(style_row, column), target_ws.cell(row, column))

    _clear_template_values(
        target_ws,
        header_row=TEMPLATE_HEADER_ROW,
        last_row=last_row,
        last_column=last_column,
    )

    for offset, record in enumerate(source_rows, start=1):
        target_row = TEMPLATE_HEADER_ROW + offset
        for header, target_column in target_headers.items():
            target_ws.cell(target_row, target_column).value = record.get(header)

    observation_col = target_headers.get("OBSERVACIONES")
    if observation_col:
        for row in range(TEMPLATE_HEADER_ROW + 1, last_row + 1):
            cell = target_ws.cell(row, observation_col)
            alignment = copy(cell.alignment)
            alignment.wrap_text = True
            alignment.vertical = "top"
            cell.alignment = alignment
            if normalize_text(cell.value):
                target_ws.row_dimensions[row].height = 72

    sheet_timing_headers = (
        FLOW20_TIMING_HEADERS
        if target_ws.title == SHEET_FLOW20
        else TIMING_HEADERS
    )
    for timing_header in sheet_timing_headers:
        timing_col = target_headers.get(normalize_header(timing_header))
        if timing_col:
            for row in range(TEMPLATE_HEADER_ROW + 1, last_row + 1):
                target_ws.cell(row, timing_col).number_format = EXCEL_DATETIME_FORMAT

    actual_last_row = _last_data_row(
        target_ws,
        key_header,
        header_row=TEMPLATE_HEADER_ROW,
    )
    if actual_last_row == TEMPLATE_HEADER_ROW:
        actual_last_row = TEMPLATE_HEADER_ROW + 1
    _set_total_formulas(target_ws, actual_last_row, header_row=TEMPLATE_HEADER_ROW)
    _rebuild_table(
        target_ws,
        table_name,
        last_column,
        actual_last_row,
        header_row=TEMPLATE_HEADER_ROW,
    )
    if target_ws.title in {SHEET_NEW, SHEET_OLD, SHEET_FLOW20}:
        from openpyxl.utils import get_column_letter

        target_ws.print_area = (
            f"A1:{get_column_letter(last_column)}{actual_last_row}"
        )


def _ensure_flow20_template_sheet(wb) -> None:
    if SHEET_FLOW20 in wb.sheetnames:
        return
    if SHEET_OLD not in wb.sheetnames:
        raise BitacoraError(
            "La plantilla F-AGP-04 no contiene una base para crear Flujo 2.0."
        )
    ws = wb.copy_worksheet(wb[SHEET_OLD])
    ws.title = SHEET_FLOW20
    headers = [
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
    for table_name in list(ws.tables):
        del ws.tables[table_name]
    source_style_column = 13
    for column, header in enumerate(headers, start=1):
        if column > source_style_column:
            _copy_cell_format(
                ws.cell(TEMPLATE_HEADER_ROW, source_style_column),
                ws.cell(TEMPLATE_HEADER_ROW, column),
            )
            _copy_cell_format(
                ws.cell(TEMPLATE_HEADER_ROW + 1, source_style_column),
                ws.cell(TEMPLATE_HEADER_ROW + 1, column),
            )
        ws.cell(TEMPLATE_HEADER_ROW, column).value = header
    for column in range(len(headers) + 1, ws.max_column + 1):
        ws.cell(TEMPLATE_HEADER_ROW, column).value = None
        ws.cell(TEMPLATE_HEADER_ROW + 1, column).value = None


def _build_from_coded_template(
    source_wb,
    template_bytes: bytes,
    validations: list[ValidationRow],
    email_quantities: list[EmailQuantityResult],
    unreported_cases: list[UnreportedCase],
    audit_entries: list[dict[str, Any]] | None,
    protocol_evidence: list[dict[str, Any]] | None,
) -> bytes:
    try:
        wb = load_workbook(BytesIO(template_bytes), data_only=False)
    except Exception as exc:
        raise BitacoraError("No pude abrir la plantilla F-AGP-04.") from exc

    if SHEET_FLOW20 in source_wb.sheetnames:
        _ensure_flow20_template_sheet(wb)

    sheet_specs = {
        SHEET_NEW: ("Nuevo", "QR"),
        SHEET_OLD: ("Viejo", "QR"),
        "Odm": ("Correos6", "QR"),
        "Correo-Rotulacion": ("Correos", "ASUNTO"),
        SHEET_PNC: ("Tabla4", "QR/ASUNTO"),
    }
    if SHEET_FLOW20 in source_wb.sheetnames:
        sheet_specs[SHEET_FLOW20] = ("Flujo20", "QR")

    missing_template_sheets = [
        sheet_name for sheet_name in sheet_specs if sheet_name not in wb.sheetnames
    ]
    if missing_template_sheets:
        raise BitacoraError(
            "La plantilla F-AGP-04 no contiene las hojas requeridas: "
            + ", ".join(missing_template_sheets)
            + "."
        )

    for sheet_name, (table_name, key_header) in sheet_specs.items():
        target_ws = wb[sheet_name]
        if sheet_name in source_wb.sheetnames:
            _copy_rows_into_template_sheet(
                source_wb[sheet_name],
                target_ws,
                table_name=table_name,
                key_header=key_header,
            )
        else:
            _copy_rows_into_template_sheet(
                None,
                target_ws,
                table_name=table_name,
                key_header=key_header,
            )

    _add_validation_result_sheet(wb, validations)
    _add_email_quantity_result_sheet(wb, email_quantities)
    _add_unreported_cases_sheet(wb, unreported_cases)
    if audit_entries is not None:
        _add_guided_audit_sheet(wb, audit_entries)
    _add_protocol_evidence_sheet(wb, protocol_evidence or [])
    _add_oca_logos_from_template(wb, template_bytes)

    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass

    output = BytesIO()
    wb.save(output)
    wb.close()
    return output.getvalue()


def build_upload_workbook(
    workbook_bytes: bytes,
    validations: list[ValidationRow],
    *,
    matches_only: bool = True,
    template_bytes: bytes | None = None,
    email_quantity_results: list[EmailQuantityResult] | None = None,
    unreported_cases: list[UnreportedCase] | None = None,
    audit_entries: list[dict[str, Any]] | None = None,
    protocol_evidence: list[dict[str, Any]] | None = None,
) -> bytes:
    try:
        wb = load_workbook(BytesIO(workbook_bytes), data_only=False)
    except Exception as exc:
        raise BitacoraError("No pude preparar el Excel de salida.") from exc

    email_quantities = email_quantity_results or []
    _apply_cleaning_and_enrichment(wb, validations, email_quantities)

    validation_sheet_names = [SHEET_NEW, SHEET_OLD]
    if SHEET_FLOW20 in wb.sheetnames:
        validation_sheet_names.append(SHEET_FLOW20)
    by_sheet = {
        sheet: [item for item in validations if item.sheet == sheet]
        for sheet in validation_sheet_names
    }

    table_specs = {
        SHEET_NEW: ("Nuevo", "QR"),
        SHEET_OLD: ("Viejo", "QR"),
        "Odm": ("Correos6", "QR"),
        "Correo-Rotulacion": ("Correos", "ASUNTO"),
        SHEET_PNC: ("Tabla4", "QR/ASUNTO"),
    }
    if SHEET_FLOW20 in wb.sheetnames:
        table_specs[SHEET_FLOW20] = ("Flujo20", "QR")

    for sheet_name in validation_sheet_names:
        ws = wb[sheet_name]
        header_row = _detect_header_row(ws, {"QR", "ESTADO DE LA OT"})
        if matches_only:
            _remove_unvalidated_rows(ws, by_sheet[sheet_name])
        _delete_columns_by_header(ws, REMOVED_UPLOAD_COLUMNS, header_row=header_row)

    for sheet_name in ("Odm", "Correo-Rotulacion"):
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            header_row = _try_detect_header_row(ws, {"ESTADO DE LA OT"}) or 1
            _delete_columns_by_header(
                ws,
                REMOVED_UPLOAD_COLUMNS,
                header_row=header_row,
            )

    if template_bytes:
        return _build_from_coded_template(
            wb,
            template_bytes,
            validations,
            email_quantities,
            unreported_cases or [],
            audit_entries,
            protocol_evidence,
        )

    for sheet_name, (table_name, key_header) in table_specs.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        header_row = _detect_header_row(ws, {key_header})
        last_row = _last_data_row(ws, key_header, header_row=header_row)
        headers = _header_map(ws, header_row=header_row)
        last_column = max(headers.values(), default=1)
        _set_total_formulas(ws, last_row, header_row=header_row)
        _rebuild_table(
            ws,
            table_name,
            last_column,
            last_row,
            header_row=header_row,
        )

    _add_validation_result_sheet(wb, validations)
    _add_email_quantity_result_sheet(wb, email_quantities)
    _add_unreported_cases_sheet(wb, unreported_cases or [])
    if audit_entries is not None:
        _add_guided_audit_sheet(wb, audit_entries)
    _add_protocol_evidence_sheet(wb, protocol_evidence or [])

    for ws in wb.worksheets:
        for pivot in getattr(ws, "_pivots", []):
            try:
                pivot.cacheDefinition.refreshOnLoad = True
            except Exception:
                pass

    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass

    output = BytesIO()
    wb.save(output)
    wb.close()
    return output.getvalue()


def output_filename(production_date: date, *, matches_only: bool = True) -> str:
    suffix = "validada" if matches_only else "revision"
    return f"Produccion {production_date:%d-%m-%Y} - {suffix}.xlsx"
