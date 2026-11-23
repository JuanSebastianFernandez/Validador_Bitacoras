"""Keep native vector header logos that openpyxl cannot round-trip.

The corporate template uses DrawingML groups, not PNGs. Copy those original
parts after serialization; do not rasterize, recreate, or substitute the brand.
Raster logos in older templates are still handled by the existing image path.
"""
from __future__ import annotations

from io import BytesIO
from posixpath import basename, dirname, normpath
import re
import zipfile
from xml.etree import ElementTree as ET

SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DRAW_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
DRAW_TYPE = "application/vnd.openxmlformats-officedocument.drawing+xml"
X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
XM_NS = "http://schemas.microsoft.com/office/excel/2006/main"


def restore_template_data_validations(wb, template: bytes) -> None:
    """Keep template dropdowns stored as x14 extensions as standard Excel rules.

    openpyxl drops x14 validations during import. These rules have the same list
    semantics in the base schema; ranges are metadata, not millions of cells.
    """
    from openpyxl.worksheet.datavalidation import DataValidation

    fields = {"type", "allowBlank", "showInputMessage", "showErrorMessage", "showDropDown",
              "errorStyle", "errorTitle", "error", "promptTitle", "prompt", "operator", "imeMode"}
    with zipfile.ZipFile(BytesIO(template)) as source:
        for name, path in _sheet_paths(source).items():
            if name not in wb.sheetnames:
                continue
            ws = wb[name]
            existing = {(str(v.sqref), v.type, v.formula1, v.formula2) for v in ws.data_validations.dataValidation}
            for rule in ET.fromstring(source.read(path)).iter(f"{{{X14_NS}}}dataValidation"):
                sqref = rule.findtext(f"{{{XM_NS}}}sqref")
                if not sqref:
                    continue
                formulas = {key: rule.findtext(f"{{{X14_NS}}}{key}/{{{XM_NS}}}f")
                            for key in ("formula1", "formula2")}
                validation = DataValidation(
                    sqref=sqref, **formulas,
                    **{key: value for key, value in rule.attrib.items() if key in fields},
                )
                key = (str(validation.sqref), validation.type, validation.formula1, validation.formula2)
                if key not in existing:
                    ws.add_data_validation(validation)
                    existing.add(key)


def _package_xml(root: ET.Element, namespace: str) -> bytes:
    # OPC readers (including Excel/OpenXML SDK) expect package metadata to use
    # an unprefixed default namespace, not ElementTree's generated ns0 prefix.
    plain = ET.Element(root.tag.rsplit("}", 1)[-1], {"xmlns": namespace})
    for child in root:
        ET.SubElement(plain, child.tag.rsplit("}", 1)[-1], dict(child.attrib))
    return ET.tostring(plain, encoding="utf-8", xml_declaration=True)


def _relationships_path(part: str) -> str:
    return f"{dirname(part)}/_rels/{basename(part)}.rels"


def _target(part: str, target: str) -> str:
    return target.lstrip("/") if target.startswith("/") else normpath(f"{dirname(part)}/{target}")


def _sheet_paths(archive: zipfile.ZipFile) -> dict[str, str]:
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {r.attrib["Id"]: _target("xl/workbook.xml", r.attrib["Target"]) for r in rels}
    book = ET.fromstring(archive.read("xl/workbook.xml"))
    return {
        sheet.attrib["name"]: targets[sheet.attrib[f"{{{REL_NS}}}id"]]
        for sheet in book.findall(f"{{{SHEET_NS}}}sheets/{{{SHEET_NS}}}sheet")
    }


def restore_template_vector_logos(output: bytes, template: bytes, sheets: tuple[str, ...]) -> bytes:
    """Restore pure, self-contained vector drawings only on coded header sheets.

    Drawings with images/charts/other relationships are not blindly transplanted:
    that could overwrite generated charts or protocol evidence. The supplied
    company logo consists solely of self-contained grpSp vector shapes.
    """
    with zipfile.ZipFile(BytesIO(template)) as reference, zipfile.ZipFile(BytesIO(output)) as result:
        reference_paths, result_paths = _sheet_paths(reference), _sheet_paths(result)
        parts: dict[str, bytes] = {}
        types = ET.fromstring(result.read("[Content_Types].xml"))
        for index, sheet in enumerate(sheets, start=1):
            source_path, target_path = reference_paths.get(sheet), result_paths.get(sheet)
            if not source_path or not target_path:
                continue
            rel_path = _relationships_path(source_path)
            if rel_path not in reference.namelist():
                continue
            rels = ET.fromstring(reference.read(rel_path))
            drawing_rel = next((r for r in rels if r.attrib.get("Type") == f"{REL_NS}/drawing"), None)
            if drawing_rel is None:
                continue
            drawing_path = _target(source_path, drawing_rel.attrib["Target"])
            drawing_bytes = reference.read(drawing_path)
            drawing = ET.fromstring(drawing_bytes)
            if drawing.find(f".//{{{DRAW_NS}}}grpSp") is None:
                continue
            if _relationships_path(drawing_path) in reference.namelist():
                raise ValueError(f"El logo vectorial de '{sheet}' tiene dependencias no admitidas.")
            sheet_xml = result.read(target_path)
            if re.search(rb"<(?:\w+:)?drawing\b", sheet_xml):
                raise ValueError(f"La hoja '{sheet}' ya contiene un dibujo; no se reemplazará silenciosamente.")
            restored_drawing = f"xl/drawings/template_logo_{index}.xml"
            if restored_drawing in result.namelist():
                raise ValueError("Ya existe la parte reservada para el logo de la plantilla.")
            parts[restored_drawing] = drawing_bytes  # Exact original geometry and anchor.

            result_rel_path = _relationships_path(target_path)
            sheet_rels = (
                ET.fromstring(result.read(result_rel_path))
                if result_rel_path in result.namelist()
                else ET.Element(f"{{{PKG_NS}}}Relationships")
            )
            used_ids = {r.attrib["Id"] for r in sheet_rels}
            rel_id = "rIdTemplateLogo"
            while rel_id in used_ids:
                rel_id += "x"
            ET.SubElement(sheet_rels, f"{{{PKG_NS}}}Relationship", {
                "Id": rel_id, "Type": f"{REL_NS}/drawing",
                "Target": f"/xl/drawings/template_logo_{index}.xml",
            })
            parts[result_rel_path] = _package_xml(sheet_rels, PKG_NS)
            drawing_tag = f'<drawing xmlns:r="{REL_NS}" r:id="{rel_id}"/>'.encode()
            # Schema order: drawing precedes legacyDrawing, tableParts and extLst.
            position = re.search(
                rb"<(?:\w+:)?(?:legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b|</(?:\w+:)?worksheet>",
                sheet_xml,
            )
            if position is None:
                raise ValueError("No se encontró el cierre XML de una hoja codificada.")
            parts[target_path] = sheet_xml[:position.start()] + drawing_tag + sheet_xml[position.start():]
            ET.SubElement(types, f"{{{CT_NS}}}Override", {
                "PartName": f"/{restored_drawing}", "ContentType": DRAW_TYPE,
            })
        if not parts:
            return output
        parts["[Content_Types].xml"] = _package_xml(types, CT_NS)
        buffer = BytesIO()
        with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as target:
            for info in result.infolist():
                target.writestr(info, parts.pop(info.filename, result.read(info.filename)))
            for path, data in parts.items():
                target.writestr(path, data)
        return buffer.getvalue()
