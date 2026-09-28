from __future__ import annotations

import copy
import json
import posixpath
import re
import struct
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .catalog import Slot, TemplateCatalog

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
EP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
NS = {"p": P, "a": A, "r": R, "pr": PR, "ct": CT, "ep": EP}

for prefix, uri in (("p", P), ("a", A), ("r", R), ("", PR)):
    ET.register_namespace(prefix, uri)


class BuildError(RuntimeError):
    pass


def _xml(data: bytes) -> ET.Element:
    return ET.fromstring(data)


def _bytes(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _shape_by_id(root: ET.Element, shape_id: int) -> ET.Element:
    candidates = {
        f"{{{P}}}sp": f"./{{{P}}}nvSpPr/{{{P}}}cNvPr",
        f"{{{P}}}pic": f"./{{{P}}}nvPicPr/{{{P}}}cNvPr",
        f"{{{P}}}graphicFrame": f"./{{{P}}}nvGraphicFramePr/{{{P}}}cNvPr",
        f"{{{P}}}cxnSp": f"./{{{P}}}nvCxnSpPr/{{{P}}}cNvPr",
        f"{{{P}}}grpSp": f"./{{{P}}}nvGrpSpPr/{{{P}}}cNvPr",
    }
    for element in root.iter():
        path = candidates.get(element.tag)
        if path:
            marker = element.find(path)
            if marker is not None and marker.get("id") == str(shape_id):
                return element
    raise BuildError(f"Shape id {shape_id} not found in cloned slide")


def _set_paragraph_text(paragraph: ET.Element, text: str) -> None:
    text_nodes = {f"{{{A}}}r", f"{{{A}}}fld", f"{{{A}}}br"}
    first_run = paragraph.find(f"{{{A}}}r")
    run_props = copy.deepcopy(first_run.find(f"{{{A}}}rPr")) if first_run is not None else None
    for child in list(paragraph):
        if child.tag in text_nodes:
            paragraph.remove(child)
    run = ET.Element(f"{{{A}}}r")
    if run_props is not None:
        run.append(run_props)
    node = ET.SubElement(run, f"{{{A}}}t")
    if text[:1].isspace() or text[-1:].isspace():
        node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    node.text = text
    end_props = paragraph.find(f"{{{A}}}endParaRPr")
    insert_at = list(paragraph).index(end_props) if end_props is not None else len(paragraph)
    paragraph.insert(insert_at, run)


def _replace_text(shape: ET.Element, slot: Slot, value: Any) -> None:
    if not isinstance(value, str):
        raise BuildError(f"Slot {slot.key} requires a string")
    paragraphs = list(shape.iter(f"{{{A}}}p"))
    if not paragraphs:
        raise BuildError(f"Text slot {slot.key} points to a shape without paragraphs")
    if slot.paragraph_indices is not None:
        for index in slot.paragraph_indices:
            if index >= len(paragraphs):
                raise BuildError(f"Paragraph {index} does not exist for slot {slot.key}")
            _set_paragraph_text(paragraphs[index], value)
        return

    lines = value.splitlines() or [""]
    tx_body = shape.find(f".//{{{P}}}txBody")
    if tx_body is None:
        tx_body = shape.find(f".//{{{A}}}txBody")
    if tx_body is None:
        raise BuildError(f"Text body not found for slot {slot.key}")
    direct_paragraphs = tx_body.findall(f"{{{A}}}p")
    if not direct_paragraphs:
        raise BuildError(f"Text body has no paragraphs for slot {slot.key}")
    while len(direct_paragraphs) < len(lines):
        clone = copy.deepcopy(direct_paragraphs[-1])
        tx_body.append(clone)
        direct_paragraphs.append(clone)
    while len(direct_paragraphs) > len(lines):
        tx_body.remove(direct_paragraphs.pop())
    for paragraph, line in zip(direct_paragraphs, lines, strict=True):
        _set_paragraph_text(paragraph, line)


def _replace_table(shape: ET.Element, slot: Slot, value: Any) -> None:
    if not isinstance(value, list) or not value or not all(isinstance(row, list) for row in value):
        raise BuildError(f"Slot {slot.key} requires a non-empty two-dimensional array")
    table = shape.find(f".//{{{A}}}tbl")
    if table is None:
        raise BuildError(f"Table object not found for slot {slot.key}")
    rows = table.findall(f"{{{A}}}tr")
    if len(value) > len(rows):
        raise BuildError(f"Table slot {slot.key} has {len(rows)} template rows, got {len(value)}")
    for row_index, row in enumerate(rows[: len(value)]):
        cells = row.findall(f"{{{A}}}tc")
        incoming = value[row_index]
        if len(incoming) > len(cells):
            raise BuildError(f"Table slot {slot.key} row {row_index} has {len(cells)} columns")
        for column_index, cell in enumerate(cells):
            text = str(incoming[column_index]) if column_index < len(incoming) else ""
            paragraph = cell.find(f".//{{{A}}}p")
            if paragraph is not None:
                _set_paragraph_text(paragraph, text)
    for row in rows[len(value) :]:
        table.remove(row)


def _image_size(data: bytes, suffix: str) -> tuple[int, int] | None:
    if suffix == ".png" and data.startswith(b"\x89PNG"):
        return struct.unpack(">II", data[16:24])
    if suffix in {".jpg", ".jpeg"} and data.startswith(b"\xff\xd8"):
        offset = 2
        while offset + 9 < len(data):
            if data[offset] != 0xFF:
                offset += 1
                continue
            marker = data[offset + 1]
            length = int.from_bytes(data[offset + 2 : offset + 4], "big")
            if marker in range(0xC0, 0xC4):
                return int.from_bytes(data[offset + 7 : offset + 9], "big"), int.from_bytes(data[offset + 5 : offset + 7], "big")
            offset += max(length + 2, 2)
    return None


def _set_center_crop(shape: ET.Element, image_size: tuple[int, int], box_size: tuple[int, int]) -> None:
    image_ratio = image_size[0] / image_size[1]
    box_ratio = box_size[0] / box_size[1]
    crop = {"l": "0", "r": "0", "t": "0", "b": "0"}
    if image_ratio > box_ratio:
        amount = round((1 - box_ratio / image_ratio) * 50000)
        crop["l"] = crop["r"] = str(amount)
    else:
        amount = round((1 - image_ratio / box_ratio) * 50000)
        crop["t"] = crop["b"] = str(amount)
    blip_fill = shape.find(f"{{{P}}}blipFill")
    if blip_fill is None:
        return
    src_rect = blip_fill.find(f"{{{A}}}srcRect")
    if src_rect is None:
        src_rect = ET.Element(f"{{{A}}}srcRect")
        blip = blip_fill.find(f"{{{A}}}blip")
        blip_fill.insert((list(blip_fill).index(blip) + 1) if blip is not None else 0, src_rect)
    src_rect.attrib.clear()
    src_rect.attrib.update(crop)


class PresentationBuilder:
    def __init__(self, package_dir: str | Path):
        self.catalog = TemplateCatalog(package_dir)
        self.package_dir = self.catalog.package_dir
        self.source_path = self.package_dir / self.catalog.template_data["source"].get("path", "source.pptx")
        if not self.source_path.exists():
            self.source_path = self.package_dir / "source.pptx"

    def build(self, plan: dict[str, Any], output_path: str | Path, *, clear_unfilled: bool = True) -> dict[str, Any]:
        errors = self.catalog.validate_plan(plan)
        if errors:
            raise BuildError("; ".join(errors))
        with zipfile.ZipFile(self.source_path) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}

        presentation = _xml(parts["ppt/presentation.xml"])
        presentation_rels = _xml(parts["ppt/_rels/presentation.xml.rels"])
        content_types = _xml(parts["[Content_Types].xml"])
        slide_list = presentation.find(f"{{{P}}}sldIdLst")
        if slide_list is None:
            raise BuildError("Source presentation has no slide list")
        old_slide_ids = [int(node.get("id", "255")) for node in slide_list]
        for node in list(slide_list):
            slide_list.remove(node)

        existing_numbers = [int(match.group(1)) for name in parts if (match := re.fullmatch(r"ppt/slides/slide(\d+)\.xml", name))]
        next_part_number = max(existing_numbers, default=0) + 1
        rel_numbers = [int(match.group(1)) for rel in presentation_rels if (match := re.fullmatch(r"rId(\d+)", rel.get("Id", "")))]
        next_rel_number = max(rel_numbers, default=0) + 1
        next_slide_id = max(old_slide_ids, default=255) + 1
        generated_media = 1
        report_slides: list[dict[str, Any]] = []

        for output_index, requested in enumerate(plan["slides"], start=1):
            slide_id = requested["template_slide_id"]
            template = self.catalog.get(slide_id)
            source_part = template.source_part
            source_rel_part = posixpath.join(posixpath.dirname(source_part), "_rels", posixpath.basename(source_part) + ".rels")
            new_part = f"ppt/slides/slide{next_part_number}.xml"
            new_rel_part = f"ppt/slides/_rels/slide{next_part_number}.xml.rels"
            slide_root = _xml(parts[source_part])
            slide_rels = _xml(parts[source_rel_part])
            for rel in list(slide_rels):
                rel_type = rel.get("Type", "")
                if rel_type.endswith("/notesSlide") or rel_type.endswith("/comments") or rel_type.endswith("/commentAuthors"):
                    slide_rels.remove(rel)

            supplied = requested.get("values", {})
            changed: list[str] = []
            for key, slot in template.slots.items():
                if key not in supplied and not (clear_unfilled and slot.content_type == "text"):
                    continue
                value = supplied.get(key, "")
                obj = self.catalog.object(slide_id, slot.object_id)
                if obj["source_part"].lstrip("/") != source_part:
                    raise BuildError(f"Editable slot {key} unexpectedly targets shared master/layout content")
                shape = _shape_by_id(slide_root, int(obj["shape_id"]))
                if slot.content_type == "text":
                    _replace_text(shape, slot, value)
                elif slot.content_type == "table":
                    _replace_table(shape, slot, value)
                elif slot.content_type == "image":
                    image_path = Path(value["path"] if isinstance(value, dict) else value)
                    if not image_path.is_absolute():
                        content_root = Path(plan.get("metadata", {}).get("content_root", self.package_dir))
                        image_path = (content_root / image_path).resolve()
                    if not image_path.exists():
                        raise BuildError(f"Image for {key} does not exist: {image_path}")
                    image_data = image_path.read_bytes()
                    suffix = image_path.suffix.lower()
                    if suffix not in {".png", ".jpg", ".jpeg"}:
                        raise BuildError(f"Image slot {key} only supports PNG and JPEG")
                    blip = shape.find(f".//{{{A}}}blip")
                    if blip is None or f"{{{R}}}embed" not in blip.attrib:
                        raise BuildError(f"Image relationship not found for {key}")
                    rel_id = blip.get(f"{{{R}}}embed")
                    relation = next((item for item in slide_rels if item.get("Id") == rel_id), None)
                    if relation is None:
                        raise BuildError(f"Relationship {rel_id} not found for {key}")
                    media_name = f"generated_{next_part_number}_{generated_media}{suffix}"
                    generated_media += 1
                    parts[f"ppt/media/{media_name}"] = image_data
                    relation.set("Target", f"../media/{media_name}")
                    geometry = obj["geometry"]
                    size = _image_size(image_data, suffix)
                    if size and geometry.get("width") and geometry.get("height"):
                        _set_center_crop(shape, size, (int(geometry["width"]), int(geometry["height"])))
                    self._ensure_content_type(content_types, suffix)
                else:
                    raise BuildError(f"Unsupported slot type {slot.content_type!r} for {key}")
                changed.append(key)

            parts[new_part] = _bytes(slide_root)
            parts[new_rel_part] = _bytes(slide_rels)
            rel_id = f"rId{next_rel_number}"
            ET.SubElement(
                presentation_rels,
                f"{{{PR}}}Relationship",
                {
                    "Id": rel_id,
                    "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
                    "Target": f"slides/slide{next_part_number}.xml",
                },
            )
            ET.SubElement(slide_list, f"{{{P}}}sldId", {"id": str(next_slide_id), f"{{{R}}}id": rel_id})
            ET.SubElement(
                content_types,
                f"{{{CT}}}Override",
                {
                    "PartName": f"/ppt/slides/slide{next_part_number}.xml",
                    "ContentType": "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
                },
            )
            report_slides.append({"number": output_index, "template_slide_id": slide_id, "changed_slots": changed})
            next_part_number += 1
            next_rel_number += 1
            next_slide_id += 1

        parts["ppt/presentation.xml"] = _bytes(presentation)
        parts["ppt/_rels/presentation.xml.rels"] = _bytes(presentation_rels)
        # OPC validators require Content_Types children in its default namespace,
        # not behind a generated ns0 prefix.
        ET.register_namespace("", CT)
        parts["[Content_Types].xml"] = _bytes(content_types)
        ET.register_namespace("", PR)
        if "docProps/app.xml" in parts:
            app = _xml(parts["docProps/app.xml"])
            slides_node = app.find(f"{{{EP}}}Slides")
            if slides_node is not None:
                slides_node.text = str(len(plan["slides"]))
            parts["docProps/app.xml"] = _bytes(app)

        output = Path(output_path).resolve()
        output.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for name, data in parts.items():
                archive.writestr(name, data)
        return {
            "output": str(output),
            "slide_count": len(plan["slides"]),
            "slides": report_slides,
            "source": str(self.source_path),
        }

    @staticmethod
    def _ensure_content_type(root: ET.Element, suffix: str) -> None:
        extension = suffix.lstrip(".").lower()
        if any(node.get("Extension", "").lower() == extension for node in root.findall(f"{{{CT}}}Default")):
            return
        mime = "image/png" if extension == "png" else "image/jpeg"
        ET.SubElement(root, f"{{{CT}}}Default", {"Extension": extension, "ContentType": mime})
