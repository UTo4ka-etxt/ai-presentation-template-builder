from __future__ import annotations

import collections
import hashlib
import json
import math
import posixpath
import re
import shutil
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .ooxml import A, P, PR, R, _xml

C = "http://schemas.openxmlformats.org/drawingml/2006/chart"


class DecompositionError(RuntimeError):
    pass


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _shape_marker(element: ET.Element) -> ET.Element | None:
    paths = {
        f"{{{P}}}sp": f"./{{{P}}}nvSpPr/{{{P}}}cNvPr",
        f"{{{P}}}pic": f"./{{{P}}}nvPicPr/{{{P}}}cNvPr",
        f"{{{P}}}graphicFrame": f"./{{{P}}}nvGraphicFramePr/{{{P}}}cNvPr",
        f"{{{P}}}cxnSp": f"./{{{P}}}nvCxnSpPr/{{{P}}}cNvPr",
        f"{{{P}}}grpSp": f"./{{{P}}}nvGrpSpPr/{{{P}}}cNvPr",
    }
    path = paths.get(element.tag)
    return element.find(path) if path else None


def _kind(element: ET.Element) -> str:
    if element.tag == f"{{{P}}}pic":
        return "image"
    if element.tag == f"{{{P}}}graphicFrame":
        if element.find(f".//{{{A}}}tbl") is not None:
            return "table"
        if element.find(f".//{{{C}}}chart") is not None:
            return "chart"
        return "graphic"
    if element.tag == f"{{{P}}}grpSp":
        return "group"
    if element.tag == f"{{{P}}}cxnSp":
        return "connector"
    return "shape"


def _geometry(element: ET.Element) -> dict[str, Any]:
    xfrm = None
    if element.tag == f"{{{P}}}graphicFrame":
        xfrm = element.find(f"{{{P}}}xfrm")
    elif element.tag == f"{{{P}}}grpSp":
        xfrm = element.find(f"{{{P}}}grpSpPr/{{{A}}}xfrm")
    else:
        xfrm = element.find(f"{{{P}}}spPr/{{{A}}}xfrm")
    off = xfrm.find(f"{{{A}}}off") if xfrm is not None else None
    ext = xfrm.find(f"{{{A}}}ext") if xfrm is not None else None
    return {
        "x": int(off.get("x", "0")) if off is not None else 0,
        "y": int(off.get("y", "0")) if off is not None else 0,
        "width": int(ext.get("cx", "0")) if ext is not None else 0,
        "height": int(ext.get("cy", "0")) if ext is not None else 0,
        "rotation": float(xfrm.get("rot", "0")) / 60000 if xfrm is not None else 0.0,
        "units": "EMU",
        "coordinate_space": "slide",
    }


def _paragraphs(element: ET.Element) -> list[dict[str, Any]]:
    output = []
    for paragraph in element.iter(f"{{{A}}}p"):
        runs = []
        for run in paragraph.findall(f"{{{A}}}r"):
            props = run.find(f"{{{A}}}rPr")
            font = props.find(f"{{{A}}}latin") if props is not None else None
            runs.append(
                {
                    "text": "".join(node.text or "" for node in run.iter(f"{{{A}}}t")),
                    "font_name": font.get("typeface") if font is not None else None,
                    "font_size_pt": int(props.get("sz")) / 100 if props is not None and props.get("sz") else None,
                    "bold": props.get("b") == "1" if props is not None and props.get("b") is not None else None,
                    "italic": props.get("i") == "1" if props is not None and props.get("i") is not None else None,
                }
            )
        text = "".join(node.text or "" for node in paragraph.iter(f"{{{A}}}t"))
        ppr = paragraph.find(f"{{{A}}}pPr")
        output.append({"text": text, "level": int(ppr.get("lvl", "0")) if ppr is not None else 0, "runs": runs})
    return output


def _placeholder(element: ET.Element) -> dict[str, Any] | None:
    node = element.find(f".//{{{P}}}nvPr/{{{P}}}ph")
    if node is None:
        return None
    return {"type": node.get("type", "body"), "idx": node.get("idx")}


def _table_values(element: ET.Element) -> list[list[str]]:
    rows = []
    for row in element.findall(f".//{{{A}}}tbl/{{{A}}}tr"):
        rows.append(["".join(node.text or "" for node in cell.iter(f"{{{A}}}t")) for cell in row.findall(f"{{{A}}}tc")])
    return rows


class PptxDecomposer:
    def decompose(self, template_path: str | Path, output_dir: str | Path) -> dict[str, Any]:
        source = Path(template_path).resolve()
        output = Path(output_dir).resolve()
        if source.suffix.lower() != ".pptx" or not source.exists():
            raise DecompositionError(f"Template must be an existing .pptx file: {source}")
        output.mkdir(parents=True, exist_ok=True)
        (output / "semantics").mkdir(exist_ok=True)
        (output / "assets").mkdir(exist_ok=True)
        source_data = source.read_bytes()
        shutil.copyfile(source, output / "source.pptx")

        with zipfile.ZipFile(source) as archive:
            names = set(archive.namelist())
            presentation = _xml(archive.read("ppt/presentation.xml"))
            presentation_rels = _xml(archive.read("ppt/_rels/presentation.xml.rels"))
            relation_map = {rel.get("Id"): rel for rel in presentation_rels}
            slide_size_node = presentation.find(f"{{{P}}}sldSz")
            width = int(slide_size_node.get("cx", "12192000")) if slide_size_node is not None else 12192000
            height = int(slide_size_node.get("cy", "6858000")) if slide_size_node is not None else 6858000
            slide_list = presentation.find(f"{{{P}}}sldIdLst")
            if slide_list is None:
                raise DecompositionError("Template has no slides")

            slides: list[dict[str, Any]] = []
            text_presence: collections.Counter[str] = collections.Counter()
            image_presence: collections.Counter[str] = collections.Counter()
            for index, slide_ref in enumerate(slide_list, start=1):
                rel_id = slide_ref.get(f"{{{R}}}id")
                relation = relation_map.get(rel_id)
                if relation is None:
                    raise DecompositionError(f"Missing presentation relationship {rel_id}")
                slide_part = posixpath.normpath(posixpath.join("ppt", relation.get("Target", "")))
                rel_part = posixpath.join(posixpath.dirname(slide_part), "_rels", posixpath.basename(slide_part) + ".rels")
                slide_rels = _xml(archive.read(rel_part)) if rel_part in names else ET.Element(f"{{{PR}}}Relationships")
                slide_rel_map = {rel.get("Id"): rel for rel in slide_rels}
                root = _xml(archive.read(slide_part))
                objects = []
                for z_order, element in enumerate(root.iter()):
                    marker = _shape_marker(element)
                    if marker is None:
                        continue
                    shape_id = int(marker.get("id", "0"))
                    object_kind = _kind(element)
                    paragraphs = _paragraphs(element)
                    text = "\n".join(item["text"] for item in paragraphs).strip() if paragraphs else None
                    geometry = _geometry(element)
                    placeholder = _placeholder(element)
                    asset = None
                    if object_kind == "image":
                        blip = element.find(f".//{{{A}}}blip")
                        embed = blip.get(f"{{{R}}}embed") if blip is not None else None
                        image_rel = slide_rel_map.get(embed)
                        if image_rel is not None and image_rel.get("TargetMode") != "External":
                            media_part = posixpath.normpath(posixpath.join(posixpath.dirname(slide_part), image_rel.get("Target", "")))
                            if media_part in names:
                                media = archive.read(media_part)
                                digest = _sha256(media)
                                suffix = Path(media_part).suffix.lower() or ".bin"
                                asset_path = output / "assets" / f"{digest}{suffix}"
                                if not asset_path.exists():
                                    asset_path.write_bytes(media)
                                asset = {"source_part": media_part, "asset_id": digest, "path": f"assets/{asset_path.name}", "sha256": digest}
                    item = {
                        "id": f"s{index}_o{shape_id}",
                        "shape_id": shape_id,
                        "name": marker.get("name", ""),
                        "kind": object_kind,
                        "parent_id": None,
                        "z_order": z_order,
                        "geometry": geometry,
                        "source_part": slide_part,
                        "xml": ET.tostring(element, encoding="unicode"),
                        "text": text,
                        "paragraphs": paragraphs,
                        "assets": [asset] if asset else [],
                        "origin": {"scope": "slide", "part_id": f"s{index}", "source_part": slide_part},
                    }
                    if placeholder:
                        item["placeholder"] = placeholder
                    if object_kind == "table":
                        item["table"] = _table_values(element)
                    objects.append(item)
                slide_text_keys = {self._text_key(item.get("text")) for item in objects if item.get("text")}
                text_presence.update(key for key in slide_text_keys if key)
                slide_image_keys = {item["assets"][0]["sha256"] for item in objects if item.get("assets")}
                image_presence.update(slide_image_keys)
                slides.append({"id": f"s{index}", "index": index, "source_part": slide_part, "objects": objects})

        available = []
        unavailable = []
        total_slides = len(slides)
        repeat_threshold = max(2, math.ceil(total_slides * 0.5))
        template_slides = []
        for slide in slides:
            components = []
            for item in slide["objects"]:
                slot = self._slot_for(item, width, height, text_presence, image_presence, repeat_threshold)
                if slot is None:
                    continue
                component_id, component_type, role, content_type = slot
                components.append(
                    {
                        "id": component_id,
                        "type": component_type,
                        "members": [item["id"]],
                        "slots": [{"object_id": item["id"], "role": role, "content_type": content_type, "paragraph_indices": None}],
                        "preserve": [],
                    }
                )
            slide_role = self._slide_role(components)
            template_slides.append(
                {
                    "id": slide["id"],
                    "index": slide["index"],
                    "source_part": "/" + slide["source_part"],
                    "layout_id": None,
                    "objects": slide["objects"],
                }
            )
            if not components:
                unavailable.append({"slide_id": slide["id"], "status": "unavailable", "reasons": [{"code": "NO_EDITABLE_SLOTS", "object_ids": []}]})
                continue
            semantics = {
                "schema_version": "0.4",
                "source_sha256": _sha256(source_data),
                "model": "deterministic-ooxml",
                "prompt_version": "generic_decompose_v1",
                "preview_sha256": None,
                "result": {
                    "slide_id": slide["id"],
                    "slide_role": slide_role,
                    "usage": "generation_template",
                    "components": components,
                    "needs_review": False,
                    "unassigned": [],
                    "warnings": [],
                },
                "object_scope": "slide",
                "object_catalog": slide["objects"],
            }
            semantics_path = output / "semantics" / f"{slide['id']}.json"
            semantics_path.write_text(json.dumps(semantics, ensure_ascii=False, indent=2), encoding="utf-8")
            available.append(
                {
                    "slide_id": slide["id"],
                    "semantics": f"semantics/{slide['id']}.json",
                    "preview": None,
                    "components": components,
                    "usage": "generation_template",
                    "needs_review": False,
                    "unassigned": [],
                    "warnings": [],
                }
            )

        template = {
            "schema_version": "0.4",
            "source": {"name": source.name, "path": "source.pptx", "sha256": _sha256(source_data)},
            "slide_size": {"width": width, "height": height, "units": "EMU"},
            "assets": {},
            "asset_catalog": {},
            "themes": [],
            "masters": [],
            "layouts": [],
            "slides": template_slides,
            "diagnostics": [],
            "external_media": [],
        }
        catalog = {
            "schema_version": "0.4",
            "source": "source.pptx",
            "template": "template.json",
            "source_sha256": _sha256(source_data),
            "total_slides": total_slides,
            "partial": bool(unavailable),
            "slides": available,
        }
        validation = {
            "schema_version": "0.4",
            "source_sha256": _sha256(source_data),
            "counts": {"valid": len(available), "unavailable": len(unavailable)},
            "valid_slide_ids": [item["slide_id"] for item in available],
            "slides": [*[{"slide_id": item["slide_id"], "status": "valid"} for item in available], *unavailable],
        }
        (output / "template.json").write_text(json.dumps(template, ensure_ascii=False, indent=2), encoding="utf-8")
        (output / "catalog.json").write_text(json.dumps(catalog, ensure_ascii=False, indent=2), encoding="utf-8")
        (output / "validation_report.json").write_text(json.dumps(validation, ensure_ascii=False, indent=2), encoding="utf-8")
        return {
            "package": str(output),
            "source": str(source),
            "slides_total": total_slides,
            "slides_available": len(available),
            "slides_unavailable": len(unavailable),
        }

    @staticmethod
    def _text_key(text: str | None) -> str:
        return re.sub(r"\s+", " ", (text or "").strip().lower())

    def _slot_for(
        self,
        item: dict[str, Any],
        slide_width: int,
        slide_height: int,
        text_presence: collections.Counter[str],
        image_presence: collections.Counter[str],
        repeat_threshold: int,
    ) -> tuple[str, str, str, str] | None:
        kind = item["kind"]
        shape_id = item["shape_id"]
        geometry = item["geometry"]
        placeholder_type = (item.get("placeholder") or {}).get("type")
        if kind == "table":
            return f"table_{shape_id}", "table", "data_table", "table"
        if kind == "image":
            asset_hash = item["assets"][0]["sha256"] if item.get("assets") else None
            area = geometry["width"] * geometry["height"]
            slide_area = max(slide_width * slide_height, 1)
            if placeholder_type in {"pic", "obj"}:
                return f"image_{shape_id}", "image", "content_image", "image"
            if asset_hash and image_presence[asset_hash] >= repeat_threshold:
                return None
            if area / slide_area >= 0.035:
                return f"image_{shape_id}", "image", "content_image", "image"
            return None
        if kind != "shape" or not item.get("text"):
            return None
        text = item["text"]
        key = self._text_key(text)
        font_sizes = [run.get("font_size_pt") for paragraph in item["paragraphs"] for run in paragraph.get("runs", []) if run.get("font_size_pt")]
        max_font = max(font_sizes, default=0)
        is_title = placeholder_type in {"title", "ctrTitle", "subTitle"} or (geometry["y"] < slide_height * 0.3 and max_font >= 22)
        if placeholder_type in {"dt", "ftr", "sldNum"}:
            return None
        if text_presence[key] >= repeat_threshold and not is_title:
            return None
        if geometry["y"] > slide_height * 0.88 and max_font and max_font < 18:
            return None
        if is_title:
            role = "slide_title" if placeholder_type != "subTitle" else "slide_subtitle"
            return f"text_{shape_id}", "title", role, "text"
        role = "body_text" if len(text) > 45 or "\n" in text else "label_text"
        return f"text_{shape_id}", "text_block", role, "text"

    @staticmethod
    def _slide_role(components: list[dict[str, Any]]) -> str:
        types = [component["type"] for component in components]
        text_count = sum(component["type"] in {"title", "text_block"} for component in components)
        if "table" in types:
            return "table_slide"
        if "image" in types and text_count <= 2:
            return "visual_slide"
        if text_count <= 2 and types.count("title") >= 1:
            return "cover_or_section"
        if text_count >= 5:
            return "content_grid"
        return "content_slide"

