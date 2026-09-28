from __future__ import annotations

import json
import posixpath
import re
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .ooxml import A, P, PR, R, _xml

PLACEHOLDERS = re.compile(r"(?:lorem ipsum|\bXXX\b|\bTODO\b|вставьте текст)", re.IGNORECASE)


def audit_pptx(path: str | Path) -> dict[str, Any]:
    pptx = Path(path).resolve()
    issues: list[dict[str, Any]] = []
    try:
        with zipfile.ZipFile(pptx) as archive:
            bad_member = archive.testzip()
            if bad_member:
                issues.append({"severity": "error", "code": "CORRUPT_ZIP_ENTRY", "detail": bad_member})
            names = set(archive.namelist())
            presentation = _xml(archive.read("ppt/presentation.xml"))
            rels = _xml(archive.read("ppt/_rels/presentation.xml.rels"))
            relation_map = {rel.get("Id"): rel for rel in rels}
            slide_list = presentation.find(f"{{{P}}}sldIdLst")
            slide_nodes = list(slide_list) if slide_list is not None else []
            if not slide_nodes:
                issues.append({"severity": "error", "code": "EMPTY_DECK", "detail": "No slides are referenced"})
            for number, slide_node in enumerate(slide_nodes, start=1):
                rel_id = slide_node.get(f"{{{R}}}id")
                rel = relation_map.get(rel_id)
                if rel is None:
                    issues.append({"severity": "error", "code": "MISSING_SLIDE_REL", "slide": number, "detail": rel_id})
                    continue
                target = posixpath.normpath(posixpath.join("ppt", rel.get("Target", "")))
                if target not in names:
                    issues.append({"severity": "error", "code": "MISSING_SLIDE_PART", "slide": number, "detail": target})
                    continue
                slide = _xml(archive.read(target))
                text = " ".join(node.text or "" for node in slide.iter(f"{{{A}}}t"))
                match = PLACEHOLDERS.search(text)
                if match:
                    issues.append({"severity": "warning", "code": "PLACEHOLDER_TEXT", "slide": number, "detail": match.group(0)})
                editable = sum(1 for _ in slide.iter(f"{{{P}}}sp")) + sum(1 for _ in slide.iter(f"{{{P}}}graphicFrame"))
                pictures = sum(1 for _ in slide.iter(f"{{{P}}}pic"))
                if editable == 0 and pictures <= 1:
                    issues.append({"severity": "warning", "code": "RASTER_ONLY_SLIDE", "slide": number, "detail": "No editable text/table/chart objects"})
    except (OSError, zipfile.BadZipFile, KeyError, ET.ParseError) as exc:
        issues.append({"severity": "error", "code": "FILE_OPEN_FAILED", "detail": str(exc)})
        slide_nodes = []
    return {
        "file": str(pptx),
        "slide_count": len(slide_nodes),
        "ok": not any(issue["severity"] == "error" for issue in issues),
        "issues": issues,
    }


def write_audit(report: dict[str, Any], path: str | Path) -> None:
    Path(path).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

