from __future__ import annotations

import csv
import base64
import hashlib
import json
import mimetypes
import re
import os
import time
import urllib.error
import urllib.request
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .ooxml import _image_size

TEXT_EXTENSIONS = {".txt", ".md", ".rst", ".html", ".htm"}
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg"}


class ContentPackage:
    def __init__(self, root: str | Path, *, captions_path: str | Path | None = None):
        self.root = Path(root).resolve()
        if not self.root.exists():
            raise ValueError(f"Content package does not exist: {self.root}")
        self.captions = self._load_captions(captions_path)
        self.items = self._scan()

    @staticmethod
    def _load_captions(path: str | Path | None) -> dict[str, str]:
        if not path:
            return {}
        raw = json.loads(Path(path).read_text(encoding="utf-8"))
        if not isinstance(raw, dict) or not all(isinstance(key, str) and isinstance(value, str) for key, value in raw.items()):
            raise ValueError("Image captions must be a JSON object mapping relative paths to captions")
        return {key.replace("\\", "/"): value for key, value in raw.items()}

    def _scan(self) -> list[dict[str, Any]]:
        files = [self.root] if self.root.is_file() else sorted(path for path in self.root.rglob("*") if path.is_file())
        items: list[dict[str, Any]] = []
        for path in files:
            relative = path.name if self.root.is_file() else path.relative_to(self.root).as_posix()
            suffix = path.suffix.lower()
            try:
                if suffix in IMAGE_EXTENSIONS:
                    data = path.read_bytes()
                    size = _image_size(data, suffix)
                    items.append(
                        {
                            "id": f"asset-{len(items) + 1}",
                            "kind": "image",
                            "path": relative,
                            "mime_type": mimetypes.guess_type(path.name)[0],
                            "width": size[0] if size else None,
                            "height": size[1] if size else None,
                            "caption": self.captions.get(relative),
                            "sha256": hashlib.sha256(data).hexdigest(),
                        }
                    )
                elif suffix in TEXT_EXTENSIONS:
                    items.append(self._text_item(relative, path.read_text(encoding="utf-8", errors="replace")))
                elif suffix == ".json":
                    parsed = json.loads(path.read_text(encoding="utf-8"))
                    items.append(self._text_item(relative, json.dumps(parsed, ensure_ascii=False, indent=2)))
                elif suffix in {".csv", ".tsv"}:
                    delimiter = "\t" if suffix == ".tsv" else ","
                    with path.open("r", encoding="utf-8-sig", newline="") as stream:
                        rows = list(csv.reader(stream, delimiter=delimiter))
                    items.append({"id": f"asset-{len(items) + 1}", "kind": "table", "path": relative, "rows": rows[:100]})
                elif suffix == ".docx":
                    items.append(self._text_item(relative, self._extract_docx(path)))
                elif suffix == ".pptx":
                    items.append(self._text_item(relative, self._extract_pptx(path)))
                elif suffix == ".xlsx":
                    items.extend(self._extract_xlsx(path, relative, len(items)))
                elif suffix == ".pdf":
                    text = self._extract_pdf(path)
                    items.append(self._text_item(relative, text))
            except (OSError, ValueError, KeyError, zipfile.BadZipFile, ET.ParseError) as exc:
                items.append({"id": f"asset-{len(items) + 1}", "kind": "unavailable", "path": relative, "error": str(exc)})
        return items

    def _text_item(self, relative: str, text: str) -> dict[str, Any]:
        return {"id": "pending", "kind": "text", "path": relative, "text": text[:60000]}

    @staticmethod
    def _extract_docx(path: Path) -> str:
        with zipfile.ZipFile(path) as archive:
            root = ET.fromstring(archive.read("word/document.xml"))
        return "\n".join(node.text or "" for node in root.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"))

    @staticmethod
    def _extract_pptx(path: Path) -> str:
        with zipfile.ZipFile(path) as archive:
            slide_names = sorted(
                (name for name in archive.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)),
                key=lambda name: int(re.search(r"\d+", Path(name).stem).group()),
            )
            chunks = []
            for index, name in enumerate(slide_names, start=1):
                root = ET.fromstring(archive.read(name))
                text = " ".join(node.text or "" for node in root.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}t"))
                chunks.append(f"Slide {index}: {text}")
        return "\n".join(chunks)

    @staticmethod
    def _extract_xlsx(path: Path, relative: str, offset: int) -> list[dict[str, Any]]:
        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())
            shared: list[str] = []
            if "xl/sharedStrings.xml" in names:
                root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
                ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                shared = ["".join(node.text or "" for node in item.iter(f"{{{ns}}}t")) for item in root.findall(f"{{{ns}}}si")]
            sheets = sorted(name for name in names if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", name))
            output = []
            ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            for index, name in enumerate(sheets[:5], start=1):
                root = ET.fromstring(archive.read(name))
                rows = []
                for row in root.findall(f".//{{{ns}}}row")[:100]:
                    values = []
                    for cell in row.findall(f"{{{ns}}}c"):
                        value = cell.find(f"{{{ns}}}v")
                        raw = value.text if value is not None else ""
                        if cell.get("t") == "s" and raw.isdigit() and int(raw) < len(shared):
                            raw = shared[int(raw)]
                        values.append(raw)
                    rows.append(values)
                output.append({"id": f"asset-{offset + index}", "kind": "table", "path": f"{relative}#sheet{index}", "rows": rows})
        return output

    @staticmethod
    def _extract_pdf(path: Path) -> str:
        try:
            from pypdf import PdfReader
        except ImportError:
            return "PDF text extraction unavailable; install pypdf or provide a .txt companion file."
        return "\n".join(page.extract_text() or "" for page in PdfReader(str(path)).pages)

    def model_view(self) -> dict[str, Any]:
        items = []
        for index, item in enumerate(self.items, start=1):
            normalized = dict(item)
            normalized["id"] = f"asset-{index}"
            if normalized.get("kind") == "text":
                normalized["text"] = normalized.get("text", "")[:20000]
            if normalized.get("kind") == "table":
                normalized["rows"] = normalized.get("rows", [])[:25]
            items.append(normalized)
        return {"root_name": self.root.name, "items": items}

    def uncaptioned_images(self) -> list[str]:
        return [item["path"] for item in self.items if item.get("kind") == "image" and not item.get("caption")]

    def caption_images(
        self,
        *,
        model: str,
        base_url: str | None = None,
        api_key: str | None = None,
    ) -> dict[str, str]:
        """Caption images through an OpenAI-compatible VLM endpoint."""
        endpoint = (base_url or os.getenv("PRESENTATION_VLM_BASE_URL") or os.getenv("PRESENTATION_LLM_BASE_URL", "")).rstrip("/")
        token = (
            api_key
            or os.getenv("PRESENTATION_VLM_API_KEY")
            or os.getenv("PRESENTATION_LLM_API_KEY")
            or os.getenv("OPENROUTER_API_KEY", "")
        )
        if not endpoint:
            raise ValueError("PRESENTATION_VLM_BASE_URL or PRESENTATION_LLM_BASE_URL is required")
        produced: dict[str, str] = {}
        for item in self.items:
            if item.get("kind") != "image" or item.get("caption"):
                continue
            source = self.root if self.root.is_file() else self.root / item["path"]
            data = source.read_bytes()
            if len(data) > 8 * 1024 * 1024:
                item["caption"] = "Image omitted from VLM captioning because it exceeds 8 MB."
                continue
            data_url = f"data:{item['mime_type']};base64,{base64.b64encode(data).decode('ascii')}"
            payload = {
                "model": model,
                "temperature": 0,
                "messages": [
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "text",
                                "text": "Describe this image in one concise sentence for choosing whether it belongs on a business presentation slide. Mention the subject, visible data if any, and orientation. Do not infer facts that are not visible.",
                            },
                            {"type": "image_url", "image_url": {"url": data_url}},
                        ],
                    }
                ],
            }
            fallbacks = [
                value.strip()
                for value in os.getenv("PRESENTATION_VLM_FALLBACK_MODELS", "").split(",")
                if value.strip()
            ]
            models = list(dict.fromkeys([model, *fallbacks]))
            last_error = "unknown VLM error"
            caption = ""
            for attempt, candidate_model in enumerate(models, start=1):
                payload["model"] = candidate_model
                request = urllib.request.Request(
                    f"{endpoint}/chat/completions",
                    data=json.dumps(payload).encode("utf-8"),
                    headers={
                        "Content-Type": "application/json",
                        **({"Authorization": f"Bearer {token}"} if token else {}),
                    },
                    method="POST",
                )
                try:
                    with urllib.request.urlopen(request, timeout=180) as response:
                        result = json.load(response)
                    if result.get("error"):
                        raise ValueError(f"Provider error: {result['error']}")
                    caption = result["choices"][0]["message"]["content"].strip()
                    break
                except urllib.error.HTTPError as exc:
                    details = exc.read().decode("utf-8", errors="replace")[:1000]
                    last_error = f"HTTP {exc.code}: {details}"
                    if attempt < len(models):
                        time.sleep(5 * attempt)
                        continue
                    raise ValueError(f"VLM captioning failed for {item['path']}: {last_error}") from exc
                except (urllib.error.URLError, TimeoutError, json.JSONDecodeError, KeyError, IndexError, TypeError, ValueError) as exc:
                    last_error = str(exc)
                    if attempt < len(models):
                        time.sleep(2 * attempt)
                        continue
                    raise ValueError(f"VLM captioning failed for {item['path']}: {last_error}") from exc
            if not caption:
                raise ValueError(f"VLM captioning failed for {item['path']}: {last_error}")
            item["caption"] = caption
            self.captions[item["path"]] = caption
            produced[item["path"]] = caption
        return produced

    def write_captions(self, path: str | Path) -> None:
        target = Path(path)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(json.dumps(self.captions, ensure_ascii=False, indent=2), encoding="utf-8")
