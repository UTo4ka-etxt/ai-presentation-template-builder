from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class Slot:
    key: str
    object_id: str
    role: str
    content_type: str
    paragraph_indices: tuple[int, ...] | None


@dataclass(frozen=True)
class SlideTemplate:
    slide_id: str
    slide_role: str
    needs_review: bool
    source_part: str
    slots: dict[str, Slot]


class TemplateCatalog:
    def __init__(self, package_dir: str | Path):
        self.package_dir = Path(package_dir).resolve()
        self.catalog_data = self._read_json(self.package_dir / "catalog.json")
        self.template_data = self._read_json(self.package_dir / "template.json")
        self.templates: dict[str, SlideTemplate] = {}
        self.object_catalogs: dict[str, dict[str, dict[str, Any]]] = {}
        # Freshly decomposed customer templates contain no trustworthy runtime
        # model hint. Use the production planner recommended for this service;
        # callers can still override it through PRESENTATION_LLM_MODEL.
        self.model = "Qwen/Qwen3-30B-A3B-Instruct-2507"
        self.prompt_version = "decompose_v5"
        self._load()

    @staticmethod
    def _read_json(path: Path) -> dict[str, Any]:
        with path.open("r", encoding="utf-8") as stream:
            return json.load(stream)

    def _load(self) -> None:
        slide_parts = {
            slide["id"]: slide["source_part"].lstrip("/")
            for slide in self.template_data["slides"]
        }
        for entry in self.catalog_data["slides"]:
            semantics = self._read_json(self.package_dir / entry["semantics"])
            result = semantics["result"]
            detected_model = semantics.get("model")
            if detected_model and not detected_model.startswith("deterministic-"):
                self.model = detected_model
            self.prompt_version = semantics.get("prompt_version", self.prompt_version)
            objects = semantics["object_catalog"]
            if isinstance(objects, list):
                objects = {item["id"]: item for item in objects}
            self.object_catalogs[entry["slide_id"]] = objects
            slots: dict[str, Slot] = {}
            for component in result["components"]:
                role_counts: dict[str, int] = {}
                for raw in component.get("slots", []):
                    role = raw["role"]
                    role_counts[role] = role_counts.get(role, 0) + 1
                    suffix = f"#{role_counts[role]}" if role_counts[role] > 1 else ""
                    key = f"{component['id']}.{role}{suffix}"
                    indices = raw.get("paragraph_indices")
                    slots[key] = Slot(
                        key=key,
                        object_id=raw["object_id"],
                        role=role,
                        content_type=raw["content_type"],
                        paragraph_indices=tuple(indices) if indices is not None else None,
                    )
            slide_id = entry["slide_id"]
            self.templates[slide_id] = SlideTemplate(
                slide_id=slide_id,
                slide_role=result["slide_role"],
                needs_review=bool(entry.get("needs_review")),
                source_part=slide_parts[slide_id],
                slots=slots,
            )

    def get(self, slide_id: str) -> SlideTemplate:
        try:
            return self.templates[slide_id]
        except KeyError as exc:
            raise ValueError(f"Unknown or unavailable template slide: {slide_id}") from exc

    def object(self, slide_id: str, object_id: str) -> dict[str, Any]:
        try:
            return self.object_catalogs[slide_id][object_id]
        except KeyError as exc:
            raise ValueError(f"Object {object_id!r} is not present on {slide_id}") from exc

    def planner_view(self) -> dict[str, Any]:
        slides = []
        for template in self.templates.values():
            slides.append(
                {
                    "slide_id": template.slide_id,
                    "role": template.slide_role,
                    "needs_review": template.needs_review,
                    "slots": [
                        {
                            "key": slot.key,
                            "type": slot.content_type,
                            "original_text": self.object(template.slide_id, slot.object_id).get("text"),
                            **self.slot_capacity(template.slide_id, slot),
                        }
                        for slot in template.slots.values()
                    ],
                }
            )
        return {"schema_version": "0.1", "slides": slides}

    def slot_capacity(self, slide_id: str, slot: Slot) -> dict[str, int]:
        if slot.content_type != "text":
            return {}
        obj = self.object(slide_id, slot.object_id)
        geometry = obj.get("geometry") or {}
        sizes = [
            run.get("font_size_pt")
            for paragraph in obj.get("paragraphs", [])
            for run in paragraph.get("runs", [])
            if run.get("font_size_pt")
        ]
        font_size = max(sizes, default=32 if "title" in slot.role else 18)
        width = int(geometry.get("width") or 0)
        height = int(geometry.get("height") or 0)
        if width <= 0 or height <= 0:
            max_words = 8 if "title" in slot.role else 18
        else:
            chars_per_line = max(3, int(width / (font_size * 12700 * 0.58)))
            lines = max(1, int(height / (font_size * 12700 * 1.2)))
            max_words = max(2, int(chars_per_line * lines / 7))
            if "title" in slot.role:
                max_words = min(max_words, 12)
            elif "label" in slot.role:
                max_words = min(max_words, 8)
            else:
                max_words = min(max_words, 42)
        return {"max_words": max_words, "max_chars": max_words * 8}

    def validate_plan(self, plan: dict[str, Any]) -> list[str]:
        errors: list[str] = []
        slides = plan.get("slides")
        if not isinstance(slides, list) or not slides:
            return ["Plan must contain a non-empty slides array"]
        for index, slide in enumerate(slides, start=1):
            if not isinstance(slide, dict):
                errors.append(f"Slide {index} must be an object")
                continue
            slide_id = slide.get("template_slide_id")
            if slide_id not in self.templates:
                errors.append(f"Slide {index}: unknown template_slide_id {slide_id!r}")
                continue
            values = slide.get("values", {})
            if not isinstance(values, dict):
                errors.append(f"Slide {index}: values must be an object")
                continue
            valid_keys = self.templates[slide_id].slots
            for key in values:
                if key not in valid_keys:
                    errors.append(f"Slide {index}: unknown slot {key!r} for {slide_id}")
        return errors
