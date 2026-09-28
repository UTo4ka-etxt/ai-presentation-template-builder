from __future__ import annotations

import json
import http.client
import os
import time
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any
from collections import defaultdict

from .catalog import TemplateCatalog
from .content import ContentPackage


class PlannerError(RuntimeError):
    pass


class LlmPlanner:
    """Small OpenAI-compatible client, suitable for Qwen/OpenRouter or VK inference."""

    def __init__(
        self,
        catalog: TemplateCatalog,
        *,
        base_url: str | None = None,
        api_key: str | None = None,
        model: str | None = None,
        prompt_path: str | Path | None = None,
    ):
        self.catalog = catalog
        self.base_url = (base_url or os.getenv("PRESENTATION_LLM_BASE_URL", "")).rstrip("/")
        self.api_key = api_key or os.getenv("PRESENTATION_LLM_API_KEY") or os.getenv("OPENROUTER_API_KEY", "")
        self.model = model or os.getenv("PRESENTATION_LLM_MODEL") or catalog.model
        prompt_file = Path(prompt_path) if prompt_path else Path(__file__).parent / "prompts" / "planner_system.txt"
        self.system_prompt = prompt_file.read_text(encoding="utf-8")

    def plan(
        self,
        brief: str,
        content: ContentPackage,
        *,
        variant: int = 1,
        slide_count: int = 10,
        previous_plans: list[dict[str, Any]] | None = None,
    ) -> dict[str, Any]:
        if not self.base_url:
            raise PlannerError("PRESENTATION_LLM_BASE_URL is required for AI planning")
        user_payload = {
            "brief": brief,
            "variant": variant,
            "requested_slide_count": slide_count,
            "template_catalog": self._variant_catalog(variant),
            "content_package": content.model_view(),
            "previous_variants": [
                [slide.get("template_slide_id") for slide in previous.get("slides", [])]
                for previous in (previous_plans or [])
            ],
        }
        messages = [
            {"role": "system", "content": self.system_prompt},
            {"role": "user", "content": json.dumps(user_payload, ensure_ascii=False)},
        ]
        payload: dict[str, Any] = {
            "model": self.model,
            "temperature": float(os.getenv("PRESENTATION_LLM_TEMPERATURE", "0.2")),
            "max_tokens": int(os.getenv("PRESENTATION_LLM_MAX_TOKENS", "8192")),
            "messages": messages,
        }
        if "openrouter.ai" in self.base_url and os.getenv("PRESENTATION_LLM_REASONING", "0") == "0":
            payload["reasoning"] = {"enabled": False}
        if os.getenv("PRESENTATION_LLM_JSON_MODE", "1") != "0":
            payload["response_format"] = {"type": "json_object"}
        fallback_models = [
            item.strip()
            for item in os.getenv("PRESENTATION_LLM_FALLBACK_MODELS", "").split(",")
            if item.strip()
        ]
        models = list(dict.fromkeys([self.model, *fallback_models]))
        model_index = 0
        used_model = self.model
        last_error = "unknown planning error"
        plan: dict[str, Any] | None = None
        attempts = max(3, len(models) + 2)
        for attempt in range(1, attempts + 1):
            payload["messages"] = messages
            payload["model"] = models[model_index]
            request = urllib.request.Request(
                f"{self.base_url}/chat/completions",
                data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
                headers={
                    "Content-Type": "application/json",
                    **({"Authorization": f"Bearer {self.api_key}"} if self.api_key else {}),
                },
                method="POST",
            )
            try:
                with urllib.request.urlopen(request, timeout=240) as response:
                    result = json.load(response)
                if result.get("error"):
                    last_error = f"Provider error: {result['error']}"
                    if model_index + 1 < len(models):
                        model_index += 1
                        continue
                    raise ValueError(last_error)
                raw_content = result["choices"][0]["message"]["content"]
                if isinstance(raw_content, list):
                    raw_content = "".join(part.get("text", "") for part in raw_content if isinstance(part, dict))
                if not isinstance(raw_content, str):
                    raise ValueError("message content is not text")
                start, end = raw_content.find("{"), raw_content.rfind("}")
                if start < 0 or end < start:
                    raise ValueError("JSON object not found")
                candidate = json.loads(raw_content[start : end + 1])
                errors = self._validate_grounded_plan(candidate, content, slide_count)
                if not errors:
                    plan = candidate
                    used_model = models[model_index]
                    break
                last_error = "; ".join(errors)
                messages.extend(
                    [
                        {"role": "assistant", "content": raw_content},
                        {
                            "role": "user",
                            "content": (
                                "Пересобери и верни ВЕСЬ JSON-документ целиком, а не исправленный фрагмент. "
                                f"Массив slides обязан содержать ровно {slide_count} объектов. "
                                "Не сокращай и не обрывай массив. Ошибки проверки: " + last_error
                            ),
                        },
                    ]
                )
            except urllib.error.HTTPError as exc:
                details = exc.read().decode("utf-8", errors="replace")[:1000]
                last_error = f"HTTP {exc.code}: {details}"
                if exc.code in {402, 408, 429, 500, 502, 503, 504} and model_index + 1 < len(models):
                    model_index += 1
                if attempt < attempts:
                    time.sleep(5 * attempt if exc.code == 429 else attempt)
            except (
                urllib.error.URLError,
                http.client.HTTPException,
                TimeoutError,
                json.JSONDecodeError,
                KeyError,
                IndexError,
                TypeError,
                ValueError,
            ) as exc:
                last_error = str(exc)
                if attempt < attempts:
                    time.sleep(attempt)
        if plan is None:
            raise PlannerError(f"LLM could not produce a valid plan after {attempts} attempts: {last_error}")
        plan.setdefault("metadata", {})
        plan["metadata"].update(
            {"model": used_model, "variant": variant, "content_root": str(content.root)}
        )
        return plan

    def _variant_catalog(self, variant: int) -> dict[str, Any]:
        """Keep prompts compact while rotating equivalent layouts across variants."""
        view = self.catalog.planner_view()
        slides = view.get("slides", [])
        limit = max(8, int(os.getenv("PRESENTATION_LLM_LAYOUT_LIMIT", "18")))
        if len(slides) <= limit:
            return view
        groups: dict[str, list[dict[str, Any]]] = defaultdict(list)
        for slide in slides:
            groups[str(slide.get("role") or "other")].append(slide)
        selected: list[dict[str, Any]] = []
        offset = max(0, variant - 1)
        ordered_groups = list(groups.values())
        for round_index in range(max(len(group) for group in ordered_groups)):
            for group in ordered_groups:
                if len(selected) >= limit:
                    break
                if round_index >= len(group):
                    continue
                index = (round_index + offset) % len(group)
                candidate = group[index]
                if candidate not in selected:
                    selected.append(candidate)
            if len(selected) >= limit:
                break
        return {**view, "slides": selected}

    def _validate_grounded_plan(self, plan: dict[str, Any], content: ContentPackage, slide_count: int) -> list[str]:
        errors = self.catalog.validate_plan(plan)
        slides = plan.get("slides", [])
        if isinstance(slides, list) and len(slides) != slide_count:
            errors.append(f"Expected exactly {slide_count} slides, got {len(slides)}")
        content_view = content.model_view()
        asset_ids = {item["id"] for item in content_view["items"]}
        image_paths = {item["path"] for item in content_view["items"] if item.get("kind") == "image"}
        for index, slide in enumerate(slides if isinstance(slides, list) else [], start=1):
            refs = slide.get("source_refs")
            if not isinstance(refs, list) or not refs:
                errors.append(f"Slide {index}: source_refs must cite at least one content asset")
            elif any(ref not in asset_ids for ref in refs):
                errors.append(f"Slide {index}: source_refs contains an unknown asset")
            template_id = slide.get("template_slide_id")
            if template_id not in self.catalog.templates:
                continue
            template = self.catalog.get(template_id)
            for key, value in slide.get("values", {}).items():
                if isinstance(value, dict) and "path" in value and value["path"] not in image_paths:
                    errors.append(f"Slide {index}: image path {value['path']!r} is not in the content package")
                slot = template.slots.get(key)
                if slot and slot.content_type == "text" and isinstance(value, str):
                    maximum = self.catalog.slot_capacity(template_id, slot).get("max_words")
                    if maximum and len(value.split()) > maximum:
                        errors.append(f"Slide {index}: slot {key} exceeds max_words={maximum}")
        return errors
