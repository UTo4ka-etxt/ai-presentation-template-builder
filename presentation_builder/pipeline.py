from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Callable

from .audit import audit_pptx, write_audit
from .catalog import TemplateCatalog
from .content import ContentPackage
from .decompose import PptxDecomposer
from .ooxml import PresentationBuilder
from .planner import LlmPlanner

ProgressCallback = Callable[[str, int, str], None]


class GenerationPipeline:
    def __init__(self, *, model: str | None = None, vision_model: str | None = None):
        self.model = model
        self.vision_model = vision_model

    def run(
        self,
        *,
        template_path: str | Path,
        content_path: str | Path,
        brief: str,
        output_dir: str | Path,
        slide_count: int = 10,
        variants: int = 3,
        captions_path: str | Path | None = None,
        progress: ProgressCallback | None = None,
    ) -> dict[str, Any]:
        if variants != 3:
            raise ValueError("The service produces exactly three presentation variants")
        if not 3 <= slide_count <= 20:
            raise ValueError("slide_count must be between 3 and 20")
        notify = progress or (lambda stage, percent, message: None)
        output = Path(output_dir).resolve()
        output.mkdir(parents=True, exist_ok=True)
        package_dir = output / ".template-package"

        notify("decompose", 5, "Разбираем структуру шаблона")
        decomposition = PptxDecomposer().decompose(template_path, package_dir)
        if not decomposition["slides_available"]:
            raise ValueError("The template contains no editable slide layouts")

        notify("content", 20, "Индексируем контент-пакет")
        content = ContentPackage(content_path, captions_path=captions_path)
        if self.vision_model:
            notify("vision", 25, "Нейросеть анализирует изображения")
            content.caption_images(model=self.vision_model)
            content.write_captions(output / "content-captions.json")
        elif content.uncaptioned_images():
            raise ValueError(
                "The content package contains images, but no vision model or captions were provided. "
                "Configure PRESENTATION_VLM_MODEL so images can be selected by meaning."
            )

        catalog = TemplateCatalog(package_dir)
        planner = LlmPlanner(catalog, model=self.model)
        builder = PresentationBuilder(package_dir)
        plans: list[dict[str, Any]] = []
        results = []
        for variant in range(1, variants + 1):
            base = 30 + (variant - 1) * 20
            notify("planning", base, f"Нейросеть проектирует вариант {variant} из {variants}")
            plan = planner.plan(
                brief,
                content,
                variant=variant,
                slide_count=slide_count,
                previous_plans=plans,
            )
            plans.append(plan)
            plan_path = output / f"variant-{variant}.plan.json"
            plan_path.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
            notify("build", base + 10, f"Собираем вариант {variant} в PPTX")
            pptx_path = output / f"variant-{variant}.pptx"
            build_report = builder.build(plan, pptx_path)
            audit = audit_pptx(pptx_path)
            audit_path = output / f"variant-{variant}.audit.json"
            write_audit(audit, audit_path)
            if not audit["ok"]:
                raise ValueError(f"Variant {variant} failed deterministic audit")
            results.append(
                {
                    "variant": variant,
                    "pptx": str(pptx_path),
                    "plan": str(plan_path),
                    "audit": str(audit_path),
                    "build": build_report,
                }
            )

        notify("complete", 100, "Три варианта презентации готовы")
        report = {
            "status": "complete",
            "template": str(Path(template_path).resolve()),
            "content_package": str(Path(content_path).resolve()),
            "decomposition": decomposition,
            "model": planner.model,
            "vision_model": self.vision_model,
            "slide_count": slide_count,
            "variants": results,
        }
        (output / "run-report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        return report
