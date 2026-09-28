from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path
from typing import Any

from dotenv import load_dotenv

from .audit import audit_pptx, write_audit
from .catalog import TemplateCatalog
from .content import ContentPackage
from .decompose import DecompositionError, PptxDecomposer
from .ooxml import BuildError, PresentationBuilder
from .planner import LlmPlanner, PlannerError
from .pipeline import GenerationPipeline

load_dotenv()


def _read_json(path: str | Path) -> dict[str, Any]:
    return json.loads(Path(path).read_text(encoding="utf-8"))


def _dump(value: Any) -> None:
    print(json.dumps(value, ensure_ascii=False, indent=2))


def _prepare_content(args: argparse.Namespace, captions_output: Path) -> ContentPackage:
    content = ContentPackage(args.content_package, captions_path=args.captions)
    vision_model = args.vision_model or os.getenv("PRESENTATION_VLM_MODEL")
    if vision_model:
        content.caption_images(model=vision_model)
        content.write_captions(captions_output)
    elif content.uncaptioned_images():
        raise ValueError(
            "The content package contains images, but no vision model or captions were provided. "
            "Configure PRESENTATION_VLM_MODEL so images can be selected by meaning."
        )
    return content


def cmd_inspect(args: argparse.Namespace) -> int:
    catalog = TemplateCatalog(args.package)
    view = catalog.planner_view()
    if args.output:
        Path(args.output).write_text(json.dumps(view, ensure_ascii=False, indent=2), encoding="utf-8")
    else:
        _dump(view)
    return 0


def cmd_plan(args: argparse.Namespace) -> int:
    brief = Path(args.brief).read_text(encoding="utf-8")
    content = _prepare_content(args, Path(args.output).with_suffix(".captions.json"))
    catalog = TemplateCatalog(args.package)
    planner = LlmPlanner(catalog, model=args.model)
    plan = planner.plan(brief, content, variant=args.variant, slide_count=args.slides)
    Path(args.output).write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
    _dump({"output": str(Path(args.output).resolve()), "slides": len(plan["slides"])})
    return 0


def cmd_build(args: argparse.Namespace) -> int:
    builder = PresentationBuilder(args.package)
    report = builder.build(_read_json(args.plan), args.output, clear_unfilled=not args.keep_placeholders)
    audit = audit_pptx(args.output)
    audit_path = str(Path(args.output).resolve()) + ".audit.json"
    write_audit(audit, audit_path)
    report["audit"] = audit_path
    report["audit_ok"] = audit["ok"]
    _dump(report)
    return 0 if audit["ok"] else 2


def cmd_generate(args: argparse.Namespace) -> int:
    output_dir = Path(args.output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    brief = Path(args.brief).read_text(encoding="utf-8")
    content = _prepare_content(args, output_dir / "content-captions.json")
    catalog = TemplateCatalog(args.package)
    planner = LlmPlanner(catalog, model=args.model)
    outputs = []
    previous_plans = []
    for variant in range(1, args.variants + 1):
        plan = planner.plan(brief, content, variant=variant, slide_count=args.slides, previous_plans=previous_plans)
        previous_plans.append(plan)
        plan_path = output_dir / f"variant-{variant}.plan.json"
        pptx_path = output_dir / f"variant-{variant}.pptx"
        plan_path.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
        build = PresentationBuilder(args.package).build(plan, pptx_path)
        audit = audit_pptx(pptx_path)
        write_audit(audit, str(pptx_path) + ".audit.json")
        outputs.append({"variant": variant, "plan": str(plan_path), "pptx": str(pptx_path), "audit_ok": audit["ok"]})
    _dump({"outputs": outputs})
    return 0 if all(item["audit_ok"] for item in outputs) else 2


def cmd_decompose(args: argparse.Namespace) -> int:
    report = PptxDecomposer().decompose(args.template, args.output_dir)
    _dump(report)
    return 0


def cmd_run(args: argparse.Namespace) -> int:
    brief = Path(args.brief).read_text(encoding="utf-8")
    pipeline = GenerationPipeline(
        model=args.model,
        vision_model=args.vision_model or os.getenv("PRESENTATION_VLM_MODEL"),
    )
    report = pipeline.run(
        template_path=args.template,
        content_path=args.content_package,
        brief=brief,
        output_dir=args.output_dir,
        slide_count=args.slides,
        variants=args.variants,
        captions_path=args.captions,
    )
    _dump(report)
    return 0


def cmd_serve(args: argparse.Namespace) -> int:
    import uvicorn

    uvicorn.run("presentation_builder.web:app", host=args.host, port=args.port, reload=args.reload)
    return 0


def cmd_audit(args: argparse.Namespace) -> int:
    report = audit_pptx(args.pptx)
    if args.output:
        write_audit(report, args.output)
    _dump(report)
    return 0 if report["ok"] else 2


def parser() -> argparse.ArgumentParser:
    root = argparse.ArgumentParser(description="Assemble native PPTX files from a decomposed template package")
    commands = root.add_subparsers(dest="command", required=True)
    inspect = commands.add_parser("inspect", help="Print the LLM-facing template catalog")
    inspect.add_argument("--package", default=".")
    inspect.add_argument("--output")
    inspect.set_defaults(func=cmd_inspect)

    decompose = commands.add_parser("decompose", help="Decompose an arbitrary PPTX template")
    decompose.add_argument("--template", required=True)
    decompose.add_argument("--output-dir", required=True)
    decompose.set_defaults(func=cmd_decompose)

    plan = commands.add_parser("plan", help="Create a slide plan through an OpenAI-compatible model")
    plan.add_argument("--package", default=".")
    plan.add_argument("--brief", required=True)
    plan.add_argument("--content-package", required=True)
    plan.add_argument("--captions", help="Optional JSON mapping image paths to captions")
    plan.add_argument("--vision-model", help="Optional OpenAI-compatible VLM used to caption images")
    plan.add_argument("--output", required=True)
    plan.add_argument("--slides", type=int, default=10)
    plan.add_argument("--variant", type=int, choices=(1, 2, 3), default=1)
    plan.add_argument("--model")
    plan.set_defaults(func=cmd_plan)

    build = commands.add_parser("build", help="Build a PPTX from an existing JSON plan")
    build.add_argument("--package", default=".")
    build.add_argument("--plan", required=True)
    build.add_argument("--output", required=True)
    build.add_argument("--keep-placeholders", action="store_true")
    build.set_defaults(func=cmd_build)

    generate = commands.add_parser("generate", help="Plan and build exactly three AI variants")
    generate.add_argument("--package", default=".")
    generate.add_argument("--brief", required=True)
    generate.add_argument("--content-package", required=True)
    generate.add_argument("--captions", help="Optional JSON mapping image paths to captions")
    generate.add_argument("--vision-model", help="Optional OpenAI-compatible VLM used to caption images")
    generate.add_argument("--output-dir", required=True)
    generate.add_argument("--slides", type=int, default=10)
    generate.add_argument("--variants", type=int, choices=(3,), default=3)
    generate.add_argument("--model")
    generate.set_defaults(func=cmd_generate)

    run = commands.add_parser("run", help="Full cycle: decompose template, plan and build variants")
    run.add_argument("--template", required=True)
    run.add_argument("--content-package", required=True)
    run.add_argument("--brief", required=True)
    run.add_argument("--captions", help="Optional JSON mapping image paths to captions")
    run.add_argument("--vision-model", help="Optional OpenAI-compatible VLM used to caption images")
    run.add_argument("--output-dir", required=True)
    run.add_argument("--slides", type=int, default=10)
    run.add_argument("--variants", type=int, choices=(3,), default=3)
    run.add_argument("--model")
    run.set_defaults(func=cmd_run)

    serve = commands.add_parser("serve", help="Start the web service")
    serve.add_argument("--host", default=os.getenv("PRESENTATION_HOST", "0.0.0.0"))
    serve.add_argument("--port", type=int, default=int(os.getenv("PRESENTATION_PORT", "8000")))
    serve.add_argument("--reload", action="store_true")
    serve.set_defaults(func=cmd_serve)

    audit = commands.add_parser("audit", help="Run deterministic package checks")
    audit.add_argument("pptx")
    audit.add_argument("--output")
    audit.set_defaults(func=cmd_audit)
    return root


def main(argv: list[str] | None = None) -> int:
    try:
        args = parser().parse_args(argv)
        return args.func(args)
    except (BuildError, DecompositionError, PlannerError, ValueError, OSError, json.JSONDecodeError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
