from __future__ import annotations

import hashlib
import json
import os
import tempfile
import threading
import unittest
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from xml.etree import ElementTree as ET

from presentation_builder.audit import audit_pptx
from presentation_builder.catalog import TemplateCatalog
from presentation_builder.content import ContentPackage
from presentation_builder.decompose import PptxDecomposer
from presentation_builder.ooxml import A, P, R, PresentationBuilder
from presentation_builder.pipeline import GenerationPipeline

ROOT = Path(__file__).parents[1]


class CatalogTests(unittest.TestCase):
    def test_all_valid_slides_are_loaded(self) -> None:
        catalog = TemplateCatalog(ROOT)
        self.assertEqual(len(catalog.templates), 20)
        self.assertIn("title_block.title", catalog.get("s1").slots)
        self.assertIn("table_1.data_table", catalog.get("s14").slots)

    def test_invalid_slot_is_rejected(self) -> None:
        catalog = TemplateCatalog(ROOT)
        errors = catalog.validate_plan(
            {"slides": [{"template_slide_id": "s1", "values": {"missing.slot": "x"}}]}
        )
        self.assertTrue(errors)

    def test_content_package_extracts_text_and_table(self) -> None:
        package = ContentPackage(ROOT / "examples" / "content-package")
        kinds = {item["kind"] for item in package.model_view()["items"]}
        self.assertIn("text", kinds)
        self.assertIn("table", kinds)


class BuilderTests(unittest.TestCase):
    def test_clones_duplicate_layouts_and_keeps_native_table(self) -> None:
        source = ROOT / "source.pptx"
        before = hashlib.sha256(source.read_bytes()).hexdigest()
        plan = {
            "title": "Test deck",
            "slides": [
                {
                    "template_slide_id": "s1",
                    "values": {
                        "title_block.title": "Первая новая презентация",
                        "speaker_block.speaker_name_title": "Команда тестирования",
                    },
                },
                {
                    "template_slide_id": "s1",
                    "values": {
                        "title_block.title": "Повторное использование макета",
                        "speaker_block.speaker_name_title": "Нативный PPTX",
                    },
                },
                {
                    "template_slide_id": "s14",
                    "values": {
                        "title_1.slide_title": "Сводная таблица",
                        "table_1.data_table": [["Показатель", "Значение"], ["Скорость", "до 5 минут"]],
                    },
                },
            ],
        }
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "result.pptx"
            result = PresentationBuilder(ROOT).build(plan, output)
            self.assertEqual(result["slide_count"], 3)
            report = audit_pptx(output)
            self.assertTrue(report["ok"], report)
            self.assertEqual(report["slide_count"], 3)
            with zipfile.ZipFile(output) as archive:
                presentation = ET.fromstring(archive.read("ppt/presentation.xml"))
                rels = ET.fromstring(archive.read("ppt/_rels/presentation.xml.rels"))
                rel_map = {rel.get("Id"): rel.get("Target") for rel in rels}
                slide_nodes = list(presentation.find(f"{{{P}}}sldIdLst"))
                targets = [rel_map[node.get(f"{{{R}}}id")] for node in slide_nodes]
                self.assertEqual(len(set(targets)), 3)
                slide_xml = archive.read("ppt/" + targets[2]).decode("utf-8")
                self.assertIn("Сводная таблица", slide_xml)
                self.assertIn("до 5 минут", slide_xml)
                self.assertIn("<a:tbl", slide_xml)
        self.assertEqual(hashlib.sha256(source.read_bytes()).hexdigest(), before)

    def test_replaces_image_slot_with_content_asset(self) -> None:
        asset = "assets/66be632c9ce53342875ef332ec7e780e9749ced850a09bb0f878356347d1fcb6.png"
        plan = {
            "metadata": {"content_root": str(ROOT)},
            "slides": [
                {
                    "template_slide_id": "s4",
                    "values": {
                        "title.slide_title": "Контент-пакет",
                        "section_heading.section_heading": "Материалы",
                        "list_item_1.item_icon": {"path": asset},
                        "list_item_1.item_text": "Изображение выбрано по смыслу",
                    },
                }
            ],
        }
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "image.pptx"
            PresentationBuilder(ROOT).build(plan, output)
            with zipfile.ZipFile(output) as archive:
                generated = [name for name in archive.namelist() if name.startswith("ppt/media/generated_")]
                self.assertEqual(len(generated), 1)
                self.assertEqual(archive.read(generated[0]), (ROOT / asset).read_bytes())

    def test_full_core_cycle_from_raw_template_and_ai_plan(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            package_dir = Path(directory) / "package"
            report = PptxDecomposer().decompose(ROOT / "source.pptx", package_dir)
            self.assertEqual(report["slides_total"], 29)
            self.assertGreater(report["slides_available"], 0)
            catalog = TemplateCatalog(package_dir)
            slide_id = next(iter(catalog.templates))
            template = catalog.get(slide_id)
            values = {
                slot.key: "Тестовый контент"
                for slot in template.slots.values()
                if slot.content_type == "text"
            }
            plan = {"slides": [{"template_slide_id": slide_id, "values": values}]}
            output = Path(directory) / "variant.pptx"
            PresentationBuilder(package_dir).build(plan, output)
            audit = audit_pptx(output)
            self.assertTrue(audit["ok"], audit)

    def test_network_ai_pipeline_produces_exactly_three_variants(self) -> None:
        requests_seen: list[int] = []

        class Handler(BaseHTTPRequestHandler):
            def do_POST(self) -> None:  # noqa: N802 - stdlib API
                length = int(self.headers["Content-Length"])
                request = json.loads(self.rfile.read(length))
                user = json.loads(request["messages"][1]["content"])
                requests_seen.append(user["variant"])
                templates = user["template_catalog"]["slides"]
                assets = user["content_package"]["items"]
                slides = []
                for index in range(user["requested_slide_count"]):
                    template = templates[(index + user["variant"] - 1) % len(templates)]
                    values = {}
                    for slot in template["slots"]:
                        if slot["type"] == "text":
                            values[slot["key"]] = f"Тест-{user['variant']}-{index + 1}"
                    slides.append(
                        {
                            "template_slide_id": template["slide_id"],
                            "values": values,
                            "source_refs": [assets[0]["id"]],
                        }
                    )
                response = {
                    "choices": [
                        {"message": {"content": json.dumps({"title": "AI test", "slides": slides}, ensure_ascii=False)}}
                    ]
                }
                body = json.dumps(response, ensure_ascii=False).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)

            def log_message(self, format: str, *args: object) -> None:
                return

        server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        previous = os.environ.get("PRESENTATION_LLM_BASE_URL")
        os.environ["PRESENTATION_LLM_BASE_URL"] = f"http://127.0.0.1:{server.server_port}/v1"
        try:
            with tempfile.TemporaryDirectory() as directory:
                output = Path(directory) / "result"
                report = GenerationPipeline(model="test-network-model").run(
                    template_path=ROOT / "source.pptx",
                    content_path=ROOT / "examples" / "content-package",
                    brief="Собрать тестовую презентацию",
                    output_dir=output,
                    slide_count=3,
                    variants=3,
                )
                self.assertEqual(requests_seen, [1, 2, 3])
                self.assertEqual(len(report["variants"]), 3)
                for variant in range(1, 4):
                    path = output / f"variant-{variant}.pptx"
                    self.assertTrue(path.exists())
                    self.assertTrue(audit_pptx(path)["ok"])
        finally:
            server.shutdown()
            server.server_close()
            if previous is None:
                os.environ.pop("PRESENTATION_LLM_BASE_URL", None)
            else:
                os.environ["PRESENTATION_LLM_BASE_URL"] = previous


if __name__ == "__main__":
    unittest.main()
