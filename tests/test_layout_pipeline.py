import hashlib
import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import zipfile
import xml.etree.ElementTree as ET
import pytest
from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "skills/origin-word-figure-layout/scripts"
sys.path.insert(0, str(SCRIPTS))
import compose_origin_word_figure as core

def test_source_edge_is_detected(tmp_path):
    source = tmp_path / "edge.png"
    image = Image.new("RGB", (800,650), "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle((100,100,700,500), outline="black", width=4)
    draw.line((0,0,40,0), fill="black", width=4)
    image.save(source)
    assert core.analyze_panel(source, source, None).touches_source_edge

def test_grid_rejects_incomplete_rows(tmp_path):
    class Args:
        input_dir = tmp_path
        files = None
        cols = 2
    Image.new("RGB", (50,50), "white").save(tmp_path / "one.png")
    with pytest.raises(RuntimeError, match="rectangular grid"):
        core.resolve_sources(Args())

def test_synthetic_composition_preserves_sources_common_scale_and_anchors(tmp_path):
    spec = importlib.util.spec_from_file_location("demo_generator", ROOT / "examples/generate_panels.py")
    generator = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(generator)
    panels = tmp_path / "panels"
    generator.generate(panels)
    before = {p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in panels.iterdir()}
    preflight = tmp_path / "preflight"
    subprocess.run([sys.executable, str(SCRIPTS / "preflight_origin_panels.py"), "--input-dir", str(panels), "--cols", "2", "--qa-dir", str(preflight)], check=True, capture_output=True)
    assert (preflight / "preflight-contact-sheet.png").is_file()
    qa = tmp_path / "qa"
    docx = tmp_path / "figure.docx"
    subprocess.run([sys.executable, str(SCRIPTS / "compose_origin_word_figure.py"), "--input-dir", str(panels), "--cols", "2", "--output", str(docx), "--qa-dir", str(qa)], check=True, capture_output=True)
    report = json.loads((qa / "layout-analysis.json").read_text(encoding="utf-8"))
    assert len(report["panels"]) == 4
    assert report["edge_content_warnings"] == []
    assert report["common_scale_in_per_preview_pixel"] > 0
    assert before == {p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in panels.iterdir()}
    with zipfile.ZipFile(docx) as package:
        document = ET.fromstring(package.read("word/document.xml"))
        ns = {"wp":"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"}
        anchors = document.findall(".//wp:anchor", ns)
        assert len(anchors) == 4
        assert not document.findall(".//wp:inline", ns)
        for anchor, panel in zip(anchors, report["panels"]):
            extent = anchor.find("wp:extent", ns)
            assert int(extent.attrib["cx"]) / panel["visible_size_px"][0] == pytest.approx(report["common_scale_in_per_preview_pixel"] * 914400, abs=1)
        assert len([name for name in package.namelist() if name.startswith("word/media/")]) == 4
