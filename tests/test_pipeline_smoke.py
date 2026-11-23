"""Smoke tests proving the pipeline runs against the pinned dependency
versions, and that the pins stay exact."""

import importlib.util
import tempfile
import unittest
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent


def load(module_file):
    spec = importlib.util.spec_from_file_location(
        module_file.replace("-", "_"), REPO / ("%s.py" % module_file))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


class RequirementsTests(unittest.TestCase):
    def test_all_dependencies_pinned_exactly(self):
        lines = [line.strip()
                 for line in (REPO / "requirements.txt").read_text(
                     encoding="utf-8").splitlines()
                 if line.strip() and not line.startswith("#")]
        self.assertTrue(lines)
        for line in lines:
            self.assertIn("==", line, "unpinned dependency: %s" % line)


class InjectTests(unittest.TestCase):
    def test_inject_text_into_template_roundtrip(self):
        from pptx import Presentation
        from pptx.util import Inches

        inject = load("inject_lesson")
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "template.pptx"
            dst = Path(tmp) / "out.pptx"
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            box = slide.shapes.add_textbox(Inches(1), Inches(1),
                                           Inches(4), Inches(1))
            box.text_frame.text = "placeholder"
            prs.save(str(src))

            inject.inject_text_into_template(
                str(src), {0: {0: "درس نهم"}}, str(dst))

            out = Presentation(str(dst))
            shape = list(out.slides[0].shapes)[0]
            self.assertEqual(shape.text_frame.text, "درس نهم")
            p_pr = shape.text_frame.paragraphs[0]._p.get_or_add_pPr()
            self.assertEqual(p_pr.get("rtl"), "1")


class ParseTests(unittest.TestCase):
    def test_extract_and_parse_tags_from_docx(self):
        import docx

        parser = load("02_parse_word_to_json")
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "lesson.docx"
            document = docx.Document()
            document.add_paragraph("[S1-B1]")
            document.add_paragraph("النص الأول")
            document.add_paragraph("[S1-B2]")
            document.add_paragraph("النص الثاني")
            document.save(str(path))

            opened = docx.Document(str(path))
            full_text = parser.extract_full_text_from_docx(opened)
            mapping = parser.parse_tags_from_text(full_text)

        self.assertEqual(mapping["S1-B1"], "النص الأول")
        self.assertEqual(mapping["S1-B2"], "النص الثاني")


if __name__ == "__main__":
    unittest.main()
