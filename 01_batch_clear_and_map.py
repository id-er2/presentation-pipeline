"""Replace eligible slide text with tags and save templates and text maps."""

import re
import json
from pathlib import Path
from pptx import Presentation
from pptx.oxml.ns import qn
from lxml import etree


INPUT_DIR = Path("input_pptx")
TEMP_DIR  = Path("temp_work")
MAP_DIR   = Path("mapping_files")


# A match protects the whole text box, not just the matching word.
PROTECTED_WORDS = {
    "Quran", "Hadith", "Allah", "Islam", "Muslim",
    "Sunnah", "Fiqh", "Shariah", "Halal", "Haram",
}


# Only a full-box bracketed reference is protected by this pattern.
PROTECTED_REF_PATTERN = re.compile(
    r'\['
    r'[^\[\]]*'
    r'\d+'
    r'[^\[\]]*'
    r'\]',
    re.UNICODE
)


MIN_TEXT_LENGTH = 2


def is_protected(text: str) -> bool:
    """Keep short text, numeric references, and boxes containing protected words."""
    stripped = text.strip()
    if not stripped:
        return True

    if re.fullmatch(r'[\d\.\,\:\;\-\(\)\[\]\/\\%\+\=\*\#\@\!\?\"\']+', stripped):
        return True

    if len(stripped) < MIN_TEXT_LENGTH:
        return True

    if PROTECTED_REF_PATTERN.fullmatch(stripped):
        return True

    for word in PROTECTED_WORDS:
        if re.search(rf'\b{re.escape(word)}\b', stripped, re.IGNORECASE):
            return True

    return False


def enable_rtl_on_paragraph(para_elem):
    """Set right-to-left reading order and right alignment."""
    pPr = para_elem.find(qn('a:pPr'))
    if pPr is None:
        pPr = etree.SubElement(para_elem, qn('a:pPr'))
        para_elem.insert(0, pPr)
    pPr.set('rtl', '1')
    pPr.set('algn', 'r')


def clear_textframe_runs(tf_xml):
    """Remove runs and fields without removing the paragraphs."""
    for para in tf_xml.findall('.//' + qn('a:p')):
        for run in para.findall(qn('a:r')):
            para.remove(run)
        for fld in para.findall(qn('a:fld')):
            para.remove(fld)


def process_single_pptx(pptx_path: Path, output_pptx_path: Path,
                         map_json_path: Path, map_txt_path: Path) -> int:
    """Save a tagged template and both maps; return the number of tags."""
    prs = Presentation(pptx_path)
    file_stem = pptx_path.stem

    mapping = {}
    txt_lines = [f"=== خريطة النصوص: {file_stem} ===\n"]

    for slide_idx, slide in enumerate(prs.slides, start=1):
        box_idx = 0
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue

            tf = shape.text_frame
            full_text = tf.text.strip()

            if not full_text or is_protected(full_text):
                continue

            box_idx += 1
            tag = f"S{slide_idx}-B{box_idx}"

            mapping[tag] = full_text
            preview = full_text[:100] + ('...' if len(full_text) > 100 else '')
            txt_lines.append(f"[{tag}] → {preview}")

            tf_xml = tf._txBody
            clear_textframe_runs(tf_xml)

            paragraphs = tf_xml.findall(qn('a:p'))
            if not paragraphs:
                continue

            for para in paragraphs:
                enable_rtl_on_paragraph(para)

            first_para = paragraphs[0]
            new_run = etree.SubElement(first_para, qn('a:r'))

            rPr = etree.SubElement(new_run, qn('a:rPr'))
            rPr.set('lang', 'fa-IR')
            rPr.set('altLang', 'en-US')
            rPr.set('dirty', '0')
            cs_font = etree.SubElement(rPr, qn('a:cs'))
            cs_font.set('typeface', 'Bahij TheSansArabic Bold')

            t_elem = etree.SubElement(new_run, qn('a:t'))
            t_elem.text = f"[{tag}]"

    prs.save(output_pptx_path)
    print(f"  ✅ القالب المفرغ:  {output_pptx_path.name}")

    with open(map_json_path, 'w', encoding='utf-8') as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2)
    print(f"  ✅ خريطة JSON:     {map_json_path.name}")

    txt_lines.append(f"\nإجمالي النصوص: {len(mapping)}")
    with open(map_txt_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(txt_lines))
    print(f"  ✅ خريطة TXT:      {map_txt_path.name}")

    return len(mapping)


def main():
    print("=" * 60)
    print("🔄 السكربت الأول: تفريغ وتعيين الوسوم (Batch Clear & Map)")
    print("=" * 60)

    TEMP_DIR.mkdir(parents=True, exist_ok=True)
    MAP_DIR.mkdir(parents=True, exist_ok=True)

    pptx_files = sorted(INPUT_DIR.glob("*.pptx"))
    if not pptx_files:
        print(f"⚠️  لا توجد ملفات PPTX في: {INPUT_DIR}/")
        return

    print(f"\n📂 عدد الملفات المكتشفة: {len(pptx_files)}\n")

    total_tags = 0
    for pptx_path in pptx_files:
        stem = pptx_path.stem
        print(f"{'─'*50}")
        print(f"📄 معالجة: {pptx_path.name}")

        count = process_single_pptx(
            pptx_path,
            TEMP_DIR / f"{stem}_template.pptx",
            MAP_DIR  / f"{stem}_map.json",
            MAP_DIR  / f"{stem}_map.txt",
        )
        total_tags += count
        print(f"  📊 وسوم هذا الملف: {count}")

    print("\n" + "=" * 60)
    print(f"🎉 اكتملت المعالجة | إجمالي الوسوم: {total_tags}")
    print(f"   القوالب في: {TEMP_DIR}/   |   الخرائط في: {MAP_DIR}/")
    print("=" * 60)


if __name__ == "__main__":
    main()
