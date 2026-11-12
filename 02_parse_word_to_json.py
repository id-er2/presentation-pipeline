"""Read tagged Persian DOCX files and save one JSON payload per document."""

import re
import json
from pathlib import Path
from docx import Document


WORD_DIR    = Path("farsi_word_docs")
PAYLOAD_DIR = Path("farsi_payload")


TAG_PATTERN = re.compile(r'\[S(\d+)-B(\d+)\]')


def extract_full_text_from_docx(doc: Document) -> str:
    """Collect body paragraphs first, then table-cell paragraphs."""
    lines = []

    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            lines.append(text)

    # Tables are collected after body paragraphs, not in document order.
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    text = para.text.strip()
                    if text:
                        lines.append(text)

    return '\n'.join(lines)


def parse_tags_from_text(full_text: str) -> dict:
    """Map each tag to the text before the next tag; later duplicates win."""
    mapping = {}

    matches = list(TAG_PATTERN.finditer(full_text))

    if not matches:
        return mapping

    for i, match in enumerate(matches):
        tag_key = f"S{match.group(1)}-B{match.group(2)}"

        content_start = match.end()

        if i + 1 < len(matches):
            content_end = matches[i + 1].start()
        else:
            content_end = len(full_text)

        raw_content = full_text[content_start:content_end]
        cleaned_content = raw_content.strip()

        cleaned_content = re.sub(r'\n{3,}', '\n\n', cleaned_content)

        if cleaned_content:
            mapping[tag_key] = cleaned_content
        else:
            print(f"  ⚠️  تحذير: الوسم [{tag_key}] لا يحتوي على نص مقابل")
            mapping[tag_key] = ""

    return mapping


def validate_and_report(mapping: dict, doc_name: str) -> None:
    """Print tag counts and list empty values."""
    total = len(mapping)
    empty = sum(1 for v in mapping.values() if not v.strip())
    filled = total - empty

    print(f"  📊 إجمالي الوسوم المكتشفة: {total}")
    print(f"  ✅ وسوم بها نص:              {filled}")
    if empty > 0:
        print(f"  ⚠️  وسوم فارغة:              {empty}")
        empty_tags = [k for k, v in mapping.items() if not v.strip()]
        for t in empty_tags:
            print(f"       → [{t}] فارغ")


def process_single_docx(docx_path: Path, output_json_path: Path) -> int:
    """Extract a document into a JSON payload and return its tag count."""
    doc = Document(docx_path)

    full_text = extract_full_text_from_docx(doc)

    if not full_text.strip():
        print(f"  ⚠️  الملف فارغ أو لا يحتوي على نص: {docx_path.name}")
        return 0

    mapping = parse_tags_from_text(full_text)

    validate_and_report(mapping, docx_path.stem)

    with open(output_json_path, 'w', encoding='utf-8') as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2)
    print(f"  ✅ تم حفظ: {output_json_path.name}")

    return len(mapping)


def main():
    print("=" * 60)
    print("🔄 السكربت الثاني: تحليل الوورد الفارسي → JSON")
    print("=" * 60)

    PAYLOAD_DIR.mkdir(parents=True, exist_ok=True)

    docx_files = sorted(WORD_DIR.glob("*.docx"))

    if not docx_files:
        print(f"⚠️  لا توجد ملفات DOCX في: {WORD_DIR}/")
        print("     الرجاء وضع ملفات الوورد الفارسية داخل farsi_word_docs/")
        return

    print(f"\n📂 عدد الملفات المكتشفة: {len(docx_files)}\n")

    total_tags = 0
    for docx_path in docx_files:
        stem = docx_path.stem
        output_json = PAYLOAD_DIR / f"{stem}.json"

        print(f"{'─'*50}")
        print(f"📄 معالجة: {docx_path.name}")

        count = process_single_docx(docx_path, output_json)
        total_tags += count
        print(f"  📊 وسوم مستخرجة: {count}")

    print("\n" + "=" * 60)
    print(f"🎉 اكتمل التحليل | إجمالي الوسوم: {total_tags}")
    print(f"   ملفات JSON الجاهزة في: {PAYLOAD_DIR}/")
    print("=" * 60)


if __name__ == "__main__":
    main()
