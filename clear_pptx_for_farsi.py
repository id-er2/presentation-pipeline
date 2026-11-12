import os
from pptx import Presentation


ENGLISH_FILE = "Lesson_1.pptx"


OUTPUT_FILE = "Empty_Farsi_Lesson_1.pptx"


MAP_FILE = "map_Lesson_1.txt"


PROTECTED_KEYWORDS = [
    "سنا",
    "سنع",
    "©",
    "www.",
    "http",
]


def is_protected(text):
    """Keep boxes containing a protected keyword."""
    text_stripped = text.strip()

    for keyword in PROTECTED_KEYWORDS:
        if keyword.lower() in text_stripped.lower():
            return True

    return False


def get_full_shape_text(shape):
    """Join nonempty paragraphs with a visible separator for the text map."""
    if not shape.has_text_frame:
        return ""
    full_text = []
    for para in shape.text_frame.paragraphs:
        para_text = "".join([run.text for run in para.runs])
        if para_text.strip():
            full_text.append(para_text.strip())
    return " | ".join(full_text)


def clear_shape_text(shape, placeholder_text):
    """Replace the first run and blank the remaining runs."""
    tf = shape.text_frame
    first_para = tf.paragraphs[0]

    if first_para.runs:
        first_run = first_para.runs[0]
        first_run.text = placeholder_text

        for run in first_para.runs[1:]:
            run.text = ""
    else:
        run = first_para.add_run()
        run.text = placeholder_text

    for para in tf.paragraphs[1:]:
        for run in para.runs:
            run.text = ""


def clear_presentation(input_file, output_file, map_file):
    """Save a tagged copy of one presentation and a text map."""
    if not os.path.exists(input_file):
        print(f"❌ خطأ: لم يتم العثور على الملف: {input_file}")
        print(f"   تأكد من وجود الملف في نفس المجلد: {os.getcwd()}")
        return False

    prs = Presentation(input_file)
    content_map = []

    content_map.append("=" * 70)
    content_map.append(f"خريطة تفريغ الملف: {input_file}")
    content_map.append(f"الملف المفرّغ: {output_file}")
    content_map.append("=" * 70)
    content_map.append("")
    content_map.append("الرمز [S=شريحة, B=صندوق] ← النص الأصلي الإنجليزي")
    content_map.append("-" * 70)
    content_map.append("")

    total_cleared = 0
    total_protected = 0

    for slide_idx, slide in enumerate(prs.slides, 1):
        box_counter = 0
        slide_has_content = False

        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue

            full_text = get_full_shape_text(shape)
            if not full_text.strip():
                continue

            if is_protected(full_text):
                total_protected += 1
                continue

            box_counter += 1
            placeholder = f"[S{slide_idx}-B{box_counter}]"

            if not slide_has_content:
                content_map.append(f"{'═' * 50}")
                content_map.append(f"  📄 شريحة {slide_idx}")
                content_map.append(f"{'═' * 50}")
                slide_has_content = True

            display_text = full_text[:150] + "..." if len(full_text) > 150 else full_text
            content_map.append(f"  {placeholder} ← \"{display_text}\"")
            content_map.append("")

            clear_shape_text(shape, placeholder)
            total_cleared += 1

        if slide_has_content:
            content_map.append("")

    content_map.append("")
    content_map.append("=" * 70)
    content_map.append(f"📊 الإحصائيات:")
    content_map.append(f"   عدد الشرائح: {len(prs.slides)}")
    content_map.append(f"   صناديق تم تفريغها: {total_cleared}")
    content_map.append(f"   صناديق محمية: {total_protected}")
    content_map.append("=" * 70)

    prs.save(output_file)

    map_content = "\n".join(content_map)
    with open(map_file, 'w', encoding='utf-8') as f:
        f.write(map_content)

    print("\n" + "=" * 60)
    print("✅ تم التفريغ بنجاح!")
    print("=" * 60)
    print(f"📁 الملف المفرّغ: {output_file}")
    print(f"🗺️  ملف الخريطة: {map_file}")
    print("=" * 60)
    return True


if __name__ == "__main__":
    clear_presentation(ENGLISH_FILE, OUTPUT_FILE, MAP_FILE)
