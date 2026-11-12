"""Fill tagged templates with Persian text and apply fonts and RTL settings."""

import re
import json
from pathlib import Path
from pptx import Presentation
from pptx.oxml.ns import qn
from lxml import etree
import copy


TEMP_DIR    = Path("temp_work")
PAYLOAD_DIR = Path("farsi_payload")
OUTPUT_DIR  = Path("output_farsi")


TAG_PATTERN = re.compile(r'\[S(\d+)-B(\d+)\]')


FONT_HEADING = "Bahij TheSansArabic Bold"


FONT_BODY    = "ae_AlMohanad"


FONT_FALLBACK = FONT_HEADING


def get_tag_from_textframe(tf) -> str | None:
    """Return the first tag in a text frame, without brackets."""
    full_text = tf.text.strip()
    match = TAG_PATTERN.search(full_text)
    if match:
        return f"S{match.group(1)}-B{match.group(2)}"
    return None


def clone_run_properties(source_run_elem) -> etree._Element | None:
    """Copy the run properties that are still present in the template."""
    rPr = source_run_elem.find(qn('a:rPr'))
    if rPr is not None:
        return copy.deepcopy(rPr)
    return None


def enable_word_wrap(tf_xml):
    """Enable wrapping without changing autofit settings."""
    bodyPr = tf_xml.find(qn('a:bodyPr'))
    if bodyPr is not None:
        bodyPr.set('wrap', 'square')



def enable_rtl_on_paragraph(para_elem):
    """Set right-to-left reading order and right alignment."""
    pPr = para_elem.find(qn('a:pPr'))
    if pPr is None:
        pPr = etree.SubElement(para_elem, qn('a:pPr'))
        para_elem.insert(0, pPr)
    pPr.set('rtl', '1')
    pPr.set('algn', 'r')


def apply_rtl_on_run_properties(rPr_elem):
    """Set Persian language metadata and disable kumimoji."""
    if rPr_elem is not None:
        rPr_elem.set('lang', 'fa-IR')
        rPr_elem.set('altLang', 'en-US')

        rPr_elem.set('kumimoji', '0')


def inject_farsi_text_into_shape(shape, farsi_text: str, source_rPr=None,
                                  font_override: str = None):
    """Replace the tag with nonblank lines, one paragraph per line."""
    tf = shape.text_frame
    tf_xml = tf._txBody

    for para in tf_xml.findall(qn('a:p')):
        for run in para.findall(qn('a:r')):
            para.remove(run)
        for fld in para.findall(qn('a:fld')):
            para.remove(fld)

    enable_word_wrap(tf_xml)

    farsi_lines = farsi_text.strip().split('\n')
    farsi_lines = [line for line in farsi_lines if line.strip()]

    existing_paras = tf_xml.findall(qn('a:p'))

    effective_font = font_override if font_override else FONT_FALLBACK

    for line_idx, line_text in enumerate(farsi_lines):
        if line_idx < len(existing_paras):
            para = existing_paras[line_idx]
        else:
            para = etree.SubElement(tf_xml, qn('a:p'))

        enable_rtl_on_paragraph(para)

        new_run = etree.SubElement(para, qn('a:r'))

        if source_rPr is not None:
            rPr = copy.deepcopy(source_rPr)

            apply_rtl_on_run_properties(rPr)

            cs = rPr.find(qn('a:cs'))
            if cs is None:
                cs = etree.SubElement(rPr, qn('a:cs'))
            cs.set('typeface', effective_font)

            latin = rPr.find(qn('a:latin'))
            if latin is None:
                latin = etree.SubElement(rPr, qn('a:latin'))
            latin.set('typeface', effective_font)
            new_run.insert(0, rPr)
        else:
            rPr = etree.SubElement(new_run, qn('a:rPr'))

            apply_rtl_on_run_properties(rPr)
            rPr.set('dirty', '0')
            rPr.set('kumimoji', '0')

            cs = etree.SubElement(rPr, qn('a:cs'))
            cs.set('typeface', effective_font)
            latin = etree.SubElement(rPr, qn('a:latin'))
            latin.set('typeface', effective_font)

        t_elem = etree.SubElement(new_run, qn('a:t'))
        t_elem.text = line_text.strip()

    all_paras = tf_xml.findall(qn('a:p'))
    for extra_para in all_paras[len(farsi_lines):]:
        for run in extra_para.findall(qn('a:r')):
            extra_para.remove(run)


def process_single_file(template_path: Path, payload_path: Path, output_path: Path) -> dict:
    """Fill one template, save the presentation, and return tag statistics."""
    prs = Presentation(template_path)

    with open(payload_path, 'r', encoding='utf-8') as f:
        farsi_data = json.load(f)

    stats = {
        'total_shapes': 0,
        'injected': 0,
        'missing_in_payload': [],
        'no_tag_found': 0,
    }

    for slide_idx, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue

            tf = shape.text_frame
            tag = get_tag_from_textframe(tf)

            if tag is None:
                stats['no_tag_found'] += 1
                continue

            stats['total_shapes'] += 1

            farsi_text = farsi_data.get(tag, None)

            if farsi_text is None:
                print(f"  ⚠️  لم يُعثر على نص فارسي للوسم: [{tag}]")
                stats['missing_in_payload'].append(tag)
                continue

            if not farsi_text.strip():
                print(f"  ⚠️  النص الفارسي فارغ للوسم: [{tag}]")
                stats['missing_in_payload'].append(tag)
                continue

            tf_xml = tf._txBody
            first_para = tf_xml.find(qn('a:p'))
            source_rPr = None
            if first_para is not None:
                first_run = first_para.find(qn('a:r'))
                if first_run is not None:
                    source_rPr = clone_run_properties(first_run)

            font_to_use = FONT_BODY
            if source_rPr is not None:
                sz_val = source_rPr.get('sz')
                # DrawingML font sizes are in hundredths of a point.
                if sz_val and int(sz_val) >= 2400:
                    font_to_use = FONT_HEADING

            if shape.is_placeholder:
                ph_idx = shape.placeholder_format.idx
                if ph_idx == 0:
                    font_to_use = FONT_HEADING

            inject_farsi_text_into_shape(shape, farsi_text, source_rPr,
                                         font_override=font_to_use)
            stats['injected'] += 1

    prs.save(output_path)

    return stats


def main():
    print("=" * 60)
    print("🔄 السكربت الثالث: حقن النصوص الفارسية (Inject & Format)")
    print("=" * 60)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    template_files = sorted(TEMP_DIR.glob("*_template.pptx"))

    if not template_files:
        print(f"⚠️  لا توجد قوالب في: {TEMP_DIR}/")
        print("     الرجاء تشغيل السكربت الأول أولاً.")
        return

    print(f"\n📂 عدد القوالب المكتشفة: {len(template_files)}\n")

    total_injected = 0
    total_missing = 0

    for template_path in template_files:
        original_stem = template_path.stem.replace('_template', '')

        payload_path = PAYLOAD_DIR / f"{original_stem}.json"

        if not payload_path.exists():
            print(f"{'─'*50}")
            print(f"⚠️  لا يوجد ملف JSON مقابل لـ: {template_path.name}")
            print(f"     المتوقع: {payload_path}")
            continue

        output_path = OUTPUT_DIR / f"{original_stem}_farsi.pptx"

        print(f"{'─'*50}")
        print(f"📄 معالجة: {template_path.name}")
        print(f"   📥 الحمولة: {payload_path.name}")
        print(f"   📤 الإخراج: {output_path.name}")

        stats = process_single_file(template_path, payload_path, output_path)

        total_injected += stats['injected']
        total_missing  += len(stats['missing_in_payload'])

        print(f"  ✅ وسوم محقونة:  {stats['injected']}")
        print(f"  ✅ إجمالي الصناديق: {stats['total_shapes']}")
        if stats['missing_in_payload']:
            print(f"  ⚠️  وسوم مفقودة في JSON: {len(stats['missing_in_payload'])}")
            for t in stats['missing_in_payload']:
                print(f"       → [{t}]")

    print("\n" + "=" * 60)
    print(f"🎉 اكتمل الحقن!")
    print(f"   إجمالي الوسوم المحقونة: {total_injected}")
    print(f"   إجمالي الوسوم المفقودة: {total_missing}")
    print(f"   الملفات النهائية في: {OUTPUT_DIR}/")
    print("=" * 60)


if __name__ == "__main__":
    main()
