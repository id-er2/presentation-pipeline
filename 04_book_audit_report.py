"""Report leftover tags and short or empty text boxes in output presentations."""

import re
from pathlib import Path
from datetime import datetime
from pptx import Presentation


OUTPUT_DIR   = Path("output_farsi")
REPORT_PATH  = Path("audit_report.txt")


TAG_PATTERN  = re.compile(r'\[S\d+-B\d+\]')


MIN_CONTENT_LENGTH = 2


def check_shape_content(shape) -> dict:
    """Check for a leftover tag and text shorter than the configured limit."""
    full_text = shape.text_frame.text.strip()
    tag_match = TAG_PATTERN.search(full_text)

    return {
        'has_tag':      tag_match is not None,
        'is_empty':     len(full_text) < MIN_CONTENT_LENGTH,
        'text_preview': full_text[:80] if full_text else '(فارغ)',
        'tag_found':    tag_match.group(0) if tag_match else None,
    }


def audit_single_file(pptx_path: Path) -> dict:
    """Collect tag failures and short-text warnings for one presentation."""
    prs = Presentation(pptx_path)

    file_report = {
        'filename':         pptx_path.name,
        'total_slides':     len(prs.slides),
        'total_shapes':     0,
        'shapes_ok':        0,
        'tags_remaining':   [],
        'empty_shapes':     [],
        'issues':           [],
        'passed':           True,
    }

    for slide_idx, slide in enumerate(prs.slides, start=1):
        for shape_idx, shape in enumerate(slide.shapes, start=1):
            if not shape.has_text_frame:
                continue

            file_report['total_shapes'] += 1
            result = check_shape_content(shape)

            if result['has_tag']:
                issue = {
                    'slide': slide_idx,
                    'shape': shape_idx,
                    'type':  'وسم_متبقٍّ',
                    'tag':   result['tag_found'],
                    'text':  result['text_preview'],
                }
                file_report['tags_remaining'].append(issue)
                file_report['issues'].append(issue)
                file_report['passed'] = False
                continue

            if result['is_empty']:
                issue = {
                    'slide': slide_idx,
                    'shape': shape_idx,
                    'type':  'صندوق_فارغ',
                    'tag':   None,
                    'text':  '(فارغ)',
                }
                # Empty decorative boxes are warnings, not failures.
                file_report['empty_shapes'].append(issue)

                continue

            file_report['shapes_ok'] += 1

    return file_report


def format_file_report(file_report: dict) -> list:
    """Build the per-file report lines."""
    lines = []
    status = "✅ نجح" if file_report['passed'] else "❌ يحتاج مراجعة"

    lines.append(f"\n{'═'*55}")
    lines.append(f"📄 الملف: {file_report['filename']}")
    lines.append(f"   الحالة: {status}")
    lines.append(f"   الشرائح: {file_report['total_slides']} | "
                 f"الصناديق الكلية: {file_report['total_shapes']} | "
                 f"الصناديق السليمة: {file_report['shapes_ok']}")

    if file_report['tags_remaining']:
        lines.append(f"\n   ❌ وسوم لم تُحقن ({len(file_report['tags_remaining'])} وسم):")
        for issue in file_report['tags_remaining']:
            lines.append(f"      [شريحة {issue['slide']} | صندوق {issue['shape']}] "
                         f"الوسم: {issue['tag']}")
            lines.append(f"        النص: {issue['text']}")

    if file_report['empty_shapes']:
        lines.append(f"\n   ⚠️  صناديق فارغة ({len(file_report['empty_shapes'])} صندوق) "
                     f"[قد تكون ديكورية]:")
        for issue in file_report['empty_shapes'][:5]:
            lines.append(f"      [شريحة {issue['slide']} | صندوق {issue['shape']}] فارغ")
        if len(file_report['empty_shapes']) > 5:
            lines.append(f"      ... و{len(file_report['empty_shapes'])-5} صندوق آخر")

    if file_report['passed']:
        lines.append(f"\n   🎉 الملف نظيف تماماً - جميع النصوص الفارسية حُقنت بنجاح!")

    return lines


def generate_summary_block(all_reports: list) -> list:
    """Build summary lines for all inspected files."""
    total_files   = len(all_reports)
    passed_files  = sum(1 for r in all_reports if r['passed'])
    failed_files  = total_files - passed_files
    total_shapes  = sum(r['total_shapes'] for r in all_reports)
    total_tags    = sum(len(r['tags_remaining']) for r in all_reports)
    total_empty   = sum(len(r['empty_shapes']) for r in all_reports)

    lines = []
    lines.append("=" * 60)
    lines.append("📊 ملخص تدقيق المشروع الإجمالي")
    lines.append("=" * 60)
    lines.append(f"  📁 إجمالي الملفات المفحوصة:  {total_files}")
    lines.append(f"  ✅ ملفات ناجحة:              {passed_files}")
    lines.append(f"  ❌ ملفات تحتاج مراجعة:        {failed_files}")
    lines.append(f"  📦 إجمالي الصناديق:           {total_shapes}")
    lines.append(f"  🔴 وسوم لم تُحقن:             {total_tags}")
    lines.append(f"  ⚠️  صناديق فارغة:              {total_empty}")
    lines.append("─" * 60)

    if failed_files == 0 and total_tags == 0:
        lines.append("  🏆 المشروع نظيف 100% - جاهز للاستخدام!")
    else:
        lines.append("  ⚠️  يوجد مشكلات تحتاج إلى معالجة قبل الاستخدام.")
        if total_tags > 0:
            lines.append(f"       الأولوية: حقن {total_tags} وسم متبقٍّ")

    lines.append("=" * 60)
    return lines


def main():
    print("=" * 60)
    print("🔍 السكربت الرابع: تدقيق الملفات النهائية (Audit Report)")
    print(f"   التاريخ والوقت: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    output_files = sorted(OUTPUT_DIR.glob("*.pptx"))

    if not output_files:
        print(f"⚠️  لا توجد ملفات PPTX في: {OUTPUT_DIR}/")
        print("     الرجاء تشغيل السكربت الثالث أولاً.")
        return

    print(f"\n📂 عدد الملفات للفحص: {len(output_files)}\n")

    all_reports   = []
    all_log_lines = [
        f"تقرير التدقيق الشامل - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"المشروع: {OUTPUT_DIR.resolve()}",
        "=" * 60,
    ]

    for pptx_path in output_files:
        print(f"{'─'*50}")
        print(f"🔍 فحص: {pptx_path.name} ...")

        report = audit_single_file(pptx_path)
        all_reports.append(report)

        report_lines = format_file_report(report)
        for line in report_lines:
            print(line)
        all_log_lines.extend(report_lines)

    print("\n")
    summary_lines = generate_summary_block(all_reports)
    for line in summary_lines:
        print(line)
    all_log_lines.append("\n")
    all_log_lines.extend(summary_lines)

    with open(REPORT_PATH, 'w', encoding='utf-8') as f:
        f.write('\n'.join(all_log_lines))

    print(f"\n💾 تم حفظ تقرير التدقيق في: {REPORT_PATH}")


if __name__ == "__main__":
    main()
