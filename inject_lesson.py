import sys
import os
from pptx import Presentation
from pptx.enum.text import PP_ALIGN

def inject_text_into_template(template_path, data_mapping, output_path):
    if not os.path.exists(template_path):
        print(f"❌ خطأ: الملف '{template_path}' غير موجود في المجلد الحالي!")
        return

    prs = Presentation(template_path)

    for slide_idx, slide_data in data_mapping.items():
        if slide_idx < len(prs.slides):
            slide = prs.slides[slide_idx]

            for shape_idx, shape in enumerate(slide.shapes):
                if shape.has_text_frame and shape_idx in slide_data:
                    new_text = slide_data[shape_idx]
                    tf = shape.text_frame
                    tf.word_wrap = True

                    if tf.paragraphs:
                        p = tf.paragraphs[0]

                        p.alignment = PP_ALIGN.RIGHT
                        p_pr = p._p.get_or_add_pPr()
                        p_pr.set('rtl', '1')

                        if p.runs:
                            p.runs[0].text = new_text

                            for r in p.runs[1:]:
                                r.text = ""
                        else:
                            run = p.add_run()
                            run.text = new_text

    prs.save(output_path)
    print(f"🎉 تم حقن البيانات بنجاح! الملف الناتج: {output_path}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("\n💡 طريقة التشغيل الصحيحة:")
        print("python inject_lesson.py اسم_الملف_المفرغ.pptx\n")
        sys.exit(1)

    TEMPLATE_FILE = sys.argv[1]

    file_name, file_ext = os.path.splitext(TEMPLATE_FILE)
    OUTPUT_FILE = f"{file_name}_Result{file_ext}"

    # Example lesson data: zero-based slide index, then shape index.
    lesson_data = {
        0: {0: "فقه - سطح (۱)"},
        1: {0: "درس نهم", 1: "مسح بر خُفَّین (موزه و پاپوش)"},
        2: {0: "مباحث درس", 1: "اول: دلیل مشروعیت مسح بر موزه‌ها", 2: "دوم: شروط مسح بر موزه‌ها"},
        3: {0: "مباحث درس", 1: "سوم: روش مسح", 2: "چهارم: باطل کننده‌های مسح"},
        4: {0: "مبحث اول: دلیل مشروعیت مسح بر خُف‌ها\n\nاز مُغِیره بن شُعبه رضی الله عنه روایت است که گفت: در سفری همراه پیامبر ﷺ بودم، پس خواستم موزه‌های ایشان را بیرون آورم، اما ایشان فرمودند: «آنان را بگذار، چون آن دو را در حال طهارت [و وضو] پوشیده‌ام» سپس بر آن دو مسح کردند. (به روایت بخاری: ۲۰۶ و مسلم: ۲۷۴)"},
    }

    inject_text_into_template(TEMPLATE_FILE, lesson_data, OUTPUT_FILE)
