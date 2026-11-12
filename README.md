# Automated Presentation Pipeline

A Python workflow for preparing Persian versions of existing PowerPoint lessons. It replaces eligible English text with stable tags, reads manually prepared Persian text from Word documents, fills the tagged slides, and reports leftover tags.

The project was built around Fiqh teaching material. It automates file handling and text placement, not translation. Persian text must be supplied and reviewed by a person.

## Architecture

The workflow has five stages implemented in four batch scripts. Clearing and mapping happen together in the first script.

| Stage | Script | Input | Output |
| --- | --- | --- | --- |
| 1. Clear eligible text | `01_batch_clear_and_map.py` | `input_pptx/*.pptx` | `temp_work/*_template.pptx` with tags such as `[S1-B1]` |
| 2. Map original text | Same script | Original text boxes | `mapping_files/*_map.json` and `*_map.txt` |
| 3. Parse Persian content | `02_parse_word_to_json.py` | `farsi_word_docs/*.docx` | `farsi_payload/*.json` |
| 4. Inject and format | `03_batch_inject_and_format.py` | Tagged templates and matching JSON | `output_farsi/*_farsi.pptx` |
| 5. Audit | `04_book_audit_report.py` | Presentations in `output_farsi/` | Console output and `audit_report.txt` |

A tag identifies a slide and an eligible text box on that slide, both numbered from 1. For example, `S2-B3` is the third eligible text box on slide 2. It is not the raw shape index.

## Tech stack

- Python 3.10 or newer, required by the type annotations.
- `python-pptx` for presentations and text frames.
- `python-docx` for Word paragraphs and tables.
- `lxml` for DrawingML edits, including paragraph direction and font metadata.
- Standard-library `json`, `re`, and `pathlib` for payloads and file matching.

No API key, translation service, or network connection is required by these scripts.

## Setup

Run commands from this directory. All configured paths are relative to the current working directory, not the script location.

```bash
python -m venv .venv
```

Activate the environment:

```bash
# macOS / Linux
source .venv/bin/activate

# Windows PowerShell
.venv\Scripts\Activate.ps1
```

Install dependencies and create the two input folders:

```bash
python -m pip install -r requirements.txt
python -c "from pathlib import Path; [Path(p).mkdir(exist_ok=True) for p in ('input_pptx', 'farsi_word_docs')]"
```

For the intended appearance, install `Bahij TheSansArabic Bold` and `ae_AlMohanad` on the computer used to open the slides. Font files are not included or embedded by this project. Obtain them under suitable licenses.

## Run the workflow

### 1. Prepare templates and maps

Place source presentations in `input_pptx/`, then run:

```bash
python 01_batch_clear_and_map.py
```

The script writes tagged copies and maps without editing the source presentations. The JSON map contains the full extracted text. The text map truncates each text-box preview to 100 characters, so use the JSON map and original slides when preparing translations.

By default, the script leaves these boxes unchanged:

- Empty or very short text.
- Text made entirely of the configured numeric and punctuation characters.
- A full-box bracketed reference containing digits, such as `[النساء: 116]`.
- Any box containing a whole-word match from `PROTECTED_WORDS`, including `Quran`, `Hadith`, and `Allah`.

A protected word preserves the entire box, including surrounding English text. Review the list in script 01 for your material.

### 2. Prepare and parse Persian Word documents

Use matching base names across input files:

```text
input_pptx/chapter01.pptx
farsi_word_docs/chapter01.docx
```

In the Word document, place each tag before its Persian text:

```text
[S1-B1]
عنوان درس

[S1-B2]
متن فارسی درس
```

Then run:

```bash
python 02_parse_word_to_json.py
```

This produces `farsi_payload/chapter01.json`, with unbracketed keys:

```json
{
  "S1-B1": "عنوان درس",
  "S1-B2": "متن فارسی درس"
}
```

The parser collects body paragraphs first, then table-cell paragraphs. Avoid mixing related tags between paragraphs and tables if their reading order matters. Text after a tag belongs to that tag until the next tag or the end of the document. Later duplicate tags overwrite earlier values. Empty values are reported.

### 3. Fill the templates

```bash
python 03_batch_inject_and_format.py
```

The script matches `chapter01_template.pptx` with `chapter01.json` and writes `chapter01_farsi.pptx`. Missing or empty payload values leave their tags in place and produce warnings. A template without a matching JSON file is skipped.

The injector enables word wrapping, right alignment, and right-to-left paragraph direction. It sets Persian run-language metadata and applies the configured fonts. Nonblank payload lines become separate paragraphs; blank lines are discarded. The default font is the body font, with a heading-font heuristic based on available run size or placeholder index. Ordinary text boxes are filled too; placeholder metadata is checked only for actual placeholders.

### 4. Audit and review

```bash
python 04_book_audit_report.py
```

Review the console output and `audit_report.txt`, then open every output presentation in PowerPoint or another suitable viewer. Check wrapping, overflow, fonts, references, and reading order yourself.

The audit flags leftover tags as failures. Text shorter than two characters, including empty decorative boxes, is only a warning. It does not verify translation accuracy, Persian language, visual layout, or completeness against the original maps. Its success messages describe these limited checks, not publication readiness.

## Directory layout

```text
pptx_automation/
  01_batch_clear_and_map.py
  02_parse_word_to_json.py
  03_batch_inject_and_format.py
  04_book_audit_report.py
  clear_pptx_for_farsi.py
  inject_lesson.py
  requirements.txt
  README.md
  input_pptx/          # source presentations supplied locally
  farsi_word_docs/    # tagged translations supplied locally
  temp_work/          # generated templates
  mapping_files/      # generated original-text maps
  farsi_payload/      # generated Persian JSON
  output_farsi/       # generated presentations
  audit_report.txt   # generated report
```

Generated folders are created by the scripts that use them. Input documents, maps, payloads, outputs, and local environments are excluded by `.gitignore`. Review all files before publishing, especially documents you do not own.

## Earlier single-lesson helpers

These are separate utilities, not additional batch stages:

- `clear_pptx_for_farsi.py` reads `Lesson_1.pptx` by default and writes `Empty_Farsi_Lesson_1.pptx` and `map_Lesson_1.txt`. Edit its constants to use different names. It has a different protected-keyword list from the batch script.
- `inject_lesson.py` accepts a template path and uses the small Persian lesson mapping embedded in the script. Its mapping uses zero-based slide and raw shape indexes, not the batch tags. The output filename ends in `_Result.pptx`.

```bash
python clear_pptx_for_farsi.py
python inject_lesson.py Empty_Farsi_Lesson_1.pptx
```

## Known limitations

- **Original run styling is not fully preserved.** Script 01 deletes source runs and creates a new tag run. Script 03 can only copy properties still present in that template. Original font size and color from deleted runs are not recoverable through this step, and the size-based heading heuristic may not have a source size to inspect.
- **Only top-level text-frame shapes are processed.** Grouped shapes, presentation table cells, images, charts, and other non-text-frame content are not traversed by the batch scripts.
- **No automatic translation or digit conversion.** Supplied Persian wording and digits are used as written, apart from whitespace handling.
- **No automatic text fitting.** Word wrapping does not guarantee that translated text fits its box.
- **No whole-book completeness check.** Files skipped during injection may not appear in the audit at all. Old output files can also remain between runs. Check expected filenames and use a separate working copy for each run.
- **Existing generated files can be overwritten.** There is no backup or interactive overwrite prompt.

The source archive contains code and documentation only, with no original lesson files or reference outputs. Automated checks on synthetic files do not establish visual fidelity for the real book.
