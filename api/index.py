import re
from flask import Flask, request, send_file
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree
import copy

app = Flask(__name__)

# ─────────────────────────────────────────────────────────────────
# EXACT STYLE CONSTANTS extracted from premium reference documents
# ─────────────────────────────────────────────────────────────────
HEADING1_COLOR   = RGBColor(0x2E, 0x74, 0xB5)  # #2E74B5 — sz=32 (16pt)
HEADING2_COLOR   = RGBColor(0x2E, 0x74, 0xB5)  # #2E74B5 — sz=26 (13pt)
HEADING3_COLOR   = RGBColor(0x1F, 0x4D, 0x78)  # #1F4D78
SECTION_LBL_COL  = "888888"                     # Gray section labels (Arial, Bold, ALL CAPS, 10pt)
BODY_COLOR       = "1A1A1A"                      # Near-black body text
BORDER_COLOR     = "CCCCCC"                      # Separator line color
KEY_TAKEAWAY_COL = "1A1A1A"                      # Key takeaway box text

# ─────────────────────────────────────────────────────────────────
# UTILITY: build a border XML element
# ─────────────────────────────────────────────────────────────────
def make_border(side, color=BORDER_COLOR, sz="4", space="12", val="single"):
    el = OxmlElement(f'w:{side}')
    el.set(qn('w:val'), val)
    el.set(qn('w:sz'), sz)
    el.set(qn('w:space'), space)
    el.set(qn('w:color'), color)
    return el

def add_paragraph_border(paragraph, top=False, bottom=False):
    pPr = paragraph._p.get_or_add_pPr()
    # Remove old pBdr if exists
    old = pPr.find(qn('w:pBdr'))
    if old is not None:
        pPr.remove(old)
    pBdr = OxmlElement('w:pBdr')
    if top:
        pBdr.append(make_border('top'))
    if bottom:
        pBdr.append(make_border('bottom'))
    pPr.append(pBdr)

# ─────────────────────────────────────────────────────────────────
# UTILITY: set paragraph spacing via XML
# ─────────────────────────────────────────────────────────────────
def set_spacing(paragraph, before=None, after=None, line=None, line_rule=None):
    pPr = paragraph._p.get_or_add_pPr()
    spacing = pPr.find(qn('w:spacing'))
    if spacing is None:
        spacing = OxmlElement('w:spacing')
        pPr.append(spacing)
    if before is not None:
        spacing.set(qn('w:before'), str(before))
    if after is not None:
        spacing.set(qn('w:after'), str(after))
    if line is not None:
        spacing.set(qn('w:line'), str(line))
    if line_rule is not None:
        spacing.set(qn('w:lineRule'), line_rule)

# ─────────────────────────────────────────────────────────────────
# STEP 1: XML-level cleanup (track changes, comments)
# ─────────────────────────────────────────────────────────────────
def clean_doc_xml(doc):
    for tag in ['w:commentRangeStart', 'w:commentRangeEnd', 'w:commentReference', 'w:del']:
        for el in doc.element.xpath(f'//{tag}'):
            el.getparent().remove(el)
    for ins in doc.element.xpath('//w:ins'):
        parent = ins.getparent()
        idx = parent.index(ins)
        for child in reversed(ins):
            parent.insert(idx, child)
        parent.remove(ins)

# ─────────────────────────────────────────────────────────────────
# STEP 2: Deep run-level font stripping and cleanup
# ─────────────────────────────────────────────────────────────────
def deep_clean_runs(doc, font_name, font_size_pt):
    for para in doc.paragraphs:
        if para.style.name.startswith('Heading'):
            continue
        for run in para.runs:
            run.font.name = font_name
            run.font.size = Pt(font_size_pt)
            run.font.highlight_color = None
            # Preserve bold/italic — these are intentional author choices
            if run.font.color and run.font.color.type:
                run.font.color.rgb = None
            if '\t' in run.text:
                run.text = run.text.replace('\t', '')

# ─────────────────────────────────────────────────────────────────
# STEP 3: Inject / configure premium heading styles
# ─────────────────────────────────────────────────────────────────
def ensure_premium_styles(doc, center_h1, center_h2, border_h1):
    styles = doc.styles

    def get_or_create(name, style_id):
        try:
            return styles[name]
        except KeyError:
            s = styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            s.name = name
            return s

    # ── Heading 1 ──
    h1 = get_or_create('Heading 1', 'Heading1')
    h1.font.bold = True
    h1.font.size = Pt(16)
    h1.font.color.rgb = HEADING1_COLOR
    h1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER if center_h1 else WD_ALIGN_PARAGRAPH.LEFT
    h1.paragraph_format.space_before = Pt(24)
    h1.paragraph_format.space_after = Pt(12)
    h1.paragraph_format.page_break_before = True

    # ── Heading 2 ──
    h2 = get_or_create('Heading 2', 'Heading2')
    h2.font.bold = True
    h2.font.size = Pt(13)
    h2.font.color.rgb = HEADING2_COLOR
    h2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER if center_h2 else WD_ALIGN_PARAGRAPH.LEFT
    h2.paragraph_format.space_before = Pt(12)
    h2.paragraph_format.space_after = Pt(6)

    # ── Heading 3 ──
    h3 = get_or_create('Heading 3', 'Heading3')
    h3.font.bold = True
    h3.font.size = Pt(11)
    h3.font.color.rgb = HEADING3_COLOR
    h3.paragraph_format.space_before = Pt(8)
    h3.paragraph_format.space_after = Pt(4)

    return h1, h2

# ─────────────────────────────────────────────────────────────────
# STEP 4: Apply body paragraph formatting
# ─────────────────────────────────────────────────────────────────
def apply_body_formatting(doc, genre, line_spacing_val):
    for para in doc.paragraphs:
        if para.style.name.startswith('Heading') or para.style.name == 'Title':
            continue
        if not para.text.strip():
            continue

        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        # Line spacing as multiplier → twips (240 = single, 360 = 1.5x)
        line_twips = int(line_spacing_val * 240)
        set_spacing(para, after=(0 if genre == 'F' else 160), line=line_twips, line_rule='auto')

        if genre == 'F':
            # Fiction: first-line indent, no space between paragraphs
            para.paragraph_format.first_line_indent = Inches(0.3)
        else:
            # Non-Fiction: block paragraphs, no indent
            para.paragraph_format.first_line_indent = None

# ─────────────────────────────────────────────────────────────────
# STEP 5: Smart structure detection
# ─────────────────────────────────────────────────────────────────
CHAPTER_PATTERN = re.compile(
    r'^(chapter|part|book|section|unit|introduction|conclusion|epilogue|prologue|'
    r'acknowledgments|acknowledgements|foreword|preface|afterword|appendix|'
    r'table of contents|bibliography|references|about the author)\b',
    re.IGNORECASE
)

SECTION_LABEL_PATTERN = re.compile(
    r'^(why it works|what you should notice|key takeaway|real world example|'
    r'how to recognize|how to respond|warning signs|the psychology|in practice|'
    r'the strategy|case study|the pattern|the technique)\s*$',
    re.IGNORECASE
)

def is_chapter_heading(text):
    if not text:
        return False
    if CHAPTER_PATTERN.match(text.strip()):
        return True
    # Numbered chapter e.g. "1.", "1 -", "Chapter 1:"
    if re.match(r'^(chapter\s*)?\d+[\.\:\-\s]', text.strip(), re.IGNORECASE):
        return True
    return False

def is_section_label(text):
    return bool(SECTION_LABEL_PATTERN.match(text.strip()))

def apply_section_label_style(para):
    """
    Apply the premium section label style: Arial Bold, #888888, 10pt, ALL CAPS,
    letter spacing 100, space before 360, after 120 — exactly as in the reference.
    """
    for run in para.runs:
        run.font.name = 'Arial'
        run.font.size = Pt(10)
        run.font.bold = True
        rPr = run._r.get_or_add_rPr()
        # Set color
        color_el = rPr.find(qn('w:color'))
        if color_el is None:
            color_el = OxmlElement('w:color')
            rPr.append(color_el)
        color_el.set(qn('w:val'), SECTION_LBL_COL)
        # Set letter spacing
        spacing_el = rPr.find(qn('w:spacing'))
        if spacing_el is None:
            spacing_el = OxmlElement('w:spacing')
            rPr.append(spacing_el)
        spacing_el.set(qn('w:val'), '100')
        # ALL CAPS via w:caps
        if rPr.find(qn('w:caps')) is None:
            rPr.append(OxmlElement('w:caps'))
    set_spacing(para, before=360, after=120)

def apply_key_takeaway_style(para):
    """
    Premium key takeaway box: top+bottom #CCCCCC border, centered, italic, #1A1A1A,
    space before/after 360.
    """
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_spacing(para, before=360, after=360)
    add_paragraph_border(para, top=True, bottom=True)
    for run in para.runs:
        run.font.italic = True
        rPr = run._r.get_or_add_rPr()
        color_el = rPr.find(qn('w:color'))
        if color_el is None:
            color_el = OxmlElement('w:color')
            rPr.append(color_el)
        color_el.set(qn('w:val'), KEY_TAKEAWAY_COL)

def structure_document(doc, center_h1, center_h2, border_h1, border_h2):
    # Remove excessive blank lines (keep at most 1 consecutive)
    empty_streak = 0
    for para in list(doc.paragraphs):
        if not para.text.strip():
            empty_streak += 1
            if empty_streak > 1:
                p_el = para._element
                p_el.getparent().remove(p_el)
        else:
            empty_streak = 0

    # Detect and apply heading/section styles
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        if is_chapter_heading(text):
            para.style = 'Heading 1'
            # Clear direct run formatting so Heading 1 style applies cleanly
            for run in para.runs:
                run.font.size = None
                run.font.color.rgb = None
                run.font.bold = None
                run.font.name = None
            if border_h1:
                add_paragraph_border(para, bottom=True)

        elif is_section_label(text):
            apply_section_label_style(para)

        elif text.lower().startswith('key takeaway') or \
             (para.style.name not in ('Heading 1', 'Heading 2', 'Heading 3') and
              text.startswith('"') and text.endswith('"') and len(text) < 200):
            apply_key_takeaway_style(para)

# ─────────────────────────────────────────────────────────────────
# STEP 6: TOC injection with auto-update
# ─────────────────────────────────────────────────────────────────
def inject_toc(doc):
    # Insert TOC heading at top
    toc_title_para = doc.paragraphs[0].insert_paragraph_before('Table of Contents')
    toc_title_para.style = 'Heading 1'
    toc_title_para.paragraph_format.page_break_before = True

    # Insert TOC field
    toc_para = doc.paragraphs[0].insert_paragraph_before('')
    run = toc_para.add_run()

    fldChar_begin = OxmlElement('w:fldChar')
    fldChar_begin.set(qn('w:fldCharType'), 'begin')

    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'TOC \\o "1-3" \\h \\z \\u'

    fldChar_sep = OxmlElement('w:fldChar')
    fldChar_sep.set(qn('w:fldCharType'), 'separate')

    fldChar_end = OxmlElement('w:fldChar')
    fldChar_end.set(qn('w:fldCharType'), 'end')

    run._r.append(fldChar_begin)
    run._r.append(instrText)
    run._r.append(fldChar_sep)
    run._r.append(fldChar_end)

    # Force Word to update fields on open
    settings = doc.settings.element
    existing = settings.find(qn('w:updateFields'))
    if existing is None:
        update_fields = OxmlElement('w:updateFields')
        update_fields.set(qn('w:val'), 'true')
        settings.append(update_fields)

# ─────────────────────────────────────────────────────────────────
# MAIN API ROUTE
# ─────────────────────────────────────────────────────────────────
@app.route('/api/format', methods=['POST'])
def format_document():
    if 'file' not in request.files:
        return 'No file part', 400
    file = request.files['file']
    if file.filename == '':
        return 'No selected file', 400

    # User options
    genre       = request.form.get('genre', 'N')          # F = Fiction, N = Non-Fiction
    font        = request.form.get('font', 'Georgia')
    spacing     = float(request.form.get('spacing', '1.5'))
    center_h1   = request.form.get('center_h1', 'true') == 'true'
    center_h2   = request.form.get('center_h2', 'false') == 'true'
    border_h1   = request.form.get('border_h1', 'true') == 'true'
    border_h2   = request.form.get('border_h2', 'false') == 'true'

    font_size = 12 if font == 'Georgia' else 11

    try:
        doc = Document(file)

        # Run in correct dependency order
        clean_doc_xml(doc)
        deep_clean_runs(doc, font, font_size)
        ensure_premium_styles(doc, center_h1, center_h2, border_h1)
        apply_body_formatting(doc, genre, spacing)
        structure_document(doc, center_h1, center_h2, border_h1, border_h2)
        inject_toc(doc)

        buf = BytesIO()
        doc.save(buf)
        buf.seek(0)

        return send_file(
            buf,
            as_attachment=True,
            download_name=f"KDP_Formatted_{file.filename}",
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        import traceback
        return traceback.format_exc(), 500

if __name__ == '__main__':
    app.run(debug=True)
