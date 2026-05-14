"""
KDP Smart Formatter — api/index.py
Phase 1: Intelligent Document Analysis
Phase 2: Premium Formatting Application

Style reference: 30_manipulation_techniques_premium (2).docx
"""

import re
from flask import Flask, request, send_file
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

app = Flask(__name__)

# ═══════════════════════════════════════════════════════════════════
# CONSTANTS — extracted from premium reference document
# ═══════════════════════════════════════════════════════════════════

H1_COLOR         = RGBColor(0x2E, 0x74, 0xB5)   # #2E74B5  Heading 1
H2_COLOR         = RGBColor(0x2E, 0x74, 0xB5)   # #2E74B5  Heading 2
H3_COLOR         = RGBColor(0x1F, 0x4D, 0x78)   # #1F4D78  Heading 3
BODY_COLOR       = "1A1A1A"                       # Near-black body
SECTION_LBL_COL  = "888888"                       # Gray section labels
BORDER_COLOR     = "CCCCCC"                       # Separator line
QUOTE_COLOR      = "555555"                       # Block quotes

# Paragraph classification labels
class PType:
    TITLE        = "TITLE"
    SUBTITLE     = "SUBTITLE"
    AUTHOR       = "AUTHOR"
    CHAPTER_H1   = "CHAPTER_H1"
    SECTION_H2   = "SECTION_H2"
    SUBSECTION_H3= "SUBSECTION_H3"
    SECTION_LABEL= "SECTION_LABEL"  # e.g. "WHY IT WORKS"
    KEY_TAKEAWAY = "KEY_TAKEAWAY"
    BLOCK_QUOTE  = "BLOCK_QUOTE"
    BODY         = "BODY"
    LIST_ITEM    = "LIST_ITEM"
    EMPTY        = "EMPTY"
    DISCLAIMER   = "DISCLAIMER"
    DIVIDER      = "DIVIDER"

# ═══════════════════════════════════════════════════════════════════
# PHASE 1 — INTELLIGENT DOCUMENT ANALYZER
# ═══════════════════════════════════════════════════════════════════

# Patterns for chapter detection
_CHAPTER_KW = re.compile(
    r'^(chapter|part|book|section|unit|introduction|conclusion|epilogue|'
    r'prologue|acknowledgments|acknowledgements|foreword|preface|afterword|'
    r'appendix|bibliography|references|about the author|table of contents|'
    r'final thoughts|closing thoughts|summary)\b',
    re.IGNORECASE
)

# Numbered section: "1.", "1 —", "01.", "Chapter 1:", "Part II"
_NUMBERED = re.compile(
    r'^(chapter\s*)?(\d+|i{1,3}v?|vi{0,3}|ix|x{1,3}|[IVX]+)[\.\:\-–—\s]',
    re.IGNORECASE
)

# Section label: short ALL-CAPS lines that act as sub-section labels
_SECTION_LABEL_KW = re.compile(
    r'^(why it works|what you should notice|key takeaway|real.?world example|'
    r'how to recognize|how to respond|warning signs|the psychology|in practice|'
    r'the strategy|case study|the pattern|the technique|practical application|'
    r'real life example|signs to watch|what this looks like|bottom line|'
    r'the science|quick tip|remember|note|important|caution|'
    r'definition|overview|background)\s*$',
    re.IGNORECASE
)

# Block quote markers
_QUOTE_MARKERS = re.compile(r'^["""\'\']|["""\'\']$')

def _word_count(text):
    return len(text.split())

def _is_all_caps(text):
    # True if all letters are uppercase (ignoring punctuation/numbers)
    letters = [c for c in text if c.isalpha()]
    return len(letters) > 0 and all(c.isupper() for c in letters)

def _has_bold_run(para):
    return any(run.bold for run in para.runs if run.text.strip())

def _has_italic_run(para):
    return any(run.italic for run in para.runs if run.text.strip())

def _existing_style(para):
    return para.style.name if para.style else "Normal"

def analyze_document(doc):
    """
    Read every paragraph and assign a PType classification.
    Returns: list of (paragraph, PType)
    """
    paragraphs = doc.paragraphs
    total = len(paragraphs)
    classified = []

    # ── Pass 1: Basic classification ──────────────────────────────
    for i, para in enumerate(paragraphs):
        text = para.text.strip()
        wc   = _word_count(text)
        ex   = _existing_style(para)

        # Empty
        if not text:
            classified.append((para, PType.EMPTY))
            continue

        # Already a heading style from Word — respect it
        if ex.startswith("Heading 1"):
            classified.append((para, PType.CHAPTER_H1))
            continue
        if ex.startswith("Heading 2"):
            classified.append((para, PType.SECTION_H2))
            continue
        if ex.startswith("Heading 3"):
            classified.append((para, PType.SUBSECTION_H3))
            continue

        # Explicit keyword section labels (short ALL CAPS or matched phrase)
        if _SECTION_LABEL_KW.match(text):
            classified.append((para, PType.SECTION_LABEL))
            continue

        # Short ALL CAPS line (not a full sentence — no period at end)
        if _is_all_caps(text) and 1 < wc <= 8 and not text.endswith(('.', '!', '?')):
            classified.append((para, PType.SECTION_LABEL))
            continue

        # Key Takeaway patterns
        is_quoted = bool(_QUOTE_MARKERS.search(text))
        is_short  = wc < 35
        if text.lower().startswith('key takeaway'):
            classified.append((para, PType.KEY_TAKEAWAY))
            continue
        if is_quoted and is_short and _has_italic_run(para):
            classified.append((para, PType.KEY_TAKEAWAY))
            continue

        # Block quote — longer italic passage
        if _has_italic_run(para) and is_quoted and wc < 80:
            classified.append((para, PType.BLOCK_QUOTE))
            continue

        # Chapter headings
        if _CHAPTER_KW.match(text):
            classified.append((para, PType.CHAPTER_H1))
            continue
        if _NUMBERED.match(text) and wc <= 12:
            classified.append((para, PType.CHAPTER_H1))
            continue

        # Divider lines (e.g. "---", "***", "· · ·")
        stripped_symbols = re.sub(r'[\s\*\-–—·\.\_~=]', '', text)
        if len(stripped_symbols) == 0 and len(text) >= 3:
            classified.append((para, PType.DIVIDER))
            continue

        # Everything else = body
        classified.append((para, PType.BODY))

    # ── Pass 2: Context-aware heuristics ──────────────────────────
    # Look at the first ~10% of the document for title/subtitle/author
    title_zone = max(10, int(total * 0.06))
    title_found = False
    subtitle_found = False

    for idx, (para, ptype) in enumerate(classified[:title_zone]):
        text = para.text.strip()
        wc   = _word_count(text)
        if ptype in (PType.BODY, PType.CHAPTER_H1):
            if not title_found and wc <= 12:
                classified[idx] = (para, PType.TITLE)
                title_found = True
                continue
            if title_found and not subtitle_found and wc <= 20:
                classified[idx] = (para, PType.SUBTITLE)
                subtitle_found = True
                continue
            if title_found and subtitle_found and wc <= 6:
                classified[idx] = (para, PType.AUTHOR)
                continue

    # ── Pass 3: Infer H2 from body context ────────────────────────
    # Short bold paragraphs surrounded by body paragraphs → H2
    for idx, (para, ptype) in enumerate(classified):
        if ptype != PType.BODY:
            continue
        text = para.text.strip()
        wc   = _word_count(text)
        if wc <= 10 and _has_bold_run(para) and not text.endswith(('.', '?', '!')):
            classified[idx] = (para, PType.SECTION_H2)

    # ── Pass 4: Collapse consecutive empties ──────────────────────
    streak = 0
    for idx, (para, ptype) in enumerate(classified):
        if ptype == PType.EMPTY:
            streak += 1
            if streak > 1:
                classified[idx] = (para, "DELETE")
        else:
            streak = 0

    return classified

# ═══════════════════════════════════════════════════════════════════
# XML HELPERS
# ═══════════════════════════════════════════════════════════════════

def _make_border_el(side, color=BORDER_COLOR, sz="4", space="12"):
    el = OxmlElement(f'w:{side}')
    el.set(qn('w:val'), 'single')
    el.set(qn('w:sz'), sz)
    el.set(qn('w:space'), space)
    el.set(qn('w:color'), color)
    return el

def _set_para_border(para, top=False, bottom=False):
    pPr = para._p.get_or_add_pPr()
    old = pPr.find(qn('w:pBdr'))
    if old is not None:
        pPr.remove(old)
    if not top and not bottom:
        return
    pBdr = OxmlElement('w:pBdr')
    if top:    pBdr.append(_make_border_el('top'))
    if bottom: pBdr.append(_make_border_el('bottom'))
    pPr.append(pBdr)

def _set_spacing(para, before=None, after=None, line=None, line_rule='auto'):
    pPr  = para._p.get_or_add_pPr()
    sp   = pPr.find(qn('w:spacing'))
    if sp is None:
        sp = OxmlElement('w:spacing')
        pPr.append(sp)
    if before is not None: sp.set(qn('w:before'), str(before))
    if after  is not None: sp.set(qn('w:after'),  str(after))
    if line   is not None:
        sp.set(qn('w:line'),     str(line))
        sp.set(qn('w:lineRule'), line_rule)

def _set_run_color(run, hex_color):
    rPr = run._r.get_or_add_rPr()
    col = rPr.find(qn('w:color'))
    if col is None:
        col = OxmlElement('w:color')
        rPr.append(col)
    col.set(qn('w:val'), hex_color)

def _clear_run_direct_fmt(run, font_name, font_size_pt):
    """Strip run-level direct formatting and apply clean base style."""
    run.font.name = font_name
    run.font.size = Pt(font_size_pt)
    run.font.highlight_color = None
    if run.font.color and run.font.color.type:
        run.font.color.rgb = None

# ═══════════════════════════════════════════════════════════════════
# PHASE 2 — SMART FORMATTER
# ═══════════════════════════════════════════════════════════════════

def _ensure_styles(doc, center_h1, center_h2, font_name, font_size):
    """Inject or update all required styles."""
    styles = doc.styles

    def get_or_create(name):
        try:
            return styles[name]
        except KeyError:
            s = styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            s.name = name
            return s

    # Normal / body
    normal = get_or_create('Normal')
    normal.font.name = font_name
    normal.font.size = Pt(font_size)

    # Heading 1
    h1 = get_or_create('Heading 1')
    h1.font.bold  = True
    h1.font.size  = Pt(16)
    h1.font.color.rgb = H1_COLOR
    h1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER if center_h1 else WD_ALIGN_PARAGRAPH.LEFT
    h1.paragraph_format.space_before = Pt(24)
    h1.paragraph_format.space_after  = Pt(12)
    h1.paragraph_format.page_break_before = True

    # Heading 2
    h2 = get_or_create('Heading 2')
    h2.font.bold  = True
    h2.font.size  = Pt(13)
    h2.font.color.rgb = H2_COLOR
    h2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER if center_h2 else WD_ALIGN_PARAGRAPH.LEFT
    h2.paragraph_format.space_before = Pt(12)
    h2.paragraph_format.space_after  = Pt(6)
    h2.paragraph_format.page_break_before = False

    # Heading 3
    h3 = get_or_create('Heading 3')
    h3.font.bold  = True
    h3.font.size  = Pt(11)
    h3.font.color.rgb = H3_COLOR
    h3.paragraph_format.space_before = Pt(8)
    h3.paragraph_format.space_after  = Pt(4)


def _format_title(para, font_name):
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_spacing(para, before=1440, after=240)
    for run in para.runs:
        run.font.name = font_name
        run.font.size = Pt(32)
        run.font.bold = True
        _set_run_color(run, BODY_COLOR)
    # Add decorative bottom border
    _set_para_border(para, bottom=True)


def _format_subtitle(para, font_name):
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_spacing(para, before=0, after=720)
    for run in para.runs:
        run.font.name = font_name
        run.font.size = Pt(14)
        run.font.italic = True
        _set_run_color(run, QUOTE_COLOR)


def _format_author(para, font_name):
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_spacing(para, before=0, after=480)
    for run in para.runs:
        run.font.name = 'Arial'
        run.font.size = Pt(11)
        run.font.bold = True
        _set_run_color(run, SECTION_LBL_COL)
        # Letter spacing
        rPr = run._r.get_or_add_rPr()
        sp  = rPr.find(qn('w:spacing'))
        if sp is None:
            sp = OxmlElement('w:spacing'); rPr.append(sp)
        sp.set(qn('w:val'), '80')
        # ALL CAPS
        if rPr.find(qn('w:caps')) is None:
            rPr.append(OxmlElement('w:caps'))


def _format_chapter_h1(para, border_below, font_name):
    para.style = 'Heading 1'
    for run in para.runs:
        run.font.size  = None
        run.font.name  = None
        run.font.color.rgb = None
        run.font.bold  = None
    if border_below:
        _set_para_border(para, bottom=True)


def _format_section_h2(para, border_below, font_name):
    para.style = 'Heading 2'
    for run in para.runs:
        run.font.size  = None
        run.font.name  = None
        run.font.color.rgb = None
        run.font.bold  = None
    if border_below:
        _set_para_border(para, bottom=True)


def _format_subsection_h3(para):
    para.style = 'Heading 3'
    for run in para.runs:
        run.font.size  = None
        run.font.name  = None
        run.font.color.rgb = None


def _format_section_label(para):
    """
    Arial Bold, #888888, 10pt, ALL CAPS, letter-spacing 100.
    Space before 360, after 120. Matches premium reference exactly.
    """
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _set_spacing(para, before=360, after=120)
    for run in para.runs:
        run.font.name = 'Arial'
        run.font.size = Pt(10)
        run.font.bold = True
        rPr = run._r.get_or_add_rPr()
        col = rPr.find(qn('w:color'))
        if col is None:
            col = OxmlElement('w:color'); rPr.append(col)
        col.set(qn('w:val'), SECTION_LBL_COL)
        sp = rPr.find(qn('w:spacing'))
        if sp is None:
            sp = OxmlElement('w:spacing'); rPr.append(sp)
        sp.set(qn('w:val'), '100')
        if rPr.find(qn('w:caps')) is None:
            rPr.append(OxmlElement('w:caps'))


def _format_key_takeaway(para):
    """
    Centered, italic, #1A1A1A, top+bottom #CCCCCC border.
    Space before/after 360. Matches premium reference exactly.
    """
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_spacing(para, before=360, after=360)
    _set_para_border(para, top=True, bottom=True)
    for run in para.runs:
        run.font.italic = True
        _set_run_color(run, BODY_COLOR)


def _format_block_quote(para, font_name, font_size_pt):
    """
    Left-indented, italic, #555555, slightly smaller font.
    """
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    para.paragraph_format.left_indent  = Inches(0.5)
    para.paragraph_format.right_indent = Inches(0.5)
    _set_spacing(para, before=180, after=180)
    for run in para.runs:
        run.font.name   = font_name
        run.font.size   = Pt(font_size_pt - 1)
        run.font.italic = True
        _set_run_color(run, QUOTE_COLOR)


def _format_body(para, genre, font_name, font_size_pt, line_twips):
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _set_spacing(para,
                 after=(0 if genre == 'F' else 160),
                 line=line_twips,
                 line_rule='auto')
    if genre == 'F':
        para.paragraph_format.first_line_indent = Inches(0.3)
    else:
        para.paragraph_format.first_line_indent = None

    for run in para.runs:
        _clear_run_direct_fmt(run, font_name, font_size_pt)


def _format_divider(para):
    """Replace visual dividers with a proper #CCCCCC border line."""
    # Clear text
    for run in para.runs:
        run.text = ''
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_spacing(para, before=240, after=240)
    _set_para_border(para, top=True)


def _inject_toc(doc):
    toc_title = doc.paragraphs[0].insert_paragraph_before('Table of Contents')
    toc_title.style = 'Heading 1'
    toc_title.paragraph_format.page_break_before = True

    toc_para = doc.paragraphs[0].insert_paragraph_before('')
    run = toc_para.add_run()
    for ftype, label in [('begin', None), (None, 'TOC \\o "1-3" \\h \\z \\u'), ('separate', None), ('end', None)]:
        if ftype:
            el = OxmlElement('w:fldChar')
            el.set(qn('w:fldCharType'), ftype)
            run._r.append(el)
        else:
            el = OxmlElement('w:instrText')
            el.set(qn('xml:space'), 'preserve')
            el.text = label
            run._r.append(el)

    settings = doc.settings.element
    if settings.find(qn('w:updateFields')) is None:
        uf = OxmlElement('w:updateFields')
        uf.set(qn('w:val'), 'true')
        settings.append(uf)


def _clean_xml(doc):
    for tag in ['w:commentRangeStart', 'w:commentRangeEnd', 'w:commentReference', 'w:del']:
        for el in doc.element.xpath(f'//{tag}'):
            el.getparent().remove(el)
    for ins in doc.element.xpath('//w:ins'):
        parent = ins.getparent()
        idx = parent.index(ins)
        for child in reversed(ins):
            parent.insert(idx, child)
        parent.remove(ins)


# ═══════════════════════════════════════════════════════════════════
# MAIN PIPELINE
# ═══════════════════════════════════════════════════════════════════

@app.route('/api/format', methods=['POST'])
def format_document():
    if 'file' not in request.files:
        return 'No file part', 400
    file = request.files['file']
    if not file.filename:
        return 'No file selected', 400

    genre      = request.form.get('genre',      'N')
    font       = request.form.get('font',       'Georgia')
    spacing    = float(request.form.get('spacing', '1.5'))
    center_h1  = request.form.get('center_h1',  'true')  == 'true'
    center_h2  = request.form.get('center_h2',  'false') == 'true'
    border_h1  = request.form.get('border_h1',  'true')  == 'true'
    border_h2  = request.form.get('border_h2',  'false') == 'true'

    font_size  = 12 if font == 'Georgia' else 11
    line_twips = int(spacing * 240)

    try:
        doc = Document(file)

        # ── Step 1: XML-level cleanup ──────────────────────────────
        _clean_xml(doc)

        # ── Step 2: Analyze ────────────────────────────────────────
        classified = analyze_document(doc)

        # ── Step 3: Delete excess paragraphs ──────────────────────
        to_delete = [para for para, ptype in classified if ptype == "DELETE"]
        for para in to_delete:
            p_el = para._element
            p_el.getparent().remove(p_el)

        # Rebuild classified without deleted entries
        classified = [(p, t) for p, t in classified if t != "DELETE"]

        # ── Step 4: Inject styles ─────────────────────────────────
        _ensure_styles(doc, center_h1, center_h2, font, font_size)

        # ── Step 5: Apply formatting per paragraph type ────────────
        for para, ptype in classified:
            if ptype == PType.EMPTY:
                _set_spacing(para, before=0, after=0)
                continue
            if ptype == PType.TITLE:
                _format_title(para, font)
            elif ptype == PType.SUBTITLE:
                _format_subtitle(para, font)
            elif ptype == PType.AUTHOR:
                _format_author(para, font)
            elif ptype == PType.CHAPTER_H1:
                _format_chapter_h1(para, border_h1, font)
            elif ptype == PType.SECTION_H2:
                _format_section_h2(para, border_h2, font)
            elif ptype == PType.SUBSECTION_H3:
                _format_subsection_h3(para)
            elif ptype == PType.SECTION_LABEL:
                _format_section_label(para)
            elif ptype == PType.KEY_TAKEAWAY:
                _format_key_takeaway(para)
            elif ptype == PType.BLOCK_QUOTE:
                _format_block_quote(para, font, font_size)
            elif ptype == PType.DIVIDER:
                _format_divider(para)
            elif ptype in (PType.BODY, PType.LIST_ITEM):
                _format_body(para, genre, font, font_size, line_twips)

        # ── Step 6: TOC ────────────────────────────────────────────
        _inject_toc(doc)

        # ── Step 7: Save ───────────────────────────────────────────
        buf = BytesIO()
        doc.save(buf)
        buf.seek(0)
        return send_file(
            buf,
            as_attachment=True,
            download_name=f"KDP_Formatted_{file.filename}",
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    except Exception:
        import traceback
        return traceback.format_exc(), 500


if __name__ == '__main__':
    app.run(debug=True)
