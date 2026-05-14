import re
from flask import Flask, request, send_file
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

app = Flask(__name__)

H1_COLOR        = RGBColor(0x2E, 0x74, 0xB5)
H2_COLOR        = RGBColor(0x2E, 0x74, 0xB5)
H3_COLOR        = RGBColor(0x1F, 0x4D, 0x78)
BODY_COLOR      = "1A1A1A"
SECTION_LBL_COL = "888888"
BORDER_COLOR    = "CCCCCC"
QUOTE_COLOR     = "555555"

KDP_SIZES = {
    '5x8':     (Inches(5),   Inches(8)),
    '5.5x8.5': (Inches(5.5), Inches(8.5)),
    '6x9':     (Inches(6),   Inches(9)),
}

class PType:
    TITLE=        'TITLE'
    SUBTITLE=     'SUBTITLE'
    AUTHOR=       'AUTHOR'
    CHAPTER_H1=   'CHAPTER_H1'
    SECTION_H2=   'SECTION_H2'
    SUBSECTION_H3='SUBSECTION_H3'
    SECTION_LABEL='SECTION_LABEL'
    KEY_TAKEAWAY= 'KEY_TAKEAWAY'
    BLOCK_QUOTE=  'BLOCK_QUOTE'
    BODY=         'BODY'
    DIVIDER=      'DIVIDER'
    EMPTY=        'EMPTY'
    DELETE=       'DELETE'

_CHAPTER_KW = re.compile(
    r'^(chapter|part|book|section|unit|introduction|conclusion|epilogue|prologue|'
    r'acknowledgments|acknowledgements|foreword|preface|afterword|appendix|'
    r'bibliography|references|about the author|table of contents|final thoughts|'
    r'closing thoughts|summary)\b', re.IGNORECASE)

_SECTION_LABEL_KW = re.compile(
    r'^(why it works|what you should notice|key takeaway|real.?world example|'
    r'how to recognize|how to respond|warning signs|the psychology|in practice|'
    r'the strategy|case study|the pattern|the technique|practical application|'
    r'bottom line|the science|quick tip|remember|note|important|definition|overview)\s*$',
    re.IGNORECASE)

_NUMBERED = re.compile(r'^(chapter\s*)?(\d+|[IVXivx]+)[\.\:\-\u2013\u2014\s]', re.IGNORECASE)

def _all_caps(text):
    letters = [c for c in text if c.isalpha()]
    return len(letters) > 0 and all(c.isupper() for c in letters)

def _has_bold(para):  return any(r.bold for r in para.runs if r.text.strip())
def _has_italic(para): return any(r.italic for r in para.runs if r.text.strip())

def _xml_el(tag): return OxmlElement(tag)
def _set_attr(el, attr, val): el.set(qn(attr), val)

def _border_el(side, color=BORDER_COLOR):
    el = _xml_el(f'w:{side}')
    for k,v in [('w:val','single'),('w:sz','4'),('w:space','12'),('w:color',color)]:
        _set_attr(el, k, v)
    return el

def _set_para_border(para, top=False, bottom=False):
    pPr = para._p.get_or_add_pPr()
    old = pPr.find(qn('w:pBdr'))
    if old is not None: pPr.remove(old)
    if not top and not bottom: return
    pBdr = _xml_el('w:pBdr')
    if top:    pBdr.append(_border_el('top'))
    if bottom: pBdr.append(_border_el('bottom'))
    pPr.append(pBdr)

def _set_spacing(para, before=None, after=None, line=None, rule='auto'):
    pPr = para._p.get_or_add_pPr()
    sp  = pPr.find(qn('w:spacing'))
    if sp is None: sp = _xml_el('w:spacing'); pPr.append(sp)
    if before is not None: _set_attr(sp, 'w:before', str(before))
    if after  is not None: _set_attr(sp, 'w:after',  str(after))
    if line   is not None: _set_attr(sp, 'w:line', str(line)); _set_attr(sp, 'w:lineRule', rule)

def _set_run_color(run, hex_col):
    rPr = run._r.get_or_add_rPr()
    col = rPr.find(qn('w:color'))
    if col is None: col = _xml_el('w:color'); rPr.append(col)
    _set_attr(col, 'w:val', hex_col)

def _widow_orphan(para):
    pPr = para._p.get_or_add_pPr()
    if pPr.find(qn('w:widowControl')) is None:
        wc = _xml_el('w:widowControl')
        _set_attr(wc, 'w:val', 'true')
        pPr.append(wc)

def _keep_with_next(para):
    pPr = para._p.get_or_add_pPr()
    if pPr.find(qn('w:keepNext')) is None:
        pPr.append(_xml_el('w:keepNext'))

def set_kdp_page(doc, trim_size):
    w, h = KDP_SIZES.get(trim_size, KDP_SIZES['5.5x8.5'])
    inside  = Inches(1.0)
    outside = Inches(0.75)
    top     = Inches(0.875)
    bottom  = Inches(0.875)
    for section in doc.sections:
        section.page_width    = w
        section.page_height   = h
        section.top_margin    = top
        section.bottom_margin = bottom
        section.left_margin   = inside
        section.right_margin  = outside
        section.gutter        = Emu(0)
    settings = doc.settings.element
    if settings.find(qn('w:mirrorMargins')) is None:
        settings.append(_xml_el('w:mirrorMargins'))

def add_page_numbers(doc):
    for section in doc.sections:
        footer = section.footer
        footer.is_linked_to_previous = False
        fp = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        for r in fp.runs: r.text = ''
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = fp.add_run()
        run.font.size = Pt(9)
        for ftype, txt in [('begin', None), (None, ' PAGE '), ('separate', None), ('end', None)]:
            if ftype:
                el = _xml_el('w:fldChar'); _set_attr(el, 'w:fldCharType', ftype); run._r.append(el)
            else:
                el = _xml_el('w:instrText'); _set_attr(el, 'xml:space', 'preserve'); el.text = txt; run._r.append(el)

def add_drop_cap(para):
    text = para.text
    if not text.strip(): return
    first_char = text[0]
    rest_text  = text[1:]
    for run in para.runs:
        run.text = ''
    pPr = para._p.get_or_add_pPr()
    fp = _xml_el('w:framePr')
    for k,v in [('w:dropCap','drop'),('w:lines','3'),('w:wrap','around'),
                ('w:vAnchor','text'),('w:hAnchor','text')]:
        _set_attr(fp, k, v)
    pPr.append(fp)
    dc_run = para.add_run(first_char)
    dc_run.font.size = Pt(48)
    dc_run.font.bold = True
    rest_para = para._element.getparent()
    import copy
    new_p = copy.deepcopy(para._element)
    rest_pPr = new_p.find(qn('w:pPr'))
    if rest_pPr is not None:
        old_fp = rest_pPr.find(qn('w:framePr'))
        if old_fp is not None: rest_pPr.remove(old_fp)
    for r_el in new_p.findall(qn('w:r')):
        new_p.remove(r_el)
    new_r = _xml_el('w:r')
    new_t = _xml_el('w:t')
    _set_attr(new_t, 'xml:space', 'preserve')
    new_t.text = rest_text
    new_r.append(new_t)
    new_p.append(new_r)
    para._element.addnext(new_p)

def style_tables(doc, font_name, font_size):
    for table in doc.tables:
        for i, row in enumerate(table.rows):
            for cell in row.cells:
                for para in cell.paragraphs:
                    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    for run in para.runs:
                        run.font.name = font_name
                        run.font.size = Pt(font_size - 1)
                        if i == 0: run.font.bold = True

def style_lists(doc, font_name, font_size):
    for para in doc.paragraphs:
        if para.style.name in ('List Paragraph', 'List Bullet', 'List Number'):
            para.paragraph_format.left_indent = Inches(0.35)
            _set_spacing(para, after=100)
            for run in para.runs:
                run.font.name = font_name
                run.font.size = Pt(font_size)

def analyze_document(doc):
    paragraphs = doc.paragraphs
    total      = len(paragraphs)
    classified = []

    for para in paragraphs:
        text = para.text.strip()
        wc   = len(text.split())
        ex   = para.style.name if para.style else 'Normal'

        if not text:
            classified.append((para, PType.EMPTY)); continue
        if ex.startswith('Heading 1'):
            classified.append((para, PType.CHAPTER_H1)); continue
        if ex.startswith('Heading 2'):
            classified.append((para, PType.SECTION_H2)); continue
        if ex.startswith('Heading 3'):
            classified.append((para, PType.SUBSECTION_H3)); continue
        if _SECTION_LABEL_KW.match(text):
            classified.append((para, PType.SECTION_LABEL)); continue
        if _all_caps(text) and 1 < wc <= 8 and not text.endswith(('.','!','?')):
            classified.append((para, PType.SECTION_LABEL)); continue
        if text.lower().startswith('key takeaway'):
            classified.append((para, PType.KEY_TAKEAWAY)); continue
        is_quoted = text.startswith(('"','"',"'")) or text.endswith(('"','"',"'"))
        if is_quoted and _has_italic(para) and wc < 40:
            classified.append((para, PType.KEY_TAKEAWAY)); continue
        if _has_italic(para) and is_quoted and wc < 80:
            classified.append((para, PType.BLOCK_QUOTE)); continue
        if _CHAPTER_KW.match(text):
            classified.append((para, PType.CHAPTER_H1)); continue
        if _NUMBERED.match(text) and wc <= 12:
            classified.append((para, PType.CHAPTER_H1)); continue
        stripped = re.sub(r'[\s\*\-\u2013\u2014\xb7\._~=]', '', text)
        if len(stripped) == 0 and len(text) >= 3:
            classified.append((para, PType.DIVIDER)); continue
        classified.append((para, PType.BODY))

    # Title zone
    title_zone = max(10, int(total * 0.06))
    tf = sf = af = False
    for i, (para, ptype) in enumerate(classified[:title_zone]):
        wc = len(para.text.strip().split())
        if ptype in (PType.BODY, PType.CHAPTER_H1):
            if not tf and wc <= 12:
                classified[i] = (para, PType.TITLE); tf = True; continue
            if tf and not sf and wc <= 20:
                classified[i] = (para, PType.SUBTITLE); sf = True; continue
            if tf and sf and not af and wc <= 6:
                classified[i] = (para, PType.AUTHOR); af = True; continue

    # Infer H2 from short bold body
    for i, (para, ptype) in enumerate(classified):
        if ptype != PType.BODY: continue
        text = para.text.strip()
        wc   = len(text.split())
        if wc <= 10 and _has_bold(para) and not text.endswith(('.','?','!')):
            classified[i] = (para, PType.SECTION_H2)

    # Collapse excess empties
    streak = 0
    for i, (para, ptype) in enumerate(classified):
        if ptype == PType.EMPTY:
            streak += 1
            if streak > 1: classified[i] = (para, PType.DELETE)
        else:
            streak = 0

    return classified

def ensure_styles(doc, center_h1, center_h2):
    def get_or_create(name):
        try: return doc.styles[name]
        except KeyError:
            s = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            s.name = name; return s

    h1 = get_or_create('Heading 1')
    h1.font.bold = True; h1.font.size = Pt(16); h1.font.color.rgb = H1_COLOR
    h1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER if center_h1 else WD_ALIGN_PARAGRAPH.LEFT
    h1.paragraph_format.space_before = Pt(24); h1.paragraph_format.space_after = Pt(12)
    h1.paragraph_format.page_break_before = True

    h2 = get_or_create('Heading 2')
    h2.font.bold = True; h2.font.size = Pt(13); h2.font.color.rgb = H2_COLOR
    h2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER if center_h2 else WD_ALIGN_PARAGRAPH.LEFT
    h2.paragraph_format.space_before = Pt(12); h2.paragraph_format.space_after = Pt(6)

    h3 = get_or_create('Heading 3')
    h3.font.bold = True; h3.font.size = Pt(11); h3.font.color.rgb = H3_COLOR
    h3.paragraph_format.space_before = Pt(8); h3.paragraph_format.space_after = Pt(4)

def fmt_title(para, font):
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_spacing(para, before=1440, after=240)
    _set_para_border(para, bottom=True)
    for run in para.runs:
        run.font.name = font; run.font.size = Pt(32); run.font.bold = True
        _set_run_color(run, BODY_COLOR)

def fmt_subtitle(para, font):
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_spacing(para, after=720)
    for run in para.runs:
        run.font.name = font; run.font.size = Pt(14); run.font.italic = True
        _set_run_color(run, QUOTE_COLOR)

def fmt_author(para):
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_spacing(para, after=480)
    for run in para.runs:
        run.font.name = 'Arial'; run.font.size = Pt(11); run.font.bold = True
        _set_run_color(run, SECTION_LBL_COL)
        rPr = run._r.get_or_add_rPr()
        sp = rPr.find(qn('w:spacing'))
        if sp is None: sp = _xml_el('w:spacing'); rPr.append(sp)
        _set_attr(sp, 'w:val', '80')
        if rPr.find(qn('w:caps')) is None: rPr.append(_xml_el('w:caps'))

def fmt_h1(para, border_below):
    para.style = 'Heading 1'
    for run in para.runs:
        run.font.size = None; run.font.name = None
        run.font.bold = None; run.font.color.rgb = None
    if border_below: _set_para_border(para, bottom=True)
    _keep_with_next(para)

def fmt_h2(para, border_below):
    para.style = 'Heading 2'
    for run in para.runs:
        run.font.size = None; run.font.name = None
        run.font.bold = None; run.font.color.rgb = None
    if border_below: _set_para_border(para, bottom=True)
    _keep_with_next(para)

def fmt_h3(para):
    para.style = 'Heading 3'
    for run in para.runs:
        run.font.size = None; run.font.name = None; run.font.color.rgb = None

def fmt_section_label(para):
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _set_spacing(para, before=360, after=120)
    for run in para.runs:
        run.font.name = 'Arial'; run.font.size = Pt(10); run.font.bold = True
        rPr = run._r.get_or_add_rPr()
        col = rPr.find(qn('w:color'))
        if col is None: col = _xml_el('w:color'); rPr.append(col)
        _set_attr(col, 'w:val', SECTION_LBL_COL)
        sp = rPr.find(qn('w:spacing'))
        if sp is None: sp = _xml_el('w:spacing'); rPr.append(sp)
        _set_attr(sp, 'w:val', '100')
        if rPr.find(qn('w:caps')) is None: rPr.append(_xml_el('w:caps'))

def fmt_key_takeaway(para):
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_spacing(para, before=360, after=360)
    _set_para_border(para, top=True, bottom=True)
    for run in para.runs:
        run.font.italic = True; _set_run_color(run, BODY_COLOR)

def fmt_block_quote(para, font, size):
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    para.paragraph_format.left_indent  = Inches(0.5)
    para.paragraph_format.right_indent = Inches(0.5)
    _set_spacing(para, before=180, after=180)
    for run in para.runs:
        run.font.name = font; run.font.size = Pt(size - 1)
        run.font.italic = True; _set_run_color(run, QUOTE_COLOR)

def fmt_divider(para):
    for run in para.runs: run.text = ''
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_spacing(para, before=240, after=240)
    _set_para_border(para, top=True)

def fmt_body(para, genre, font, size, line_twips, is_first_after_h1=False):
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _set_spacing(para, after=(0 if genre == 'F' else 160), line=line_twips, rule='auto')
    # First paragraph after chapter heading: no indent (publishing convention)
    if is_first_after_h1:
        para.paragraph_format.first_line_indent = None
    elif genre == 'F':
        para.paragraph_format.first_line_indent = Inches(0.3)
    else:
        para.paragraph_format.first_line_indent = None
    _widow_orphan(para)
    for run in para.runs:
        run.font.name = font; run.font.size = Pt(size)
        run.font.highlight_color = None
        if run.font.color and run.font.color.type:
            run.font.color.rgb = None
        if '\t' in run.text: run.text = run.text.replace('\t', '')

def clean_xml(doc):
    for tag in ['w:commentRangeStart','w:commentRangeEnd','w:commentReference','w:del']:
        for el in doc.element.xpath(f'//{tag}'):
            el.getparent().remove(el)
    for ins in doc.element.xpath('//w:ins'):
        parent = ins.getparent(); idx = parent.index(ins)
        for child in reversed(ins): parent.insert(idx, child)
        parent.remove(ins)

def inject_toc(doc):
    toc_h = doc.paragraphs[0].insert_paragraph_before('Table of Contents')
    toc_h.style = 'Heading 1'
    toc_h.paragraph_format.page_break_before = True
    toc_p = doc.paragraphs[0].insert_paragraph_before('')
    run   = toc_p.add_run()
    for ftype, txt in [('begin',None),(None,'TOC \\o "1-3" \\h \\z \\u'),('separate',None),('end',None)]:
        if ftype:
            el = _xml_el('w:fldChar'); _set_attr(el,'w:fldCharType',ftype); run._r.append(el)
        else:
            el = _xml_el('w:instrText'); _set_attr(el,'xml:space','preserve'); el.text = txt; run._r.append(el)
    settings = doc.settings.element
    if settings.find(qn('w:updateFields')) is None:
        uf = _xml_el('w:updateFields'); _set_attr(uf,'w:val','true'); settings.append(uf)

@app.route('/api/format', methods=['POST'])
def format_document():
    if 'file' not in request.files: return 'No file part', 400
    file = request.files['file']
    if not file.filename: return 'No file selected', 400

    genre       = request.form.get('genre',      'N')
    font        = request.form.get('font',       'Georgia')
    spacing     = float(request.form.get('spacing',    '1.5'))
    center_h1   = request.form.get('center_h1',  'true')  == 'true'
    center_h2   = request.form.get('center_h2',  'false') == 'true'
    border_h1   = request.form.get('border_h1',  'true')  == 'true'
    border_h2   = request.form.get('border_h2',  'false') == 'true'
    trim_size   = request.form.get('trim_size',  '5.5x8.5')
    drop_caps   = request.form.get('drop_caps',  'false') == 'true'
    page_nums   = request.form.get('page_nums',  'true')  == 'true'
    auto_num    = request.form.get('auto_num',   'false') == 'true'

    font_size  = 12 if font == 'Georgia' else 11
    line_twips = int(spacing * 240)

    try:
        doc = Document(file)
        clean_xml(doc)
        classified = analyze_document(doc)

        # Delete excess empties
        for para, ptype in classified:
            if ptype == PType.DELETE:
                el = para._element; el.getparent().remove(el)
        classified = [(p,t) for p,t in classified if t != PType.DELETE]

        ensure_styles(doc, center_h1, center_h2)

        # Auto-number chapters
        if auto_num:
            chap_num = 0
            skip_kw  = re.compile(r'^(introduction|conclusion|epilogue|prologue|acknowledgments|'
                                   r'foreword|preface|afterword|appendix|bibliography|references|'
                                   r'about the author|table of contents)\b', re.IGNORECASE)
            for para, ptype in classified:
                if ptype == PType.CHAPTER_H1:
                    txt = para.text.strip()
                    if not skip_kw.match(txt) and not re.match(r'^chapter\s+\d+', txt, re.IGNORECASE):
                        chap_num += 1
                        for run in para.runs: run.text = ''
                        if para.runs:
                            para.runs[0].text = f'Chapter {chap_num}: {txt}'
                        else:
                            para.add_run(f'Chapter {chap_num}: {txt}')

        # Apply formatting — track first body para after each H1
        prev_was_h1 = False
        drop_cap_targets = []
        for para, ptype in classified:
            if ptype == PType.EMPTY:
                _set_spacing(para, before=0, after=0); prev_was_h1 = False; continue
            if ptype == PType.TITLE:
                fmt_title(para, font); prev_was_h1 = False
            elif ptype == PType.SUBTITLE:
                fmt_subtitle(para, font); prev_was_h1 = False
            elif ptype == PType.AUTHOR:
                fmt_author(para); prev_was_h1 = False
            elif ptype == PType.CHAPTER_H1:
                fmt_h1(para, border_h1); prev_was_h1 = True
            elif ptype == PType.SECTION_H2:
                fmt_h2(para, border_h2); prev_was_h1 = False
            elif ptype == PType.SUBSECTION_H3:
                fmt_h3(para); prev_was_h1 = False
            elif ptype == PType.SECTION_LABEL:
                fmt_section_label(para); prev_was_h1 = False
            elif ptype == PType.KEY_TAKEAWAY:
                fmt_key_takeaway(para); prev_was_h1 = False
            elif ptype == PType.BLOCK_QUOTE:
                fmt_block_quote(para, font, font_size); prev_was_h1 = False
            elif ptype == PType.DIVIDER:
                fmt_divider(para); prev_was_h1 = False
            elif ptype in (PType.BODY,):
                is_first = prev_was_h1
                fmt_body(para, genre, font, font_size, line_twips, is_first)
                if is_first and drop_caps:
                    drop_cap_targets.append(para)
                prev_was_h1 = False

        # Apply drop caps after main loop to avoid index issues
        for para in drop_cap_targets:
            try: add_drop_cap(para)
            except Exception: pass

        # KDP page setup
        set_kdp_page(doc, trim_size)

        # Style lists and tables
        style_lists(doc, font, font_size)
        style_tables(doc, font, font_size)

        # Page numbers
        if page_nums:
            add_page_numbers(doc)

        # TOC
        inject_toc(doc)

        buf = BytesIO()
        doc.save(buf)
        buf.seek(0)
        return send_file(buf, as_attachment=True,
                         download_name=f'KDP_Formatted_{file.filename}',
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except Exception:
        import traceback
        return traceback.format_exc(), 500

if __name__ == '__main__':
    app.run(debug=True)
