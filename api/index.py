from flask import Flask, request, send_file
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re

app = Flask(__name__)

def apply_base_styles(doc, font_name, font_size):
    # Set the Document's Normal style
    style = doc.styles['Normal']
    font = style.font
    font.name = font_name
    font.size = Pt(font_size)

    # Forcefully apply to all runs to override Direct Formatting
    for paragraph in doc.paragraphs:
        if paragraph.style.name.startswith('Heading'):
            continue
            
        for run in paragraph.runs:
            run.font.name = font_name
            run.font.size = Pt(font_size)
            # Remove messy highlights and manual colors
            run.font.highlight_color = None
            run.font.color.rgb = None
            # Strip out tabs
            if '\t' in run.text:
                run.text = run.text.replace('\t', '')

def apply_formatting(doc, genre, line_spacing):
    for paragraph in doc.paragraphs:
        if paragraph.style.name.startswith('Heading'):
            continue
            
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        paragraph.paragraph_format.line_spacing = line_spacing
        
        if genre == 'F':
            # Fiction: Indented paragraphs, no space between
            paragraph.paragraph_format.first_line_indent = Inches(0.3)
            paragraph.paragraph_format.space_after = Pt(0)
        else:
            # Non-Fiction: Block paragraphs, space between
            paragraph.paragraph_format.first_line_indent = None
            paragraph.paragraph_format.space_after = Pt(6)

def clean_doc_xml(doc):
    # Rip out Track Changes and Comments at the XML level
    for element_tag in ['w:commentRangeStart', 'w:commentRangeEnd', 'w:commentReference', 'w:del']:
        for element in doc.element.xpath(f'//{element_tag}'):
            element.getparent().remove(element)
            
    # Commit insertions
    for ins_element in doc.element.xpath('//w:ins'):
        parent = ins_element.getparent()
        index = parent.index(ins_element)
        for child in reversed(ins_element):
            parent.insert(index, child)
        parent.remove(ins_element)

def add_toc(paragraph):
    run = paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'TOC \\o "1-3" \\h \\z \\u'
    
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)

def ensure_styles(doc):
    # Ensure Heading 1 exists
    if 'Heading 1' not in doc.styles:
        doc.styles.add_style('Heading 1', WD_STYLE_TYPE.PARAGRAPH)
    
    h1_style = doc.styles['Heading 1']
    h1_font = h1_style.font
    h1_font.bold = True
    h1_font.size = Pt(16)
    h1_font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)
    h1_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h1_style.paragraph_format.space_before = Pt(24)
    h1_style.paragraph_format.space_after = Pt(12)

def structure_doc(doc):
    ensure_styles(doc)

    # 1. Manage Empty Paragraphs (Preserve intentional scene breaks, delete massive gaps)
    empty_count = 0
    for paragraph in list(doc.paragraphs):
        if not paragraph.text.strip():
            empty_count += 1
            if empty_count > 2: # Delete the 3rd+ empty line
                p = paragraph._element
                p.getparent().remove(p)
                p._p = p._element = None
        else:
            empty_count = 0

    # 2. Smart Chapter Detection
    chapter_pattern = re.compile(r'^(chapter|part|book|introduction|conclusion|epilogue|prologue|acknowledgments|table of contents)\b', re.IGNORECASE)
    
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
            
        is_heading = bool(chapter_pattern.match(text))
        
        # Fallback: if it's very short, ALL CAPS, and has no punctuation at the end, it MIGHT be a heading
        # But we require length > 2 to avoid single words like "STOP"
        if not is_heading and text.isupper() and 2 < len(text) < 40 and not text.endswith(('.', '!', '?')):
            is_heading = True
            
        if is_heading:
            paragraph.style = 'Heading 1'
            paragraph.paragraph_format.page_break_before = True
            # Clear run-level direct formatting so the Heading 1 style applies correctly
            for run in paragraph.runs:
                run.font.size = None
                run.font.color.rgb = None
                run.bold = None
                run.font.name = None
                
    # 3. Inject Table of Contents
    toc_title = doc.paragraphs[0].insert_paragraph_before('Table of Contents')
    toc_title.style = 'Heading 1'
    toc_title.paragraph_format.page_break_before = True
    
    toc_paragraph = doc.paragraphs[0].insert_paragraph_before('')
    add_toc(toc_paragraph)
    
    if len(doc.paragraphs) > 2:
        doc.paragraphs[2].paragraph_format.page_break_before = True

    # 4. Force TOC to update on open
    settings_element = doc.settings.element
    update_fields = OxmlElement('w:updateFields')
    update_fields.set(qn('w:val'), 'true')
    settings_element.append(update_fields)


@app.route('/api/format', methods=['POST'])
def format_document():
    if 'file' not in request.files:
        return 'No file part', 400
    
    file = request.files['file']
    if file.filename == '':
        return 'No selected file', 400
        
    genre = request.form.get('genre', 'F')
    font = request.form.get('font', 'Georgia')
    spacing = float(request.form.get('spacing', '1.5'))
    
    try:
        doc = Document(file)
        
        # Apply strict premium formatting
        clean_doc_xml(doc)
        apply_base_styles(doc, font, 12 if font == 'Georgia' else 11)
        apply_formatting(doc, genre, spacing)
        structure_doc(doc)
        
        # Save to memory buffer
        file_stream = BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)
        
        return send_file(
            file_stream,
            as_attachment=True,
            download_name=f"Formatted_{file.filename}",
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        return str(e), 500

# Required for Vercel Serverless Functions
if __name__ == '__main__':
    app.run(debug=True)
