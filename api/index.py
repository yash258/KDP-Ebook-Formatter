from flask import Flask, request, send_file
from io import BytesIO
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

app = Flask(__name__)

def apply_font(doc, font_name, font_size):
    style = doc.styles['Normal']
    font = style.font
    font.name = font_name
    font.size = Pt(font_size)

def apply_formatting(doc, genre, line_spacing):
    for paragraph in doc.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        paragraph.paragraph_format.line_spacing = line_spacing
        
        if genre == 'F':
            paragraph.paragraph_format.first_line_indent = Inches(0.3)
            paragraph.paragraph_format.space_after = Pt(0)
        else:
            paragraph.paragraph_format.first_line_indent = None
            paragraph.paragraph_format.space_after = Pt(6)

def clean_doc(doc):
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            if '\t' in run.text:
                run.text = run.text.replace('\t', '')
            run.font.highlight_color = None
            run.font.color.rgb = None
            
    for element_tag in ['w:commentRangeStart', 'w:commentRangeEnd', 'w:commentReference', 'w:del']:
        for element in doc.element.xpath(f'//{element_tag}'):
            element.getparent().remove(element)
            
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

def structure_doc(doc):
    empty_count = 0
    for paragraph in list(doc.paragraphs):
        if not paragraph.text.strip():
            p = paragraph._element
            p.getparent().remove(p)
            p._p = p._element = None
            empty_count += 1

    try:
        h1_style = doc.styles['Heading 1']
        h1_font = h1_style.font
        h1_font.bold = True
        h1_font.size = Pt(16)
        h1_font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)
        h1_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        h2_style = doc.styles['Heading 2']
        h2_font = h2_style.font
        h2_font.size = Pt(13)
        h2_font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)
    except KeyError:
        pass 

    keywords = ['chapter', 'title page', 'copyright']
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        lower_text = text.lower()
        
        is_heading = False
        if any(lower_text.startswith(kw) for kw in keywords):
            is_heading = True
        elif text.isupper() and 0 < len(text) < 50:
            is_heading = True
            
        if is_heading:
            paragraph.style = 'Heading 1'
            paragraph.paragraph_format.page_break_before = True
                
    toc_title = doc.paragraphs[0].insert_paragraph_before('Table of Contents')
    toc_title.style = 'Heading 1'
    toc_title.paragraph_format.page_break_before = True
    
    toc_paragraph = doc.paragraphs[0].insert_paragraph_before('')
    add_toc(toc_paragraph)
    
    if len(doc.paragraphs) > 2:
        doc.paragraphs[2].paragraph_format.page_break_before = True

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
        
        # Apply standard flow
        apply_font(doc, font, 12 if font == 'Georgia' else 11)
        apply_formatting(doc, genre, spacing)
        clean_doc(doc)
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
