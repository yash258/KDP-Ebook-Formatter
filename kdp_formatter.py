import antigravity
import os
import sys

try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:
    print("Error: The 'python-docx' library is required.")
    print("Please install it using: pip install python-docx")
    sys.exit(1)

def prompt_yes_no(question):
    """Ask a yes/no question and return True/False."""
    while True:
        response = input(f"{question} (Y/N): ").strip().upper()
        if response in ['Y', 'N']:
            return response == 'Y'
        print("Please enter Y or N.")

def interactive_font_selection(doc):
    print("\n" + "="*50)
    print("1. Interactive Font Selection and Preview")
    print("="*50)
    
    fonts = ['Times New Roman', 'Garamond', 'Georgia']
    print("Select a readable industry-standard font:")
    for i, f in enumerate(fonts, 1):
        print(f"  {i}. {f}")
    
    font_choice = 0
    while font_choice not in [1, 2, 3]:
        try:
            font_choice = int(input("Enter 1, 2, or 3: "))
        except ValueError:
            pass

    font_name = fonts[font_choice - 1]
    
    print("\nSelect font size:")
    print("  1. 11 (Best for naturally larger fonts like Georgia)")
    print("  2. 12 (Extracted PREMIUM default)")
    size_choice = 0
    while size_choice not in [1, 2]:
        try:
            size_choice = int(input("Enter 1 or 2: "))
        except ValueError:
            pass
            
    font_size = 11 if size_choice == 1 else 12
    
    # Preview
    print("\n--- Font Preview Snippet ---")
    print(f"[Applying: {font_name}, Size: {font_size} pt]")
    print("Sample Text: \"The quick brown fox jumps over the lazy dog.\"")
    print("----------------------------\n")
    
    if prompt_yes_no("Confirm setting this font and size as the global default?"):
        style = doc.styles['Normal']
        font = style.font
        font.name = font_name
        font.size = Pt(font_size)
        print("[✓] Font settings applied to the Normal style.")
    else:
        print("[X] Font selection cancelled.")

def interactive_formatting(doc):
    print("\n" + "="*50)
    print("2. Interactive Formatting Confirmation (Genre-Specific)")
    print("="*50)
    
    genre = ""
    while genre not in ['F', 'N']:
        genre = input("Is your manuscript Fiction or Non-Fiction? (F/N): ").strip().upper()
        
    line_spacing = 0.0
    while line_spacing not in [1.25, 1.5]:
        try:
            line_spacing = float(input("Enter global line spacing (1.25 or 1.5 - PREMIUM default is 1.5): "))
        except ValueError:
            pass

    print("\n--- Planned Changes ---")
    if genre == 'F':
        print("* Paragraphs: 0.3-inch first-line indent, 0 point spacing between paragraphs.")
    else:
        print("* Paragraphs: No first-line indent, 6-point space after each paragraph.")
    print("* Global Alignment: Justified")
    print(f"* Line Spacing: {line_spacing}")
    print("-----------------------\n")
    
    if prompt_yes_no("Apply these core formatting changes?"):
        for paragraph in doc.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.line_spacing = line_spacing
            
            if genre == 'F':
                paragraph.paragraph_format.first_line_indent = Inches(0.3)
                paragraph.paragraph_format.space_after = Pt(0)
            else:
                paragraph.paragraph_format.first_line_indent = None
                paragraph.paragraph_format.space_after = Pt(6)
        print("[✓] Core formatting changes applied.")
    else:
        print("[X] Core formatting changes skipped.")

def clean_document(doc):
    print("\n" + "="*50)
    print("3. Document Cleaning")
    print("="*50)
    
    if prompt_yes_no("Would you like to automatically clean the document of manual errors and invisible markup?"):
        print("Cleaning document...")
        
        # Remove manual tabs, text highlights, and normalize color
        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                # Remove manual tab keystrokes
                if '\t' in run.text:
                    run.text = run.text.replace('\t', '')
                    
                # Normalize text color to automatic (black) and remove highlights
                run.font.highlight_color = None
                run.font.color.rgb = None
                
        # Remove tracking changes and comments via underlying XML
        for element_tag in ['w:commentRangeStart', 'w:commentRangeEnd', 'w:commentReference', 'w:del']:
            for element in doc.element.xpath(f'//{element_tag}'):
                element.getparent().remove(element)
                
        for ins_element in doc.element.xpath('//w:ins'):
            # Strip the w:ins wrapper but keep the text content inside
            parent = ins_element.getparent()
            index = parent.index(ins_element)
            for child in reversed(ins_element):
                parent.insert(index, child)
            parent.remove(ins_element)
            
        print("[✓] Document cleaned (manual tabs removed, colors normalized, markup stripped).")
    else:
        print("[X] Document cleaning skipped.")

def add_toc(paragraph):
    """Add a clickable Table of Contents field code to a paragraph."""
    run = paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    # \o "1-3" limits heading levels 1 to 3
    # \h makes it hyperlinks
    # \z hides page numbers (ideal for KDP ebooks)
    # \u uses applied paragraph outline levels
    instrText.text = 'TOC \\o "1-3" \\h \\z \\u'
    
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)

def structure_document(doc):
    print("\n" + "="*50)
    print("4. Document Structure and TOC")
    print("="*50)
    
    if prompt_yes_no("Do you want to automatically structure chapters and generate a Table of Contents?"):
        
        # 1. Remove excess manual 'enter' keystrokes (empty paragraphs)
        empty_count = 0
        for paragraph in list(doc.paragraphs):
            if not paragraph.text.strip():
                p = paragraph._element
                p.getparent().remove(p)
                p._p = p._element = None
                empty_count += 1
                
        print(f"Removed {empty_count} excess manual line breaks.")

        # 2. Setup Style rules based on extracted premium formatting
        try:
            h1_style = doc.styles['Heading 1']
            h1_font = h1_style.font
            h1_font.bold = True
            h1_font.size = Pt(16)
            h1_font.color.rgb = RGBColor(0x2E, 0x74, 0xB5) # Premium Blue
            h1_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Setup Heading 2
            h2_style = doc.styles['Heading 2']
            h2_font = h2_style.font
            h2_font.size = Pt(13)
            h2_font.color.rgb = RGBColor(0x2E, 0x74, 0xB5) # Premium Blue
        except KeyError:
            pass 

        # 3. Identify Structure and Chapters
        chapters_found = 0
        keywords = ['chapter', 'title page', 'copyright']
        
        for paragraph in doc.paragraphs:
            text = paragraph.text.strip()
            lower_text = text.lower()
            
            is_heading = False
            # Check for standard sections
            if any(lower_text.startswith(kw) for kw in keywords):
                is_heading = True
            # Check for standalone caps text which usually indicates a chapter title
            elif text.isupper() and 0 < len(text) < 50:
                is_heading = True
                
            if is_heading:
                chapters_found += 1
                paragraph.style = 'Heading 1'
                paragraph.paragraph_format.page_break_before = True
                    
        print(f"Structured {chapters_found} sections/chapters with 'Heading 1' and page breaks.")
        
        # 4. Generate custom clickable TOC
        print("Generating Table of Contents...")
        toc_title = doc.paragraphs[0].insert_paragraph_before('Table of Contents')
        toc_title.style = 'Heading 1'
        toc_title.paragraph_format.page_break_before = True
        
        toc_paragraph = doc.paragraphs[0].insert_paragraph_before('')
        add_toc(toc_paragraph)
        
        # Make sure the paragraph after TOC breaks appropriately
        if len(doc.paragraphs) > 2:
            doc.paragraphs[2].paragraph_format.page_break_before = True
            
        print("[✓] Table of Contents inserted. Page numbers hidden for eBook format.")
    else:
        print("[X] Document structuring skipped.")

def main():
    print("=========================================================")
    print("        Amazon KDP-Ready eBook Formatter Pipeline        ")
    print("=========================================================\n")
    
    if len(sys.argv) < 3:
        print("Usage: python kdp_formatter.py <input.docx> <output.docx>")
        print("\nSince no arguments were provided, creating a dummy test document...")
        
        doc = Document()
        doc.add_paragraph("Title Page\n")
        doc.add_paragraph("Copyright\n")
        doc.add_paragraph("Table of Contents\n")
        doc.add_paragraph("CHAPTER 1")
        doc.add_paragraph("\tThis is the first paragraph. It has a manual tab.")
        doc.add_paragraph("\n\n\n") # Excess enters
        doc.add_paragraph("CHAPTER 2")
        doc.add_paragraph("This paragraph is normal but has some highlighted text.")
        
        input_file = "test_input.docx"
        output_file = "test_output.docx"
        doc.save(input_file)
        print(f"Created {input_file}.\n")
    else:
        input_file = sys.argv[1]
        output_file = sys.argv[2]
        
    try:
        doc = Document(input_file)
    except Exception as e:
        print(f"Error loading document '{input_file}': {e}")
        sys.exit(1)

    interactive_font_selection(doc)
    interactive_formatting(doc)
    clean_document(doc)
    structure_document(doc)

    print(f"\nSaving final document to '{output_file}'...")
    try:
        doc.save(output_file)
        print(f"[✓] Document successfully saved! You are KDP ready.")
    except Exception as e:
        print(f"Error saving document: {e}")

if __name__ == "__main__":
    main()
