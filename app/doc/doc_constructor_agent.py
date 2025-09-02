import os
import re
from docx import Document
from docx.shared import Inches

def add_heading(doc, text, level):
    doc.add_heading(text, level=level)

def add_table(doc, colnames, rows):
    table = doc.add_table(rows=1, cols=len(colnames))
    table.style = "Light List"
    hdr_cells = table.rows[0].cells
    for i, col in enumerate(colnames):
        hdr_cells[i].text = str(col)
    for row_data in rows:
        row_cells = table.add_row().cells
        for i, val in enumerate(row_data):
            row_cells[i].text = str(val) if val is not None else ""
    return table

def find_section_content(content_list, section_title):
    for sec in content_list:
        if sec.get('section_name', '').lower().strip() == section_title.lower().strip():
            return sec['content']
    return None

def find_all_markdown_tables_and_text(text):
    if not text:
        return []
    table_pattern = re.compile(
        r'((?:\|.*\n)+?\|[ \t\-\|:]+\|\n(?:\|.*\n?)+)',
        re.MULTILINE
    )
    chunks = []
    last_idx = 0
    for match in table_pattern.finditer(text):
        start, end = match.span()
        if start > last_idx:
            txt = text[last_idx:start].strip()
            if txt:
                chunks.append(('text', txt))
        table_md = match.group(1).strip()
        chunks.append(('table', table_md))
        last_idx = end
    if last_idx < len(text):
        txt = text[last_idx:].strip()
        if txt:
            chunks.append(('text', txt))
    return chunks

def parse_markdown_table(table_md):
    lines = [line.strip() for line in table_md.strip().splitlines() if line.strip() and line.strip().startswith('|')]
    if not lines:
        return None, None
    rows = [[cell.strip() for cell in l.strip('|').split('|')] for l in lines]
    if len(rows) < 2:
        return None, None
    divider_row = rows[1]
    if all(re.match(r'^[-:\s]+$', c) for c in divider_row):
        del rows[1]
    colnames = rows[0]
    data_rows = rows[1:]
    return colnames, data_rows

def parse_pseudomarkdown_table(table_txt):
    """
    Handles table-like text with columns separated by | but missing markdown headers and dividers.
    Returns colnames, data_rows or (None, None) if not matching.
    """
    # Split into lines, skip empty lines
    lines = [line.strip() for line in table_txt.strip().splitlines() if line.strip()]
    # Must have at least 2 lines (header, one row)
    if len(lines) < 2:
        return None, None
    # All lines need at least 2 columns, separated by |
    if not all('|' in l for l in lines):
        return None, None
    # No line should start or end with | (then it's markdown)
    if all(not l.startswith('|') and not l.endswith('|') for l in lines):
        rows = [[cell.strip() for cell in l.split('|')] for l in lines]
        colnames = rows[0]
        data_rows = rows[1:]
        # All rows should have the same number of columns as header
        if all(len(r) == len(colnames) for r in data_rows):
            return colnames, data_rows
    return None, None

def extract_arrow_flow(text):
    if not text:
        return ""
    for line in text.splitlines():
        line = line.strip("` ").strip()
        if "->" in line and not line.lower().startswith(('diagram', 'flow', 'legend', '#')):
            return line
    if "->" in text:
        return text.strip()
    return ""

def build_document(content, sections, flow_diagram_agent=None, diagram_dir="diagrams"):
    doc = Document()

    # Add the main heading at the top
    add_heading(doc, "Technical Specification Document", 0)

    for i, section in enumerate(sections):
        title = section.get("title")
        header = f"{i+1}. {title}"
        add_heading(doc, header, 1)

        sec_content = find_section_content(content, title)

        # FLOW DIAGRAM SECTION HANDLING
        if title.strip().lower() == "flow diagram":
            diagram_img = None
            if flow_diagram_agent is not None and sec_content:
                try:
                    flow_line = extract_arrow_flow(sec_content)
                    if flow_line:
                        diagram_img = flow_diagram_agent.run(flow_line)  # <-- Returns BytesIO
                    else:
                         diagram_img = None
                except Exception as e:
                    print(f"Flow diagram agent error: {e}")
                    diagram_img = None
            if diagram_img:
                 doc.add_picture(diagram_img, width=Inches(5.5))
            else:
                 doc.add_paragraph("[Flow diagram not available]")
                 continue   # Skip remaining processing for this section

        # Universal parsing for text+tables:
        chunks = find_all_markdown_tables_and_text(sec_content)
        for typ, value in chunks:
            if typ == 'text':
                if value:
                    doc.add_paragraph(value)
            elif typ == 'table':
                colnames, rows = parse_markdown_table(value)
                # If NOT a markdown table, try pseudomarkdown table
                if not (colnames and rows):
                    colnames, rows = parse_pseudomarkdown_table(value)
                if colnames and rows:
                    add_table(doc, colnames, rows)
                else:
                    doc.add_paragraph(value)

    doc.add_paragraph("\nDocument generated by PWC AI-powered ABAP Tech Spec Assistant.")
    return doc