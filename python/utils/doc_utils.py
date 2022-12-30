from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Inches, RGBColor
import os, win32com.client

def edit_styles(doc):
    styles = doc.styles

    # Heading 1
    h1 = styles.add_style('H1', WD_STYLE_TYPE.PARAGRAPH)
    h1.paragraph_format.page_break_before = False
    h1.paragraph_format.space_after = 0
    h1.font.name = 'Arial'
    h1.font.bold = True
    h1.font.size = Pt(16)
    # h1.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    # Heading 2
    h2 = styles.add_style('H2', WD_STYLE_TYPE.PARAGRAPH)
    h2.paragraph_format.space_before = 0
    h2.paragraph_format.space_after = Pt(12)
    h2.font.name = 'Arial'
    h2.font.bold = True
    h2.font.size = Pt(11)
    # h2.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    # Table Cell
    tr = styles.add_style('Table Cell', WD_STYLE_TYPE.PARAGRAPH)
    tr.paragraph_format.space_after = 0
    tr.font.name = 'Arial'
    tr.font.size = Pt(10)

    # Footnote
    fn = styles.add_style('Footnote', WD_STYLE_TYPE.PARAGRAPH)
    fn.paragraph_format.space_after = 0
    fn.font.name = 'Arial'
    fn.font.size = Pt(8)
    
    # Text Box
    tb = styles.add_style('TB', WD_STYLE_TYPE.TABLE)
    tb.font.name = 'Arial'
    tb.font.size = Pt(12)
    tb.font.bold = True
    #tb.row.height = Pt(12)

    # Garage Headers
    tb = styles.add_style('GH', WD_STYLE_TYPE.TABLE)
    tb.font.name = 'Times New Roman'
    tb.font.size = Pt(12)
    #tb.row.height = Pt(12) 

    # Standard Paragraph
    pg = styles.add_style('Paragraph', WD_STYLE_TYPE.PARAGRAPH)
    pg.font.name = 'Times New Roman'
    pg.font.size = Pt(12)
    
    # Bold Paragraph
    pg = styles.add_style('B_Paragraph', WD_STYLE_TYPE.PARAGRAPH)
    pg.font.name = 'Times New Roman'
    pg.font.size = Pt(12)
    pg.font.bold = True
    
    # Underlind headers
    ul = styles.add_style("UL", WD_STYLE_TYPE.PARAGRAPH)
    ul.font.name = 'Time New Roman'
    ul.font.size = Pt(12)
    ul.font.underline = True
    

def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # Create tcBorders tag if it doesn't exist
    tcBorders = tcPr.first_child_found_in('w:tcBorders')
    if tcBorders == None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)

    # Iterate over available edges
    for edge in ['start', 'top', 'end', 'bottom', 'insideH', 'insideV']:
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)

            # Create tag if it doesn't exist
            element = tcBorders.find(qn(tag))
            if element == None:
                element = OxmlElement(tag)
                tcBorders.append(element)
            for key in ['sz', 'val', 'color', 'space', 'shadow']:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))


def set_cell_shading(cell, fill_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # Create shd tag if it doesn't exist
    shd = tcPr.first_child_found_in('w:shd')
    if shd == None:
        shd = OxmlElement('w:shd')
        tcPr.append(shd)

        # Create tags
        atts = {'val': 'clear', 'color': 'auto', 'fill': fill_color}
        for att in atts:
            shd.set(qn('w:{}'.format(att)), atts[att])


def set_col_widths(table, widths):
    for row in table.rows:
        for i, width in enumerate(widths):
            row.cells[i].width = width

def set_margins(doc, width):
    for s in doc.sections:
        s.top_margin = width
        s.bottom_margin = width
        s.left_margin = width
        s.right_margin = width


# Sets left and right cell margins to zero
def set_cell_margins(table):
    tblPr = table._tblPr
    tblCellMar = OxmlElement('w:tblCellMar')
    for m in ['start', 'end']:
        node = OxmlElement('w:{}'.format(m))
        node.set(qn('w:w'), '0')
        node.set(qn('w:type'), 'dxa')
        tblCellMar.append(node)
    tblPr.append(tblCellMar)


"""
def write_pdf(doc_paths):
    word = win32com.client.Dispatch('Word.Application')

    for doc_path in doc_paths:
        try:
            full_path = os.path.abspath(doc_path)
            doc = word.Documents.Open(full_path)
            doc.SaveAs(full_path.replace('.docx', '.pdf'), FileFormat=17)
            doc.Close()
        except:
            print('Error converting {} to PDF.'.format(doc_path))
        
    word.Quit()
"""


# Formats a table cell or a list/tuple of cells
def format_cells(cell, bold=False, italic=False, border=None, shading=None,
                 halign=None, valign=None, size=None):
    if type(cell) in (list, tuple):
        for c in cell:
            format_cells(c, bold, italic, None, shading, halign, valign)
            
    else:
        for p in cell.paragraphs:
            # Horizontal alignment
            if halign in ['left', 'center', 'right']:
                p.paragraph_format.alignment = getattr(WD_ALIGN_PARAGRAPH,
                                                       halign.upper())
            for r in p.runs:
                # Bold and italic
                if bold:
                    r.font.bold = True
                if italic:
                    r.font.italic = True
                if size != None:
                    r.font.size = size
                    
        # Vertical alignment
        if valign in ['top', 'center', 'bottom']:
            cell.vertical_alignment = getattr(WD_ALIGN_VERTICAL, valign.upper())

        # Borders - TO DO
        if border:
            pass

        # Shading
        if shading:
            set_cell_shading(cell, shading)
                
