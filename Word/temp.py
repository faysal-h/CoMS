import os

from docx import Document
from docx.shared import Inches, Pt, Mm, Emu
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL

class Reports():
    def __init__(self):
        self.document = Document()
    
    def PageLayout(self, size):
        self.size = size
        if self.size == "A4":
            sections = self.document.sections
            sectionMain = sections[0]
            # Page dimension and header footer distance
            sectionMain.page_height = Mm(297)
            sectionMain.page_width = Mm(210)
            sectionMain.top_margin = Inches(2.14)
            sectionMain.bottom_margin = Inches(0.87)
            sectionMain.left_margin = Inches(0.75)
            sectionMain.right_margin = Inches(0.7)
            sectionMain.header_distance = Inches(.39)
            sectionMain.footer_distance = Inches(1.18)

            return 'First Section of A4 pages size is created.'
        else:
            return 'Page size not supported.'

report1 = Reports()