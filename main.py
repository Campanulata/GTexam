import paper
from os.path import split
from typing import Pattern
from docx import Document
import docx
from docx.enum.text import WD_LINE_SPACING
from docx.shared import Pt
from docx.shared import RGBColor
from docx.shared import Inches
import re
gaokaoPaper=paper.Paper()
#手动配置信息
docxPath='/Users/tylor/Desktop/1.docx'
gaokaoPaper.choiceNum=9
gaokaoPaper.maxNum=23


gaokaoPaper.document=docx.Document('/Users/tylor/Desktop/1.docx')
gaokaoPaper.par = gaokaoPaper.document.paragraphs

gaokaoPaper.image_to_img()
gaokaoPaper.add_image()

gaokaoPaper.get_ABCD_adn_sub()
gaokaoPaper.get_list_question()
gaokaoPaper.get_list_ABCD()
gaokaoPaper.for_i_in_par()

gaokaoPaper.get_choice_all()
gaokaoPaper.get_unchoice_all()
gaokaoPaper.latex_will_work()
gaokaoPaper.write_to_tex()
print(gaokaoPaper.listQuestion)
print(gaokaoPaper.listABCD)
print(1)