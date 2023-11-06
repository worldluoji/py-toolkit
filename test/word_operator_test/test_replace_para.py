import sys
sys.path.append('../src')
from word import word_operator
from docx import Document
import os

def test_replace_para():
    current_file_dir = os.path.dirname(os.path.abspath(__file__))
    doc = Document(os.path.join(current_file_dir, 'data', 'demo.docx'))
    TEXT = 'Change leve 1'
    ret = word_operator.replace_para(doc, 'Heading, level 1', TEXT, save_to="text.docx")
    os.remove("text.docx")
    assert ret == 0
