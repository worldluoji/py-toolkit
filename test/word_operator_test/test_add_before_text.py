import sys
sys.path.append('../src')
from word import word_operator
from docx import Document
import os

def test_add_before_text():
    current_file_dir = os.path.dirname(os.path.abspath(__file__))
    doc = Document(os.path.join(current_file_dir, 'data', 'demo.docx'))
    TEXT = 'Before leve 1'
    newPara = word_operator.add_before_text(doc, 'Heading, level 1', TEXT)
    assert newPara.text == 'Before leve 1'
