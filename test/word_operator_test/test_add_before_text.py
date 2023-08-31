import sys
sys.path.append('../../src')
from word import word_operator
from docx import Document

def test_add_before_text():
    doc = Document('./data/demo.docx')
    TEXT = 'Before leve 1'
    newPara = word_operator.add_before_text(doc, 'Heading, level 1', TEXT)
    assert newPara.text == 'Before leve 1'
