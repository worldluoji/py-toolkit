import sys
sys.path.append('../../src')
from word import word_operator
from docx import Document

def test_add_to_end():
    doc = Document('./data/demo.docx')

    TEXT = 'Hello Python!!!'
    word_operator.add_to_end(doc, TEXT)

    assert doc.paragraphs[-1].text == TEXT
