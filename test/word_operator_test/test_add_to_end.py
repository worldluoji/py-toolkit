import sys
sys.path.append('../src')
from word import word_operator
from docx import Document
import os

def test_add_to_end():
    current_file_dir = os.path.dirname(os.path.abspath(__file__))
    doc = Document(os.path.join(current_file_dir, 'data', 'demo.docx'))

    TEXT = 'Hello Python!!!'
    word_operator.add_to_end(doc, TEXT)

    assert doc.paragraphs[-1].text == TEXT
