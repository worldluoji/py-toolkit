import sys
sys.path.append("..") 
from word_operator import add_before_text 
from docx import Document

doc = Document('./demo.docx')

add_before_text(doc, 'Heading, level 1', 'Before leve 1')

for para in doc.paragraphs:
    print(para.text)