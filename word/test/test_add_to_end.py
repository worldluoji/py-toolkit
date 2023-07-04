import sys
sys.path.append("..") 
from word_operator import add_to_end
from docx import Document

doc = Document('./demo.docx')

add_to_end(doc, 'Hello Python!!!')

for para in doc.paragraphs:
    print(para.text)