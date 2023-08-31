from docx import Document
from docx.shared import RGBColor


'''add content in end of the doc
   :param doc: the word document to open
   :param content: the content to be insert
   :param font: the content's font style
   :param underline: if content with underline
   :param save_to: indicates where to save the modified doc
'''
def add_to_end(doc, content, font='宋体', underline = False, color='', save_to= ''):
    para = doc.add_paragraph().add_run(content)
    # set font style
    para.font.name = font
    # set underline
    para.font.underline = underline
    # set color
    if color != '':
        para.font.color.rgb = color # RGBColor(255,128,128)
    
    if (save_to is not None) and len(save_to) > 0:
        try:
            doc.save(save_to)
        except Exception as e:
            print(e)

'''add content before text in doc
   :param doc: the word document to open
   :param keyword: the content to be insert before keyword
   :param content: the content to be insert
   :param stop: when found the first keyword then break if stop = True
   :param font: the content's font style
   :param underline: if content with underline
   :param save_to: indicates where to save the modified doc
'''
def add_before_text(doc, keyword, content, stop = True, font='宋体', underline = False, color='', save_to= ''):
    ret = None
    for para in doc.paragraphs:
        if para.text == keyword:
            ret = para.insert_paragraph_before(content)
            if stop:
                break
    
    if (save_to is not None) and len(save_to) > 0:
        try:
            doc.save(save_to)
        except Exception as e:
            print(e)
    
    return ret