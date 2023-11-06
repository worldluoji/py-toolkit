from docx import Document
from docx.shared import RGBColor
from enum import Enum  

# https://python-docx.readthedocs.io/en/latest/#

class RETUENED_STATUS(Enum):  
    SUCCESS = 0
    NOT_CHANGED = 1
    FAIL_TO_SAVE = 2

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
            return RETUENED_STATUS.FAIL_TO_SAVE.value

    return RETUENED_STATUS.SUCCESS.value

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
    ret = RETUENED_STATUS.NOT_CHANGED.value
    for para in doc.paragraphs:
        if para.text == keyword:
            para.insert_paragraph_before(content)
            ret = RETUENED_STATUS.SUCCESS.value
            if stop:
                break
    
    if (save_to is not None) and len(save_to) > 0:
        try:
            doc.save(save_to)
        except Exception as e:
            print(e)
            ret = RETUENED_STATUS.FAIL_TO_SAVE.value
    
    return ret


def replace_para(doc, keyword, content, stop = True, underline = False, color='', save_to= ''):
    ret = RETUENED_STATUS.NOT_CHANGED.value
    for para in doc.paragraphs:
        if para.text == keyword:
            para.text = content
            ret = RETUENED_STATUS.SUCCESS.value
            if stop:
                break
        
    if (save_to is not None) and len(save_to) > 0:
        try:
            doc.save(save_to)
        except Exception as e:
            print(e)
            ret = RETUENED_STATUS.FAIL_TO_SAVE.value

    return ret