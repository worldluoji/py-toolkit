from docx import Document
from docx.shared import RGBColor
from docx.shared import Pt
from enum import Enum  

# https://python-docx.readthedocs.io/en/latest/#

class RETUENED_STATUS(Enum):  
    SUCCESS = 0
    NOT_CHANGED = 1
    FAIL_TO_SAVE = 2


default_styles = {'font': {'name': '宋体', 'color': RGBColor(0,0,0), 'bold': False, 'underline': False, 'size': Pt(12)}}

'''set the paragraph styles
    :param para: the paragraph style object to be set
    :param styles: a dict accroding to the reference: https://python-docx.readthedocs.io/en/latest/api/style.html#docx.styles.style.ParagraphStyle
'''
def set_paragraph_styles(para_style, styles):
    if styles['font'] is None:
        return
    para_style.font.name = styles['font']['name'] if styles['font']['name'] is not None else default_styles['font']['name']
    para_style.font.color.rgb = styles['font']['color'] if styles['font']['color'] is not None else default_styles['font']['color']
    para_style.font.bold = styles['font']['bold'] if styles['font']['bold'] is not None else default_styles['font']['bold']
    para_style.font.underline = styles['font']['underline'] if styles['font']['underline'] is not None else default_styles['font']['underline']
    para_style.font.size = styles['font']['size'] if styles['font']['size'] is not None else default_styles['font']['size']

'''add content in end of the doc
   :param doc: the word document to open
   :param content: the content to be insert
   :param font: the content's font style
   :param underline: if content with underline
   :param save_to: indicates where to save the modified doc
'''
def add_to_end(doc, content, save_to= '', styles=default_styles):
    para = doc.add_paragraph().add_run(content)
    set_paragraph_styles(para, styles)
    
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
   :param save_to: indicates where to save the modified doc
   :param styles: a dict accroding to the reference: https://python-docx.readthedocs.io/en/latest/api/style.html#docx.styles.style.ParagraphStyle
   :return RETUENED_STATUS
'''
def add_before_text(doc, keyword, content, stop = True, save_to= '', styles = default_styles):
    ret = RETUENED_STATUS.NOT_CHANGED.value
    for para in doc.paragraphs:
        if para.text == keyword:
            para.insert_paragraph_before(content)
            ret = RETUENED_STATUS.SUCCESS.value
            set_paragraph_styles(para.style, styles)
            if stop:
                break
    
    if (save_to is not None) and len(save_to) > 0:
        try:
            doc.save(save_to)
        except Exception as e:
            print(e)
            ret = RETUENED_STATUS.FAIL_TO_SAVE.value
    
    return ret


'''replace paragraph's content which content is equal to text
   :param doc: the word document to open
   :param keyword: the content to be insert before keyword
   :param content: the content to be insert
   :param stop: when found the first keyword then break if stop = True
   :param save_to: indicates where to save the modified doc
   :param styles: a dict accroding to the reference: https://python-docx.readthedocs.io/en/latest/api/style.html#docx.styles.style.ParagraphStyle
   :return RETUENED_STATUS
'''
def replace_para(doc, keyword, content, stop = True, save_to= '', styles = default_styles):
    ret = RETUENED_STATUS.NOT_CHANGED.value
    for para in doc.paragraphs:
        if para.text == keyword:
            para.text = content
            ret = RETUENED_STATUS.SUCCESS.value
            set_paragraph_styles(para.style, styles)
            if stop:
                break
        
    if (save_to is not None) and len(save_to) > 0:
        try:
            doc.save(save_to)
        except Exception as e:
            print(e)
            ret = RETUENED_STATUS.FAIL_TO_SAVE.value

    return ret