import sys 
sys.path.append('../../src')
import config
from excel import excel_operator
import os

def test_read_successive_cells():
    filepath = os.path.join(os.path.join(os.getcwd(), config.template_dir), config.project_completion_status_cfg["src_file"])
    sheetname = config.project_completion_status_cfg['src_sheetname']
    expected = '''AC米兰\r\n国际米兰'''
    assert excel_operator.read_successive_cells(filepath, sheetname, "2:3", "B") == expected
    