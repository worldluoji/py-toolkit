import sys 
sys.path.append('../../src')
import config
from excel import excel_operator
import os

def test_write_to_assigned_cell():
    filepath = os.path.join(os.path.join(os.getcwd(), config.template_dir), "demo.xlsx")
    sheetname = 'Sheet1'
    excel_operator.write_to_assigned_cell(filepath, sheetname, 3, 6, "pig")
    assert excel_operator.read_successive_cells(filepath, sheetname, "3", "F") == "pig"