import sys 
sys.path.append('../src')
import config
from excel import excel_operator
import os

def test_write_to_assigned_cell():
    current_file_dir = os.path.dirname(os.path.abspath(__file__))
    filepath = os.path.join(current_file_dir, config.template_dir, "demo.xlsx")
    print(filepath)
    sheetname = 'Sheet1'
    excel_operator.write_to_assigned_cell(filepath, sheetname, 3, 5, "pig")
    assert excel_operator.read_successive_cells(filepath, sheetname, "3", "E") == "pig"