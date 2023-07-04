import sys 
sys.path.append("..") 
import config
from excel_operator import write_to_assigned_cell
import os

if __name__ == '__main__':
    filepath = os.path.join(config.template_dir, "demo.xlsx")
    sheetname = 'Sheet1'
    write_to_assigned_cell(filepath, sheetname, 5, 5, "pig")