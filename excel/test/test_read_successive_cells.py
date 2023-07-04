import sys 
sys.path.append("..") 
import config
from excel_operator import read_successive_cells
import os

if __name__ == '__main__':
    filepath = os.path.join(config.template_dir, config.project_completion_status_cfg["src_file"])
    sheetname = config.project_completion_status_cfg['src_sheetname']
    print(read_successive_cells(filepath, sheetname, "2:3", "B"))