import sys 
sys.path.append('../../src')
import config
from excel import excel_operator
import os


def test_copy_sheet():
    src_dir = os.path.join(os.getcwd(), config.template_dir, config.project_completion_status_cfg["src_file"])
    src_sheetname = config.project_completion_status_cfg['src_sheetname']
    dst_dir = os.path.join(os.getcwd(), config.template_dir, config.project_completion_status_cfg["dst_file"])
    dst_sheetname = config.project_completion_status_cfg["dst_sheetname"]
    
    excel_operator.copy_sheet(src_dir, src_sheetname,
        config.project_completion_status_cfg["src_copy_row"],
        config.project_completion_status_cfg["src_copy_column"],
        dst_dir, dst_sheetname,
        config.project_completion_status_cfg["dst_start_row"],
        config.project_completion_status_cfg["dst_start_column"]
    )
    
    filepath = os.path.join(os.path.join(os.getcwd(), config.template_dir), "demo.xlsx")
    assert excel_operator.read_successive_cells(filepath, 'Sheet1', "2", "F") == 'test'