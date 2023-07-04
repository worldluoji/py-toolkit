import sys 
sys.path.append("..") 
import config
from excel_operator import copy_sheet
import os


if __name__ == '__main__':
    src_dir = os.path.join(config.template_dir, config.project_completion_status_cfg["src_file"])
    src_sheetname = config.project_completion_status_cfg['src_sheetname']
    dst_dir = os.path.join(config.template_dir, config.project_completion_status_cfg["dst_file"])
    dst_sheetname = config.project_completion_status_cfg["dst_sheetname"]
    
    copy_sheet(src_dir, dst_dir, src_sheetname, dst_sheetname,
        config.project_completion_status_cfg["src_copy_row"],
        config.project_completion_status_cfg["src_copy_column"],
        config.project_completion_status_cfg["dst_start_row"],
        config.project_completion_status_cfg["dst_start_column"],
        config.project_completion_status_cfg["src_header_row"],
        config.project_completion_status_cfg["dst_header_row"],
    )
    