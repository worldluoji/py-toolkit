import pandas as pd

def copy_sheet(src_dir, dst_dir, src_sheetname, dst_sheetname,
        src_copy_row, src_copy_column, dst_start_row, dst_start_column, 
        src_header_row=0, dst_header_row = 0, export_file_name="demo.xlsx"):
    
    nrows, src_copy_row_start, src_row_start_index = get_src_row_info(src_copy_row)
    ncolumns, src_column_start_index = get_src_column_info(src_copy_column)

    pdSrc = pd.read_excel(src_dir,  
                            sheet_name=src_sheetname, 
                            usecols=src_copy_column, 
                            skiprows=lambda x:x in range(0, src_copy_row_start - 1),
                            nrows=nrows,
                            header=src_header_row
                        )

    pdDst =  pd.read_excel(dst_dir, sheet_name=dst_sheetname, header=dst_header_row)
    
    # dst file may be shorter than src file，then lead iloc out of bounds. The next code ensure pdDst will never out of bounds
    last_row_index = dst_start_row + nrows - 1
    last_columns_index =  dst_start_column + ncolumns - 1
    current_index = len(pdDst.index)
    column_real_len = max(last_columns_index, pdDst.shape[1])
    while current_index < last_row_index:
        # add row
        pdDst.loc[current_index] = ['' for _ in range(column_real_len)]
        current_index += 1
    
    
    for r in range(dst_start_row - 1, last_row_index):
        for c in range(dst_start_column - 1, last_columns_index):
            pdDst.iloc[r, c] = pdSrc.iloc[src_row_start_index, src_column_start_index]
            src_column_start_index += 1
        src_row_start_index += 1
        src_column_start_index = src_copy_row_start - 1
    
    pdDst.to_excel(excel_writer=export_file_name, sheet_name=dst_sheetname, index=False)


def read_successive_cells(filepath, sheetname, rows, columns, seperator = "\r\n"):
    nrows, _, row_start_index = get_src_row_info(rows)
    pds = pd.read_excel(filepath,  
                sheet_name=sheetname, 
                usecols=columns, 
                skiprows=lambda x:x in range(0, row_start_index),
                nrows=nrows,
                header=None
    )

    ret = ''
    r,c = pds.shape[0], pds.shape[1]
    for i in range(r):
        for j in range(c):
            ret += str(pds.iloc[i, j]) + seperator
    
    return ret

def write_to_assigned_cell(filepath, sheetname, row, column, content):
    pds = pd.read_excel(filepath,  
                sheet_name=sheetname, 
                header=None
    )
    r,c = pds.shape[0], pds.shape[1]
    if row > r or column > c:
        raise ValueError("input param error")
    
    pds.iloc[row-1, column-1] = content
    pds.to_excel(excel_writer=filepath, sheet_name=sheetname, index=False, header=False)
        
    
def get_src_row_info(src_copy_row):
    src_copy_row_array = src_copy_row.split(":")
    src_copy_row_start = int(src_copy_row_array[0])
    src_copy_row_end = int(src_copy_row_array[1])
    nrows = src_copy_row_end - src_copy_row_start + 1
    return nrows, src_copy_row_start, src_copy_row_start - 1
    


excel_col_alphbet_num_map = {
    'A': 1, 'B': 2, 'C': 3, 'D': 4,
    'E': 5, 'F': 6, 'G': 7, 'H': 8,
    'I': 9, 'J': 10, 'K': 11, 'L': 12,
    'M': 13, 'N': 14, 'O': 15, 'P': 16,
    'Q': 17, 'R': 18, 'S': 19, 'T': 20,
    'U': 21, 'V': 22, 'W': 23, 'X': 24,
    'Y': 25, 'Z': 26,
}

BASE = 26

def get_src_column_info(src_copy_column):
    total = 0
    start = -1
    for part in src_copy_column.split(","):
        src_copy_columns = part.split(":")
        if len(src_copy_columns) == 1:
            total += 1
        else:
            total += excel_column_alphbet_to_num(src_copy_columns[1]) - excel_column_alphbet_to_num(src_copy_columns[0]) + 1
        if start == -1:
            start = excel_column_alphbet_to_num(src_copy_columns[0]) - 1
    return total, start

def excel_column_alphbet_to_num(s):
    c = 0
    for i in range(0, len(s)):
        c += excel_col_alphbet_num_map[s[i]] * pow(BASE, len(s) - i - 1)
    return c