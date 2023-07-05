import sys 
sys.path.append("..") 
from excel_operator import get_src_column_info, excel_column_alphbet_to_num


print(excel_column_alphbet_to_num("A"))
print(excel_column_alphbet_to_num("AB"))
print(excel_column_alphbet_to_num("ABA"))

total, start = get_src_column_info("A:B")
print(total == 2, start == 0)