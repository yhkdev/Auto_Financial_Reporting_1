import os
import openpyxl
from openpyxl import Workbook
from openpyxl.utils import column_index_from_string as get_col_idx

# os.chdir(os.getcwd() + r"\test excel files")  # Change working directory to folder with the excel files
os.chdir(os.getcwd() + "/local_test")  # Change working directory to folder with the excel files

# Set up - Import files (Identify main XL and 3 other XL to copy from)
copy_wb = openpyxl.load_workbook(filename='CopyFrom.xlsx', data_only=True)
paste_wb = openpyxl.load_workbook(filename='PasteTo.xlsx', data_only=True)  # 'data_only' makes sure you only get data, Not formula in the cells

copy_ws = copy_wb['Sheet1']
paste_ws = paste_wb['Sheet1']


def copy_paste_cell(copy_ws, paste_ws, copy_loc, paste_loc):
    """ Hardcoded Copy & Paste -- 'B2' -> 'B2' in other WS """
    paste_ws[paste_loc].value = copy_ws[copy_loc].value

def copy_paste_range(copy_ws, paste_ws, copy_range, paste_range):
    """ Hardcoded Copy & Paste -- "B2:B4" -> "B2:B4" in other WS """
    paste_ws["B2"].value = copy_ws["B2"].value

# ----- Test -----
test_wb = Workbook()

# copy_paste_cell()
test_ws = test_wb.create_sheet()
copy_loc, paste_loc = "B2", "B2"
assert test_ws[paste_loc].value != copy_ws[copy_loc].value
copy_paste_cell(copy_ws, test_ws, copy_loc, paste_loc)
assert test_ws[paste_loc].value == copy_ws[copy_loc].value


# copy_paste_range()
test_ws = test_wb.create_sheet()
copy_range, paste_range = "B2:B4", "B2:B4"
# for row in range(3):
#     print(paste_ws.cell(row=row+2, column=2).value)
#     print(copy_ws.cell(row=row+2, column=2).value)
    # assert paste_ws.cell(row=row+2, column=2).value != copy_ws.cell(row=row+2, column=2).value
# copy_paste_range(copy_loc, paste_loc)
# for row in range(3):
#     assert paste_ws.cell(row=row+1, column=1).value == copy_ws.cell(row=row+1, column=1).value



def clear_cells(ws, loc):
    """ Clear specified cell range of a ws """
    for row in ws[loc]:
        for cell in row:
            cell.value = None

# TEST clear_cells()
test_paste_ws = test_wb.create_sheet()
for row in test_paste_ws["A1:C3"]:
    for cell in row:
        cell.value = "1"
assert test_paste_ws["A1:C3"] != None




# Check if XL format of reports are still the same (ex: no extra column item)
col_items = ['Fundraising', 'G&A', 'Programs', 'Arconic Grant', 'Disaster Relief (Restricted Fd)', 'Total Programs',
             'Restricted Funds', 'TOTAL']  # Cell[B5:I5]
row_items = ['Leadership Contributions', 'Individual Contrilbutions', 'Corporate Contributions']  # Cell[A6:A9]
#
# def check_wb(wb, col_items, row_items):
#     ws = A_wb['Sheet1']
#
#     try:  # Check Columns
#         for i in range(2, len(col_items)):
#             assert ws.cell(row=5, column=i).value == col_items[i-2]
#     except:
#         raise Exception("Columns of the QBO report sheet has changed!")
#
#     try:  # Check Rows
#         for i in range(len(row_items)):
#             assert ws.cell(row=i+7, column=1).value == row_items[i]
#     except:
#         raise Exception("Rows of the QBO report sheet has changed!")
# # check_wb(A_wb, col_items, row_items)
#
#
# # If pass: Run the default hardcoded copy-paste algo
#
# copy_start, copy_end = "I7", "I9"
# paste_start, paste_end = "C9", "C11"
#
# copy_loc = ("I7","I9")
# paste_loc = ("C9","C11")
#
# def copy_paste_cells(copy_wb, main_wb):
#     pass
#
# ws = A_wb['Sheet1']
# main_ws = main_wb['Sheet1']
#
# for idx in range(3):
#     print(ws.cell(row=idx+int(copy_loc[0][1]), column=get_col_idx(copy_loc[0][0])).value, end=" ")
#     print(main_ws.cell(row=idx+int(paste_loc[0][1]), column=get_col_idx(paste_loc[0][0])).value)
#
#     main_ws.cell(row=idx+int(paste_loc[0][1]), column=get_col_idx(paste_loc[0][0])).value = ws.cell(row=idx+int(copy_loc[0][1]), column=get_col_idx(copy_loc[0][0])).value
#
#     print(ws.cell(row=idx + int(copy_loc[0][1]), column=get_col_idx(copy_loc[0][0])).value, end=" ")
#     print(main_ws.cell(row=idx + int(paste_loc[0][1]), column=get_col_idx(paste_loc[0][0])).value)



# Else: Identify each col/row items and find cells to copy/paste from. Then run copy/paste algo

# Save file?
