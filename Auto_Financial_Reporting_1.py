import os
import openpyxl
from openpyxl.utils import column_index_from_string as get_col_idx

os.chdir(os.getcwd() + r"\test excel files")  # Change working directory to folder with the excel files

# Set up - Import files (Identify main XL and 3 other XL to copy from)
main_wb = openpyxl.load_workbook(filename='paste_to_TEST.xlsx')
A_wb = openpyxl.load_workbook(filename='A.xlsx', data_only=True)  # 'data_only' makes sure you only get data, Not formula in the cells
B_wb = openpyxl.load_workbook(filename='B.xlsx', data_only=True)
D_wb = openpyxl.load_workbook(filename='D.xlsx', data_only=True)

# Check if XL format of reports are still the same (ex: no extra column item)
col_items = ['Fundraising', 'G&A', 'Programs', 'Arconic Grant', 'Disaster Relief (Restricted Fd)', 'Total Programs',
             'Restricted Funds', 'TOTAL']  # Cell[B5:I5]
row_items = ['Leadership Contributions', 'Individual Contrilbutions', 'Corporate Contributions']  # Cell[A6:A9]

def check_wb(wb, col_items, row_items):
    ws = A_wb['Sheet1']

    try:  # Check Columns
        for i in range(2, len(col_items)):
            assert ws.cell(row=5, column=i).value == col_items[i-2]
    except:
        raise Exception("Columns of the QBO report sheet has changed!")

    try:  # Check Rows
        for i in range(len(row_items)):
            assert ws.cell(row=i+7, column=1).value == row_items[i]
    except:
        raise Exception("Rows of the QBO report sheet has changed!")
# check_wb(A_wb, col_items, row_items)



# If pass: Run the default hardcoded copy-paste algo

copy_start, copy_end = "I7", "I9"
paste_start, paste_end = "C9", "C11"


copy_loc = ("I7","I9")
paste_loc = ("C9","C11")

def copy_paste_cells(copy_wb, main_wb):
    pass

ws = A_wb['Sheet1']
main_ws = main_wb['Sheet1']

for idx in range(3):
    print(ws.cell(row=idx+int(copy_loc[0][1]), column=get_col_idx(copy_loc[0][0])).value, end=" ")
    print(main_ws.cell(row=idx+int(paste_loc[0][1]), column=get_col_idx(paste_loc[0][0])).value)

    main_ws.cell(row=idx+int(paste_loc[0][1]), column=get_col_idx(paste_loc[0][0])).value = ws.cell(row=idx+int(copy_loc[0][1]), column=get_col_idx(copy_loc[0][0])).value

    print(ws.cell(row=idx + int(copy_loc[0][1]), column=get_col_idx(copy_loc[0][0])).value, end=" ")
    print(main_ws.cell(row=idx + int(paste_loc[0][1]), column=get_col_idx(paste_loc[0][0])).value)



# Else: Identify each col/row items and find cells to copy/paste from. Then run copy/paste algo

# Save file?
