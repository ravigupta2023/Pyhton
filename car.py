import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Border,Font,Alignment,Side,PatternFill

path = r"c:\Users\Ravi Gupta\OneDrive\Desktop\Book1.xlsx"

df = pd.read_excel(path)

wb = load_workbook(path)
ws = wb['Sheet1']


header_fill = PatternFill(start_color = 'FFFF00',fill_type = 'solid')
total_fill = PatternFill(start_color = 'FFFF00',fill_type = 'solid')
center_align = Alignment(horizontal = 'center',vertical = 'center')
font_color = Font(bold = True)
border = Border(
        left = Side(style = 'thin'),
        right = Side(style = 'thin'),
        top = Side(style = 'thin'),
        bottom = Side(style = 'thin')
)

for cell in ws[1]:
    if cell.value:
        cell.font = font_color
        cell.alignment = center_align 
        cell.fill = header_fill
        cell.border = border

max_row = ws.max_row
max_col = ws.max_column
        
for row in ws.iter_rows(min_row = 1,max_row = max_row,max_col= max_col)        :
    for cell in row:
        cell.border = border
        
for col in ws.columns:
    max_length = 0
    col_letter = col[0].column_letter
    for cell in col:
        if cell.value:
            cell_value = len(str(cell.value))
            if cell_value>max_length:
                max_length = cell_value
                
    ws.column_dimensions[col_letter].width = max_length + 4
    
for row in ws.iter_rows():
    for cell in row:
        if cell.value:
            cell.alignment = center_align 

for col in ws.iter_columns():
    
         
    
wb.save(path)
print(f"{max_row} and {max_col}")
print("save successfull")


# print(df.columns)

