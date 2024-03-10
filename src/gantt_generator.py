import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

csv_file_path = '/Users/darren/Desktop/gantt/oricsv.csv' # YOU NEED TO REPLACE WITH oricsv.csv PATH
xlsx_file_path = '/Users/darren/Desktop/gantt/xlsFromCsv.xlsx' # YOU NEED TO REPLACE TARGET EXCEL PATH
task_id_col = 'Task Id' 
task_name_col = 'Task Name'
date_start_col = 'Date Start' 
date_end_col = 'Date End'

df = pd.read_csv(csv_file_path, parse_dates=[date_start_col, date_end_col])
date_range = pd.date_range(df[date_start_col].min(), df[date_end_col].max())

wb = Workbook()
ws = wb.active

font_size_8 = Font(size=8, name='Arial')
light_blue_fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid')

column_headers = [task_id_col, task_name_col, date_start_col, date_end_col] + [date.strftime('%Y-%m-%d') for date in date_range]
ws.append(column_headers)
for cell in ws[1]:
    cell.font = font_size_8
    cell.alignment = Alignment(horizontal='center')

for _, task in df.iterrows():
    task_id = task[task_id_col]
    task_name = task[task_name_col]
    task_start = task[date_start_col].date()
    task_end = task[date_end_col].date()

    row_num = task_id + 1
    cell1 = ws.cell(row=row_num, column=1, value=task_id)
    cell1.font = font_size_8
    cell2 = ws.cell(row=row_num, column=2, value=task_name)
    cell2.font = font_size_8
    cell3 = ws.cell(row=row_num, column=3, value=task_start)
    cell3.font = font_size_8
    cell4 = ws.cell(row=row_num, column=4, value=task_end)
    cell4.font = font_size_8

    for col_num, date in enumerate(date_range, start=5):
        if task_start <= date.date() <= task_end:
            ws.cell(row=row_num, column=col_num).fill = light_blue_fill
        ws.cell(row=row_num, column=col_num).font = font_size_8
ws.freeze_panes = 'E2' 
wb.save(xlsx_file_path)
print("File saved at:", xlsx_file_path)
