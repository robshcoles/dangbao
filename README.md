# dangbao


from pathlib import Path
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.styles import Font
from openpyxl.styles import Side,Border

p = Path('D:\\POAOCode\\POAOT2112\\excel_2\\各组上报')
# file_path = Path()
p1 = p.iterdir()
for i in p1:
    f = i.name
    f_path = Path('D:\\POAOCode\\POAOT2112\\excel_2\\各组上报\\' + f)
    f_e_path = f_path.iterdir()
    cou = 1
    for every_exc in f_e_path:
        book = openpyxl.load_workbook(every_exc)
        book_hui = openpyxl.load_workbook('D:\\POAOCode\\POAOT2112\\excel_2\\全省汇总.xlsx')
        sheet = book['Sheet1']
        sheet_hui = book_hui['Sheet1']
        values_sheet = list(sheet.iter_rows(min_row=7,max_row=7,values_only=True))
        #print(values_sheet)
        eve_rows = sheet_hui.max_row + 1
        e = str(cou) + "月"
        sheet_hui['b'+str(eve_rows)].value = e
        sheet_hui['c'+str(eve_rows)].value = values_sheet[0][1]
        sheet_hui['d'+str(eve_rows)].value = values_sheet[0][2]
        sheet_hui['e'+str(eve_rows)].value = values_sheet[0][3]
        sheet_hui['f'+str(eve_rows)].value = values_sheet[0][4]
        sheet_hui['g'+str(eve_rows)].value = values_sheet[0][5]
        sheet_hui['h'+str(eve_rows)].value = values_sheet[0][6]
        sheet_hui['i'+str(eve_rows)].value = values_sheet[0][7]
        sheet_hui['j'+str(eve_rows)].value = values_sheet[0][8]
        sheet_hui['k'+str(eve_rows)].value = values_sheet[0][9]
        sheet_hui['l'+str(eve_rows)].value = values_sheet[0][10]
        sheet_hui['m'+str(eve_rows)].value = values_sheet[0][11]
        cou += 1
        book_hui.save('D:\\POAOCode\\POAOT2112\\excel_2\\全省汇总.xlsx')
    max_max_row = sheet_hui.max_row
    zong_row = sheet_hui.max_row + 1
    small_row = zong_row - 12
    for small_col_count in range(3,14):
        res = 0
        for small_row_count in range(small_row,zong_row):
            qq = int(sheet_hui.cell(small_row_count,small_col_count).value)
            res += qq
        sheet_hui.cell(zong_row,small_col_count).value = res    
        book_hui.save('D:\\POAOCode\\POAOT2112\\excel_2\\全省汇总.xlsx')   

    sheet_hui.cell(zong_row,2).value = "单组汇总"
    sheet_hui['a'+str(small_row)].value = values_sheet[0][0]
    sheet_hui.merge_cells(start_row=small_row,start_column=1,end_row=zong_row,end_column=1)
    area = sheet_hui['a7:m'+str(zong_row)]
    style = Side(border_style='thin',color='000000')
    for area_row in area:
        for area_range in area_row:
            area_range.alignment = Alignment(horizontal='center',vertical='center')
            area_range.font = Font(name='宋体',size=14,bold=True)
            area_range.border = Border(left=style,right=style,top=style,bottom=style)
           
    book_hui.save('D:\\POAOCode\\POAOT2112\\excel_2\\全省汇总.xlsx')
print('已汇总成功')
