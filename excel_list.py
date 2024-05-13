from openpyxl import Workbook
from openpyxl.styles import Alignment
import db_select

def create_excel_list():
    wb = Workbook()
    ws = wb.active

    ws['A1'] = 'Номенклатура'
    ws['B1'] = 'начальный остаток'
    ws['C1'] = ''
    ws['D1'] = 'приход'
    ws['E1'] = ''
    ws['F1'] = 'расход'
    ws['G1'] = ''
    ws['H1'] = 'конечный остаток'
    ws['I1'] = ''

    ws['A2'] = ''
    ws['B2'] = 'кол-во'
    ws['C2'] = 'сумма'
    ws['D2'] = 'кол-во'
    ws['E2'] = 'сумма'
    ws['F2'] = 'кол-во'
    ws['G2'] = 'сумма'
    ws['H2'] = 'кол-во'
    ws['I2'] = 'сумма'

    abc = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']

    for i in range(len(db_select.tovari)):
        for j in range(9):
            if abc[j] == 'A':
                ws[f'{abc[j]}{2+i+1}'] = db_select.tovari[i]
            if abc[j] == 'B':
                ws[f'{abc[j]}{2+i+1}'] = db_select.nach_o[i][1]
            if abc[j] == 'C':
                ws[f'{abc[j]}{2+i+1}'] = db_select.nach_o[i][2]
            if abc[j] == 'D':
                ws[f'{abc[j]}{2+i+1}'] = db_select.prihod[i][1]
            if abc[j] == 'E':
                ws[f'{abc[j]}{2+i+1}'] = db_select.prihod[i][2]
            if abc[j] == 'F':
                ws[f'{abc[j]}{2+i+1}'] = db_select.rashod[i][1]
            if abc[j] == 'G':
                ws[f'{abc[j]}{2+i+1}'] = (db_select.sebestoimost[i][1] * db_select.rashod[i][1])
            if abc[j] == 'H':
                ws[f'{abc[j]}{2+i+1}'] = f'=B{2+i+1} + D{2+i+1} - F{2+i+1}'
            if abc[j] == 'I':
                ws[f'{abc[j]}{2+i+1}'] = f'=C{2+i+1} + E{2+i+1} - G{2+i+1}'

    ws[f'A{2 + len(db_select.tovari) + 1}'] = 'итого'
    ws[f'B{2 + len(db_select.tovari) + 1}'] = ''
    ws[f'C{2 + len(db_select.tovari) + 1}'] = f'{sum(item[2] for item in db_select.nach_o)}'
    ws[f'D{2 + len(db_select.tovari) + 1}'] = ''
    ws[f'E{2 + len(db_select.tovari) + 1}'] = f'{sum(item[2] for item in db_select.prihod)}'
    ws[f'F{2 + len(db_select.tovari) + 1}'] = ''
    ws[f'G{2 + len(db_select.tovari) + 1}'] = f'{db_select.sebestoimost_m}'
    ws[f'H{2 + len(db_select.tovari) + 1}'] = ''
    ws[f'I{2 + len(db_select.tovari) + 1}'] = f'{db_select.konech_o}'

    ws.merge_cells('A1:A2')
    ws.merge_cells('B1:C1')
    ws.merge_cells('D1:E1')
    ws.merge_cells('F1:G1')
    ws.merge_cells('H1:I1')

    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws['B1'].alignment = Alignment(horizontal='center', vertical='center')
    ws['D1'].alignment = Alignment(horizontal='center', vertical='center')
    ws['F1'].alignment = Alignment(horizontal='center', vertical='center')
    ws['H1'].alignment = Alignment(horizontal='center', vertical='center')

    ws[f'A{2 + len(db_select.tovari) + 1}'].alignment = Alignment(horizontal='right', vertical='center')

    ws[f'C{2 + len(db_select.tovari) + 1}'].alignment = Alignment(horizontal='right', vertical='center')
    ws[f'E{2 + len(db_select.tovari) + 1}'].alignment = Alignment(horizontal='right', vertical='center')
    ws[f'G{2 + len(db_select.tovari) + 1}'].alignment = Alignment(horizontal='right', vertical='center')
    ws[f'I{2 + len(db_select.tovari) + 1}'].alignment = Alignment(horizontal='right', vertical='center')

    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 10
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 10
    ws.column_dimensions['F'].width = 10
    ws.column_dimensions['G'].width = 10
    ws.column_dimensions['H'].width = 10
    ws.column_dimensions['I'].width = 10

    wb.save('report.xlsx')
