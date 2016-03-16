from xlrd import open_workbook
import json
name = 'SGF_2013_SGF003.xls'
bigdata = {}

book = open_workbook('SGF.xlsx')
sheet = book.sheet_by_index(0)

for col_i in range(sheet.ncols):
    if col_i > 1:
        data = {}
        genrev = {}
        for row_i in range(3, 10):
            genrev[sheet.cell_value(row_i, 0).strip()] = sheet.cell_value(row_i, col_i)
        insrev = {}
        for row_i in range(11, 15):
            insrev[sheet.cell_value(row_i, 0).strip()] = sheet.cell_value(row_i, col_i)
        totalrev = {}
        totalrev['General Revenue'] = genrev
        totalrev['Insurance Trust Revenue'] = insrev
        genexp = {}
        for row_i in range(17, 29):
          genexp[sheet.cell_value(row_i, 0).strip()] = sheet.cell_value(row_i, col_i)
        insexp = {}
        for row_i in range(30, 34):
            insexp[sheet.cell_value(row_i, 0).strip()] = sheet.cell_value(row_i, col_i)
        totalexp = {}
        totalexp['General Expenditure'] = genexp
        totalexp['Insurance Trust Expenditure'] = insexp
        data['Total Revenue'] = totalrev
        data['Total Expenditure'] = totalexp
        bigdata[sheet.cell_value(0, col_i).strip()] = data

with open('data.json', 'w') as outfile:
    json.dump(bigdata, outfile, indent=4, sort_keys=True)



