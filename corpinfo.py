import xml.etree.ElementTree as ET
import xlsxwriter 

workbook = xlsxwriter.Workbook('corp_codes.xlsx')
worksheet = workbook.add_worksheet()

tree = ET.parse('CORPCODE.XML')
root= tree.getroot()
numbers = []
names = []
i = 0
for child in root:
    if "보험" in root[i][1].text:
        numbers.append(root[i][0].text)
        names.append(root[i][1].text)
    i += 1

cell_format = workbook.add_format()
cell_format.set_align('right')

worksheet.set_column(0, 0, 20)
worksheet.write('A1', '회사명')
worksheet.write('B1', '고유번호')
worksheet.write('C1', '계:', cell_format)
worksheet.write_number('C2', len(numbers))
row = 1
col = 0
for name, code in zip(names, numbers):
    worksheet.write_string(row, col, name)
    worksheet.write_number(row, col + 1, int(code))
    row += 1
    
workbook.close()