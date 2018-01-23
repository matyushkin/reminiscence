import re
import xlwt

f = open('Data.txt')
text = f.read()

from xlwt import Workbook
excel_file = xlwt.Workbook()
sheet = excel_file.add_sheet('Memories')

memory_col = sheet.col(2)
memory_col.width = 256*72   # 72 symbols width

alignment = xlwt.Alignment()
alignment.wrap = 1
alignment.horz = xlwt.Alignment.HORZ_LEFT # May be: HORZ_GENERAL, HORZ_LEFT, HORZ_CENTER, HORZ_RIGHT, HORZ_FILLED, HORZ_JUSTIFIED, HORZ_CENTER_ACROSS_SEL, HORZ_DISTRIBUTED
alignment.vert = xlwt.Alignment.VERT_TOP # May be: VERT_TOP, VERT_CENTER, VERT_BOTTOM, VERT_JUSTIFIED, VERT_DISTRIBUTED

style = xlwt.XFStyle()
style.alignment = alignment

sheet.write(0, 0, '№')
sheet.write(0, 1, 'Имя')
sheet.write(0, 2, 'Воспоминание')

codes = re.findall('\[(\d*)\]', text)
names  = re.findall('\][\s\t]([^:.]*):?\s?\n', text)
texts = []

for i in range(len(codes)):
    start = text.index('[' + codes[i] + ']')
    start = text.index('\n', start) + len('\n«')

    if i != len(codes)-1:
        end = text.index('[' + codes[i+1] + ']')
    else:
        end = len(text)
    
    line = text[start:end].rstrip()[:-2] + '.'
    texts.append(line)

for code, name, t in zip(codes, names, texts):
    code = int(code)
    sheet.write(code, 0, code, style)
    sheet.write(code, 1, name, style)
    sheet.write(code, 2, t, style)

excel_file.save("Data.xls")    
