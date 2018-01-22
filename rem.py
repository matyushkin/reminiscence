import re
import xlwt

f = open('Data.txt')
text = f.read()

from xlwt import Workbook
excel_file = xlwt.Workbook()
sheet = excel_file.add_sheet('Memories')

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
    sheet.write(code, 0, code)
    sheet.write(code, 1, name)
    sheet.write(code, 2, t)

excel_file.save("Data.xls")    
