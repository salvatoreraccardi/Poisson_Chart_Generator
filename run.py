#
# An example of generating Poisson chart with Python and Excel
#
# 05/03/2020 - Salvatore Raccardi, info@objex-iot.tech
#

import xlsxwriter
import math

def poisson(n, mi):
    value = math.exp(-mi) * ( math.pow(mi, n) / math.factorial(n) )
    return value

workbook = xlsxwriter.Workbook('poisson.xlsx')
worksheet = workbook.add_worksheet('Poisson')
formatTitle = workbook.add_format({'bold': 1, 'italic': 1, 'align': 'center', 'bg_color': 'yellow', 'align': 'center','valign': 'vcenter', 'bottom': 1, 'left': 1, 'right': 1})
formatRef = workbook.add_format({'bold': 1, 'align': 'center', 'bg_color': 'yellow', 'bottom': 1, 'left': 1, 'right': 1})
formatData = workbook.add_format({'align': 'center', 'left': 1, 'right': 1})
formatEndDataset = workbook.add_format({'bold': 1, 'italic': 1, 'align': 'center', 'bg_color': 'red', 'align': 'center','valign': 'vcenter', 'top': 1, 'bottom': 1, 'left': 1, 'right': 1})

# Parametri di Poisson
poissonN = 33
poissonMi = 3

# Adattamento parametri
array_offset = 1
poissonK = poissonN + array_offset
array_end = poissonK + 3

# Array 2D - dinamico
data = [
    [None] * poissonK,
    [None] * poissonK,
]

for i in range(poissonK):
    data[0][i] = i
    data[1][i] = poisson(i, poissonMi)
    print(data[1][i])


# Parametri - xlsxwriter settings
worksheet.merge_range('A1:B1', 'PARAMETRI', formatTitle)
worksheet.write('A2', 'N', formatData)
worksheet.write('B2', poissonN, formatData)
worksheet.write('A3', 'Mi', formatData)
worksheet.write('B3', poissonMi, formatData)
worksheet.set_column('A:B', 16)

# Dataset - xlsxwriter settings
headings = ['X', 'P(X)']

worksheet.merge_range('E1:F1', 'DATASET', formatTitle)
worksheet.write_row('E2', headings, formatRef)
worksheet.write_column('E3', data[0], formatData)
worksheet.write_column('F3', data[1], formatData)
worksheet.merge_range('E'+ str(array_end) + ':F' + str(array_end), 'END - DATASET', formatEndDataset)
worksheet.set_column('E:F', 15)

# Poisson char
chart5 = workbook.add_chart({'type': 'scatter',
                             'subtype': 'smooth'})

test1 = '=Poisson!$E$3:$E$' + str(poissonK)
test2 = '=Poisson!$F$3:$F$' + str(poissonK)

chart5.add_series({
    'name':       '=Poisson!$F$2',
    'categories': test1,
    'values':     test2,
})

chart5.set_title ({'name': 'Modello di Poisson - P(X)'})
chart5.set_style(15)
chart5.set_size({'width': 920, 'height': 576})

worksheet.insert_chart('I2', chart5)

workbook.close()