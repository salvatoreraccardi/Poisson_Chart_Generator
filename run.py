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
bold = workbook.add_format({'bold': 1})

array_offset = 1
poissonK = 20 + array_offset
poissonMi = 0.5

data = [
    [None] * poissonK,
    [None] * poissonK,
]

for i in range(poissonK):
    data[0][i] = i
    data[1][i] = poisson(i, poissonMi)
    print(data[1][i])


headings = ['X', 'P(X)']

worksheet.write_row('A1', headings, bold)
worksheet.write_column('A2', data[0])
worksheet.write_column('B2', data[1])

chart5 = workbook.add_chart({'type': 'scatter',
                             'subtype': 'smooth'})

test1 = '=Poisson!$A$2:$A$' + str(poissonK)
test2 = '=Poisson!$B$2:$B$' + str(poissonK)

chart5.add_series({
    'name':       '=Poisson!$B$1',
    'categories': test1,
    'values':     test2,
})

chart5.set_title ({'name': 'Modello di Poisson - P(X)'})
chart5.set_style(15)
chart5.set_size({'width': 920, 'height': 576})

worksheet.insert_chart('E2', chart5)

workbook.close()