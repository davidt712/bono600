import xlsxwriter

workbook = xlsxwriter.Workbook("Otra-forma.xlsx")
worksheet = workbook.add_worksheet()

título1 = ['VARONES']
título2 = ['MUJERES']
headings = ['Departamentos', 'Cantidades']

data = [
    ['ANCASH', 'APURIMAC', 'CALLAO', 'HUANCAVELICA', 'HUANUCO', 'ICA', 'JUNIN', 'LIMA', 'PASCO'],
    [1614, 699, 900, 663, 1256, 792, 1644, 8542, 352],
    [1668, 710, 1013, 681, 1237, 837, 1737, 9287, 367],
]


worksheet.write_row('A1', título1)
worksheet.write_row('D1', título2)
worksheet.write_row('A2', headings)
worksheet.write_row('D2', headings)

worksheet.write_column('A3', data[0])
worksheet.write_column('B3', data[1])
worksheet.write_column('D3', data[0])
worksheet.write_column('E3', data[1])

chart1 = workbook.add_chart({'type': 'pie'})
chart2 = workbook.add_chart({'type': 'pie'})

chart1.add_series({
    'name': 'Bono Varones',
    'categories': ['Sheet1', 2, 0, 10, 0],
    'values': ['Sheet1', 2, 1, 10, 1],
})

chart2.add_series({
    'name': 'Bono Mujeres',
    'categories': ['Sheet1', 2, 3, 10, 3],
    'values': ['Sheet1', 2, 4, 10, 4],
})


chart1.set_title({'name': 'Bonos Hombres'})
chart2.set_title({'name': 'Bonos Mujeres'})

chart1.set_style(10)
chart2.set_style(2)

worksheet.insert_chart('A13', chart1)
worksheet.insert_chart('H2', chart2)

workbook.close()