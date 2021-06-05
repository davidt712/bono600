import xlsxwriter

workbook = xlsxwriter.Workbook("Bono_600.xlsx")
worksheet = workbook.add_worksheet()

headings = ['Departamentos', 'Cantidades']

data = [
    ['ANCASH', 'APURIMAC', 'CALLAO', 'HUANCAVELICA', 'HUANUCO', 'ICA', 'JUNIN', 'LIMA', 'PASCO'],
    [1614, 699, 900, 663, 1256, 792, 1644, 8542, 352],
    [1668, 710, 1013, 681, 1237, 837, 1737, 9287, 367],
]

worksheet.write_row('A2', headings)
worksheet.write_row('I2', headings)

worksheet.write_column('A3', data[0])
worksheet.write_column('B3', data[1])
worksheet.write_column('I3', data[0])
worksheet.write_column('J3', data[1])

chart1 = workbook.add_chart({'type': 'pie'})
chart2 = workbook.add_chart({'type': 'pie'})


#Gráfico 1
chart1.add_series({
    'name': 'Bono Varones',
    'categories': ['Sheet1', 2, 0, 10, 0],
    'values': ['Sheet1', 2, 1, 10, 1],
})
chart1.set_title({'name': 'Bonos Hombres'}) #Nombre del gráfico
chart1.set_style(10) #Estilo del gráfico
worksheet.insert_chart('A13', chart1) #Ubicación del gráfico

#Gráfico 2
chart2.add_series({
    'name': 'Bono Mujeres',
    'categories': ['Sheet1', 2, 8, 10, 8],
    'values': ['Sheet1', 2, 9, 10, 9],
})
chart2.set_title({'name': 'Bonos Mujeres'})
chart2.set_style(2)
worksheet.insert_chart('I13', chart2)


#Formato a la celda
merge_format = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': 'yellow'})
worksheet.merge_range('A1:B1', "Varones", merge_format)
worksheet.merge_range('I1:J1', "Mujeres", merge_format)
worksheet.set_column('A:B', 15)
worksheet.set_column('I:J', 15)



workbook.close()
