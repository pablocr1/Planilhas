import xlsxwriter

from excel.format import text_tittle, image_tittle, tittles

workbook = xlsxwriter.Workbook('hello.xlsx')
worksheet1 = workbook.add_worksheet()

texto_titulo = workbook.add_format(text_tittle)
titulos = workbook.add_format(tittles)
celula = workbook.add_format()

celula.set_fg_color('#00b0f0')

worksheet1.set_landscape()
worksheet1.insert_image('A1', 'logos2.png', image_tittle)

worksheet1.merge_range('A1:A6', '')
worksheet1.merge_range('A7:A8', 'EQUIPE', titulos)
worksheet1.merge_range('B7:B8', 'CARGO', titulos)
worksheet1.merge_range('C7:C8', 'HORÁRIO', titulos)
worksheet1.merge_range('B1:AG6', '')

worksheet1.set_column('A:A', 74)
worksheet1.set_column('B:B', 18.30)
worksheet1.set_column('C:C', 18.30)
worksheet1.set_column('D1:AG1', 4.4)

worksheet1.set_row(0, 12.62)
worksheet1.set_row(1, 12.62)
worksheet1.set_row(2, 12.62)
worksheet1.set_row(3, 12.62)
worksheet1.set_row(4, 12.62)
worksheet1.set_row(5, 12.62)

worksheet1.write('D9:', '', celula)

worksheet1.write_string('B1', 'ESCALA EQUIPE MULTIDISCIPLINAR \n HOSPITAL ESTADUAL DE PIRENÓPOLIS ERNESTINA LOPES JAIME \n JULHO / 2019', texto_titulo)










workbook.close()
