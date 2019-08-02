import calendar
import xlsxwriter

from testes_pablo.config import days, doctor
from testes_pablo.excel.format import text_tittle_style, tittles_style, period_style, data_title_rule, \
    merge_format_month_rule, merge_empty_line_rule, days_number_format_rule, days_number_bg_rule, image_tittle_style
from testes_pablo.excel.functions import collaborator_write

workbook = xlsxwriter.Workbook('hello.xlsx')
worksheet1 = workbook.add_worksheet()

worksheet1.set_landscape()
worksheet1.set_paper(9)
worksheet1.center_horizontally()

text_tittle = workbook.add_format(text_tittle_style)
titulos = workbook.add_format(tittles_style)
celula = workbook.add_format()
period = workbook.add_format(period_style)
# -
data_title = workbook.add_format(data_title_rule)
merge_format_month = workbook.add_format(merge_format_month_rule)
merge_empty_line = workbook.add_format(merge_empty_line_rule)
days_number_format = workbook.add_format(days_number_format_rule)
days_number_bg = workbook.add_format(days_number_bg_rule)

worksheet1.insert_image('A1', 'logos2.png', image_tittle_style)

worksheet1.merge_range('A1:A6', '')
worksheet1.merge_range('A7:A8', 'EQUIPE', titulos)
worksheet1.merge_range('B7:B8', 'HORÁRIO', titulos)
worksheet1.merge_range('B1:AG6', '')

worksheet1.set_column('A:A', 74)
worksheet1.set_column('B:B', 18.30)
worksheet1.set_column('C1:AG1', 4.4)

worksheet1.set_row(0, 12.62)
worksheet1.set_row(1, 12.62)
worksheet1.set_row(2, 12.62)
worksheet1.set_row(3, 12.62)
worksheet1.set_row(4, 12.62)
worksheet1.set_row(5, 12.62)
worksheet1.set_row(8, 15)

worksheet1.write_string('B1',
                        'ESCALA MÉDICA CENTRO CIRÚRGICO\n HOSPITAL ESTADUAL DE PIRENÓPOLIS ERNESTINA LOPES JAIME\n JULHO / 2019',
                        text_tittle)

year = 2019
month = 7

monthRange = calendar.monthrange(year, month)

row_number_day = 7
col_number_day = 2
week_day = monthRange[0]
i = 0
row_number = 6
col_number = 2
month_days = 1

while i < monthRange[1]:
    # Cria a coluna onde fica os dias(1,2,3..)
    worksheet1.write(row_number, col_number, month_days, days_number_format)

    # Cria a coluna onde fica os dias da semana(S,D,Q...)
    if week_day == 5 or week_day == 6:

        worksheet1.write_column(row_number_day, col_number_day, str(days[week_day]), days_number_bg)
        worksheet1.set_column(col_number_day, col_number_day, 3, days_number_bg)
    else:
        worksheet1.write_column(row_number_day, col_number_day, str(days[week_day]), days_number_format)
        worksheet1.set_column(col_number_day, col_number_day, 3, days_number_format)

    week_day = week_day + 1
    if week_day > 6:
        week_day = 0

    i += 1
    col_number_day += 1
    col_number += 1
    month_days += 1

a = collaborator_write(9, 0, doctor, worksheet1, period, data_title, 'Diurno')
b = a + 1
c = collaborator_write(b, 0, doctor, worksheet1, period, data_title, 'Noturno')


worksheet1.set_default_row(hide_unused_rows=True)
worksheet1.set_column('AM:XFD', None, None, {'hidden': True})

workbook.close()
