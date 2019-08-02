import xlsxwriter
import calendar
from config import doctor, days
from excel.format import image_header_rule, merge_title_format_rule, data_title_rule, \
    name_data_title_rule, merge_format_month_rule, merge_empty_line_rule, days_number_format_rule, days_number_bg_rule

workbook = xlsxwriter.Workbook('file.xlsx')

worksheet = workbook.add_worksheet('PLANTONISTA')

# Formatação de estilo
image_header = workbook.add_format(image_header_rule)
merge_format_title = workbook.add_format(merge_title_format_rule)
data_title = workbook.add_format(data_title_rule)
name_data_title = workbook.add_format(name_data_title_rule)
merge_format_month = workbook.add_format(merge_format_month_rule)
merge_empty_line = workbook.add_format(merge_empty_line_rule)
days_number_format = workbook.add_format(days_number_format_rule)
days_number_bg = workbook.add_format(days_number_bg_rule)

# Formatação de tamanho
worksheet.set_column('B1:G1', 12)
worksheet.set_column(0, 0, 32)
worksheet.set_row(5, 19)
# Inserção da imagem
worksheet.insert_image('A1', 'logos2.png')

# Mesclagens
worksheet.merge_range('A1:AL5', '', image_header)
worksheet.merge_range('A6:AL6', '', merge_empty_line)
worksheet.merge_range('A8:A9', 'Nome', merge_format_title)
worksheet.merge_range('B8:B9', 'Função', merge_format_title)
worksheet.merge_range('C8:C9', 'Setor', merge_format_title)
worksheet.merge_range('D8:D9', 'CRM', merge_format_title)
worksheet.merge_range('E8:E9', 'Vinculo', merge_format_title)
worksheet.merge_range('F8:F9', 'CH Semanal', merge_format_title)
worksheet.merge_range('G8:G9', 'Intervalo', merge_format_title)
worksheet.merge_range('A7:G7', '', merge_format_month)
worksheet.merge_range('H7:AL7', 'MÊS/ANO: JULHO 2019', merge_format_month)

# Criando a sequencia semanal correta para cada mês.
# Com os dias do mês e os dias da semana
year = 2019
month = 6

monthRange = calendar.monthrange(year, month)

row_number_day = 8
col_number_day = 7
j = monthRange[0]
i = 0
row_number = 7
col_number = 7
k = 1

while i < monthRange[1]:
    # Cria a coluna onde fica os dias(1,2,3..)
    worksheet.write(row_number, col_number, k, days_number_format)

    # Cria a coluna onde fica os dias da semana(S,D,Q...)
    if j == 5 or j == 6:

        worksheet.write_column(row_number_day, col_number_day, str(days[j]), days_number_bg)
        worksheet.set_column(col_number_day, col_number_day, 3, days_number_bg)
    else:
        worksheet.write_column(row_number_day, col_number_day, str(days[j]), days_number_format)
        worksheet.set_column(col_number_day, col_number_day, 3, days_number_format)

    j = j + 1
    if j > 6:
        j = 0

    i += 1
    col_number_day += 1
    col_number += 1
    k += 1

# Cria os valores trazidos pelo array

row_array = 9
col_array = 0
len_array = len(doctor)
number_init_row = 10
p = 0

while p < len_array:
    for item in doctor:
        worksheet.write_row(f'A{number_init_row}', item.values(), data_title)
    number_init_row += 1
    p += 1

worksheet.set_default_row(hide_unused_rows = True)
worksheet.set_column('AM:XFD', None, None, {'hidden': True})
workbook.close()