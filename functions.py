def collaborator_write(row, col, collaborator, worksheet, tittle_style, cell_style, tittle):
    row_init = row
    col_init = col
    worksheet.merge_range(f'A{row_init}:AG{row_init}', tittle, tittle_style)
    for item in collaborator:
        count = (len(item))
        for i in item.values():
            worksheet.write_string(row_init, col_init, str(i), cell_style)
            col_init += 1
            if col_init == count:
                row_init += 1
                col_init = 0
    return row_init
