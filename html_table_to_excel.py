import xlwt

def html_table_to_excel(table):

    """ html_table_to_excel(table): Takes an HTML table of data and formats it so that it can be inserted into an Excel Spreadsheet.
    """

    data = {}

    table = table[table.index('<tr>'):table.index('</table>')]

    rows = table.strip('\n').split('</tr>')[:-1]
    for (x, row) in enumerate(rows):   
        columns = row.strip('\n').split('</td>')[:-1] 
        data[x] = {}
        for (y, col) in enumerate(columns):
            data[x][y] = col.replace('<tr>', '').replace('<td>', '').strip()

    return data

def export_to_xls(data, title='Sheet1', filename='export.xls'):

    """ export_to_xls(data, title, filename): Exports data to an Excel Spreadsheet.
    Data should be a dictionary with rows as keys; the values of which should be a dictionary with columns as keys; the value should be the value at the x, y coordinate.
    """

    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet(title)

    for x in sorted(data.iterkeys()):
        for y in sorted(data[x].iterkeys()):
            try:
                if float(data[x][y]).is_integer():
                    worksheet.write(x, y, int(float(data[x][y])))
                else:
                    worksheet.write(x, y, float(data[x][y]))
            except ValueError:
                worksheet.write(x, y, data[x][y])

    workbook.save(filename)

    return
