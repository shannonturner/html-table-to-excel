import xlwt

def html_table_to_excel(table):

    """ html_table_to_excel(table): Takes an HTML table of data and formats it so that it can be inserted into an Excel Spreadsheet.
    """

    data = {}

    table = table[table.index('<tr>'):table.index('</table>')]

    rows = table.split('</tr>')
    for (x, row) in enumerate(rows):   
        columns = row.split('</td>')
        data[x] = {}
        for (y, col) in enumerate(columns):
            data[x][y] = col.replace('<tr>', '').replace('<td>', '')

    return data

def export_to_xls(data, title='Sheet1', filename='export.xls'):

    """ export_to_xls(data, title, filename): Exports data to an Excel Spreadsheet.
    Data should be a dictionary with rows as keys; the values of which should be a dictionary with columns as keys; the value should be the value at the x, y coordinate.
    """

    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet(title)

    for x in sorted(data.iterkeys()):
        for y in sorted(data[x].iterkeys()):
            worksheet.write(x, y, data[x][y])

    workbook.save(filename)

    return
