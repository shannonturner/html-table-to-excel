from xlsxwriter.workbook import Workbook

def export_to_xlsx(data, chart={}, filename='export.xlsx'):

    """ export_to_xlsx(data, chart, filename): Exports data to an Excel Spreadsheet.
    Data should be a dictionary with rows as keys; the values of which should be a dictionary with columns as keys; the value should be the value at the x, y coordinate.
    Chart should be either True, in which case the values will be generated, or a dictionary with the following keys:
        type: string, choose from: line, ...
        series: list of dictionaries with the following: name, categories, values
        title, xaxis_title, yaxis_title: strings
        style: int
        cell
        x_offset
        y_offset
    """

    workbook = Workbook(filename)
    worksheet = workbook.add_worksheet()
   
    for x in sorted(data.iterkeys()):
        for y in sorted(data[x].iterkeys()):
            try:
                if float(data[x][y]).is_integer():
                    worksheet.write(x, y, int(float(data[x][y])))
                else:
                    worksheet.write(x, y, float(data[x][y]))
            except ValueError:
                worksheet.write(x, y, data[x][y])

    if chart is not {}:

        generated_chart = generate_chart(data)
        for key in generated_chart.iterkeys():
            if chart.get(key) is None:
                chart[key] = generated_chart[key]

        new_chart = workbook.add_chart({'type': chart['type']})

        for each_series in chart['series']:
            new_chart.add_series({'name': each_series['name'], 'categories': each_series['categories'], 'values': each_series['values']})

        new_chart.set_title({'name': chart['title']})
        new_chart.set_x_axis({'name': chart['xaxis_title']})
        new_chart.set_y_axis({'name': chart['yaxis_title']})

        new_chart.set_style(chart['style'])

        worksheet.insert_chart(chart['cell'], new_chart, {'x_offset': chart['x_offset'], 'y_offset': chart['y_offset']})

    workbook.close()
   
    return

def convert_number_to_excel_colname(n):

    """ Converts a number to an excel column name, A through IV (1-256)
    """

    assert 0 < n <= 256

    alphabet = [chr(x) for x in xrange(65, 91)]

    if n > 26:
        return '{0}{1}'.format(alphabet[(n/26) - 1], alphabet[(n%26) - 1])
    else:
        return alphabet[(n%26) - 1]

def generate_chart(data, sheet='Sheet1', chart_type='line', chart_title='', chart_xaxis_title='', chart_yaxis_title='', chart_x_offset=5, chart_y_offset=5, chart_style=19):

    """ generate_chart(): From the table-as-dictionary, Generates the values needed to create a chart by xlsxwriter.
    """

    rows = max(data.keys()) + 1
    cols = max(data[0].keys()) + 1

    chart = {}

    chart['type'] = chart_type
    chart['title'] = chart_title
    chart['xaxis_title'] = chart_xaxis_title
    chart['yaxis_title'] = chart_yaxis_title
    chart['x_offset'] = chart_x_offset
    chart['y_offset'] = chart_y_offset
    chart['style'] = chart_style

    if rows > cols or cols > 254:
        chart['cell'] = '{0}2'.format(convert_number_to_excel_colname(cols + 2))
    else:
        chart['cell'] = 'B{0}'.format(rows + 2)
    
    chart['series'] = []
    categories = [sheet, 1, 0, rows - 1, 0] # first row, first col, last row, last col
    for col in sorted(data[0].keys()[1:]):
        chart['series'].append({'name': data[0][col], 'categories': categories, 'values': [sheet, 1, col, rows - 1,col]})

    return chart
