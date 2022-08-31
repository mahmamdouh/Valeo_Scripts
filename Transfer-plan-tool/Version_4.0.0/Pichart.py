import xlsxwriter
from datetime import datetime


# Create a new Format object to formats cells
# in worksheets using add_format() method .
def Pi_Chart_creation(workbook,worksheet,data,info_list):
    # here we create bold format object .
    bold = workbook.add_format({'bold': 1})

    # create a data list .
    headings = ['Title', 'Number']



    # Write a row of data starting from 'A1'
    # with bold format.
    worksheet.write_row('A1', headings, bold)

    # Write a column of data starting from
    # A2, B2, C2 respectively.
    worksheet.write_column('A2', data[0])
    worksheet.write_column('B2', data[1])

    # Create a chart object that can be added
    # to a worksheet using add_chart() method.

    # here we create a pie chart object .
    chart1 = workbook.add_chart({'type': 'pie'})
    chart2 = workbook.add_chart({'type': 'pie'})
    chart3 = workbook.add_chart({'type': 'pie'})
    # Add a data series to a chart
    # using add_series method.
    # Configure the first series.
    # [sheetname, first_row, first_col, last_row, last_col].
    chart1.add_series({
        'name': 'SWR Coverage',
        'categories': ['KPI', 1, 0, 2, 0],
        'values': ['KPI', 1, 1, 2, 1],
    })
    chart2.add_series({
        'name': 'DWI Coverage',
        'categories': ['KPI', 3, 0, 4, 0],
        'values': ['KPI', 3, 1, 4, 1],
    })

    chart3.add_series({
        'name': 'SWC Coverage',
        'categories': ['KPI', 5, 0, 6, 0],
        'values': ['KPI', 5, 1, 6, 1],
    })
    # Add a chart title
    chart1.set_title({'name': 'SWR Coverage'})
    chart2.set_title({'name': 'DWI Coverage'})
    chart3.set_title({'name': 'SWC Coverage'})
    # Set an Excel chart style. Colors with white outline and shadow.
    chart1.set_style(10)
    chart2.set_style(10)
    chart3.set_style(10)

    # Insert the chart into the worksheet(with an offset).
    # the top-left corner of a chart is anchored to cell C2.
    worksheet.insert_chart('C2', chart1, {'x_offset': 25, 'y_offset': 10})
    worksheet.insert_chart('L2', chart2, {'x_offset': 25, 'y_offset': 10})
    worksheet.insert_chart('L18', chart3, {'x_offset': 25, 'y_offset': 10})
    # Finally, close the Excel file
    # via the close() method.
    return workbook


