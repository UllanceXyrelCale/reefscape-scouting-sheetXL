# This is responsible for creating various bar graphs that are needed for every sheets

import openpyxl as xl
from openpyxl.chart import BarChart, Reference

def bar_graph(wb, sheet, start_row, start_column, graph_location):
    wb_sheet = wb[sheet]

    print('setup: done')

    values = Reference(wb_sheet,
                       min_row=start_row,
                       max_row=wb_sheet.max_row,
                       min_col=start_column,
                       max_col=start_column)
    print('values: done')

    graph = BarChart()
    print('graphcreated: done')

    graph.add_data(values)
    print('valuesadded: done')

    wb_sheet.add_chart(graph, graph_location)
    print('graphdeployed: done')