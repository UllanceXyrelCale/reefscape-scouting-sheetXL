def safety_multiplier(cell_value, multiplier):
    try:
        return cell_value * multiplier
    except TypeError:
        return 0

def safety_divider(cell_value, divisor):
    try:
        return cell_value / divisor
    except TypeError:
        return 0

def safety_adder(a, b, c, d, e):
    try:
        return a + b + c + d + e
    except TypeError:
        return 0

def assign_value(sheet, column, variable):
    sheet.cell(sheet.max_row, column).value = variable