"""

Wrapper functions for xlwings

"""

import pandas as pd
from xlwings import Workbook, Sheet, Range

# Utililty functions for dataframes
def clear(range_name, ws=None):
    """Clear a table started in a range cell, ex: clear('a1').

    clear(range_name, ws=None)
    """
    ws = Sheet.active() if ws is None else ws
    Range(ws.name, range_name).table.clear()

def clear_ws(ws=None):
    """Clear a table started in a range cell, ex: clearcontents('a1').

    clear_ws(ws=None)
    """
    ws = Sheet.active() if ws is None else ws
    ws.clear()

def clearcontents(range_name, ws=None):
    """Clear a table started in a range cell, ex: clearcontents('a1').

    clearcontents(range_name, ws=None)
    """
    ws = Sheet.active() if ws is None else ws
    Range(ws.name, range_name).table.clear_contents()

def clearcontents_ws(ws=None):
    """Clear entire worksheet, ex: clearcontents_ws(ws=worksheetobj).

    clearcontents_ws(ws=None)
    """
    ws = Sheet.active() if ws is None else ws
    ws.clear_contents()

def list_write(range_name, value_list, ws=None):
    """Write a list vertially.

    list_write(range_name, value_list, ws=None)
    """
    ws = Sheet.active() if ws is None else ws
    # Turn list into a column
    value_column = [[e] for e in value_list]
    Range(ws.name, range_name).value = value_column

def list_read(range_name, ws=None):
    """Read a list vertially.

    list_read(range_name, value_list, ws=None)
    """
    ws = Sheet.active() if ws is None else ws
    # Range(range_name).options(transpose=True).value = value_list
    datalist = Range(ws.name, range_name).vertical.value
    return datalist

def df_read(range_name, ws=None):
    """Return dataframe from range name and Sheet.; ex: df = df_read('a1').

    df_read(range_name, ws=None)
    """
    ws = Sheet.active() if ws is None else ws
    data = Range(ws.name, range_name).table.value
    df = pd.DataFrame(data[1:], columns=data[0])
    return df

def df_write(df, range_name, ws=None):
    """Write a dataframe to a cell.

    df_write(df, range_name, ws=None)
    """
    ws = Sheet.active() if ws is None else ws
    Range(ws.name, range_name, index=False).value = df # without pd indices

def autofit(ws=None):
    """Autofit columns of worksheet
    """
    ws = Sheet.active() if ws is None else ws
    ws.autofit(axis='columns')
