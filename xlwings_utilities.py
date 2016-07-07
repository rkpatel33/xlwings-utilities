"""
To use:

import os, sys
sys.path.append(os.path.abspath('/Users/rishi/git/xlwings-utilities/'))
from xlwingswrapper import clear, clearcontents, list_write, list_read, df_read, df_write

"""

# Wrapper functions for xlwings

import pandas as pd
from xlwings import Workbook, Sheet, Range

# Utililty functions for dataframes
def clear(range_name, ws=None):
    """Clear a table started in a range cell, ex: clear('a1')"""
    ws = Sheet.active() if ws is None else ws
    Range(ws.name, range_name).table.clear()

def clear_ws(ws=None):
    """Clear a table started in a range cell, ex: clearcontents('a1')"""
    ws = Sheet.active() if ws is None else ws
    Range(ws.name).clear()

def clearcontents(range_name, ws=None):
    """Clear a table started in a range cell, ex: clearcontents('a1')"""
    ws = Sheet.active() if ws is None else ws
    Range(ws.name, range_name).table.clear_contents()

def clearcontents_ws(ws=None):
    """Clear a table started in a range cell, ex: clearcontents('a1')"""
    ws = Sheet.active() if ws is None else ws
    Range(ws.name).clear_contents()

def list_write(range_name, value_list, ws=None):
    """Write a list vertially"""
    ws = Sheet.active() if ws is None else ws
    # Turn list into a column
    value_column = [[e] for e in value_list]
    Range(ws.name, range_name).value = value_column

def list_read(range_name, value_list, ws=None):
    """Read a list vertially"""
    ws = Sheet.active() if ws is None else ws
    # Range(range_name).options(transpose=True).value = value_list
    datalist = Range(ws.name, range_name).vertical.value

def df_read(range_name, ws=None):
    """Return dataframe from range name and Sheet.; ex: df = df_read('a1')"""
    ws = Sheet.active() if ws is None else ws
    data = Range(ws.name, range_name).table.value
    df = pd.DataFrame(data[1:], columns=data[0])
    return df

def df_write(range_name, df, ws=None):
    """Write a dataframe to a cell."""
    ws = Sheet.active() if ws is None else ws
    Range(ws.name, range_name, index=False).value = df # without pd indices
