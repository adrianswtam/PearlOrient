#!/usr/bin/env

#
# Given Excel files in _census_ca/
# Convert all file into a SQLite database
#

# require version >= 2.3.0 of openpyxl
import openpyxl
assert(tuple(int(x) for x in openpyxl.__version__.split('.')[:2]) >= (2,3))

from openpyxl import load_workbook
import os
import sqlite3
import contextlib

def scanExcelFiles(dirname='_census_ca'):
    for filename in os.listdir(dirname):
        filename = os.path.join(dirname, filename)
        if not os.path.isfile(filename) or not filename.endswith('.xlsx'):
            continue
        yield filename

def flatten2D(array):
    "Yield each element in a 2D array"
    for row in array:
        for item in row:
            yield item

def readExcelData(filename):
    '''
    Return a sequence of attributes and data from the Excel file.

    Each item of data is a tuple of attributes and a numerical value
    Data in return is in the format of
      (area, question, answer, data_qualifier, value)
    Example:
      ('Z34', 'Ethnicity', 'White', 'Male', 1.23)
    '''

    # Prepare worksheet
    wb = load_workbook(filename=filename, read_only=True)
    sheets = wb.get_sheet_names()
    ws = wb.get_sheet_by_name(sheets[-1])
    area = os.path.basename(filename).split('.')[0]
    colLetter = dict((i+1, c) for i,c in enumerate('ABCDEFGHIJKLMNO'))

    # Read all data into array
    cells = []
    for row in ws.rows:
        cells.append([])
        for cell in row:
            cells[-1].append(cell)

    # Look for numeric data and figure out what it is
    #   Explanation of cellOk, diffColour, sameColour: Excel may have empty cells, which cell.style is not
    #   available if cell.value is None. However, it does not mean the cell has no background colour. Therefore
    #   the trick here is to be carefully distinguish the boundary that we are going to include or to exclude.
    #   The latter should use diffColour or sameColour while the former should you not sameColour or not diffColour.
    cellOk = lambda _cell: _cell.value is not None
    diffColour = lambda _cell, _colour: cellOk(_cell) and _cell.style.fill.fgColor.value != _colour
    sameColour = lambda _cell, _colour: cellOk(_cell) and _cell.style.fill.fgColor.value == _colour
    for cell in flatten2D(cells):
        if cell.data_type != 'n' or cell.value is None:
            continue # we take only numeric data
        # Get cell coordinates
        c = cell.column
        r = cell.row
        assert(r < len(cells)) # should not hit the last row of Excel
        assert(colLetter[c] in 'CDELMN') # numerical data should only in these cols

        # Find row description: At column A or H of same row, find adjacent rows of same style
        descCol = ord('A' if colLetter[c] < 'H' else 'H') - ord('A')
        assert(cells[r-1][descCol].value is None or isinstance(cells[r-1][descCol].value, basestring))
        thisRowColour = cells[r-1][descCol].style.fill.fgColor.value
        thisRowBold = cells[r-1][descCol].style.font.bold
        leftCol = [cells[i][descCol] for i in range(len(cells))] # get the whole column
        start = max(i for i in range(r-1) if diffColour(leftCol[i], thisRowColour))
        end = min(i for i in range(r, len(cells)) if not sameColour(leftCol[i], thisRowColour))
        colSplice = [x for x in leftCol[start+1:end] if x.value and x.style.font.bold == thisRowBold]
        rowText = " ".join(flatten2D(x.value.split() for x in colSplice if x.value)).replace(u'\u2267', '>=')

        # Find table description: At column A or H, find closest rows above me that is green
        goodIdx = [i for i in range(r) if sameColour(leftCol[i], 'FFCCFFCC')] # up to the current cell
        if r-1 in goodIdx: # if this cell is green, it is a footer. Crop goodIdx up to the header
            end = max(i for i in range(r-1) if i not in goodIdx)
            goodIdx = [i for i in goodIdx if i < end]
        end = max(goodIdx)
        start = max(i for i in range(end) if i not in goodIdx)
        tableText = " ".join(filter(None, [x.value.strip() for x in leftCol[start+1:end+1] if x.value]))

        # Find column description: At same column, find closest rows above me that is green
        thisCol = [cells[i][c-1] for i in range(len(cells))] # get the whole column
        goodIdx = [i for i in range(r) if sameColour(thisCol[i], 'FFCCFFCC')] # up to the current cell
        if r-1 in goodIdx: # if this cell is green, it is a footer. Crop goodIdx up to the header
            end = max(i for i in range(r-1) if i not in goodIdx)
            goodIdx = [i for i in goodIdx if i < end]
        end = max(goodIdx)
        start = max(i for i in range(end) if i not in goodIdx)
        colSplice = [x for x in thisCol[start+1:end+1] if x.value]
        while len(colSplice) > 1:
            for i, x in enumerate(colSplice):
                if i > 0 and (colSplice[i-1].style.border.bottom.style or x.style.border.top.style):
                    colSplice = colSplice[i:]
                    break
            else:
                break # break out of while if can't find any border line in colSplice
        columnText = " ".join(filter(None, [x.value.strip() for x in colSplice]))
        if '(' in cell.style.number_format:
            columnText += ' (excluding foreign domestic helpers)'
        
        # return attribute and data
        yield  (area, tableText, rowText, columnText, cell.value)

def convertToSqlite(dirname='_census_ca', sqlitefile="aggregate.db"):
    'Read all census Excel files under the directory and convert the data into one big table in SQLite'

    # Prepare SQLite file
    sqlitepath = os.path.join(dirname, sqlitefile)
    try:
        os.unlink(sqlitepath) # remove file if exists
    except:
        pass
    with contextlib.closing(sqlite3.connect(sqlitepath)) as conn:
        cur = conn.cursor()
        cur.execute("CREATE TABLE aggregate(" \
                         "area TEXT NOT NULL" \
                       ", category TEXT NOT NULL" \
                       ", statistics TEXT NOT NULL" \
                       ", qualifier TEXT NOT NULL" \
                       ", value NUMERIC NOT NULL"
                       ", PRIMARY KEY(area, category, statistics, qualifier))")
        # Collect data tuples
        for filename in scanExcelFiles(dirname):
            tuplelist = list(readExcelData(filename))
            print "%d entries from %s" % (len(tuplelist), filename)
            cur.executemany("INSERT INTO aggregate(area, category, statistics, qualifier, value) VALUES(?,?,?,?,?)", tuplelist)
        conn.commit()

if __name__ == "__main__":
    convertToSqlite()

# vim:set ts=4 sw=4 sts=4 et:
