#!/usr/bin/env

# Given Excel files in _census_ca/
# Check if all of them are in same format

from openpyxl import load_workbook
import os

def scanExcelFiles(dirname='_census_ca'):
    for filename in os.listdir(dirname):
        filename = os.path.join(dirname, filename)
        if not os.path.isfile(filename) or not filename.endswith('.xlsx'):
            continue
        yield filename

def getExcelSkeleton(filename):
    wb = load_workbook(filename=filename, read_only=True)
    sheets = wb.get_sheet_names()
    if len(sheets) != 3: raise("%s has (!=3) sheets: %s" % (filename, sheets))
    if not sheets[-1].endswith('e'): raise("Last sheet of %s not ends with 'e' [%s]" % (filename, sheets))

    ws = wb.get_sheet_by_name(sheets[-1])
    textcells = {}
    for row in ws.rows:
        for cell in row:
            if cell.row < 3:
                break # get data start from 3rd row
            if not isinstance(cell.value, basestring):
                continue # we don't check numeric or blank cells
            if all(not c.isalpha() for c in cell.value):
                continue # skip if no alphabet in this string, probably blank
            #colCode = chr(ord('A')+cell.column-1)
            #cellCode = '%s%d' % (colCode, cell.row)
            textcells[cell.coordinate] = cell.value.strip()
    return textcells

def checkExcelFiles():
    refSkeleton = None
    refFile = None
    refCode = None
    for filename in scanExcelFiles('_census_ca'):
        print "Checking %s" % filename
        skeleton = getExcelSkeleton(filename)
        if refSkeleton is None:
            refSkeleton = skeleton
            refFile = filename
            refCode = [k for k in refSkeleton if refSkeleton[k]=='Proportion of population of Chinese ethnicity (%)'][0]
        elif cmp(skeleton, refSkeleton):
            # Skeleton does not match! Print the detailed differences
            # Note: B198 is a footnote and it is known to slightly vary.
            unmatched1 = set(k for k,_ in set(skeleton.items()) - set(refSkeleton.items()) if k != 'B198')
            unmatched2 = set(k for k,_ in set(refSkeleton.items()) - set(skeleton.items()) if k != 'B198')
            myCode = [k for k in skeleton if skeleton[k]=='Proportion of population of Chinese ethnicity (%)']
            if myCode:
                # Skip the demographic characteristics section as the items are known to vary
                myCode = myCode[0]
                unmatched1 = set(k for k in unmatched1 if (len(k)==2 and k < 'A7') or (len(k) > 2 and k > myCode))
                unmatched2 = set(k for k in unmatched2 if (len(k)==2 and k < 'A7') or (len(k) > 2 and k > refCode))
                if not unmatched1 and not unmatched2:
                    continue
            print "File %s does not match file %s" % (filename, refFile)
            for k in sorted(unmatched1 | unmatched2):
                if k in skeleton and k in refSkeleton:
                    print "Cell %s: %s vs %s" % (k, repr(skeleton[k]), repr(refSkeleton[k]))
                elif k in skeleton and k not in refSkeleton:
                    print "Cell %s only in %s: %s" % (k, filename, repr(skeleton[k]))
                elif k not in skeleton and k in refSkeleton:
                    print "Cell %s not in %s: %s" % (k, filename, repr(refSkeleton[k]))

if __name__ == "__main__":
    checkExcelFiles()

# vim:set ts=4 sw=4 sts=4 et:
