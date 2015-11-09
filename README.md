The Pearl of the Orient
=======================

This is the repository I created to keep track of the nuts and blots that handles a variety of data for Hong Kong.


Census data
-----------
Census data is downloaded from Census and Statistics Department,
http://www.census2011.gov.hk. There are Excel XML files (`*.xlsx`) at the
level of District Board Constituency Area.

The Excel files are verified to be in the same format for each of the 18
districts, 412 constituency areas. Briefly, the file is in the format of
following:

* The file is in "two-column" format, whereas columns A-E and H-N are two
  separate *tables*
* Data description in text string is in columns A and H
* Numerical values are in columns C,D,E and L,M,N
* Tables are categorised in different sections. Section labels are in column A
  with style matched by regex `\d\. .+`
* Each table is started by a header in green background and bold face
   + Background fill color: `cell.style.fill.fgColor.value`
   + Green colour code in ARGB: FFCCFFCC
   + Boldface: `cell.style.font.bold == True`
   + Header text may span in rows. Rules to check for span
      * Cells vertically adjacent in same green background
      * Cells vertically adjacent are text
      * Lower cell has no top border and upper cell has no bottom border,
        checked by `cell.style.border.top == ''` and `cell.style.border.bottom == ''`
   + Header must begin by the row atop is blank (columns A-F or H-O)
* Table may end with a green footer as summary row. In same format as the
  header
   + Footer must end by the row below is blank (columns A-F or H-O)
* In the middle of the table, rows are in alternate color of no-fill or light
  green
   + Light green colour code in ARGB: FFF0FFF0
* Numbers are mostly integers
   + Number format defined iin `cell.style.number_format`
   + Usual format string: `'#\\ ###\\ ###;\\-#\\ ###\\ ###;\\-'`
   + Format for income number at summary row: `'#,###,###;\\-#,###,###;\\-'`
   + Number for income table excluding foreign domestic helpers: `'\\(#,###,###\\);\\(\\-#,###,###\\);\\(\\-\\)'`
   + Format for percentage data (floating point with 1 d.p.): `'0.0;\\-0.0;\\-'`
* Mostly, each number can be associated with the description on left (column A
  or H of same row) and top (nearest green-backgrounded cell atop in the same column)
   + Exception for incomes table
   + Numbers can be bracketed to mean figures excluding foreign domestic helpers
   + Labels in column A/H need to check for upper and lower adjacent cells are
     in different colour. Otherwise are continuation of text spanning across
     rows.
