# Datarange
datarange function accepts the name of a named range as input and returns an object with methods to work with data placed on a Google sheet.

<br>

## getRange()

**Description:**
Returns range of dataset

**Returns:**
[range](https://developers.google.com/apps-script/reference/spreadsheet/range)

<br>

## getRow()

**Description:**
Returns starting row number of dataset range

**Returns:**
int

<br>

## getCol()

**Description:**
Returns starting column number of dataset range

**Returns:**
int

<br>

## all()

**Description:**
Returns two dimensional array of headers and all rows

**Returns:**
[[header],[row1],[row2]]

<br>

## datarray()

**Description:**
Returns datarray object of dataset

**Returns:**
[datarray](https://github.com/BryanJacques/Datify/blob/main/Datarray.md)

<br>

## print()

**Description:**
Logs header and data to console

<br>

## prettyPrint()

**Description:**
Logs header and data to console in semi formatted way

<br>

## getHeader()

**Description:**
Returns array of column headers

**Returns:**
[string]

<br>

## getData()

**Description:**
Returns two dimensional array of data rows

**Returns:**
[[row1],[row2]]

<br>

## getHeaderRange()

**Description:**
Returns range of column headers

**Returns:**
[range](https://developers.google.com/apps-script/reference/spreadsheet/range)

<br>

## getDataRange()

**Description:**
Returns range of data only

**Returns:**
[range](https://developers.google.com/apps-script/reference/spreadsheet/range)

<br>

## getNumCols()

**Description:**
Returns number of columns

**Returns:**
int

<br>

## getNumRows()

**Description:**
Returns number of rows

**Returns:**
int

<br>

## shape()

**Description:**
Returns object listing row and column count

**Returns:**
{rows : int, cols : int}

<br>

## getSheet()

**Description:**
Returns sheet datarange is on

**Returns:**
[sheet](https://developers.google.com/apps-script/reference/spreadsheet/sheet)

