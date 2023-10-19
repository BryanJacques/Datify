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

<br>

## filterize()

**Description:**
Adds [filter](https://developers.google.com/apps-script/reference/spreadsheet/filter) to datarange

**Returns:**
[datarange](https://github.com/BryanJacques/Datify/blob/main/Datarange.md)

<br>

## unFilterize()

**Description:**
Removes [filter](https://developers.google.com/apps-script/reference/spreadsheet/filter) from datarange

**Returns:**
[datarange](https://github.com/BryanJacques/Datify/blob/main/Datarange.md)

<br>

## cols()

**Description:**
Returns object with properties corresponding to each column name, and each property has additional methods pertaining to that column.

**Returns:**
```
{
  fieldName :
    {
      index() : returns index of column
      ,vals() : returns array of column vals
      ,distinct() : returns array of distinct col vals
      ,name() : returns col name
      ,mean() : returns average of numerical column
      ,range() : returns range of selected column
      ,headerRange() : returns header range of selected column
    }
  }
```

<br>

## formatCols(formats)

**Description:**
Applies formatting to columns.

**Parameters:**  
*formats* : object with properties specified below. For each property, the value is the array of columns you want the property applied to.  
```javascript
  {
    integer                 : [columns]  // sets [columns] number format to integer w/ comma
    ,numeric2dec            : [columns]  // sets [columns] number format to numeric to 2 decimals
    ,numeric1dec            : [columns]  // sets [columns] number format to numeric to 1 decimals
    ,currency2dec           : [columns]  // sets [columns] number format to currency rounded to two decimals
    ,currencyRound          : [columns]  // sets [columns] number format to currency rounded to nearest whole number
    ,accounting2dec         : [columns]  // sets [columns] number format to accounting 2 decimals
    ,pcntWhole              : [columns]  // sets [columns] number format to whole percent
    ,pcnt1dec               : [columns]  // sets [columns] number format to percent with 1 decimals
    ,pcnt2dec               : [columns]  // sets [columns] number format to percent with 2 decimals
    ,yyyymmdd               : [columns]  // sets [columns] date format to YYYY-MM-DD
    ,mmddyyyy               : [columns]  // sets [columns] date format to MM/DD/YYYY
    ,datetime               : [columns]  // sets [columns] date to M/D/YYYY H:MM:SS
    ,hhmm                   : [columns]  // sets [columns] to HH:MM format
    ,hhmmampm               : [columns]  // sets [columns] to HH:MM am/pm format
    ,bold                   : [columns]  // sets [columns] font weight to bold
    ,italic                 : [columns]  // sets [columns] font style to italic
    ,leftAlign              : [columns]  // sets [columns] horizontal alignment left
    ,centerAlign            : [columns]  // sets [columns] horizontal alignment center
    ,rightAlign             : [columns]  // sets [columns] horizontal alignment right
    ,topAlign               : [columns]  // sets [columns] vertical alignment top
    ,middleAlign            : [columns]  // sets [columns] vertical alignment middle
    ,bottomAlign            : [columns]  // sets [columns] vertical alignment bottom
    ,overflow               : [columns]  // sets [columns] wrap strategy to overflow
    ,clip                   : [columns]  // sets [columns] wrap strategy to clip
    ,wrap                   : [columns]  // sets [columns] wrap strategy to wrap
    ,rightBorderSolid       : [columns]  // sets [columns] right border to SOLID
    ,rightBorderSolidMedium : [columns]  // sets [columns] right border to SOLID_MEDIUM
}
```

**Returns:**
[datarange](https://github.com/BryanJacques/Datify/blob/main/Datarange.md)

<br>

## applyColorGradient(args)

**Description:**
Applies either value or percent based color gradient to columns selected

**Parameters:**  
*args* : object with properties specified below:
```javascript
  {
    cols : array of column names to apply gradient to
    ,gradientType : 'PERCENT'/'VALUE'
    ,minColor : hex value or color name for minimum of gradient
    ,midColor : hex value or color name for middle point of gradient
    ,maxColor : hex value or color name for maximum of gradient
    ,minVal : value for minimum of gradient
    ,midVal : value for middle point of gradient
    ,maxVal : value for maximum of gradient
  }
```

**Returns:**
[datarange](https://github.com/BryanJacques/Datify/blob/main/Datarange.md)

<br>

## setHeaderNotes(notePairs)

**Description:**
Sets notes on headers

**Parameters:**  
*namePairs* : two dimensional array of col names with name to rename as : [['colName1','Note1'],['colName2','Note2']]

**Returns:**
[datarange](https://github.com/BryanJacques/Datify/blob/main/Datarange.md)

<br>

## clear(includeHeader, content, dataValidations, format, note)

**Description:**
Clears datarange of existing data/formatting

**Parameters:**  
*includeHeader* : boolean, clear header data, default false  
*content* : boolean, clear range content, default true  
*dataValidations* : boolean, clear data validations, default true  
*format* : boolean, clear formatting, default true  
*note* : boolean, clear notes, default true  

**Returns:**
[datarange](https://github.com/BryanJacques/Datify/blob/main/Datarange.md)

<br>  

## clearFormatting()

**Description:**
Clears datarange of existing formatting only

**Returns:**
[datarange](https://github.com/BryanJacques/Datify/blob/main/Datarange.md)

<br>  

## hideSheet()

**Description:**
hides sheet datarange is on

**Returns:**
[datarange](https://github.com/BryanJacques/Datify/blob/main/Datarange.md)

<br>  

## showSheet()

**Description:**
shows sheet datarange is on

**Returns:**
[datarange](https://github.com/BryanJacques/Datify/blob/main/Datarange.md)

<br>  

## activateSheet()

**Description:**
shows and activates sheet datarange is on

**Returns:**
[datarange](https://github.com/BryanJacques/Datify/blob/main/Datarange.md)

<br>  

## freezeHeader()

**Description:**
freeze header row

**Returns:**
[datarange](https://github.com/BryanJacques/Datify/blob/main/Datarange.md)

<br>  

## unFreezeHeader()

**Description:**
unfreeze header row

**Returns:**
[datarange](https://github.com/BryanJacques/Datify/blob/main/Datarange.md)

<br>  

## freezeCol(afterCol)

**Description:**
freezes sheet after specified column

**Parameters:**  
*afterCol* : name of column to freeze 

**Returns:**
[datarange](https://github.com/BryanJacques/Datify/blob/main/Datarange.md)

<br> 

## unFreezeCol()

**Description:**
unfreeze sheet columns

**Returns:**
[datarange](https://github.com/BryanJacques/Datify/blob/main/Datarange.md)

<br> 

## resizeDataRows(height)

**Description:**
alters heigh of data rows

**Parameters:**  
*height* : height in pixels to set data rows to 

**Returns:**
[datarange](https://github.com/BryanJacques/Datify/blob/main/Datarange.md)

<br> 

## requireValsInList(col,list,allowInvalid)

**Description:**
sets data validation on specified column to a selection dropdown

**Parameters:**  
*col* : name of column to apply data dropdown 
*list* : array of selection options for dropdown  
*allowInvalid* : boolean, determines if invalid input is allowed or not, default true  

**Returns:**
[datarange](https://github.com/BryanJacques/Datify/blob/main/Datarange.md)

<br> 

## hideCols(cols)

**Description:**
hides specified columns

**Parameters:**  
*cols* : array of column names to hide

**Returns:**
[datarange](https://github.com/BryanJacques/Datify/blob/main/Datarange.md)

<br>

## filter(inFunction, returnRangeList)

**Description:**
filters datarange using inFunction to subselect row ranges following format of array method filter()

```javascript
// example use case

var dr = Datify.datarange('ExampleNamedRange')

dr.filter(row => {return row['header1'] == 'ExampleColVal'}).setBackground("blue")
  
// would find any rows where column header1 = 'ExampleColVal' and would set row background to be blue
```

**Parameters:**  
*inFunction* : function following format of array method filter()  
*returnRangeList* : boolean, if true function returns [RangeList](https://developers.google.com/apps-script/reference/spreadsheet/range-list), if false, returns array of [ranges](https://developers.google.com/apps-script/reference/spreadsheet/range#setNote(String))

**Returns:**
[RangeList](https://developers.google.com/apps-script/reference/spreadsheet/range-list), if false, returns array of [ranges](https://developers.google.com/apps-script/reference/spreadsheet/range#setNote(String)) or array of [ranges](https://developers.google.com/apps-script/reference/spreadsheet/range#setNote(String))

<br>

## autoResizeCols()

**Description:**
auto resizes column widths to fit data


**Returns:**
[datarange](https://github.com/BryanJacques/Datify/blob/main/Datarange.md)

<br>

## colorCols(inputObjArray)

**Description:**
colors columns

**Parameters:**  
*inputObjArray* : array of objects with the following format:

```javascript
[
  {
    color : hex or color name to color cols
    ,cols : array of column names to apply color to
  }
]
```

**Returns:**
[datarange](https://github.com/BryanJacques/Datify/blob/main/Datarange.md)

<br>

## colorColHeaders(inputObjArray)

**Description:**
colors column headers

**Parameters:**  
*inputObjArray* : array of objects with the following format:

```javascript
[
  {
    color : hex or color name to color cols
    ,cols : array of column names to apply color to
  }
]
```

**Returns:**
[datarange](https://github.com/BryanJacques/Datify/blob/main/Datarange.md)

<br>

## alternateColorByGuideCols(color1,color2,guideCols,applyBorder,borderStyle,borderColor)

**Description:**
alternates background colors of entire rows based on grouping selected

**Parameters:**  
*color1* : hex or color name of first and other odd number groups  
*color2* : hex or color name of second and other even number groups  
*guideCols* : array of column names to determine when to alternate colors  
*applyBorder* : boolean, true to place cell border between groups, default true  
*borderStyle* : [BorderStyle](https://developers.google.com/apps-script/reference/spreadsheet/border-style) if applyBorder is true, default SpreadsheetApp.BorderStyle.SOLID  
*borderColor* : hex or color name of border color  

**Returns:**
[datarange](https://github.com/BryanJacques/Datify/blob/main/Datarange.md)

<br>

## applyRowBanding(options)

**Description:**
applies [Banding](https://developers.google.com/apps-script/reference/spreadsheet/banding) to rows to alternate colors

**Parameters:**  
*options* : object with below properties:
```javascript
{
  headerColor     : hex or color name for header
  ,firstRowColor  : hex or color name for odd rows
  ,secondRowColor : hex or color name for even rows
  ,showHeader     : boolean, include banding on header
  ,showFooter     : boolean, show footer
  ,bandingTheme   : default SpreadsheetApp.BandingTheme.LIGHT_GREY
}
```
see [BandingTheme](https://developers.google.com/apps-script/reference/spreadsheet/banding-theme#:~:text=An%20enumeration%20of%20banding%20themes,For%20example%2C%20SpreadsheetApp.) for BandingTheme enum

**Returns:**
[datarange](https://github.com/BryanJacques/Datify/blob/main/Datarange.md)

<br>

## clearRowBanding(options)

**Description:**
clears [Banding](https://developers.google.com/apps-script/reference/spreadsheet/banding) from rows

**Returns:**
[datarange](https://github.com/BryanJacques/Datify/blob/main/Datarange.md)

<br>

## setColNotes(col,notes)

**Description:**
sets notes on entire column

**Parameters:**  
*col* : name of column to apply notes  
*notes* : array of values to set notes as, much be same length as column

**Returns:**
[datarange](https://github.com/BryanJacques/Datify/blob/main/Datarange.md)

<br>

## styleBody(theme,options)

**Description:**
styles body of dataset. Pick from standard theme or supply your own options

**Parameters:**  
*theme* : curated themes, pick from the following  
1. light grey
2. gray contrast
3. dark
4. cornflower
5. green
6. orange
7. white
8. purple
9. solarized light
10. wizpig
11. mario

*options* : instead of using a theme you can supply the following options, examples provided:
```javascript
{
  background        : '#F7141A'
  ,fontFamily       : 'Arial'
  ,fontSize         : 10
  ,textColor        : '#fff7f8'
  ,wrapStrategy     : SpreadsheetApp.WrapStrategy.WRAP
  ,fontWeight       : 'bold'
  ,setBorder        : true
  ,borderTop        : true
  ,borderLeft       : true
  ,borderBottom     : true
  ,borderRight      : true
  ,borderVertical   : true
  ,borderHorizontal : true
  ,borderColor      : '#D77B58'
  ,borderStyle      : SpreadsheetApp.BorderStyle.SOLID
}
```

**Returns:**
[datarange](https://github.com/BryanJacques/Datify/blob/main/Datarange.md)

<br>

## styleHeader(theme,options)

**Description:**
styles headers of dataset. Pick from standard theme or supply your own options

**Parameters:**  
*theme* : curated themes, pick from the following  
1. light grey
2. gray contrast
3. dark
4. cornflower
5. green
6. orange
7. white
8. purple
9. solarized light
10. wizpig
11. mario

*options* : instead of using a theme you can supply the following options, examples provided:
```javascript
{
  background        : '#F7141A'
  ,fontFamily       : 'Arial'
  ,fontSize         : 10
  ,textColor        : '#fff7f8'
  ,wrapStrategy     : SpreadsheetApp.WrapStrategy.WRAP
  ,fontWeight       : 'bold'
  ,setBorder        : true
  ,borderTop        : true
  ,borderLeft       : true
  ,borderBottom     : true
  ,borderRight      : true
  ,borderVertical   : true
  ,borderHorizontal : true
  ,borderColor      : '#D77B58'
  ,borderStyle      : SpreadsheetApp.BorderStyle.SOLID
}
```

**Returns:**
[datarange](https://github.com/BryanJacques/Datify/blob/main/Datarange.md)

<br>

## style(theme,options)

**Description:**
styles headers and body of dataset. Pick from standard theme

**Parameters:**  
*theme* : curated themes, pick from the following  
1. light grey
2. gray contrast
3. dark
4. cornflower
5. green
6. orange
7. white
8. purple
9. solarized light
10. wizpig
11. mario

**Returns:**
[datarange](https://github.com/BryanJacques/Datify/blob/main/Datarange.md)
