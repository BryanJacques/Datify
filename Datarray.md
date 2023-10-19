# Datarray
<br>

## getHeader()

** Description:**
Returns array of column headers

#### Returns:
[string]

<br>

## all()

#### Description:
Returns two dimensional array of headers and all rows

#### Returns:
[[header],[row1],[row2]]

<br>

## getData()

#### Description:
Returns two dimensional array of data rows

#### Returns:
[[row1],[row2]]

<br>

## getNumCols()

#### Description:
Returns number of columns

#### Returns:
int

<br>

## getNumRows()

#### Description:
Returns number of rows

#### Returns:
int

<br>

## insertRecords(records, afterDataRow)

#### Description:
Inserts data into datarray

#### Parameters:
records: two dimensional array of row values to insert  
afterDataRow : location to insert data, defaults to end of dataset

#### Returns:
datarray{}

<br>

## print()

#### Description:
Logs header and data to console

<br>

## prettyPrint()

#### Description:
Logs header and data to console in semi formatted way

<br>

## shape()

#### Description:
Returns object listing row and column count

#### Returns:
{rows : int, cols : int}

<br>

## emailCsv(fileName, recipients, subject, body, msgOptions, archiveFolderId)

#### Description:
Emails dataset to user(s)

#### Parameters:
fileName : name of file to email  
recipients : comma separated list of emails to send to  
subject : subject of email  
body : body of email  
msgOptions : additional email options (same options found [here](https://developers.google.com/apps-script/reference/mail/mail-app#sendEmail(String,String,String,Object)))

#### Returns:
datarray{}

<br>

## rows()

#### Description:
returns array of row values where each row is object, and each object has properties corresponding to field names. Useful for working with data without having to get index of headers

#### Returns:
[{rowObject}]

<br>

## cols()

#### Description:
returns object with properties corresponding to each column name, and each property has additional methods pertaining to that column

#### Returns:
```
{
  fieldName :
    {
      index() : returns index of column
      ,vals() : returns array of column vals
      ,distinct() : returns array of distinct col vals
      ,name() : returns col name
      ,mean() : returns average of numerical column
    }
  }
```

<br>

## renameCols(namePairs)

#### Description:
renames column headers

#### Parameters:
namePairs : two dimensional array of col names with name to rename as : [['col1','renamedCol1'],['col2','renamedCol2']]

#### Returns:
datarray{}

<br>

## map(inFunction)

#### Description:
iterates over rows() objects to alter row values and create new columns if they didn't exist already. One of the most useful functions in datarray.  

```javascript
// example use case

var da = Datify.datarray([['Header1','Header2'],['Val1','Val2']])

da = da.map(row => {
    row['Header1'] = row['Header1'] + '--'
    row['Header3'] = 'Val3' 
    return row
  })

da.prettyPrint()

/* OUTPUT
Header1|Header2|Header3
Val1--|Val2|Val3
*/
```

#### Parameters:
inFunction : function following format of array method map() 


#### Returns:
datarray{}

<br>

## filter(inFunction)

#### Description:
filters datarray using inFunction to subselect rows following format of array method filter()

```javascript
// example use case

var da = Datify.datarray([['Header1','Header2'],['Val1','Val2'],['Val3','Val4']])

  da = da.filter(row => {
    return row['Header1'] == 'Val3'
  })
  
  da.prettyPrint()

/* OUTPUT
Header1|Header2
Val3|Val4
*/
```

#### Parameters:
inFunction : function following format of array method map() 

#### Returns:
datarray{}

<br>

## sort(inFunction)

#### Description:
iterates over rows() objects to resort datarray  

#### Parameters:
inFunction : function following format of array method sort() 

#### Returns:
datarray{}

<br>

## select(cols)

#### Description:
selects subset of columns from datarray in order of cols selected

#### Parameters:
cols : array of column names 

#### Returns:
datarray{}

<br>

## except(cols)

#### Description:
selects subset of columns from datarray excluding columns in cols

#### Parameters:
cols : array of column names 

#### Returns:
datarray{}

<br>

## toJson()

#### Description:
returns JSON string value of entire dataset

#### Returns:
string

<br>

## head()

#### Description:
prints first 5 rows of dataset

<br>

## properCaseHeader()

#### Description:
capitalizes header names and replaces underscores with spaces, useful for renaming columns directly from database

<br>

## limit(rowCount)

#### Description:
reduces number of rows in datarray to rowCount

#### Parameters:
rowCount : integer of rows to limit to

#### Returns:
datarray{}

<br>

## place(rangeName,doc)

#### Description:
places datarray object into named range on sheet and returns datarange object to further manipulate once it's on the page

#### Parameters:
rangeName : existing named range to place data to
doc : [spreadsheet](https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet) to place to, defaults to current document

#### Returns:
datarange{}










