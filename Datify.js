/*
------------------------------------------------------------------------------------
-- Datarray
------------------------------------------------------------------------------------

datarray function accepts two dimensional array in following format:
  [
    ['col 1 header','col 2 header','col 3 header]
    ,['row 1 col 1','row 1 col 2','row 1 col 3']
    ,['row 2 col 1','row 2 col 2','row 2 col 3']
  ]

It returns an object with methods to work with tabular data.


datarray.getHeader()

  Description:
    Returns list of column headers

  Returns:
    [string]


possible additions
  .insertToBq // maybe truncate, specify fields
  .longer     // pivot data
  .simpleSort // something to sort data without having to pass function
  .distinct   // pass multiple cols, get distinct vals
*/


function datarray(vals){
  
  if(!vals){
    Logger.log("No input data found for datarray")
    return null
  }

  var returnObj = {}
  
  returnObj.header = vals[0];
  returnObj.data = vals.slice(1);

  returnObj.getHeader = function() {
    return returnObj.header.slice(0);
  }

  returnObj.all = function() {
    return [returnObj.header].concat(returnObj.data);
  }

  returnObj.getData = function() {
    return returnObj.data;
  }

  returnObj.getNumCols = function() {
    return returnObj.header.length
  }

  returnObj.getNumRows = function() {
    return returnObj.data.length;
  }

  returnObj.insertRecords = function(records,afterDataRow = returnObj.getNumRows()){
    var header = returnObj.getHeader()
    var data = returnObj.getData()
    var beforeData = data.slice(0,afterDataRow)
    var afterData = data.slice(afterDataRow,returnObj.getNumRows())
    var newData = beforeData.concat(records).concat(afterData)

    return datarray([header].concat(newData));
  }

  returnObj.toString = function(){
    return returnObj.all().map(
      row => row.map(
        cell => String(cell ?? '')
          .trim()
      ).join('|')
    ).join('\r\n');
  }

  returnObj.print = function(){
    Logger.log(returnObj.all())
  }

  returnObj.prettyPrint = function(){
    Logger.log("\n" + returnObj.toString())
  }

  returnObj.shape = function(){
    return {
      rows  : returnObj.getNumRows()
      ,cols : returnObj.getNumCols()
    }
  }

  returnObj.emailCsv = function(
    fileName = "datarrayCsv"
    ,recipients = Session.getActiveUser().getEmail()
    ,subject = "datarrayCsv"
    ,body = ""
    ,msgOptions=null
    ,archiveFolderId=null
    ){
    var file_data = returnObj.all();
    var file_data_joined = file_data.join('\n');

    var file = DriveApp.createFile(fileName,file_data_joined,MimeType.CSV);
    var file_mime_csv = file.getAs(MimeType.CSV);    

    if(msgOptions==null){var msgOptions = {}}
    if("attachments" in msgOptions){msgOptions.attachments.push(file_mime_csv)}
    else {msgOptions.attachments = [file_mime_csv]}
    
    MailApp.sendEmail(recipients,subject,body,msgOptions);

    if(archiveFolderId){file.moveTo(DriveApp.getFolderById(archiveFolderId))}
    else {file.setTrashed(true)}

    return returnObj;
  }

  returnObj.rows = function(){

    var header = returnObj.getHeader()

    var rowObjects = [];
    var rowNum = 1
    var data = returnObj.getData()

    for (row of data){
      var rowObj = {}
      for (i = 0; i < row.length; i++){
        rowObj[header[i]] = row[i]
      }
      // store general properties in _properties object
      rowObj._properties = {}
      rowObj._properties.rowNum = rowNum
      rowObj._properties.vals = row
      rowObjects.push(rowObj)
      rowNum ++
    }
    return rowObjects
  }

  returnObj.createColObjects = function() {

    var dfu = datifyUtils()
    var header = returnObj.getHeader()
    var data = returnObj.getData()
    var transposedData = dfu.transpose(data)

    var colObjects = {}
    for (var i = 0; i < header.length; i ++){

      var colObj = {}

      colObj.index_ = i

      colObj.index = function(){
        return this.index_
      }

      colObj.vals = function(){
        return transposedData[this.index_]
      }

      colObj.distinct = function(){
        return [...new Set(this.vals())]
      }

      colObj.name = function(){
        return header[this.index_]
      }

      colObj.mean = function(){
        var colValsToMean = this.vals()
        var counter = 0
        var cmltvSum = 0
        var allNumeric = true
        for (val of colValsToMean){
          numVal = + val
          if(isNaN(numVal)){
            allNumeric = false
            continue
          }
          counter ++
          cmltvSum += numVal
        }
        if(!allNumeric){Logger.log("warning: non numeric values encountered and ignored")}
        return cmltvSum/counter
      }

      colObjects[header[i]] = colObj
    }
    returnObj.columnObjects = colObjects
  }

  returnObj.cols = function(){
    return returnObj.columnObjects
  }

  returnObj.filter = function(inFunction){
    
    var rows = returnObj.rows()

    if(typeof(inFunction) == 'function'){
      var filteredData = rows.filter(inFunction)
    } else {
      throw new Error("Invalid input for datarray.filterRows()") 
    }
    var records = filteredData.map(v => v._properties.vals)
    return datarray([returnObj.header].concat(records));
  }

  returnObj.renameCols = function(namePairs){
    var header = returnObj.getHeader()
    for (namePair of namePairs){
      var before = namePair[0]
      var after = namePair[1]
      var beforeIndex = header.indexOf(before)
      if(beforeIndex >= 0){header[beforeIndex] = after}
    }
    return datarray([header].concat(returnObj.getData()))
  }

  returnObj.map = function(inFunction){

    var header = returnObj.getHeader()
    var rows = returnObj.rows()
    var newRows = rows.map(inFunction)
    
    // map function can create new fields. get first row of newrows and check for new fields. if so, add to header
    var newFields = Object.keys(newRows[0])
        .filter(v => v != '_properties' && header.indexOf(v) == -1)

    newFields.forEach(v => {header.push(v)})

    var newRecords = []

    for (newRow of newRows){
      var newRecord = []
      for (key of Object.keys(newRow)){
        var keyIndex = header.indexOf(key)
        if(keyIndex == -1){continue}
        newRecord[keyIndex] = newRow[key]
      }
      newRecords.push(newRecord)
    }

    return datarray([header].concat(newRecords))
  }

  returnObj.sort = function(inFunction){

    var header = returnObj.getHeader()
    var rows = returnObj.rows()
    var newRows = rows.sort(inFunction)
    
    var newRecords = []

    for (newRow of newRows){
      var newRecord = []
      for (key of Object.keys(newRow)){
        var keyIndex = header.indexOf(key)
        if(keyIndex == -1){continue}
        newRecord[keyIndex] = newRow[key]
      }
      newRecords.push(newRecord)
    }

    return datarray([header].concat(newRecords))
  }

  returnObj.select = function(cols){

    var header = returnObj.getHeader()
    var selectIndexes = []
    var newHeader = []

    cols.forEach((col,ix)=>{
      let colIndex = header.indexOf(col)
      if(colIndex == -1){
        Logger.log("datarange.select() : Column %s not found",col)
      } else {
        selectIndexes.push(colIndex)
        newHeader.push(header[colIndex])
      }
    })

    var data = returnObj.getData()
    var newData = []

    for (line of data){
      var newLine = []
      selectIndexes.forEach(v=>newLine.push(line[v]))
      newData.push(newLine)
    }

    var array2d = [newHeader].concat(newData)

    return datarray(array2d)
  }

  returnObj.except = function(exceptFields){
    var selectList = returnObj.getHeader()
    for (field of exceptFields){
      var index = selectList.indexOf(field) 
      if(index != null && index >= 0){selectList.splice(index,1)}
    }
    return returnObj.select(selectList)
  }

  returnObj.toJson = function(space = null){
    return JSON.stringify(returnObj.all(),null,space)
  }

  returnObj.head = function(rowCount = 5){
    returnObj.filter((v,i)=> i <= (rowCount - 1)).prettyPrint()
  }

  returnObj.properCaseHeader = function(){
    var dfu = datifyUtils()
    var header = returnObj.getHeader()
    var capHeader = header.map(colName => {
      return dfu.capitalizePhrase(colName.replace(/_/g,' '))
    })
    return datarray([capHeader].concat(returnObj.getData()))
  }

  returnObj.limit = function(rowCount){
    var rows = returnObj.rows().slice(0,rowCount)
    var records = rows.map(value => value._properties.vals)
    return datarray([returnObj.header].concat(records));
  }

  returnObj.place = function(rangeName,doc = null){

    if(!doc){doc = SpreadsheetApp.getActive()}

    var existingRange   = doc.getRangeByName(rangeName)
    if(!existingRange){
      Logger.log("Named Range %s not found, exit",rangeName)
      return
    }

    var data            = returnObj.all()
    var startCell       = existingRange.getCell(1,1)
    var sheet           = startCell.getSheet()
    var startRow        = startCell.getRow()
    var startColumn     = startCell.getColumn()
    var sheetMaxRows    = sheet.getMaxRows()
    var sheetMaxColumns = sheet.getMaxColumns()
    var dataRowCount    = data.length
    var dataColCount    = data[0].length

    // Make sure there enough rows to hold the results
    if (dataRowCount + startRow > sheetMaxRows) {
        sheet.insertRowsAfter(sheetMaxRows, (dataRowCount + startRow - sheetMaxRows ));
        sheetMaxRows = sheet.getMaxRows();    
      }

    // Make sure there enough columns to hold the results
    if (dataColCount + startColumn > sheetMaxColumns) {
        sheet.insertColumnsAfter(sheetMaxColumns, (dataColCount + startColumn - sheetMaxColumns ));
        sheetMaxColumns = sheet.getMaxColumns();
    }

    var destinationRange = sheet.getRange(startRow, startColumn, dataRowCount, dataColCount);

    // having filtered data when trying to import new data causes formatting issues, remove filter first
    var filter = sheet.getFilter();
    if(filter){filter.remove();}


    existingRange.clearContent()
      .clearDataValidations()
      .clearFormat()
      .clearNote()
      .clear({commentsOnly : true})

    destinationRange.clearContent()
      .clearDataValidations()
      .clearFormat()
      .clearNote()

    destinationRange.setValues(data)
    
    doc.setNamedRange(rangeName,destinationRange)

    return datarange(rangeName,doc)
  }

  returnObj.createColObjects()

  return returnObj
}



/*
------------------------------------------------------------------------------------
-- Datarange
------------------------------------------------------------------------------------

consider adding:
  insertRows            // some type of function to insert data after position
  trimSpaceAroundRange  // delete blank rows and cols around dataset
  colWidthsArray        // a way to programatically alter col widths?
  groupCols             //  
  insertCheckboxes      // pass [cols] argument 
  setHeaderNotes        // [[col,note]]
  setHiddenFilterVals   // for filterized dr
  appendCol             // 
  warn
  removeWarnings
  resizeToRows          // resizes range to accomodate data entered below it
  getFilterCriteria
  setFilterCriteria
  
  // add banding to themes?
*/

function datarange(rangeName,doc = null){
 
  var returnObj = {}

  returnObj.rangeName = rangeName

  returnObj.doc = doc ?? SpreadsheetApp.getActive()
  
  returnObj.getRange = function(){
    if(!returnObj.range){
      returnObj.range = returnObj.doc.getRangeByName(returnObj.rangeName)
    }
    return returnObj.range
  }

  returnObj.getRange()

  if(!returnObj.range){
    throw new Error("Range " + rangeName + " not found")
  }

  returnObj.getRow = function(){
    return returnObj.getRange().getRow()
  }

  returnObj.getCol = function(){
    return returnObj.getRange().getColumn()
  }

  returnObj.all = function(){
    return returnObj.getRange().getValues()
  }

  returnObj.datarray = function(){
    return datarray(returnObj.all())
  }

  returnObj.print = function(){
    returnObj.datarray().print()
  }

  returnObj.prettyPrint = function(){
    returnObj.datarray().prettyPrint()
  }

  returnObj.getHeader = function() {
    return returnObj.datarray().getHeader();
  }

  returnObj.getData = function() {
    return returnObj.datarray().getData();
  }

  returnObj.getHeaderRange = function(){
    var range = returnObj.getRange()
    var headerRange = range.offset(0,0,1)
    return headerRange
  }

  returnObj.getDataRange = function(){
    var range = returnObj.getRange()
    var datarange = range.offset(1,0,range.getNumRows()-1)
    return datarange
  }

  returnObj.getNumRows = function(){
    return returnObj.datarray().getNumRows()
  }

  returnObj.getNumCols = function(){
    return returnObj.datarray().getNumCols()
  }

  returnObj.shape = function(){
    return returnObj.datarray().shape()
  }

  returnObj.getSheet = function() {
    var sheet = returnObj.sheet
    if(!sheet){
      var range = returnObj.getRange();
      sheet = range.getSheet();
      returnObj.sheet = sheet
    }
    return sheet;
  }

  returnObj.filterize = function() {
    var range = returnObj.getRange();
    var sheet = returnObj.getSheet();
    var filter = sheet.getFilter();
    if(filter){filter.remove()};
    range.createFilter();
    return returnObj;
  }

  returnObj.unFilterize = function(){
    var sheet = returnObj.getSheet()
    var filter = sheet.getFilter()
    if(filter){filter.remove()}
    return returnObj
  }

  returnObj.createColObjects = function() {
    var header = returnObj.getHeader()
    var datarray = returnObj.datarray()
    var range = returnObj.getRange()
    //Logger.log("datarray.cols() = %s",datarray.cols())

    var columnObjects = datarray.cols()

    //for (ix = 0; ix < header.length; ix ++){
    for (col of header){
      
      columnObjects[col].range = function(){
        return range.offset(1,this.index(),range.getNumRows()-1,1)
      }
      columnObjects[col].headerRange = function(){
        return range.offset(0,this.index(),1,1)
      }
    }
    
    returnObj.columnObjects = columnObjects
  }

  returnObj.cols = function(){
    return returnObj.columnObjects
  }

  returnObj.formatCols = function(formats = {
    /*
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
    */
  }){
    var sheet = returnObj.getSheet()
    var dfu = datifyUtils()
    var rangeList = undefined

    
    function getRangeListFromFormatCols(cols){
      if(!cols){return null}
      var ranges = cols.map(col=>{
        var colRange = returnObj.cols()[col].range()
        return colRange
      })
      var rangeList = dfu.getRangeListFromRanges(ranges,sheet)
      return rangeList
    }
    
    rangeList = getRangeListFromFormatCols(formats.integer)
    if(rangeList){rangeList.setNumberFormat('#,##0')}

    rangeList = getRangeListFromFormatCols(formats.numeric2dec)
    if(rangeList){rangeList.setNumberFormat('#,##0.00')}

    rangeList = getRangeListFromFormatCols(formats.numeric1dec)
    if(rangeList){rangeList.setNumberFormat('#,##0.0')}

    rangeList = getRangeListFromFormatCols(formats.currency2dec)
    if(rangeList){rangeList.setNumberFormat('"$"#,##0.00')}

    rangeList = getRangeListFromFormatCols(formats.currencyRound)
    if(rangeList){rangeList.setNumberFormat('"$"#,##0')}

    rangeList = getRangeListFromFormatCols(formats.accounting2dec)
    if(rangeList){rangeList.setNumberFormat('_("$"* #,##0.00_);_("$"* \(#,##0.00\);_("$"* "-"??_);_(@_)')}

    rangeList = getRangeListFromFormatCols(formats.pcntWhole)
    if(rangeList){rangeList.setNumberFormat('0%')}

    rangeList = getRangeListFromFormatCols(formats.pcnt1dec)
    if(rangeList){rangeList.setNumberFormat('0.0%')}

    rangeList = getRangeListFromFormatCols(formats.pcnt2dec)
    if(rangeList){rangeList.setNumberFormat('0.00%')}

    rangeList = getRangeListFromFormatCols(formats.yyyymmdd)
    if(rangeList){rangeList.setNumberFormat('yyyy-mm-dd')}

    rangeList = getRangeListFromFormatCols(formats.mmddyyyy)
    if(rangeList){rangeList.setNumberFormat('mm/dd/yyyy')}

    rangeList = getRangeListFromFormatCols(formats.datetime)
    if(rangeList){rangeList.setNumberFormat('mm/dd/yyyy H:mm:ss')}

    rangeList = getRangeListFromFormatCols(formats.hhmm)
    if(rangeList){rangeList.setNumberFormat('hh":"mm')}

    rangeList = getRangeListFromFormatCols(formats.hhmmampm)
    if(rangeList){rangeList.setNumberFormat('h:mm am/pm')}

    rangeList = getRangeListFromFormatCols(formats.bold)
    if(rangeList){rangeList.setFontWeight('bold')}

    rangeList = getRangeListFromFormatCols(formats.italic)
    if(rangeList){rangeList.setFontStyle('italic')}

    rangeList = getRangeListFromFormatCols(formats.leftAlign)
    if(rangeList){rangeList.setHorizontalAlignment('left')}

    rangeList = getRangeListFromFormatCols(formats.centerAlign)
    if(rangeList){rangeList.setHorizontalAlignment('center')}

    rangeList = getRangeListFromFormatCols(formats.rightAlign)
    if(rangeList){rangeList.setHorizontalAlignment('right')}

    rangeList = getRangeListFromFormatCols(formats.topAlign)
    if(rangeList){rangeList.setVerticalAlignment('top')}

    rangeList = getRangeListFromFormatCols(formats.middleAlign)
    if(rangeList){rangeList.setVerticalAlignment('middle')}

    rangeList = getRangeListFromFormatCols(formats.bottomAlign)
    if(rangeList){rangeList.setVerticalAlignment('bottom')}

    rangeList = getRangeListFromFormatCols(formats.overflow)
    if(rangeList){rangeList.setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW)}

    rangeList = getRangeListFromFormatCols(formats.clip)
    if(rangeList){rangeList.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)}

    rangeList = getRangeListFromFormatCols(formats.wrap)
    if(rangeList){rangeList.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)}

    rangeList = getRangeListFromFormatCols(formats.rightBorderSolid)
    if(rangeList){
      rangeList.setBorder(null, null, null, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID)
    }

    rangeList = getRangeListFromFormatCols(formats.rightBorderSolidMedium)
    if(rangeList){
      rangeList.setBorder(null, null, null, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    }

    return returnObj
  }

  returnObj.applyColorGradient = function(args){
    
    var colRanges = args['cols'].map(col => returnObj.cols()[col].range())

    var ruleBuilder = SpreadsheetApp.newConditionalFormatRule()
      .setRanges(colRanges)

    if(args.gradientType.toUpperCase() == 'VALUE'){
    ruleBuilder
      .setGradientMinpointWithValue(args.minColor,SpreadsheetApp.InterpolationType.NUMBER,args.minVal)
      .setGradientMidpointWithValue(args.midColor,SpreadsheetApp.InterpolationType.NUMBER,args.midVal)
      .setGradientMaxpointWithValue(args.maxColor,SpreadsheetApp.InterpolationType.NUMBER,args.maxVal)
    }

    if(args.gradientType.toUpperCase() == 'PERCENT'){
    ruleBuilder
      .setGradientMinpointWithValue(args.minColor,SpreadsheetApp.InterpolationType.PERCENT,args.minVal)
      .setGradientMidpointWithValue(args.midColor,SpreadsheetApp.InterpolationType.PERCENT,args.midVal)
      .setGradientMaxpointWithValue(args.maxColor,SpreadsheetApp.InterpolationType.PERCENT,args.maxVal)
    }

    var rule = ruleBuilder.build()
    var sheet = returnObj.getSheet()
    var rules = sheet.getConditionalFormatRules()

    rules.push(rule)
    returnObj.sheet.setConditionalFormatRules(rules)

    return returnObj
  }

  returnObj.setHeaderNotes = function(notePairs){

    var headerRange = returnObj.getHeaderRange()
    var headerVals = headerRange.getValues()[0]
    var headerNotes = new Array

    headerNotes.length = headerVals.length

    for (pair of notePairs){

      var colName = pair[0]
      var headerIndex = headerVals.indexOf(colName)

      if(headerIndex == -1){
        Logger.log("Col %s not found in header",colName)
        continue
        } 

      var note = pair[1]
      headerNotes[headerIndex] = note
    }

    headerRange.setNotes([headerNotes])
    return returnObj
  }

  returnObj.clear = function(includeHeader = false, content = true, dataValidations = true, format = true, note = true){
    if(includeHeader){var range = returnObj.getRange()}
    else {var range = returnObj.getDataRange()}
    if(content && range){range.clearContent()}
    if(dataValidations && range){range.clearDataValidations()}
    if(format && range){range.clearFormat()}
    if(note && range){range.clearNote()}
    return returnObj
  }

  returnObj.clearFormatting = function(){
    returnObj.clear(false,false,false,true,false)
  }

  returnObj.hideSheet = function() {
    sheet = returnObj.getSheet();
    sheet.hideSheet();
    return returnObj;
  }

  returnObj.showSheet = function() {
    sheet = returnObj.getSheet();
    sheet.showSheet();
    return returnObj;
  }

  returnObj.activateSheet = function() {
    var sheet = returnObj.getSheet();
    sheet.activate();
    return returnObj;
  }

  returnObj.freezeHeader = function(){
    var sheet = returnObj.getSheet();
    var row_num = returnObj.getRow();
    sheet.setFrozenRows(row_num);
    return returnObj;
  }

  returnObj.unFreezeHeader = function(){
    var sheet = returnObj.getSheet();
    sheet.setFrozenRows(0);
    return returnObj;
  }

  returnObj.freezeCol = function(afterCol){
    var sheet = returnObj.getSheet();
    var colNum = returnObj.cols()[afterCol].range().getColumn()
    sheet.setFrozenColumns(colNum);
    return returnObj;
  }

  returnObj.unFreezeCol = function(){
    var sheet = returnObj.getSheet();
    sheet.setFrozenColumns(0);
    return returnObj;
  }

  returnObj.resizeDataRows = function(height){
    var sheet = returnObj.getSheet();
    sheet.setRowHeightsForced(
      returnObj.getRow() + 1
      ,returnObj.getNumRows()
      ,height  
      );
    return returnObj;
  }

  returnObj.requireValsInList = function(col,list,allowInvalid = true){
    var colObj = returnObj.cols()[col]

    if(!colObj){
      Logger.log("Could not find column '%s' in requireValsInList",col)
      return returnObj
    }

    var colRange = colObj.range()
    var validationRule = SpreadsheetApp.newDataValidation().requireValueInList(list).setAllowInvalid(allowInvalid).build()
    colRange.setDataValidation(validationRule)

    return returnObj
  }

  returnObj.hideCols = function(cols){
    var startCol = returnObj.getCol()
    var header = returnObj.getHeader()
    var sheet = returnObj.getSheet()

    for (col of cols){
      var colIndex = header.indexOf(col)
      if (colIndex == -1){
        Logger.log("Could not find column '%s' to hide",col)
        continue
      }
      var sheetColToHide = startCol + colIndex
      sheet.hideColumns(sheetColToHide)
    }

    return returnObj
  }

  returnObj.unHideCols = function(cols){
    var sheet = returnObj.getSheet()

    for (col of cols){
      var colObj = returnObj.cols()[col]
      if (!colObj){
        Logger.log("Could not find column '%s' to unhide",col)
        continue
      }
      sheet.unhideColumn(colObj.range())
    }

    return returnObj
  }

  returnObj.rows = function(){
    var da = returnObj.datarray()
    var rows = da.rows()
    var headerRange = returnObj.getHeaderRange()

    for (row of rows){
      row.getRange = function(){
        return headerRange.offset(this['_properties']['rowNum'],0)
      }
    }
    return rows
  }

  returnObj.filter = function(inFunction, returnRangeList = true){
    var sheet = returnObj.getSheet()
    var rows = returnObj.rows()
    var filteredRows = rows.filter(inFunction)
    var rangeArray = filteredRows.map(row => row.getRange())
    var returnVal = undefined

    if(!returnRangeList){
      returnVal = rangeArray
    } else {
      var rangeA1s = rangeArray.map(v => v.getA1Notation())
      if(rangeA1s.length == 0){return null}
      var returnVal = sheet.getRangeList(rangeA1s)
    }

    if(returnVal){return returnVal}
    else {return null}
  }

  returnObj.autoResizeCols = function(){

    var range = returnObj.getRange()
    var sheet = range.getSheet()
    var startCol = returnObj.getCol()
    var numCols = returnObj.getNumCols()
    sheet.autoResizeColumns(startCol,numCols)

    return returnObj
  }

  returnObj.colorCols = function(inputObjArray){
    var sheet = returnObj.getSheet()
    var dfu = datifyUtils()

    for (obj of inputObjArray){
      Logger.log(obj)
      var color = obj.color
      var rangeArray = obj.cols.filter(col => returnObj.cols()[col])
        .map(col => returnObj.cols()[col].range())

      if(rangeArray.length == 0){
        var errorMsg = "Error processing coloring object " + JSON.stringify(obj)
        throw new Error(errorMsg)
      }
      var rangeList = dfu.getRangeListFromRanges(rangeArray,sheet)
      rangeList.setBackground(color)
    }

    return returnObj
  }

  returnObj.colorColHeaders = function(inputObjArray){
    var sheet = returnObj.getSheet()
    var dfu = datifyUtils()

    for (obj of inputObjArray){
      Logger.log(obj)
      var color = obj.color
      var rangeArray = obj.cols.filter(col => returnObj.cols()[col])
        .map(col => returnObj.cols()[col].headerRange())
        
      if(rangeArray.length == 0){
        var errorMsg = "Error processing coloring object " + JSON.stringify(obj)
        throw new Error(errorMsg)
      }
      var rangeList = dfu.getRangeListFromRanges(rangeArray,sheet)
      rangeList.setBackground(color)
    }

    return returnObj
  }

  returnObj.alternateColorByGuideCols = function(
    color1
    ,color2
    ,guideCols
    ,applyBorder      = true
    ,borderStyle      = SpreadsheetApp.BorderStyle.DASHED
    ,borderColor      = "black"
    ){

    var datasetRange = returnObj.getDataRange();

    datasetRange.setBackground(color1)

    // set checkGroup = to first row group cols
    var checkGroup = JSON.stringify(guideCols.map(col => returnObj.rows()[0][col]))
    var row2switch = false

    var color2ranges = returnObj.filter(row=>{

      var currentGroup = JSON.stringify(guideCols.map(col => row[col]))

      if(currentGroup != checkGroup && !row2switch){
        row2switch = !row2switch
        checkGroup = currentGroup
        return true
      } else if(currentGroup == checkGroup && row2switch){
        return true
      } else if(currentGroup != checkGroup && row2switch){
        checkGroup = currentGroup
        row2switch = !row2switch
        return false
      } else if(currentGroup == checkGroup && !row2switch){
        return false
      }
    },false) // second argument false to get array of ranges instead of rangeList

    // color background 2
    var dfu = datifyUtils()
    var sheet = returnObj.getSheet()
    var color2rangeList = dfu.getRangeListFromRanges(color2ranges,sheet)
    color2rangeList.setBackground(color2)

    // apply border if option selected
    if(applyBorder == true){
      var topBorderRows = []
      var bottomBorderRows = []
      var checkRow = -1
      var color2RangesCount = color2ranges.length
      var finalTableRow = returnObj.getRange().getLastRow()

      for (rangeIx = 0; rangeIx < color2RangesCount; rangeIx ++){

        // get top border rows
        var currentRange = color2ranges[rangeIx]
        var currentRow = currentRange.getRow()
        if(currentRow > checkRow + 1){topBorderRows.push(currentRange)}

        // get bottom border rows
        var nextRange = color2ranges[rangeIx + 1]
        if(nextRange){
          var nextRow = nextRange.getRow()
        } else {
          var nextRow = null
        }

        if(nextRow > currentRow + 1){
          bottomBorderRows.push(currentRange)
        } else if (!nextRow && currentRow != finalTableRow){
          bottomBorderRows.push(currentRange)
        }

        checkRow = currentRow
      }

      var topBorderRangeList = dfu.getRangeListFromRanges(topBorderRows,sheet)
      topBorderRangeList.setBorder(true, null, null, null, null, null, borderColor, borderStyle)

      var bottomBorderRangeList = dfu.getRangeListFromRanges(bottomBorderRows,sheet)
      bottomBorderRangeList.setBorder(null, null, true, null, null, null, borderColor, borderStyle)
    }

    return returnObj
  }

  returnObj.sort = function(sortValPairs){

    var datarange = returnObj.getDataRange();
    var startCol = returnObj.getCol();
    var sortObjects = [];

    for (const i of sortValPairs){
      var column = i[0];
      var ascending = i[1];

      var sortObj = {};
      var colObj = returnObj.cols()[column]

      if(!colObj){
        Logger.log("column %s not found",column)
        continue
      }

      sortObj.column    = colObj.index() + startCol;
      sortObj.ascending = ascending;

      sortObjects.push(sortObj);
    }

    datarange.sort(sortObjects);
    return returnObj;
  }

  returnObj.applyRowBanding = function(options = {}){

    var defaultOptions = {
      headerColor     : null
      ,firstRowColor  : null
      ,secondRowColor : null
      ,showHeader     : true
      ,showFooter     : false
      ,bandingTheme   : SpreadsheetApp.BandingTheme.LIGHT_GREY
    }

    var applyOptions = datifyUtils().dictComplete(options,defaultOptions)

    var range = returnObj.getRange()
    var bandings = range.getBandings() 
    if(bandings){bandings.forEach(banding => banding.remove())}

    var banding = range.applyRowBanding(applyOptions.bandingTheme,applyOptions.showHeader,applyOptions.showFooter)
    if(applyOptions.headerColor){banding.setHeaderRowColor(applyOptions.headerColor)}
    if(applyOptions.firstRowColor){banding.setFirstRowColor(applyOptions.firstRowColor)}
    if(applyOptions.secondRowColor){banding.setSecondRowColor(applyOptions.secondRowColor)}
    return returnObj
  }

  returnObj.clearRowBanding = function(){
    var range = returnObj.getRange()
    var bandings = range.getBandings() 
    if(bandings){bandings.forEach(banding => banding.remove())}
    return returnObj
  }

  returnObj.setColNotes = function(col,notes){
    var colObj = returnObj.cols()[col]
    if(!colObj){
      Logger.log ("column %s not found",col)
    }
    var colRange = colObj.range()
    var notes2d = notes.map(note => [note])
    colRange.setNotes(notes2d)
    return returnObj
  }

  returnObj.styleBody = function(theme = null,options_ = {}){

    if(theme == 'light gray'){
      var defaultOptions = {
        background        : '#f8f8f8'
        ,fontFamily       : 'Arial'
        ,fontSize         : 10
        ,textColor        : '#666666'
        ,wrapStrategy     : SpreadsheetApp.WrapStrategy.CLIP
        ,fontWeight       : 'normal'
        ,setBorder        : true
        ,borderTop        : true
        ,borderLeft       : true
        ,borderBottom     : true
        ,borderRight      : true
        ,borderVertical   : true
        ,borderHorizontal : true
        ,borderColor      : '#b7b7b7'
        ,borderStyle      : SpreadsheetApp.BorderStyle.DOTTED
      }
    } else if (theme == 'gray contrast'){
      var defaultOptions = {
        background        : '#f8f8f8'
        ,fontFamily       : 'Arial'
        ,fontSize         : 10
        ,textColor        : '#666666'
        ,wrapStrategy     : SpreadsheetApp.WrapStrategy.CLIP
        ,fontWeight       : 'normal'
        ,setBorder        : true
        ,borderTop        : true
        ,borderLeft       : true
        ,borderBottom     : true
        ,borderRight      : true
        ,borderVertical   : true
        ,borderHorizontal : true
        ,borderColor      : '#b7b7b7'
        ,borderStyle      : SpreadsheetApp.BorderStyle.DOTTED
      }
    } else if (theme == 'dark'){
      var defaultOptions = {
        background        : '#999999'
        ,fontFamily       : 'Arial'
        ,fontSize         : 10
        ,textColor        : '#efefef'
        ,wrapStrategy     : SpreadsheetApp.WrapStrategy.CLIP
        ,fontWeight       : 'normal'
        ,setBorder        : true
        ,borderTop        : true
        ,borderLeft       : true
        ,borderBottom     : true
        ,borderRight      : true
        ,borderVertical   : true
        ,borderHorizontal : true
        ,borderColor      : '#c8c8c8'
        ,borderStyle      : SpreadsheetApp.BorderStyle.DOTTED
      }
    } else if (theme == 'cornflower'){
      var defaultOptions = {
        background        : '#c9daf8'
        ,fontFamily       : 'Arial'
        ,fontSize         : 10
        ,textColor        : '#666666'
        ,wrapStrategy     : SpreadsheetApp.WrapStrategy.CLIP
        ,fontWeight       : 'normal'
        ,setBorder        : true
        ,borderTop        : true
        ,borderLeft       : true
        ,borderBottom     : true
        ,borderRight      : true
        ,borderVertical   : true
        ,borderHorizontal : true
        ,borderColor      : '#b7b7b7'
        ,borderStyle      : SpreadsheetApp.BorderStyle.DOTTED
      }
    } else if (theme == 'green'){
      var defaultOptions = {
        background        : '#d9ead3'
        ,fontFamily       : 'Arial'
        ,fontSize         : 10
        ,textColor        : '#666666'
        ,wrapStrategy     : SpreadsheetApp.WrapStrategy.CLIP
        ,fontWeight       : 'normal'
        ,setBorder        : true
        ,borderTop        : true
        ,borderLeft       : true
        ,borderBottom     : true
        ,borderRight      : true
        ,borderVertical   : true
        ,borderHorizontal : true
        ,borderColor      : '#b7b7b7'
        ,borderStyle      : SpreadsheetApp.BorderStyle.DOTTED
      }
    } else if (theme == 'orange'){
      var defaultOptions = {
        background        : '#fce5cd'
        ,fontFamily       : 'Arial'
        ,fontSize         : 10
        ,textColor        : '#666666'
        ,wrapStrategy     : SpreadsheetApp.WrapStrategy.CLIP
        ,fontWeight       : 'normal'
        ,setBorder        : true
        ,borderTop        : true
        ,borderLeft       : true
        ,borderBottom     : true
        ,borderRight      : true
        ,borderVertical   : true
        ,borderHorizontal : true
        ,borderColor      : '#b7b7b7'
        ,borderStyle      : SpreadsheetApp.BorderStyle.DOTTED
      }
    } else if (theme == 'white'){
      var defaultOptions = {
        background        : 'white'
        ,fontFamily       : 'Arial'
        ,fontSize         : 10
        ,textColor        : 'black'
        ,wrapStrategy     : SpreadsheetApp.WrapStrategy.CLIP
        ,fontWeight       : 'normal'
        ,setBorder        : true
        ,borderTop        : true
        ,borderLeft       : true
        ,borderBottom     : true
        ,borderRight      : true
        ,borderVertical   : true
        ,borderHorizontal : true
        ,borderColor      : 'black'
        ,borderStyle      : SpreadsheetApp.BorderStyle.DOTTED
      }
    } else if (theme == 'purple'){
      var defaultOptions = {
        background        : '#f0eaff'
        ,fontFamily       : 'Arial'
        ,fontSize         : 10
        ,textColor        : '#666666'
        ,wrapStrategy     : SpreadsheetApp.WrapStrategy.CLIP
        ,fontWeight       : 'normal'
        ,setBorder        : true
        ,borderTop        : true
        ,borderLeft       : true
        ,borderBottom     : true
        ,borderRight      : true
        ,borderVertical   : true
        ,borderHorizontal : true
        ,borderColor      : '#cccccc'
        ,borderStyle      : SpreadsheetApp.BorderStyle.DOTTED
      }
    } else if (theme == 'solarized light'){
      var defaultOptions = {
        background        : '#fdf6e3'
        ,fontFamily       : 'Arial'
        ,fontSize         : 10
        ,textColor        : '#9a9ca1'
        ,wrapStrategy     : SpreadsheetApp.WrapStrategy.CLIP
        ,fontWeight       : 'normal'
        ,setBorder        : true
        ,borderTop        : true
        ,borderLeft       : true
        ,borderBottom     : true
        ,borderRight      : true
        ,borderVertical   : true
        ,borderHorizontal : true
        ,borderColor      : '#7bc4bd'
        ,borderStyle      : SpreadsheetApp.BorderStyle.DOTTED
      }
    } else if (theme == 'wizpig'){
      var defaultOptions = {
        background        : '#ffe5dd'
        ,fontFamily       : 'Arial'
        ,fontSize         : 10
        ,textColor        : '#824588'
        ,wrapStrategy     : SpreadsheetApp.WrapStrategy.CLIP
        ,fontWeight       : 'normal'
        ,setBorder        : true
        ,borderTop        : true
        ,borderLeft       : true
        ,borderBottom     : true
        ,borderRight      : true
        ,borderVertical   : true
        ,borderHorizontal : true
        ,borderColor      : '#f52a2a'
        ,borderStyle      : SpreadsheetApp.BorderStyle.DOTTED
      }
    } else if (theme == 'mario'){
      var defaultOptions = {
        background        : '#7197e9'
        ,fontFamily       : 'Arial'
        ,fontSize         : 10
        ,textColor        : '#fbdb2f'
        ,wrapStrategy     : SpreadsheetApp.WrapStrategy.CLIP
        ,fontWeight       : 'normal'
        ,setBorder        : true
        ,borderTop        : true
        ,borderLeft       : true
        ,borderBottom     : true
        ,borderRight      : true
        ,borderVertical   : true
        ,borderHorizontal : true
        ,borderColor      : '#D77B58'
        ,borderStyle      : SpreadsheetApp.BorderStyle.DOTTED
      }
    }

    var options = datifyUtils().dictComplete(options_,defaultOptions)

    var rangeToFormat = returnObj.getDataRange()
    if(!rangeToFormat){
      Logger.log("header range not found, exit")
      return returnObj
    }

    if(options['background']){rangeToFormat.setBackground(options['background'])}
    if(options['fontFamily']){rangeToFormat.setFontFamily(options['fontFamily'])}
    if(options['fontSize']){rangeToFormat.setFontSize(options['fontSize'])}
    if(options['textColor']){rangeToFormat.setFontColor(options['textColor'])}
    if(options['wrapStrategy']){rangeToFormat.setWrapStrategy(options['wrapStrategy'])}
    if(options['fontWeight']){rangeToFormat.setFontWeight(options['fontWeight'])}

    if(options['setBorder'] == true){
      rangeToFormat.setBorder(
        options['borderTop']
        ,options['borderLeft']
        ,options['borderBottom']
        ,options['borderRight']
        ,options['borderVertical']
        ,options['borderHorizontal']
        ,options['borderColor']
        ,options['borderStyle']
        )
    }

    return returnObj
  }

  returnObj.styleHeader = function(theme = null,options_ = {}){

    if(theme == 'light gray'){
      var defaultOptions = {
        background        : '#e3e3e3'
        ,fontFamily       : 'Arial'
        ,fontSize         : 10
        ,textColor        : '#666666'
        ,wrapStrategy     : SpreadsheetApp.WrapStrategy.WRAP
        ,fontWeight       : 'bold'
        ,setBorder        : true
        ,borderTop        : true
        ,borderLeft       : true
        ,borderBottom     : true
        ,borderRight      : true
        ,borderVertical   : true
        ,borderHorizontal : true
        ,borderColor      : '#b7b7b7'
        ,borderStyle      : SpreadsheetApp.BorderStyle.SOLID
      }
    } else if (theme == 'gray contrast'){
      var defaultOptions = {
        background        : '#666666'
        ,fontFamily       : 'Arial'
        ,fontSize         : 10
        ,textColor        : '#efefef'
        ,wrapStrategy     : SpreadsheetApp.WrapStrategy.WRAP
        ,fontWeight       : 'bold'
        ,setBorder        : true
        ,borderTop        : true
        ,borderLeft       : true
        ,borderBottom     : true
        ,borderRight      : true
        ,borderVertical   : true
        ,borderHorizontal : true
        ,borderColor      : '#c8c8c8'
        ,borderStyle      : SpreadsheetApp.BorderStyle.SOLID
      }
    } else if (theme == 'dark'){
      var defaultOptions = {
        background        : '#666666'
        ,fontFamily       : 'Arial'
        ,fontSize         : 10
        ,textColor        : '#efefef'
        ,wrapStrategy     : SpreadsheetApp.WrapStrategy.WRAP
        ,fontWeight       : 'bold'
        ,setBorder        : true
        ,borderTop        : true
        ,borderLeft       : true
        ,borderBottom     : true
        ,borderRight      : true
        ,borderVertical   : true
        ,borderHorizontal : true
        ,borderColor      : '#c8c8c8'
        ,borderStyle      : SpreadsheetApp.BorderStyle.SOLID
      }
    } else if (theme == 'cornflower'){
      var defaultOptions = {
        background        : '#4a86e8'
        ,fontFamily       : 'Arial'
        ,fontSize         : 10
        ,textColor        : '#efefef'
        ,wrapStrategy     : SpreadsheetApp.WrapStrategy.WRAP
        ,fontWeight       : 'bold'
        ,setBorder        : true
        ,borderTop        : true
        ,borderLeft       : true
        ,borderBottom     : true
        ,borderRight      : true
        ,borderVertical   : true
        ,borderHorizontal : true
        ,borderColor      : '#c8c8c8'
        ,borderStyle      : SpreadsheetApp.BorderStyle.SOLID
      }
    } else if (theme == 'green'){
      var defaultOptions = {
        background        : '#93c47d'
        ,fontFamily       : 'Arial'
        ,fontSize         : 10
        ,textColor        : '#efefef'
        ,wrapStrategy     : SpreadsheetApp.WrapStrategy.WRAP
        ,fontWeight       : 'bold'
        ,setBorder        : true
        ,borderTop        : true
        ,borderLeft       : true
        ,borderBottom     : true
        ,borderRight      : true
        ,borderVertical   : true
        ,borderHorizontal : true
        ,borderColor      : '#c8c8c8'
        ,borderStyle      : SpreadsheetApp.BorderStyle.SOLID
      }
    } else if (theme == 'orange'){
      var defaultOptions = {
        background        : '#f6b26b'
        ,fontFamily       : 'Arial'
        ,fontSize         : 10
        ,textColor        : '#efefef'
        ,wrapStrategy     : SpreadsheetApp.WrapStrategy.WRAP
        ,fontWeight       : 'bold'
        ,setBorder        : true
        ,borderTop        : true
        ,borderLeft       : true
        ,borderBottom     : true
        ,borderRight      : true
        ,borderVertical   : true
        ,borderHorizontal : true
        ,borderColor      : '#c8c8c8'
        ,borderStyle      : SpreadsheetApp.BorderStyle.SOLID
      }
    } else if (theme == 'white'){
      var defaultOptions = {
        background        : 'white'
        ,fontFamily       : 'Arial'
        ,fontSize         : 10
        ,textColor        : 'black'
        ,wrapStrategy     : SpreadsheetApp.WrapStrategy.WRAP
        ,fontWeight       : 'bold'
        ,setBorder        : true
        ,borderTop        : true
        ,borderLeft       : true
        ,borderBottom     : true
        ,borderRight      : true
        ,borderVertical   : true
        ,borderHorizontal : true
        ,borderColor      : 'black'
        ,borderStyle      : SpreadsheetApp.BorderStyle.SOLID
      }
    } else if (theme == 'purple'){
      var defaultOptions = {
        background        : '#b4a7d6'
        ,fontFamily       : 'Arial'
        ,fontSize         : 10
        ,textColor        : '#f3f3f3'
        ,wrapStrategy     : SpreadsheetApp.WrapStrategy.WRAP
        ,fontWeight       : 'bold'
        ,setBorder        : true
        ,borderTop        : true
        ,borderLeft       : true
        ,borderBottom     : true
        ,borderRight      : true
        ,borderVertical   : true
        ,borderHorizontal : true
        ,borderColor      : '#f8f8f8'
        ,borderStyle      : SpreadsheetApp.BorderStyle.SOLID
      }
    } else if (theme == 'solarized light'){
      var defaultOptions = {
        background        : '#e1ec88'
        ,fontFamily       : 'Arial'
        ,fontSize         : 10
        ,textColor        : '#d8aa46'
        ,wrapStrategy     : SpreadsheetApp.WrapStrategy.WRAP
        ,fontWeight       : 'bold'
        ,setBorder        : true
        ,borderTop        : true
        ,borderLeft       : true
        ,borderBottom     : true
        ,borderRight      : true
        ,borderVertical   : true
        ,borderHorizontal : true
        ,borderColor      : '#7bc4bd'
        ,borderStyle      : SpreadsheetApp.BorderStyle.SOLID
      }
    } else if (theme == 'wizpig'){
      var defaultOptions = {
        background        : '#b98496'
        ,fontFamily       : 'Arial'
        ,fontSize         : 10
        ,textColor        : '#6e3a2e'
        ,wrapStrategy     : SpreadsheetApp.WrapStrategy.WRAP
        ,fontWeight       : 'bold'
        ,setBorder        : true
        ,borderTop        : true
        ,borderLeft       : true
        ,borderBottom     : true
        ,borderRight      : true
        ,borderVertical   : true
        ,borderHorizontal : true
        ,borderColor      : '#f52a2a'
        ,borderStyle      : SpreadsheetApp.BorderStyle.SOLID
      }
    } else if (theme == 'mario'){
      var defaultOptions = {
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
    }

    var options = datifyUtils().dictComplete(options_,defaultOptions)

    var rangeToFormat = returnObj.getHeaderRange()
    if(!rangeToFormat){
      Logger.log("header range not found, exit")
      return returnObj
    }

    if(options['background']){rangeToFormat.setBackground(options['background'])}
    if(options['fontFamily']){rangeToFormat.setFontFamily(options['fontFamily'])}
    if(options['fontSize']){rangeToFormat.setFontSize(options['fontSize'])}
    if(options['textColor']){rangeToFormat.setFontColor(options['textColor'])}
    if(options['wrapStrategy']){rangeToFormat.setWrapStrategy(options['wrapStrategy'])}
    if(options['fontWeight']){rangeToFormat.setFontWeight(options['fontWeight'])}

    if(options['setBorder'] == true){
      rangeToFormat.setBorder(
        options['borderTop']
        ,options['borderLeft']
        ,options['borderBottom']
        ,options['borderRight']
        ,options['borderVertical']
        ,options['borderHorizontal']
        ,options['borderColor']
        ,options['borderStyle']
        )
    }

    return returnObj
  }

  returnObj.style = function(theme = 'light gray'){
    returnObj.styleHeader(theme)
    returnObj.styleBody(theme)
    return returnObj
  }

  // create column objects
  returnObj.createColObjects()

  return returnObj
}


/*
------------------------------------------------------------------------------------
-- Datify Utils
------------------------------------------------------------------------------------
*/

function datifyUtils(){
  var returnObj = {}

  returnObj.transpose = function(array){
    return Object.keys(array[0]).map(function (col) { return array.map(function (row) { return row[col]; }); });
  }

  returnObj.getRangeListFromRanges = function(ranges,sheet = null){
    if(!sheet) {sheet = ranges[0].getSheet()}
    var a1_list = [];
    var a1_list = ranges.map(range => range.getA1Notation())
    var range_list = sheet.getRangeList(a1_list);
    return range_list;
  }

  returnObj.dictComplete = function(dict, template) {

    if (!dict){return template}
    
    const dict_ = {...dict};

    for (const k of Object.keys(template)) {
      dict_[k] = dict_[k] ?? template[k]
    }

    return dict_
  }

  returnObj.capitalizePhrase = function(phrase){
    var reg = /\b(\w)/g;                              
    function replace(firstLetters) {
      return firstLetters.toUpperCase();
    }
    capitalized = phrase.replace(reg, replace);
    return capitalized;
  }

  return returnObj 
}

/*
------------------------------------------------------------------------------------
-- Datify Menu
------------------------------------------------------------------------------------
*/


function createDatifyMenu(){ 
  var menu = SpreadsheetApp.getUi().createMenu('Datify')

  //menu.addItem('Main Function', 'nonExistingFuncExample')

  menu.addSubMenu(SpreadsheetApp.getUi().createMenu('Datarange theme')

    .addItem('Cornflower'     , 'Datify.styleSelectedNamedRangeCornflower')
    .addItem('Dark'           , 'Datify.styleSelectedNamedRangeDark')
    .addItem('Gray Contrast'  , 'Datify.styleSelectedNamedRangeGrayContrast')
    .addItem('Green'          , 'Datify.styleSelectedNamedRangeGreen')
    .addItem('Light Gray'     , 'Datify.styleSelectedNamedRangeLightGray')
    .addItem('Mario'          , 'Datify.styleSelectedNamedRangeMario')
    .addItem('Orange'         , 'Datify.styleSelectedNamedRangeOrange')
    .addItem('Purple'         , 'Datify.styleSelectedNamedRangePurple')
    .addItem('Solarized Light', 'Datify.styleSelectedNamedRangeSolarizedLight')
    .addItem('White'          , 'Datify.styleSelectedNamedRangeWhite')
    .addItem('Wizpig'         , 'Datify.styleSelectedNamedRangeWizpig')
    )
  
  menu.addSubMenu(SpreadsheetApp.getUi().createMenu('Email CSV')
    .addItem('To Self', 'Datify.emailCsvToSelf')
    .addItem('To Others', 'Datify.emailCsvToOthers')
    )

  menu.addToUi()
}


function getSelectedNamedRange(){
  var doc = SpreadsheetApp.getActive()
  var sheet = doc.getActiveSheet()
  var activeRange = doc.getSelection().getActiveRange()
  var namedRanges = sheet.getNamedRanges()
  var selectedNamedRanges = namedRanges.filter(
    namedRange => isRangeContained(namedRange.getRange(),activeRange)
    )
  var selectedNamedRange = selectedNamedRanges[0]
  return selectedNamedRange
}

function emailCsvToSelf(){
  var selectedNamedRange = getSelectedNamedRange()
  var namedRangeName = selectedNamedRange.getName()
  var dr = datarange(namedRangeName)
  var da = dr.datarray()
  da.emailCsv(
    "datarrayCsv"
    ,Session.getActiveUser().getEmail()
    ,namedRangeName + " datarrayCsv"
    ,body = ""
    )
}

function emailCsvToOthers(){
  var selectedNamedRange = getSelectedNamedRange()
  var namedRangeName = selectedNamedRange.getName()
  var dr = datarange(namedRangeName)
  var da = dr.datarray()
  var otherEmails = Browser.inputBox('Enter emails separated by comma', Browser.Buttons.OK_CANCEL)
  if(otherEmails == 'cancel'){return}
  da.emailCsv(
    "datarrayCsv"
    ,otherEmails
    ,namedRangeName + " datarrayCsv"
    ,body = ""
    )
}

function styleSelectedNamedRange(theme){
  var selectedNamedRange = getSelectedNamedRange()
  var dr = datarange(selectedNamedRange.getName())
  dr.style(theme)
}

function styleSelectedNamedRangeCornflower(){styleSelectedNamedRange('cornflower')}
function styleSelectedNamedRangeDark(){styleSelectedNamedRange('dark')}
function styleSelectedNamedRangeGrayContrast(){styleSelectedNamedRange('gray contrast')}
function styleSelectedNamedRangeGreen(){styleSelectedNamedRange('green')}
function styleSelectedNamedRangeLightGray(){styleSelectedNamedRange('light gray')}
function styleSelectedNamedRangeMario(){styleSelectedNamedRange('mario')}
function styleSelectedNamedRangeOrange(){styleSelectedNamedRange('orange')}
function styleSelectedNamedRangePurple(){styleSelectedNamedRange('purple')}
function styleSelectedNamedRangeSolarizedLight(){styleSelectedNamedRange('solarized light')}
function styleSelectedNamedRangeWhite(){styleSelectedNamedRange('white')}
function styleSelectedNamedRangeWizpig(){styleSelectedNamedRange('wizpig')}


function styleSelectedNamedRangeWizpig(){styleSelectedNamedRange('wizpig')}

function isRangeContained(range1,range2) {

  // compares ranges based on run type and returns true or false if range1 contains range2

  var range1_sheet      = range1.getSheet().getName();
  var range1_start_row  = range1.getRow();
  var range1_start_col  = range1.getColumn();
  var range1_end_row    = range1.getNumRows() + range1_start_row - 1;
  var range1_end_col    = range1.getNumColumns() + range1_start_col - 1;

  var range2_sheet      = range2.getSheet().getName();
  var range2_start_row  = range2.getRow();
  var range2_start_col  = range2.getColumn();
  var range2_end_row    = range2.getNumRows() + range2_start_row - 1;
  var range2_end_col    = range2.getNumColumns() + range2_start_col - 1;

  var return_val = null


  // "contains" run type
  ////////////////////////////////////////////////////////////

  if(
    range1_sheet != range2_sheet
    || range1_start_col > range2_start_col
    || range1_start_row > range2_start_row
    || range1_end_col < range2_end_col
    || range1_end_row < range2_end_row
    ){return_val = false}
  else{return_val = true}


  return return_val
  
}

/*
------------------------------------------------------------------------------------
-- Bq Utils
------------------------------------------------------------------------------------
*/

function query(projectId, query) {
  var sqlPrefix = "#standardSQL \n  "; // This is required to run any SQL using the BigQuery API (it doesn't hurt if it is there more than once so don't even check)
  
  var request = {query: sqlPrefix + query}

  try {
      var queryResults = BigQuery.Jobs.query(request, projectId);
  } catch(e) {
    throw e;
  }

  //Logger.log("queryResults = %s",queryResults)

  var jobId = queryResults.jobReference.jobId;
 
  var sleepTimeMs = 500;
  var maxWaitMs = 60000;
  var waitedMs = 0;

  while (!queryResults.jobComplete && waitedMs < maxWaitMs) {
    Utilities.sleep(sleepTimeMs);
    waitedMs += sleepTimeMs;
    queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId);
  }

  //store the first batch of return data from BigQuery
  var rows = queryResults.rows;
  
  //Now check for more data
  while (queryResults.pageToken) {
    queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId, {pageToken: queryResults.pageToken});
    rows = rows.concat(queryResults.rows);
  }

  if (rows) {

    var headers = queryResults.schema.fields.map(v=>v['name'])

    var array2d = new Array(rows.length); 
    array2d[0] = headers; //add the header of this column to the first row in the data array

    for (var row = 0; row < rows.length; row++) {
      var rowData = rows[row].f;   
      // now add each col for this row
      array2d[row + 1] = new Array(rowData.length);
      for (var col = 0; col < rowData.length; col++) {
         array2d[row + 1][col] = rowData[col].v;
      }
    }

    return array2d;

  } else { // no rows found
    if (waitedMs >= maxWaitMs) {
      Logger.log('Results did not return within the time limit of %d seconds. \nSQL: %s\n', maxWaitMs / 1000, query);
    } 
    return null;
  }  
}
