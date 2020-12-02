// 34567890123456789012345678901234567890123456789012345678901234567890123456789

// JSHint - TODO
/* jshint asi: true */

/*******************************************************************************
 * Copyright (c) 2019 Andrew Roberts
 * 
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *   http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

// Utils_.gs
// =========

var Utils_ = {

/**
 * Get a column number
 *
 * @param {Sheet} sheet
 * @param {String}  columnName
 *
 * @return {Number} column number or -1
 */
     
getColumnNumber: function(sheet, columnName) {
    
  Log.functionEntryPoint()
  
  var numberOfHeaders = sheet.getLastColumn()
  var responseHeaders = sheet.getRange(1, 1, 1, numberOfHeaders).getValues()[0]
  var columnIndex = 0
  
  var foundHeader = responseHeaders.some(function(cellValue) {
    if (cellValue === columnName) {
      return true
    }
    columnIndex++      
  })
  
  if (!foundHeader) {
    return -1
  }
  
  Log.fine('columnIndex: ' + columnIndex)
  
  return columnIndex + 1
    
}, // Utils_.getColumnNumber() 

/**
 * Store a value in a GSheet
 *
 * @param {Sheet} sheet
 * @param {Number} rowNumber
 * @param {String} headerName
 * @param {Object} value
 */
  
storeValue: function (sheet, rowNumber, headerName, value) {
  
  Log.functionEntryPoint()
  
  Log.fine('sheet (name): ' + sheet.getName())
  Log.fine('rowNumber: ' + rowNumber)
  Log.fine('headerName: ' + headerName)
  Log.fine('value: ' + value)
  
  var columnNumber = Utils_.getColumnNumber(sheet, headerName)
  sheet.getRange(rowNumber, columnNumber).setValue(value)
  Log.info('Value - ' + value + ' - written to "' + headerName + '" column')
  
}, // Utils_.storeValue() 
  
/**
 * Get information about the active cell/row and check it
 *
 * @param {UI} ui
 *
 * @return {Object}
 */
 
getActiveCellObject: function(ui) {

  Log.functionEntryPoint() 
  
  var activeSheet = SpreadsheetApp.getActiveSheet()
  var activeSpreadsheet = activeSheet.getParent()
  var sheetName = activeSheet.getName()
  Log.fine('sheetName: ' + sheetName)
  
  if (sheetName !== ORGANISATIONS_SHEET_NAME) {
    ui.alert('This feature only works on the "' + ORGANISATIONS_SHEET_NAME + '" sheet.')
    Log.warning('User selected option outside of the "' + ORGANISATIONS_SHEET_NAME + '" sheet.')
    return null
  }
    
  var numberOfColumns = activeSheet.getLastColumn()
  Log.fine('numberOfColumns: ' + numberOfColumns)
  
  var activeRange = activeSheet.getActiveRange()
  
  if (activeRange === null) {
    ui.alert('Select a row before clicking the menu.')
    Log.warning('User selected option without selecting an organisation\'s row')    
    return null
  }
  
  var activeRowNumber = activeRange.getRow()
  
  if (activeRowNumber > activeSheet.getLastRow()) {
    ui.alert('This row is empty, select a completed one.')
    Log.warning('User selected option on an empty row')    
    return null
  }

  if (activeRowNumber === 1) {
    ui.alert('You have selected the header row, select an organisation\'s.')
    Log.warning('User selected option on header row')
    return null
  }

  return {
    spreadsheet: activeSpreadsheet,
    sheet: activeSheet,
    range: activeRange,
    rowNumber: activeRowNumber,
    numberOfColumns: numberOfColumns,
  }
  
}, // Utils_.getActiveCellObject() 

/**
 * Get a date string of the form dd-MMM-yyyy
 *
 * @return {String} date string
 */
 
getDateString: function() {

  Log.functionEntryPoint()
  var date = new Date()
  var timeZone = Session.getScriptTimeZone()
  return Utilities.formatDate(date, timeZone, 'dd-MMM-yyyy')
  
}, // Utils_.getDateString() 

} // Utils_