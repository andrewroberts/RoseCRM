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

function test_Utils() {
//   const a = Utils_.getId('https://docs.google.com/spreadsheets/d/1Ylb95IeopzCaZvMKIjeQ7oHlqi1TLwAkgrb--GR8u3I/edit?gid=1418283387#gid=1418283387')
  const a = Utils_.getId('')
  debugger
}

const Utils_ = {

alert: function(msg) {
  Log_.info(msg)  
  var spreadsheet = SpreadsheetApp.getActive()
  if (spreadsheet === null) return
  SpreadsheetApp.getUi().alert(msg)
},

toast: function(msg, timeout = null) {
  Log_.info(msg)
  var spreadsheet = SpreadsheetApp.getActive()
  if (spreadsheet === null) return
  spreadsheet.toast(msg, 'Revitaliste Project Manager', timeout)
},

// http://stackoverflow.com/questions/16840038/easiest-way-to-get-file-id-from-url-on-google-apps-script/16840612
getId: function(url) {
  const id = url.match(/[-\w]{25,}/)
  if (!id) throw new Error(`The URL ${url} does not contain an ID`)
  return id[0]
},

getConfig: function() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Config')
  const data = sheet.getDataRange().getValues()
  let config = {}
  data.forEach(([key, value]) => {
    config[key] = (value[0] === "{") ? JSON.parse(value) : value
  })
  return config
},

/**
 * Get a column number
 *
 * @param {Sheet} sheet
 * @param {String}  columnName
 *
 * @return {Number} column number or -1
 */
     
getColumnNumber: function(sheet, columnName) {
    
  const numberOfHeaders = sheet.getLastColumn()
  const responseHeaders = sheet.getRange(1, 1, 1, numberOfHeaders).getValues()[0]
  let columnIndex = 0
  
  const foundHeader = responseHeaders.some(function(cellValue) {
    if (cellValue === columnName) {
      return true
    }
    columnIndex++      
  })
  
  if (!foundHeader) {
    return -1
  }
  
  Log_.fine('columnIndex: ' + columnIndex)
  
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
  
  Log_.fine('sheet (name): ' + sheet.getName())
  Log_.fine('rowNumber: ' + rowNumber)
  Log_.fine('headerName: ' + headerName)
  Log_.fine('value: ' + value)
  
  const columnNumber = Utils_.getColumnNumber(sheet, headerName)
  sheet.getRange(rowNumber, columnNumber).setValue(value)
  Log_.info('Value - ' + value + ' - written to "' + headerName + '" column')
  
}, // Utils_.storeValue() 
  
/**
 * Get information about the active cell/row and check it
 *
 * @param {UI} ui
 *
 * @return {Object}
 */
 
getActiveCellObject: function(ui) {

  const activeSheet = SpreadsheetApp.getActiveSheet()
  const activeSpreadsheet = activeSheet.getParent()
  const sheetName = activeSheet.getName()
  Log_.fine('sheetName: ' + sheetName)

  const config = Utils_.getConfig()
  
  if (sheetName !== config.CLIENTS_SHEET_NAME) {
    ui.alert('This feature only works on the "' + config.CLIENTS_SHEET_NAME + '" sheet.')
    Log_.warning('User selected option outside of the "' + config.CLIENTS_SHEET_NAME + '" sheet.')
    return null
  }
    
  const numberOfColumns = activeSheet.getLastColumn()
  Log_.fine('numberOfColumns: ' + numberOfColumns)
  
  const activeRange = activeSheet.getActiveRange()
  
  if (activeRange === null) {
    ui.alert('Select a row before clicking the menu.')
    Log_.warning('User selected option without selecting an client\'s row')    
    return null
  }
  
  const activeRowNumber = activeRange.getRow()
  
  if (activeRowNumber > activeSheet.getLastRow()) {
    ui.alert('This row is empty, select a completed one.')
    Log_.warning('User selected option on an empty row')    
    return null
  }

  if (activeRowNumber === 1) {
    ui.alert('You have selected the header row, select an client\'s.')
    Log_.warning('User selected option on header row')
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

  const date = new Date
  const timeZone = Session.getScriptTimeZone()
  return Utilities.formatDate(date, timeZone, 'dd-MMM-yyyy')
  
}, // Utils_.getDateString() 

} // Utils_