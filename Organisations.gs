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

// Organisations.gs
// ================
//
// External interface to this script - all of the event handlers.
//
// This files contains all of the event handlers, plus miscellaneous functions 
// not worthy of their own files yet

// Public event handlers
// ---------------------
//
// All external event handlers need to be top-level function calls; they can't 
// be part of an object, and to ensure they are all processed similarily 
// for things like logging and error handling, they all go through 
// errorHandler_(). These can be called from custom menus, web apps, 
// triggers, etc
// 
// The main functionality of a call is in a function with the same name but 
// post-fixed with an underscore (to indicate it is private to the script)
//
// For debug, rather than production builds, lower level functions are exposed
// in the menu

//   :      [function() {},  '()',      'Failed to ', ],

var EVENT_HANDLERS = {

//                         Initial actions  Name                         onError Message                        Main Functionality
//                         ---------------  ----                         ---------------                        ------------------

  onFormSubmit:            [function() {},  'onFormSubmit()',            'Failed to process form submit',       onFormSubmit_],
//  onEdit:                  [function() {},  'onEdit()',                  'Failed to process edit',              onEdit_],  
  onCreateContract:        [function() {},  'onCreateContract()',        'Failed to create contract',           onCreateContract_],  
  onCreateNda:             [function() {},  'onCreateNda()',             'Failed to create NDA',                onCreateNda_],    
  onAddNote:               [function() {},  'onAddNote()',               'Failed to add note',                  onAddNote_],    
}

// function (arg)                     {return eventHandler_(EVENT_HANDLERS., arg)}

function onFormSubmit(arg)               {return eventHandler_(EVENT_HANDLERS.onFormSubmit, arg)}
//function onEdit(arg)                     {return eventHandler_(EVENT_HANDLERS.onEdit, arg)}
function onCreateContract(arg)           {return eventHandler_(EVENT_HANDLERS.onCreateContract, arg)}
function onCreateNda(arg)                {return eventHandler_(EVENT_HANDLERS.onCreateNda, arg)}
function onAddNote(arg)                  {return eventHandler_(EVENT_HANDLERS.onAddNote, arg)}

/**
 * Event handler for the sheet being opened. This is a special case
 * as all it can do is create a menu whereas the usual eventHandler_()
 * does things we don't have permission for at this stage.
 */

function onOpen() {

  Log.functionEntryPoint()
  
  var ui = SpreadsheetApp.getUi()
  var menu = ui.createMenu('Organisations')

  menu
    .addItem('Create NDA', 'onCreateNda')  
    .addItem('Create contract', 'onCreateContract')
    .addItem('Add new note...', 'onAddNote')    
    .addToUi()
    
} // onOpen()

// Private Functions
// =================

// General
// -------

/**
 * All external function calls should call this to ensure standard 
 * processing - logging, errors, etc - is always done.
 *
 * @param {array} config:
 *   [0] {function} prefunction
 *   [1] {string} eventName
 *   [2] {string} onErrorMessage
 *   [3] {function} mainFunction
 * @parma {object} arg The argument passed to the top-level event handler
 */

function eventHandler_(config, arg) {

  try {

    config[0]()

    Log.init({
      level: LOG_LEVEL, 
      sheetId: LOG_SHEET_ID,
      displayFunctionNames: LOG_DISPLAY_FUNCTION_NAMES})
    
    Log.info('Handling ' + config[1])
    
    Assert.init({
      handleError: HANDLE_ERROR, 
      sendErrorEmail: SEND_ERROR_EMAIL, 
      emailAddress: ADMIN_EMAIL_ADDRESS,
      scriptName: SCRIPT_NAME,
      scriptVersion: SCRIPT_VERSION, 
    })
    
    return config[3](arg)
    
  } catch (error) {
  
    Assert.handleError(error, config[2], Log) 
  }
  
} // eventHandler_()

// Private event handlers
// ----------------------

/**
 * Private 'on install' event handler
 */

function onInstall_() {

  // TODO - Anything that needs doing on installation

} // onInstall_()

/**
 *
 */

function onFormSubmit_(event) {

/*
  Logger.log(event)

  var range = event.range
  for (var key in range) {
    Logger.log(key + ': ' + range[key])
  }

  return
  
{
  authMode=FULL, 
  values=[09/07/2016 18:42:29, , , , , , , , , , , , , , ], 
  namedValues={Status=[], Address=[], Rate=[], Contact Title=[], Timesheet=[], Journal=[], Folder=[], Timestamp=[09/07/2016 18:42:29], Trello Board=[], Company Name=[], Contact Email=[], Contact First Name=[], Contact Last Name=[], Contract=[], Notes=[]}, 
  range={columnEnd=2, columnStart=0, rowStart=11, rowEnd=12}
  source=Spreadsheet, 
  triggerUid=988897199
}
*/

  var values = event.namedValues
  var companyName = values['Company Name'][0]
  var contactName = values['Name'][0]
  var orgsFolder = DriveApp.getFolderById(ORGS_FOLDER_ID)
  var spreadsheet = event.source
  var responseRange = event.range
  var responseSheet = responseRange.getSheet()
  var rootFolder = DriveApp.getRootFolder()
  var rowNumber = responseRange.getRow()

  // Format new responses

  responseSheet.getRange('B:H').setFontWeight('bold')

  // Set colour

  var statusColumnNumber = Utils_.getColumnNumber(responseSheet, PIPELINE_HEADER_NAME)
  var statusRange = responseSheet.getRange(rowNumber, statusColumnNumber)

//  onEdit_({
//    range: statusRange,
//  })
  
  // Create Org Folder
  
  var companyFolderName
  
  if (companyName !== '') {
    companyFolderName = companyName
  } else {
    companyFolderName = contactName 
  }

  var companyFolders = DriveApp.getFoldersByName(companyFolderName)
  var companyFolder
  
  if (!companyFolders.hasNext()) {  
  
    companyFolder = DriveApp.createFolder(companyFolderName)
    orgsFolder.addFolder(companyFolder)
    rootFolder.removeFolder(companyFolder)
    var folderUrl = companyFolder.getUrl()
    storeValue(COMPANY_FOLDER_HEADER_NAME, folderUrl)
    Log.info('Created folder: ' + companyFolderName + ' (' + folderUrl + ')')
    
  } else {
  
    companyFolder = companyFolders.next()
    
    if (companyFolders.hasNext()) {
      throw new Error('Found two folders for this organisation')
    }
    
    Log.info('Using existing folder: "' + companyFolderName + '" - assuming all files already made')
    return
  }

  // Create timesheet

  var timesheetFileUrl = createFile(TIMESHEET_TEMPLATE_SHEET_ID)
  storeValue(TIMESHEET_HEADER_NAME, timesheetFileUrl)
  Log.info('Created timesheet: ' + timesheetFileUrl)
  
  // Create journal
  
  var journalFileUrl = createFile()
  storeValue(JOURNAL_HEADER_NAME, journalFileUrl)
  Log.info('Created journal: ' + journalFileUrl)

  // Fill in Company Name column if it is empty
  
  if (companyName === '') {
    storeValue(COMPANY_NAME_HEADER_NAME, contactName)
  }

  createNotesSheet()
  GmailApp.createLabel('R/AJR Comp/' + companyName)
  return

  // Private Functions
  // -----------------

  /**
   * Store a value in the response sheet
   *
   * @param {String} headerName
   * @param {Object} value
   *
   * @return {object}
   */
   
  function storeValue(headerName, value) {
  
    Log.functionEntryPoint()    
    Utils_.storeValue(responseSheet, rowNumber, headerName, value)
    
  } // storeValue() 

  /**
   * Create a new organisation file
   *
   * @param {String} templateId 
   *
   * @return {String} new file's URL
   */
   
  function createFile(templateId) {
  
    Log.functionEntryPoint()
    
    // rootFolder
    // companyFolder
    // companyFolderName
    
    var file
    
    if (templateId === TIMESHEET_TEMPLATE_SHEET_ID) {
      
      file = SpreadsheetApp
        .openById(templateId)
        .copy('Timesheet - OPEN - ' + companyFolderName)
        
      DriveApp.getFolderById(TIMESHEETS_FOLDER_ID).createShortcut(file.getId())
        
    } else {
    
      file = DocumentApp.create('Journal - ' + companyFolderName)
    }
  
    DriveApp.getFileById(file.getId()).moveTo(companyFolder)
    return file.getUrl()
    
  } // onFormSubmit_.createFile() 
  
  /**
   * Create a new notes sheet
   */
   
  function createNotesSheet() {
  
    Log.functionEntryPoint()
    
    spreadsheet.getSheetByName(NOTES_TEMPLATE_SHEET_NAME).activate()
    spreadsheet.duplicateActiveSheet().setName(companyName)
    spreadsheet.moveActiveSheet(2) // After the organisations sheet
    responseSheet.activate()
    Log.info('Created new notes sheet for "%s"', companyName)
    
  } // onFormSubmit_.createNotesSheet() 
  
} // onFormSubmit_()

/**
 * Private event handler for "on contract create" event
 */
 
function onCreateContract_() {

  Log.functionEntryPoint()
  
  // Set up the docs and the spreadsheet access
  
  var ui = SpreadsheetApp.getUi()
  
  var copyFile = DriveApp.getFileById(CONTRACT_TEMPLATE_SHEET_ID).makeCopy()
  var copyId = copyFile.getId()
  var copyDoc = DocumentApp.openById(copyId)  
  var copyBody = copyDoc.getBody()
  var activeSheet = SpreadsheetApp.getActiveSheet()
  
  if (activeSheet.getName() !== ORGANISATIONS_SHEET_NAME) {
    ui.alert('Only works on "Organisations" sheet.')
    return
  }
    
  var numberOfColumns = activeSheet.getLastColumn()
  Log.fine('numberOfColumns: ' + numberOfColumns)
  var activeRange = activeSheet.getActiveRange()
  
  if (activeRange === null) {
    ui.alert('Select a row before clicking the menu.')
    return
  }
  
  var activeRowNumber = activeRange.getRow()
  
  if (activeRowNumber > activeSheet.getLastRow()) {
    ui.alert('This row is empty, select a completed one.')
    return
  }

  if (activeRowNumber === 1) {
    ui.alert('You have selected the header row, select an organisation\'s.')
    return
  }
    
  var activeRow = activeSheet.getRange(activeRowNumber, 1, 1, numberOfColumns).getValues()[0]
  var headerRow = activeSheet.getRange(1, 1, 1, numberOfColumns).getValues()[0]
  
  var companyName = null
  var companyFolderUrl = null
  var rate = null
  var address = null
  
  // Replace the keys with the spreadsheet values
 
  for (var columnIndex = 0; columnIndex < numberOfColumns; columnIndex++) {

    var header = headerRow[columnIndex]
    var value = activeRow[columnIndex]

    Log.finest('header: ' + header)
    Log.finest('value: ' + value)

    if (header === COMPANY_NAME) {
    
      companyName = value
      Log.fine('Found company name')
      
    } else if (header === COMPANY_FOLDER_HEADER_NAME) {
    
      companyFolderUrl = value
      Log.fine('Found company folder URL')
            
    } else if (header === RATE_NAME) {
    
      rate = value
      Log.fine('Found rate')
            
    } else if (header === COMPANY_ADDRESS_NAME) {
    
      address = value
      Log.fine('Found company address')      
    }

    copyBody.replaceText('%' + header + '%', value)    
    
  } // For each column
  
  var timeZone = Session.getScriptTimeZone()
  var date = new Date()
  var dateString = Utilities.formatDate(date, timeZone, 'dd-MMM-YYYY')
  Log.fine('dateString: ' + dateString)
  copyBody.replaceText('%ContractDate%', dateString)
    
  // Check we found all the info we need
  
  if (companyName === null || companyName === '') {
  
    ui.alert('Could not find the company name')
    return
    
  } else if (companyFolderUrl === null || companyFolderUrl === '') {
  
    ui.alert('Could not find the company folder URL')
    return
    
  } else if (rate === null || rate === '') {
  
    ui.alert('Could not find the rate')
    return
    
  } else if (address === null || address === '') {
  
    ui.alert('Could not find the company address')
    return
  }
    
  // Create the PDF file, rename it if required and delete the doc copy
    
  copyDoc.saveAndClose()
  
  var contractPdfFile = DriveApp.createFile(copyFile.getAs('application/pdf'))  
  
  copyFile.setTrashed(true)
  
  var contractName

  contractName = 'AndrewRoberts.net Contract - ' + companyName
  contractPdfFile.setName(contractName)
  Log.info('Created new contract "' + contractName +'"')
  DriveApp.getRootFolder().removeFile(contractPdfFile) // Orphaned
  
  // Put the new contract in the company folder
  
  // http://stackoverflow.com/questions/16840038/easiest-way-to-get-file-id-from-url-on-google-apps-script/16840612
  var companyFolderId = companyFolderUrl.match(/[-\w]{25,}/)
  Log.fine('companyFolderId: ' + companyFolderId)
  var companyFolder = DriveApp.getFolderById(companyFolderId)
  var companyFolderName = companyFolder.getName()
  companyFolder.addFile(contractPdfFile)
  var contractFolder = DriveApp.getFolderById(CONTRACT_FOLDER_ID)
  contractFolder.createShortcut(contractPdfFile.getId())
  Log.info('Put contract into company folder "' + companyFolder.getName() + '"')
 
  Utils_.storeValue(
    activeSheet, 
    activeRowNumber, 
    CONTRACT_HEADER_NAME, 
    contractPdfFile.getUrl())
     
  ui.alert('New contract "' + contractName + '" created in the folder "' + companyFolderName + '"')  

} // onCreateContract_() 

/**
 * Private event handler for "on NDA create" event
 */
 
function onCreateNda_() {

  Log.functionEntryPoint()
  
  // Set up the docs and the spreadsheet access
  
  var ui = SpreadsheetApp.getUi()
  
  var copyFile = DriveApp.getFileById(NDA_TEMPLATE_SHEET_ID).makeCopy()
  var copyId = copyFile.getId()
  var copyDoc = DocumentApp.openById(copyId)  
  var copyBody = copyDoc.getBody()
  var activeSheet = SpreadsheetApp.getActiveSheet()
  
  if (activeSheet.getName() !== ORGANISATIONS_SHEET_NAME) {
    ui.alert('Only works on "Organisations" sheet.')
    return
  }
    
  var numberOfColumns = activeSheet.getLastColumn()
  Log.fine('numberOfColumns: ' + numberOfColumns)
  var activeRange = activeSheet.getActiveRange()
  
  if (activeRange === null) {
    ui.alert('Select a row before clicking the menu.')
    return
  }
  
  var activeRowNumber = activeRange.getRow()
  
  if (activeRowNumber > activeSheet.getLastRow()) {
    ui.alert('This row is empty, select a completed one.')
    return
  }

  if (activeRowNumber === 1) {
    ui.alert('You have selected the header row, select an organisation\'s.')
    return
  }
    
  var activeRow = activeSheet.getRange(activeRowNumber, 1, 1, numberOfColumns).getValues()[0]
  var headerRow = activeSheet.getRange(1, 1, 1, numberOfColumns).getValues()[0]
  
  var companyName = null
  var companyFolderUrl = null
  var rate = null
  var address = null
  
  // Replace the keys with the spreadsheet values
 
  for (var columnIndex = 0; columnIndex < numberOfColumns; columnIndex++) {

    var header = headerRow[columnIndex]
    var value = activeRow[columnIndex]

    Log.finest('header: ' + header)
    Log.finest('value: ' + value)

    if (header === COMPANY_NAME) {
    
      companyName = value
      Log.fine('Found company name')
      
    } else if (header === COMPANY_FOLDER_HEADER_NAME) {
    
      companyFolderUrl = value
      Log.fine('Found company folder URL')
            
    } else if (header === RATE_NAME) {
    
      rate = value
      Log.fine('Found rate')
            
    } else if (header === COMPANY_ADDRESS_NAME) {
    
      address = value
      Log.fine('Found company address')      
    }

    copyBody.replaceText('%' + header + '%', value)    
    
  } // For each column
  
  var timeZone = Session.getScriptTimeZone()
  var date = new Date()
  var dateString = Utilities.formatDate(date, timeZone, 'dd-MMM-YYYY')
  Log.fine('dateString: ' + dateString)
  copyBody.replaceText('%ContractDate%', dateString)
    
  // Check we found all the info we need
  
  if (companyName === null || companyName === '') {
  
    ui.alert('Could not find the company name')
    return
    
  } else if (companyFolderUrl === null || companyFolderUrl === '') {
  
    ui.alert('Could not find the company folder URL')
    return
    
  } else if (address === null || address === '') {
  
    ui.alert('Could not find the company address')
    return
  }
    
  // Create the PDF file, rename it if required and delete the doc copy
    
  copyDoc.saveAndClose()
  
  var ndaPdfFile = DriveApp.createFile(copyFile.getAs('application/pdf'))  
  
  copyFile.setTrashed(true)
  
  var ndaName

  ndaName = 'Andrew Roberts NDA - ' + companyName
  ndaPdfFile.setName(ndaName)
  Log.info('Created new NDA "' + ndaName +'"')
  
  // Put the new nda in the company folder
  
  var companyFolderName
  
  // http://stackoverflow.com/questions/16840038/easiest-way-to-get-file-id-from-url-on-google-apps-script/16840612
  var companyFolderId = companyFolderUrl.match(/[-\w]{25,}/)
  Log.fine('companyFolderId: ' + companyFolderId)
  var companyFolder = DriveApp.getFolderById(companyFolderId)
  companyFolderName = companyFolder.getName()
  companyFolder.addFile(ndaPdfFile)
  var ndaFolder = DriveApp.getFolderById(CONTRACT_FOLDER_ID)
  ndaFolder.addFile(ndaPdfFile)
  DriveApp.getRootFolder().removeFile(ndaPdfFile)
  Log.info('Put NDA into company folder "' + companyFolder.getName() + '"')
 
  Utils_.storeValue(
    activeSheet, 
    activeRowNumber, 
    NDA_HEADER_NAME, 
    ndaPdfFile.getUrl())
     
  ui.alert('New NDA "' + ndaName + '" created in the folder "' + companyFolderName + '"')  

} // onCreateNda_() 

/**
 * Add a note to the organisations sheet
 *
 * @param {object} 
 *
 * @return {object}
 */
 
function onAddNote_() {

  Log.functionEntryPoint()
  
  // Get the active cell first, to check it is valid
  
  var ui = SpreadsheetApp.getUi()  
  var activeCell = Utils_.getActiveCellObject(ui)
  
  if (activeCell === null) {
    // The Error has been logged in Utils_.getActiveCellObject()
    return
  }
  
  var activeSheet = activeCell.sheet
  var activeSpreadsheet = activeCell.spreadsheet
  var activeRowNumber = activeCell.rowNumber
  
  // Get the new note
  
  var response = ui.prompt('What\'s the note?', ui.ButtonSet.OK_CANCEL)
  var note = ''
    
  if (response.getSelectedButton() == ui.Button.OK) {
  
    note = response.getResponseText()
    
  } else if (response.getSelectedButton() == ui.Button.CANCEL) {
 
    Log.warning('The user didn\'t want to provide a note.')
    return
   
  } else {
 
    Log.warning('The user clicked the close button in the dialog\'s title bar.')
    return
  }

  // Get the notes sheet for this organisation

  var companyNameColumnNumber = Utils_.getColumnNumber(
    activeSheet,
    COMPANY_NAME_HEADER_NAME)
    
  var companyName = activeSheet
    .getRange(
      activeRowNumber, 
      companyNameColumnNumber)
    .getValue()
    
  Log.fine('companyName: ' + companyName)
  
  if (companyName === '') {
    Log.warning('User selected row with no organisation name')
    ui.alert('There is no organisation named on this line, please select a row and try again')
    return
  }
  
  var notesSheet = activeSpreadsheet.getSheetByName(companyName)
  
  if (notesSheet === null) {
    throw new Error('Can not find notes sheet for "' + companyName + '"')
  }

  // Add the new note to the notes sheet

  notesSheet
    .insertRowBefore(2)
    .getRange(2, 1, 1, 2)
    .setValues([[Utils_.getDateString(), note]])
    
  // Update the formula in the "Organisations" sheet
  
  var dateColumnNumber = Utils_.getColumnNumber(activeSheet, DATE_HEADER_NAME)
  
  activeSheet
    .getRange(activeRowNumber, dateColumnNumber)
    .setFormula('\'' + companyName + '\'!A2')
    
  var noteColumnNumber = Utils_.getColumnNumber(activeSheet, NOTE_HEADER_NAME)
  
  activeSheet
    .getRange(activeRowNumber, noteColumnNumber)
    .setFormula('\'' + companyName + '\'!B2')
    
  SpreadsheetApp.flush()
    
} // onAddNote_() 
