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

// @OnlyCurrentDoc

// RoseCRM.gs
// ================
//
// External interface to this script - all of the event handlers.
//
// This files contains all of the event handlers, plus miscellaneous functions 
// not worthy of their own files yet

let Log_ = null
const Properties_ = PropertiesService.getScriptProperties() 

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

const EVENT_HANDLERS = {

//                         Name                          onError Message                        Main Functionality
//                         ----                          ---------------                        ------------------

  onSetup:                 ['onSetup()',                'Failed to setup trigger',             onSetup_],
  onFormSubmit:            ['onFormSubmit()',            'Failed to process form submit',       onFormSubmit_],
//  onEdit:                  [function() {},  'onEdit()',                  'Failed to process edit',              onEdit_],  
  onCreateContract:        ['onCreateContract()',        'Failed to create contract',           onCreateContract_],  
  onCreateNda:             ['onCreateNda()',             'Failed to create NDA',                onCreateNda_],    
  onAddNote:               ['onAddNote()',               'Failed to add note',                  onAddNote_],    
}

function onSetup(arg)                    {return eventHandler_(EVENT_HANDLERS.onSetup, arg)}
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
  
  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu('RoseCRM')

  menu
    .addItem('Create NDA', 'onCreateNda')  
    .addItem('Create contract', 'onCreateContract')
    .addItem('Add new note...', 'onAddNote')  
    .addSeparator()  

  if (Triggers_.isTriggerCreated()) {
    menu.addItem('Disable automatic file & folder creation', "onSetup")
  } else {
    menu.addItem('RUN SETUP!', "onSetup")
  }

  menu.addToUi()
} 

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

function eventHandler_(config, args) {

  const userEmail = Session.getActiveUser().getEmail()

  try {

    Log_ = BBLog.getLog({
      level:                DEBUG_LOG_LEVEL_, 
    })

    Log_.info('Handling ' + config[0] + ' from ' + (userEmail || 'unknown email') + ' (' + SCRIPT_NAME + ' ' + SCRIPT_VERSION + ')')
    
    // Call the main function
    return config[2](args)
    
  } catch (error) {
  
    console.log(error.message)

    const assertConfig = {
      error:          error,
      userMessage:    config[1],
      log:            Log_,
      handleError:    HANDLE_ERROR_, 
      sendErrorEmail: SEND_ERROR_EMAIL_, 
      // emailAddress:   userEmail || ADMIN_EMAIL_ADDRESS_,
      scriptName:     SCRIPT_NAME,
      scriptVersion:  SCRIPT_VERSION, 
    }

    Assert.handleError(assertConfig) 

    Utils_.alert(error.message)
  }
  
} // eventHandler_()

// Private event handlers
// ----------------------

function onSetup_() {Triggers_.setup()}

/**
 *
 */

function onFormSubmit_(event) {

/*
  Logger.log(event)

  const range = event.range
  for (const key in range) {
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

  const values = event.namedValues
  const config = Utils_.getConfig()
  const companyName = values[config.HEADER_CLIENT_COMPANY_NAME][0]
  const contactName = values[config.HEADER_CONTACT_NAME][0]
  const orgsFolder = DriveApp.getFolderById(config.CLIENTS_ROOT_FOLDER_ID)
  const spreadsheet = event.source
  const responseRange = event.range
  const responseSheet = responseRange.getSheet()
  const rowNumber = responseRange.getRow()

  responseSheet.getRange('B:H').setFontWeight('bold')

  const companyFolderName = companyName || contactName 

  const companyFolders = DriveApp.getFoldersByName(companyFolderName)
  let companyFolder
  
  if (!companyFolders.hasNext()) {  
  
    companyFolder = orgsFolder.createFolder(companyFolderName)
    const folderUrl = companyFolder.getUrl()
    storeValue(config.HEADER_CLIENT_FOLDER, folderUrl)
    Log_.info('Created folder: ' + companyFolderName + ' (' + folderUrl + ')')
    
  } else {
  
    companyFolder = companyFolders.next()
    
    if (companyFolders.hasNext()) {
      throw new Error('Found two folders for this client')
    }
    
    Log_.info('Using existing folder: "' + companyFolderName + '" - assuming all files already made')
    return
  }

  const timesheetFileUrl = createTimesheet()
  storeValue(config.HEADER_TIMESHEET, timesheetFileUrl)
  Log_.info('Created timesheet: ' + timesheetFileUrl)
  
  const journalFileUrl = createJournal()
  storeValue(config.HEADER_JOURNAL, journalFileUrl)
  Log_.info('Created journal: ' + journalFileUrl)

  // Fill in Company Name column if it is empty
  
  if (companyName === '') {
    storeValue(config.HEADER_CLIENT_COMPANY_NAME, contactName)
  }

  createNotesSheet()
  createGMailLabel()
  Log_.info(`Form submission processed OK!`)
  return

  // Private Functions
  // -----------------

  function createGMailLabel() {
    const parent = config.YOUR_COMPANY_GMAIL_PARENT_LABEL || ''
    GmailApp.createLabel(parent + companyName)
    Log_.info('Created GMail label: ' + journalFileUrl)  
  }

  /**
   * Store a value in the response sheet
   *
   * @param {String} headerName
   * @param {Object} value
   *
   * @return {object}
   */
   
  function storeValue(headerName, value) {
    Utils_.storeValue(responseSheet, rowNumber, headerName, value) 
  }

  function createTimesheet() {
  
    if (!config.TIMESHEET_TEMPLATE_SHEET_ID) {
      Log_.info(`No timesheet template provided.`)
      return
    }

    const timesheetSpreadsheet = SpreadsheetApp
      .openById(config.TIMESHEET_TEMPLATE_SHEET_ID)
      .copy('Timesheet - OPEN - ' + companyFolderName)
    
    const timesheetFolder = DriveApp.getFolderById(config.TIMESHEETS_FOLDER_ID)
    DriveApp.getFileById(timesheetSpreadsheet.getId()).moveTo(timesheetFolder)      
    companyFolder.createShortcut(timesheetSpreadsheet.getId())
    return timesheetSpreadsheet.getUrl() 
  }
   
  function createJournal() {
    const journalGDoc = DocumentApp.create('Journal - ' + companyFolderName)
    DriveApp.getFileById(journalGDoc.getId()).moveTo(companyFolder)
    return journalGDoc.getUrl()
  }
  
  /**
   * Create a new notes sheet
   */
   
  function createNotesSheet() {
    spreadsheet.getSheetByName(config.NOTES_TEMPLATE_SHEET_NAME).activate()
    try {
      spreadsheet.duplicateActiveSheet().setName(companyName)
    } catch (error) {
      Log_.warning(`There is already tab called ${companyName}`)
      return       
    }
    spreadsheet.moveActiveSheet(2) // After the clients sheet
    responseSheet.activate()
    Log_.info('Created new notes sheet for "%s"', companyName)
  }
  
} // onFormSubmit_()

/**
 * Private event handler for "on contract create" event
 */
 
function onCreateContract_() {

  // Set up the docs and the spreadsheet access
  
  const ui = SpreadsheetApp.getUi()
  const config = Utils_.getConfig()
  const copyFile = DriveApp.getFileById(config.CONTRACT_TEMPLATE_SHEET_ID).makeCopy()
  const copyId = copyFile.getId()
  const copyDoc = DocumentApp.openById(copyId)  
  const copyBody = copyDoc.getBody()
  const activeSheet = SpreadsheetApp.getActiveSheet()
  
  if (activeSheet.getName() !== config.CLIENTS_SHEET_NAME) {
    ui.alert('Only works on "Clients" sheet.')
    return
  }
    
  const numberOfColumns = activeSheet.getLastColumn()
  Log_.fine('numberOfColumns: ' + numberOfColumns)
  const activeRange = activeSheet.getActiveRange()
  
  if (activeRange === null) {
    ui.alert('Select a row before clicking the menu.')
    return
  }
  
  const activeRowNumber = activeRange.getRow()
  
  if (activeRowNumber > activeSheet.getLastRow()) {
    ui.alert('This row is empty, select a completed one.')
    return
  }

  if (activeRowNumber === 1) {
    ui.alert('You have selected the header row, select an client\'s.')
    return
  }
    
  const activeRow = activeSheet.getRange(activeRowNumber, 1, 1, numberOfColumns).getValues()[0]
  const headerRow = activeSheet.getRange(1, 1, 1, numberOfColumns).getValues()[0]
  
  let companyName = null
  let companyFolderUrl = null
  let rate = null
  let address = null
  
  // Replace the keys with the spreadsheet values
 
  for (let columnIndex = 0; columnIndex < numberOfColumns; columnIndex++) {

    const header = headerRow[columnIndex]
    const value = activeRow[columnIndex]

    Log_.finest('header: ' + header)
    Log_.finest('value: ' + value)

    if (header === config.HEADER_CLIENT_COMPANY_NAME) {
    
      companyName = value
      Log_.fine('Found company name')
      
    } else if (header === config.HEADER_CLIENT_FOLDER) {
    
      companyFolderUrl = value
      Log_.fine('Found company folder URL')
            
    } else if (header === config.HEADER_RATE) {
    
      rate = value
      Log_.fine('Found rate')
            
    } else if (header === config.HEADER_CLIENT_ADDRESS) {
    
      address = value
      Log_.fine('Found company address')      
    }

    const nextPlaceholder = '(?i){{' + header.replace(/[^a-z0-9\s]/gi, ".") + '}}'
    copyBody.replaceText(nextPlaceholder, value)
    
  } // For each column
  
  // Specials
  
  copyBody.replaceText('{{YOUR_COMPANY_NAME}}', config.YOUR_COMPANY_NAME)    
  copyBody.replaceText('{{YOUR_COMPANY_ADDRESS}}', config.YOUR_COMPANY_ADDRESS)    
  copyBody.replaceText('{{YOUR_COMPANY_SERVICE}}', config.YOUR_COMPANY_SERVICE)    
  copyBody.replaceText('{{YOUR_NAME}}', config.YOUR_NAME)    

  const timeZone = Session.getScriptTimeZone()
  const date = new Date()
  const dateString = Utilities.formatDate(date, timeZone, config.DATE_FORMAT)
  Log_.fine('dateString: ' + dateString)
  copyBody.replaceText('{{Contract Date}}', dateString)
    
  // Check we found all the info we need
  
  // if (companyName === null || companyName === '') {
  
  //   ui.alert('Could not find the company name')
  //   return
    
  // } else if (companyFolderUrl === null || companyFolderUrl === '') {
  
  //   ui.alert('Could not find the company folder URL')
  //   return
    
  // } else if (rate === null || rate === '') {
  
  //   ui.alert('Could not find the rate')
  //   return
    
  // } else if (address === null || address === '') {
  
  //   ui.alert('Could not find the company address')
  //   return
  // }
    
  // Create the PDF file, rename it if required and delete the doc copy
    
  copyDoc.saveAndClose()
  
  const contractPdfFile = DriveApp.createFile(copyFile.getAs('application/pdf'))  
  
  copyFile.setTrashed(true)
  
  const contractName = `${config.YOUR_COMPANY_NAME} Contract - ${companyName}`
  contractPdfFile.setName(contractName)
  Log_.info('Created new contract "' + contractName +'"')

  // Put the new contract in the contract folder, and put a shortcut in the company folder

  contractPdfFile.setName(contractName)
  const ndaFolder = DriveApp.getFolderById(config.CONTRACT_FOLDER_ID)
  contractPdfFile.moveTo(ndaFolder)
  Log_.info('Created new Contract "' + contractName +'" in ' + ndaFolder.getName())

  let companyFolderId = Utils_.getId(companyFolderUrl)
  Log_.fine('companyFolderId: ' + companyFolderId)

  const companyFolder = DriveApp.getFolderById(companyFolderId)
  companyFolder.createShortcut(contractPdfFile.getId())
  Log_.info('Put shortcut to Contract into company folder "' + companyFolder.getName() + '"')
 
  Utils_.storeValue(
    activeSheet, 
    activeRowNumber, 
    config.HEADER_CONTRACT, 
    contractPdfFile.getUrl())
     
  ui.alert('New contract "' + contractName + '" created in the folder "' + companyFolder.getName() + '"')  

} // onCreateContract_() 

/**
 * Private event handler for "on NDA create" event
 */
 
function onCreateNda_() {

  // Set up the docs and the spreadsheet access
  
  const ui = SpreadsheetApp.getUi()
  const config = Utils_.getConfig()
  const copyFile = DriveApp.getFileById(config.NDA_TEMPLATE_SHEET_ID).makeCopy()
  const copyId = copyFile.getId()
  const copyDoc = DocumentApp.openById(copyId)  
  const copyBody = copyDoc.getBody()
  const activeSheet = SpreadsheetApp.getActiveSheet()
  
  if (activeSheet.getName() !== config.CLIENTS_SHEET_NAME) {
    ui.alert(`Only works on ${config.CLIENTS_SHEET_NAME} sheet.`)
    return
  }
    
  const numberOfColumns = activeSheet.getLastColumn()
  Log_.fine('numberOfColumns: ' + numberOfColumns)
  const activeRange = activeSheet.getActiveRange()
  
  if (activeRange === null) {
    ui.alert('Select a row before clicking the menu.')
    return
  }
  
  const activeRowNumber = activeRange.getRow()
  
  if (activeRowNumber > activeSheet.getLastRow()) {
    ui.alert('This row is empty, select a completed one.')
    return
  }

  if (activeRowNumber === 1) {
    ui.alert('You have selected the header row, select an client\'s.')
    return
  }
    
  const activeRow = activeSheet.getRange(activeRowNumber, 1, 1, numberOfColumns).getValues()[0]
  const headerRow = activeSheet.getRange(1, 1, 1, numberOfColumns).getValues()[0]
  
  let companyName = null
  let companyFolderUrl = null
  let rate = null
  let address = null
  
  // Replace the keys with the spreadsheet values
 
  for (let columnIndex = 0; columnIndex < numberOfColumns; columnIndex++) {

    const header = headerRow[columnIndex]
    const value = activeRow[columnIndex]

    Log_.finest('header: ' + header)
    Log_.finest('value: ' + value)

    if (header === config.HEADER_CLIENT_COMPANY_NAME) {
    
      companyName = value
      Log_.fine('Found client company name')
      
    } else if (header === config.HEADER_CLIENT_FOLDER) {
    
      companyFolderUrl = value
      Log_.fine('Found company folder URL')
            
    } else if (header === config.HEADER_RATE) {
    
      rate = value
      Log_.fine('Found rate')
            
    } else if (header === config.HEADER_CLIENT_ADDRESS) {
    
      address = value
      Log_.fine('Found company address')      
    }

    const nextPlaceholder = '(?i){{' + header.replace(/[^a-z0-9\s]/gi, ".") + '}}'
    copyBody.replaceText(nextPlaceholder, value)
    
  } // For each column

  // Specials

  copyBody.replaceText('{{YOUR_COMPANY_NAME}}', config.YOUR_COMPANY_NAME)    
  copyBody.replaceText('{{YOUR_COMPANY_ADDRESS}}', config.YOUR_COMPANY_ADDRESS)    
  copyBody.replaceText('{{YOUR_COMPANY_SERVICE}}', config.YOUR_COMPANY_SERVICE)    
  copyBody.replaceText('{{YOUR_NAME}}', config.YOUR_NAME)    
  
  const timeZone = Session.getScriptTimeZone()
  const date = new Date()
  const dateString = Utilities.formatDate(date, timeZone, config.DATE_FORMAT)
  Log_.fine('dateString: ' + dateString)
  copyBody.replaceText('{{Contract Date}}', dateString)
    
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
  
  const ndaPdfFile = DriveApp.createFile(copyFile.getAs('application/pdf'))  
  
  copyFile.setTrashed(true)

  // Put the new nda in the NDA folder, and put a shortcut in the company folder

  const ndaName = `${config.YOUR_COMPANY_NAME} NDA - ${companyName}`
  ndaPdfFile.setName(ndaName)
  const ndaFolder = DriveApp.getFolderById(config.CONTRACT_FOLDER_ID)
  ndaPdfFile.moveTo(ndaFolder)
  Log_.info('Created new NDA "' + ndaName +'" in ' + ndaFolder.getName())

  const companyFolderId = Utils_.getId(companyFolderUrl)
  Log_.fine('companyFolderId: ' + companyFolderId)
  const companyFolder = DriveApp.getFolderById(companyFolderId)
  companyFolder.createShortcut(ndaPdfFile.getId())
  Log_.info('Put shortcut to NDA into company folder "' + companyFolder.getName() + '"')
 
  Utils_.storeValue(
    activeSheet, 
    activeRowNumber, 
    config.HEADER_NDA, 
    ndaPdfFile.getUrl())
     
  ui.alert('New NDA "' + ndaName + '" created in the folder "' + companyFolder.getName() + '"')  

} // onCreateNda_() 

/**
 * Add a note to the clients sheet
 *
 * @param {object} 
 *
 * @return {object}
 */
 
function onAddNote_() {
  
  // Get the active cell first, to check it is valid
  
  const ui = SpreadsheetApp.getUi()  
  const config = Utils_.getConfig()
  const activeCell = Utils_.getActiveCellObject(ui)
  
  if (activeCell === null) {
    // The Error has been logged in Utils_.getActiveCellObject()
    return
  }
  
  const activeSheet = activeCell.sheet
  const activeSpreadsheet = activeCell.spreadsheet
  const activeRowNumber = activeCell.rowNumber
  
  // Get the new note
  
  const response = ui.prompt('What\'s the note?', ui.ButtonSet.OK_CANCEL)
  let note = ''
    
  if (response.getSelectedButton() == ui.Button.OK) {
  
    note = response.getResponseText()
    
  } else if (response.getSelectedButton() == ui.Button.CANCEL) {
 
    Log_.warning('The user didn\'t want to provide a note.')
    return
   
  } else {
 
    Log_.warning('The user clicked the close button in the dialog\'s title bar.')
    return
  }

  // Get the notes sheet for this client

  const companyNameColumnNumber = Utils_.getColumnNumber(
    activeSheet,
    config.HEADER_CLIENT_COMPANY_NAME)
    
  const companyName = activeSheet
    .getRange(
      activeRowNumber, 
      companyNameColumnNumber)
    .getValue()
    
  Log_.fine('companyName: ' + companyName)
  
  if (companyName === '') {
    Log_.warning('User selected row with no client name')
    ui.alert('There is no client named on this line, please select a row and try again')
    return
  }
  
  const notesSheet = activeSpreadsheet.getSheetByName(companyName)
  
  if (notesSheet === null) {
    throw new Error('Can not find notes sheet for "' + companyName + '"')
  }

  // Add the new note to the notes sheet

  notesSheet
    .insertRowBefore(2)
    .getRange(2, 1, 1, 2)
    .setValues([[Utils_.getDateString(), note]])
    
  // Update the formula in the "Clients" sheet
  
  const dateColumnNumber = Utils_.getColumnNumber(activeSheet, config.HEADER_NOTE_DATE)
  
  activeSheet
    .getRange(activeRowNumber, dateColumnNumber)
    .setFormula('\'' + companyName + '\'!A2')
    
  const noteColumnNumber = Utils_.getColumnNumber(activeSheet, config.HEADER_NOTE)
  
  activeSheet
    .getRange(activeRowNumber, noteColumnNumber)
    .setFormula('\'' + companyName + '\'!B2')
    
  SpreadsheetApp.flush()
    
} // onAddNote_() 
