// 34567890123456789012345678901234567890123456789012345678901234567890123456789

// Tests.gs
// ========
//
// Code for internal/unit testing

function test_createNotesSheet() {

  var spreadsheet = SpreadsheetApp.openById('1Ylb95IeopzCaZvMKIjeQ7oHlqi1TLwAkgrb--GR8u3I')
  var companyName = '1849'
  var responseSheet = spreadsheet.getSheetByName('Organisations')

  createNotesSheet()

  // Private Functions
  // -----------------

  function createNotesSheet() {
  
    Log.functionEntryPoint()
    
    // SpreadsheetApp.getActive().duplicateActiveSheet().setName(name).activate()
    
    spreadsheet.getSheetByName(NOTES_TEMPLATE_SHEET_NAME).activate()
    spreadsheet.duplicateActiveSheet().setName(companyName)
    spreadsheet.moveActiveSheet(2)
    responseSheet.activate()
    Log.info('Created new notes sheet for "%s"', companyName)
    
  } // onFormSubmit_.createNotesSheet() 
}

function test_misc() {

  var fileId = SpreadsheetApp
  .openById(TIMESHEET_TEMPLATE_SHEET_ID)
  .copy('Timesheet - OPEN - ' + 'TEST1')
  .getId()
  
   var file = DriveApp.getFileById(fileId)
   file.moveTo(DriveApp.getFolderById('0BxRtIprIrwuzV3FwWFF2bWRHcHM'))
    
   DriveApp.getFolderById(TIMESHEETS_FOLDER_ID).createShortcut(fileId)
}