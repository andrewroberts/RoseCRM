// Tests.gs
// ========
//
// Code for internal/unit testing

function test_init() {
  Log_ = BBLog.getLog({
    sheetId:              TEST_SHEET_ID_,
    level:                BBLog.Level.FINE, 
    displayFunctionNames: BBLog.DisplayFunctionNames.NO,
  })  
}

function test_misc() {
  const EXCLUDE = ['Clients', 'Log', 'Config', 'NotesTemplate']
  const ss = SpreadsheetApp.getActive()
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName()
    if (EXCLUDE.includes(name)) return
    ss.deleteSheet(sheet)
    // const a = 1
  })
}

function test_onCreateNda() {
  test_init()
  onCreateNda_()
}

/*
  namedValues={Status=[], Address=[], Rate=[], Contact Title=[], Timesheet=[], Journal=[], Folder=[], Timestamp=[09/07/2016 18:42:29], Trello Board=[], Company Name=[], Contact Email=[], Contact First Name=[], Contact Last Name=[], Contract=[], Notes=[]}, 
  range={columnEnd=2, columnStart=0, rowStart=11, rowEnd=12}
  source=Spreadsheet, 
*/									

function test_onFormSubmit() {
  const ss = SpreadsheetApp.getActive()
  const event = {
    namedValues: {
      Status               : ['1 - ACTION'], 
      Address              : [''], 
      Rate                 : [''], 
      'Contact Title'      : [''], 
      Timesheet            : [''], 
      Journal              : [''],
      Folder               : [''], 
      Timestamp            : [new Date], 
      'Trello Board'       : [''], 
      'Company Name'       : ['TestComp1827'], 
      'Email'              : ['a@b.com'], 
      'Name'               : ['FN1 LN1'], 
      Contract             : [''], 
      Pipeline             : ['1 - CONTRACT'],
      Notes                : [''],
      Source               : ['Upwork/Another jobsite'],
    },
    range: ss.getSheetByName('Clients').getRange('A160'),
    source: ss
  }
  
  onFormSubmit_(event)
}

function test_createNotesSheet() {

  var spreadsheet = SpreadsheetApp.openById('1Ylb95IeopzCaZvMKIjeQ7oHlqi1TLwAkgrb--GR8u3I')
  var companyName = '1849'
  var responseSheet = spreadsheet.getSheetByName('Clients')

  createNotesSheet()

  // Private Functions
  // -----------------

  function createNotesSheet() {
  
    // SpreadsheetApp.getActive().duplicateActiveSheet().setName(name).activate()
    
    spreadsheet.getSheetByName(NOTES_TEMPLATE_SHEET_NAME).activate()
    spreadsheet.duplicateActiveSheet().setName(companyName)
    spreadsheet.moveActiveSheet(2)
    responseSheet.activate()
    Log_.info('Created new notes sheet for "%s"', companyName)
    
  } // onFormSubmit_.createNotesSheet() 
}

