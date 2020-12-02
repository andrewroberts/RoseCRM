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

// Code review all files - TODO
// JSHint review (see files) - TODO
// Unit Tests - TODO
// System Test (Dev) - TODO
// System Test (Prod) - TODO

// Config.gs
// =========
//
// All the constans and configuration settings

// Configuration
// =============

var SCRIPT_NAME = "Organisations"
var SCRIPT_VERSION = "v0.dev_ajr"

var PRODUCTION_VERSION = true

var ORGS_FOLDER_ID = '0BxRtIprIrwuzVjRBWjJjd0U5ZUE'

var TIMESHEET_TEMPLATE_SHEET_ID = '1wA7TRlVu6_gMm-NGD8sErOvmpQdDIfsu3aChoDmQJVM'

var TIMESHEETS_FOLDER_ID = '0BxRtIprIrwuzM0c1RDVOMEFmUlk'

var CONTRACT_TEMPLATE_SHEET_ID = '1Vja451VZ_RbCLXxUI-WC2YnQr0jpGV4DgBqf0rvKnxI'
var CONTRACT_FOLDER_ID = '0BxRtIprIrwuzV01NWG1MT2FpWlU'

var NDA_TEMPLATE_SHEET_ID = '1qw1bvHFsr5usKMEHQvZSbiVV2DVtrGKGaB2aBynqVms'

var ENABLE_CONTRACT_CREATION = true

// Log Library
// -----------

var LOG_LEVEL = PRODUCTION_VERSION ? Log.Level.INFO : Log.Level.ALL
var LOG_SHEET_ID = ''
var LOG_DISPLAY_FUNCTION_NAMES = Log.DisplayFunctionNames.YES

// Assert library
// --------------

var SEND_ERROR_EMAIL = false
var HANDLE_ERROR = Assert.HandleError.THROW
var ADMIN_EMAIL_ADDRESS = 'andrewr1969@gmail.com'

// Constants/Enums
// ===============

var COMPANY_NAME_HEADER_NAME   = 'Company Name'
var STATUS_HEADER_NAME         = 'Status'
var PIPELINE_HEADER_NAME       = 'Pipeline'
var COMPANY_FOLDER_HEADER_NAME = 'Folder'
var TIMESHEET_HEADER_NAME      = 'Timesheet'
var JOURNAL_HEADER_NAME        = 'Journal'
var CONTRACT_HEADER_NAME       = 'Contract'
var NDA_HEADER_NAME            = 'NDA'
var DATE_HEADER_NAME           = 'Date'
var NOTE_HEADER_NAME           = 'Latest Note'

var STATUS_COLOURS = Object.freeze({
  '1 - ACTION':  '#d9ead3', // light green 3
  '2 - WAITING': '#fff2cc', // light yellow 3
  '3 - CLOSED':  '#cfe2f3', // light blue 3	
})

var PIPELINE_COLOURS = Object.freeze({
  '1 - CONTRACT':      '#d9ead3', // light green 3
  '2 - PROPOSAL_SENT': '#fff2cc', // light yellow 3
  '3 - ENQUIRY':       '#f4cccc', // light red 3
  '4 - CLOSED':        '#cfe2f3', // light blue 3	
})

var COMPANY_NAME = 'Company Name'
var COMPANY_ADDRESS_NAME = 'Address'
var RATE_NAME = 'Rate'

var ORGANISATIONS_SHEET_NAME = 'Orgs'
var NOTES_TEMPLATE_SHEET_NAME = 'NotesTemplate'

// Function Template
// -----------------

/**
 *
 *
 * @param {object} 
 *
 * @return {object}
 */
 
function functionTemplate() {

  Log.functionEntryPoint()
  
  

} // functionTemplate() 
