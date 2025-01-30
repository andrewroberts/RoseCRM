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

const SCRIPT_NAME = "RoseCRM"
const SCRIPT_VERSION = "v1.0"

const PRODUCTION_VERSION = true

const TEST_SHEET_ID_ = '1Ylb95IeopzCaZvMKIjeQ7oHlqi1TLwAkgrb--GR8u3I'

// Log Library
// -----------

const DEBUG_LOG_LEVEL_ = PRODUCTION_VERSION ? BBLog.Level.INFO : BBLog.Level.FINE
const DEBUG_LOG_SHEET_ID_ = ''
const DEBUG_LOG_DISPLAY_FUNCTION_NAMES_ = BBLog.DisplayFunctionNames.NO

// Assert library
// --------------

// const SEND_ERROR_EMAIL_ = PRODUCTION_VERSION ? true : false
const SEND_ERROR_EMAIL_ = false // To restrict required scopes this is disabled
const HANDLE_ERROR_ = Assert.HandleError.THROW
const ADMIN_EMAIL_ADDRESS = ''

// Constants/Enums
// ===============

const TRIGGER_SCRIPT_NAME_ = 'onFormSubmit'

