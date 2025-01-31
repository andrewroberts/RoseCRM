// 34567890123456789012345678901234567890123456789012345678901234567890123456789

// JSHint - TODO
/* jshint asi: true */

(function() {"use strict"})()

// Triggers_.gs
// ==============
//
// Manage script's triggers

var Triggers_ = (function(ns) {
 
  ns.setup = function() {
    const trigger = Triggers_.getTrigger('TRIGGER_SCRIPT_NAME_')    
    if (trigger === null) {
      Triggers_.createTrigger()        
    } else {
      Triggers_.deleteTrigger()        
    }    
  } 

  ns.createTrigger = function() {
    let trigger = Triggers_.getTrigger()
    if (trigger !== null) throw new Error('Trying to create a trigger when there is already one: ' + trigger.getUinqueId())
    trigger = ScriptApp.newTrigger(TRIGGER_SCRIPT_NAME_).forSpreadsheet(SpreadsheetApp.getActive()).onFormSubmit().create()
    const triggerId = trigger.getUniqueId()        
    Properties_.setProperty('AUTOMATIC_TRIGGER_ID', triggerId)
    Utils_.alert('Setup complete, and new "on form submit" trigger created.', 'Manage Trigger')
    Log_.info('Created  new "on form submit" trigger ' + triggerId)
    onOpen() 
  }

  ns.getTrigger = function() {
  
    let trigger = null
    
    ScriptApp.getProjectTriggers().forEach(function(nextTrigger) {
      if (nextTrigger.getHandlerFunction() === TRIGGER_SCRIPT_NAME_) {
        if (trigger !== null) {throw new Error(`Multiple ${TRIGGER_SCRIPT_NAME_} triggers`)}
        trigger = nextTrigger
        Log_.fine('Found trigger; ' + trigger.getUniqueId())
      }
    })

    if (trigger === null && !!Properties_.getProperty('AUTOMATIC_TRIGGER_ID')) {
      throw new Error('Trigger ID stored, but no trigger')
    }
    
    return trigger
  }

  ns.deleteTrigger = function() {  
  
    const trigger = Triggers_.getTrigger()
    
    if (trigger === null) {
      throw new Error('Trying to delete a trigger when there is not one')
    }    
    
    ScriptApp.deleteTrigger(trigger)
    Properties_.deleteProperty('AUTOMATIC_TRIGGER_ID')
    Log_.info('Deleted trigger: ' + trigger.getUniqueId())
    Utils_.alert('"On form submit" trigger removed.', 'Manage Trigger') 
    onOpen()    
  }

  ns.isTriggerCreated = function() {
    return !!Properties_.getProperty('AUTOMATIC_TRIGGER_ID')
  }

  return ns

})(Triggers_ || {})
