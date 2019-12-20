function onEdit(e){
  var range = e.range;
  var columnEdited = range.getColumn();
  var rowEdited = range.getRow();
  var ss = SpreadsheetApp.getActive();
  var sheetName = ss.getActiveSheet().getName()
  
  
  
  // Check for stage or status change in New/Warm Leads tab 
  if (sheetName === 'New/Warm Leads' && (columnEdited === 7 || columnEdited === 12)){
    var cell = range.getA1Notation()
    var val = ss.getRange(cell).getValue()
    
    // Moving from New/Warm Lead to Opportunity 
    if ((val === 'Searching' || val === 'Touring' || val === 'Offering' || val === 'UC') && columnEdited === 7){
      moveToOpp(rowEdited)
    }
    
    // Moving from New/Warm Lead to Cold Lead 
    else if (val === 'Cold Lead' && columnEdited === 7){
      moveToCold(rowEdited)
    }
    
    // Moving from New/Warm Lead to Archive 
    else if ((val === 'Closed' && columnEdited === 7) || ((val === 'Lost' || val === 'Abandoned') && columnEdited === 12)){
      archive(rowEdited)
    }
  }
  
  // Check for stage or status change in Cold Leads tab
  else if (sheetName === 'Cold Leads' && (columnEdited === 7 || columnEdited === 12)){
    var cell = range.getA1Notation()
    var val = ss.getRange(cell).getValue()
    
    // Moving from Cold Lead to Opportunity 
    if ((val === 'Searching' || val === 'Touring' || val === 'Offering' || val === 'UC') && columnEdited === 7){
      moveToOpp(rowEdited)
    }
    
    // Moving from Cold Lead to New/Warm Lead 
    else if ((val === 'Warm Lead' || val === 'New Lead') && columnEdited === 7){
      moveToWarm(rowEdited)
    }
    
    // Moving from Cold Lead to Archive  
    else if ((val === 'Closed' && columnEdited === 7) || ((val === 'Lost' || val === 'Abandoned') && columnEdited === 12)){
      archive(rowEdited)
    }
  }
  
  // Check for stage or status change in Opportunity tab
  else if (sheetName === 'Opportunities' && (columnEdited === 7 || columnEdited === 12)){
    var cell = range.getA1Notation()
    var val = ss.getRange(cell).getValue()
    
    // Moving from Opportunity to Cold Lead 
    if (val === 'Cold Lead' && columnEdited === 7){
      moveToCold(rowEdited)
    }
    
    // Moving from Opportunity to New/Warm Lead 
    else if ((val === 'Warm Lead' || val === 'New Lead') && columnEdited === 7){
      moveToWarm(rowEdited)
    }
    
    // Moving from Opportunity to Archive 
    else if ((val === 'Closed' && columnEdited === 7) || ((val === 'Lost' || val === 'Abandoned') && columnEdited === 12)){
      archive(rowEdited)
    }
  }
  
  // Check for status change in Archive tab
  else if (sheetName === 'Archive' && (columnEdited === 7 || columnEdited === 12)){
    var cell = range.getA1Notation()
    var val = ss.getRange(cell).getValue()
    
    // Moving from Archive by changing stage with Open status
    if (columnEdited === 7 && ss.getRange("L"+rowEdited).getValue() === 'Open'){
      if (val === 'Searching' || val === 'Touring' || val === 'Offering' || val === 'UC'){
        ss.getRange(cell).setBackground(null)
        moveToOpp(rowEdited)
      }
      else if (val === 'Warm Lead' || val === 'New Lead'){
        ss.getRange(cell).setBackground(null)
        moveToWarm(rowEdited)
      }
      else if (val === 'Cold Lead'){
        ss.getRange(cell).setBackground(null)
        moveToCold(rowEdited)
      }
    }
    
    // Moving from Archive by changing status
    else if (val === 'Open' && columnEdited === 12){
      
      var stageCell = "G"+rowEdited+""
      var stage = ss.getRange(stageCell).getValue()
      ss.getRange("L"+rowEdited+"").setBackground(null);
      
      // Moving from Archive to New/Warm Lead 
      if (stage === 'New Lead' || stage === 'Warm Lead'){
        moveToWarm(rowEdited)
      }
      
      // Moving from Archive to Cold Lead 
      else if (stage === 'Cold Lead'){
        moveToCold(rowEdited)
      }
      
      // Moving from Archive to Opportunity 
      else if (stage === 'Searching' || stage === 'Touring' || stage === 'Offering' || stage === 'UC'){
        moveToOpp(rowEdited)
      }
    }
  }
  
//  // Create Calendar event on date input
//  else if ((columnEdited === 23 || columnEdited === 24 || columnEdited === 25) && (sheetName === 'Opportunities' || sheetName === 'New/Warm Leads' || sheetName === 'Cold Leads' || sheetName === 'Archive')){
//    var cell = range.getA1Notation()
//    var val = ss.getRange(cell).getValue()
//    var eventName = 'Default'
//    
//    ss.getRange('Y14').setValue('')
//    
//    if (columnEdited === 23){
//      eventName = ss.getRange("A"+rowEdited).getValue() + ' - Due Diligence Deadline'
//    }
//    else if (columnEdited === 24){
//      eventName = ss.getRange("A"+rowEdited).getValue() + ' - F&A Deadline'
//    } 
//    else if (columnEdited === 25){
//      eventName = ss.getRange("A"+rowEdited).getValue() + ' - Settlement Deadline'
//    }
//    
//    ss.getRange('Y14').setValue('running')
////    var calendars = CalendarApp.getAllCalendars()
////    ss.getRange('Y15').setValue(calendars[0].getName())
//    var event = CalendarApp.getCalendarById('https://calendar.google.com/calendar/b/3?cid=bWlrZS5kZWdyb290QGhvbWllLmNvbQ').createAllDayEvent('test', new Date('December 25, 2019'),{location: ''})
//    ss.getRange('Y16').setValue(event.getId())
//    ss.getRange('Y14').setValue('done')
//  }
  
}

function moveToOpp(rowEdited){
  var ss = SpreadsheetApp.getActive();
  ss.getSheetByName('Opportunities').insertRowsBefore(4,1)
  var range = "A"+rowEdited+":AB"+rowEdited+""
  ss.getRange(range).copyTo(ss.getSheetByName('Opportunities').getRange('A4:AB4'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false)
  ss.deleteRows(rowEdited, 1)
}

function moveToWarm(rowEdited){
  var ss = SpreadsheetApp.getActive();
  ss.getSheetByName('New/Warm Leads').insertRowsBefore(4,1)
  var range = "A"+rowEdited+":AB"+rowEdited+""
  ss.getRange(range).copyTo(ss.getSheetByName('New/Warm Leads').getRange('A4:AB4'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false)
  ss.deleteRows(rowEdited, 1)
}

function moveToCold(rowEdited){
  var ss = SpreadsheetApp.getActive();
  ss.getSheetByName('Cold Leads').insertRowsBefore(4,1)
  var range = "A"+rowEdited+":AB"+rowEdited+""
  ss.getRange(range).copyTo(ss.getSheetByName('Cold Leads').getRange('A4:AB4'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false)
  ss.deleteRows(rowEdited, 1)
}

function archive(rowEdited){
  var ss = SpreadsheetApp.getActive();
  ss.getSheetByName('Archive').insertRowsBefore(4,1)
  var range = "A"+rowEdited+":AB"+rowEdited+""
  ss.getRange(range).copyTo(ss.getSheetByName('Archive').getRange('A4:AB4'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false)
  ss.deleteRows(rowEdited, 1)
}

function updateCalendar(){
  var ss = SpreadsheetApp.getActive()
  var calendar = CalendarApp.getCalendarById('mike.degroot@homie.com')
  
  var val = ss.getRange('W7').getValue()
  //  var eventName = 'Default'
  ss.getRange('Y14').setValue('')
  ss.getRange('Y14').setValue('running')
  
  var id = getEventByName('test #2')
  
  calendar.getEventById(id).deleteEvent()  
  ss.getRange('Y17').setValue('done')

  //  var event = CalendarApp.getCalendarById('mike.degroot@homie.com').createAllDayEvent('test #2', new Date('December 25, 2019'),{location: ''})
  //  ss.getRange('Y16').setValue(event.getId())
}

function getEventByName(name){
  var calendar = CalendarApp.getCalendarById('mike.degroot@homie.com')
  var events = calendar.getEventsForDay(new Date('December 25, 2019'))
  var title = ''
  
  for (var i = 0; i < events.length; i++){
    title = events[i].getTitle()
    if (title === name){
      return events[i].getId()
    }
  }
}