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
  var buyerName = ss.getRange('V2').getValue()
  var dueDiligenceDate = ss.getRange('W2').setNumberFormat('mmmm" "d", "yyyy').getValue()
  var financingDate = ss.getRange('X2').setNumberFormat('mmmm" "d", "yyyy').getValue()
  var settlementDate = ss.getRange('Y2').setNumberFormat('mmmm" "d", "yyyy').getValue()
  ss.getRange('W2:Y2').setNumberFormat('m"/"d"/"yy')
  
  // If no buyer is selected, throw an error
  if (!buyerName){
    return makeBuyerRed()
  }
  
  // If no dates are selected, throw an error
  else if (!dueDiligenceDate && !financingDate && !settlementDate){
    return makeDatesRed()
  }
  
  // If buyer and at least 1 date entered, delete old events and create new ones
  else {
  
    var email = ss.getSheetByName('Dashboard').getRange('B6').getValue()
    deleteCreateEvents(email)
    
    email = 'homie.com_1cs8eji9ahpmol4rvqllcq8bco@group.calendar.google.com'
    deleteCreateEvents(email)
    
    redoFormatting()
  }
}

function deleteCreateEvents(email){
  var calendar = CalendarApp.getCalendarById(email)
  var ss = SpreadsheetApp.getActive()
  var buyerName = ss.getRange('V2').getValue()
  var dueDiligenceDate = ss.getRange('W2').setNumberFormat('mmmm" "d", "yyyy').getValue()
  var financingDate = ss.getRange('X2').setNumberFormat('mmmm" "d", "yyyy').getValue()
  var settlementDate = ss.getRange('Y2').setNumberFormat('mmmm" "d", "yyyy').getValue()
  ss.getRange('W2:Y2').setNumberFormat('m"/"d"/"yy')
  var rowNum = ss.getRange('AA1').setFormula('=IFERROR(MATCH(V2,A:A,0),"")').getValue()
  var eventName = ''
  var newEvent = ''
   
  // Run if a new Due Diligence date is entered
  if (dueDiligenceDate){ 
    
    // If previous due diligence date exists, find event and delete
    var dueDiligenceOldDate = ss.getRange('W' + rowNum + '').getValue()
    if (dueDiligenceOldDate){
      dueDiligenceOldDate = ss.getRange('W' + rowNum + '').setNumberFormat('mmmm" "d", "yyyy').getValue()
      var dueDiligenceID = getIdFromName('' + buyerName + ' - Due Diligence Deadline', dueDiligenceOldDate)
      
      // ********************
      ss.getRange('W20').setValue(dueDiligenceID)
      // ********************
      
      if (calendar.getEventById(dueDiligenceID)){
        calendar.getEventById(dueDiligenceID).deleteEvent()
      }
    }
    
    // Create new event with new date
    eventName = '' + buyerName + ' - Due Diligence Deadline'
    newEvent = calendar.createAllDayEvent(eventName, new Date(dueDiligenceDate),{location: ''})
    ss.getRange('W' + rowNum + '').setValue(dueDiligenceDate).setNumberFormat('m"/"d"/"yy')
  }
  
  // Run if a new F&A date is entered
  if (financingDate){ 
    
    // If previous F&A exists, find event and delete
    var financingOldDate = ss.getRange('X' + rowNum + '').getValue()
    if (financingOldDate){
      financingOldDate = ss.getRange('X' + rowNum + '').setNumberFormat('mmmm" "d", "yyyy').getValue()
      var financingID = getIdFromName('' + buyerName + ' - F&A Deadline', financingOldDate)
      if (calendar.getEventById(financingID)){
        calendar.getEventById(financingID).deleteEvent()
      }
    }
      
    // Create new event with new date
    eventName = '' + buyerName + ' - F&A Deadline'
    newEvent = calendar.createAllDayEvent(eventName, new Date(financingDate),{location: ''})
    ss.getRange('X' + rowNum + '').setValue(financingDate).setNumberFormat('m"/"d"/"yy')
  }
  
  // Run if a new Settlement date is entered
  if (settlementDate){ 
    
    // If previous Settlement exists, find event and delete
    var settlementOldDate = ss.getRange('Y' + rowNum + '').getValue()
    if (settlementOldDate){
      settlementOldDate = ss.getRange('Y' + rowNum + '').setNumberFormat('mmmm" "d", "yyyy').getValue()
      var settlementID = getIdFromName('' + buyerName + ' - Settlement & Closing Deadline', settlementOldDate)
      if (calendar.getEventById(settlementID)){
        calendar.getEventById(settlementID).deleteEvent()
      }
    }
    eventName = '' + buyerName + ' - Settlement & Closing Deadline'
    newEvent = calendar.createAllDayEvent(eventName, new Date(settlementDate),{location: ''})
    ss.getRange('Y' + rowNum + '').setValue(settlementDate).setNumberFormat('m"/"d"/"yy')
  }
}

function getIdFromName(name, date){
  var ss = SpreadsheetApp.getActive()
  var email = ss.getSheetByName('Dashboard').getRange('B6').getValue()
  var calendar = CalendarApp.getCalendarById(email)
  var events = calendar.getEventsForDay(new Date(date))
  var title = ''
  
  for (var i = 0; i < events.length; i++){
    title = events[i].getTitle()
    if (title === name){
      return events[i].getId()
    } else {return ''}
  }
}

function redoFormatting() {
  var ss = SpreadsheetApp.getActive()
  ss.getRange('V2:Y2')
  .clear({contentsOnly: true})
  .setBackground('#7fa0af')
  .setFontColor('#ffffff')
  .setHorizontalAlignment('center')
  .setVerticalAlignment('middle')
  .setFontFamily('Verdana')
  .setFontSize(10)
  .setNumberFormat('m"/"d"/"yy')
  ss.getRange('X1:X2').setBorder(true, true, true, true, null, null, '#999999', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  ss.getRange('W1:W2').setBorder(true, true, true, true, null, null, '#999999', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  ss.getRange('V1:V2').setBorder(true, true, true, true, null, null, '#999999', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  ss.getRange('V1:Y2').setBorder(true, true, true, true, null, null, '#ffffff', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
}

function makeBuyerRed() {
  var ss = SpreadsheetApp.getActive();
  return ss.getRange('V2').setBackground('#f4cccc').setFontColor('#303f46');
}

function makeDatesRed() {
  var ss = SpreadsheetApp.getActive();
  return ss.getRange('W2:Y2').setBackground('#f4cccc').setFontColor('#303f46');
}



function getCalendarId(){
  var calendars = CalendarApp.getCalendarsByName('Black Ops Deadlines')
  var id = calendars[0].getName()
  SpreadsheetApp.getActive().getRange('V20').setValue(id)
}