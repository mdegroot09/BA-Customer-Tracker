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
  var dueDiligenceDate = ss.getRange('W2').getValue()
  var financingDate = ss.getRange('X2').getValue()
  var settlementDate = ss.getRange('Y2').getValue()
  
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
    
    // Get row number of buyer being changed
    var rowNum = ss.getRange('AA1').setFormula('=IFERROR(MATCH(V2,A:A,0),"")').getValue()
    ss.getRange('W' + rowNum + ':Y' + rowNum + '').setNumberFormat('m"/"d"/"yy')
    
    // Capture new dates and change format
    dueDiligenceDate = ss.getRange('W2').setNumberFormat('mmmm" "d", "yyyy').getValue()
    financingDate = ss.getRange('X2').setNumberFormat('mmmm" "d", "yyyy').getValue()
    settlementDate = ss.getRange('Y2').setNumberFormat('mmmm" "d", "yyyy').getValue()
    ss.getRange('W2:Y2').setNumberFormat('m"/"d"/"yy')
    
    // Capture old dates
    var dueDiligenceOldDate = ss.getRange('W' + rowNum + '').setNumberFormat('mmmm" "d", "yyyy').getValue()
    var financingOldDate = ss.getRange('X' + rowNum + '').setNumberFormat('mmmm" "d", "yyyy').getValue()
    var settlementOldDate = ss.getRange('Y' + rowNum + '').setNumberFormat('mmmm" "d", "yyyy').getValue()
  
    var email = ss.getSheetByName('Dashboard').getRange('B6').getValue()
    deleteCreateEvents(email, dueDiligenceOldDate, financingOldDate, settlementOldDate)
    
    email = 'homie.com_1cs8eji9ahpmol4rvqllcq8bco@group.calendar.google.com'
    deleteCreateEvents(email, dueDiligenceOldDate, financingOldDate, settlementOldDate)
    
    redoFormatting()
  }
}

function deleteCreateEvents(email, dueDiligenceOldDate, financingOldDate, settlementOldDate){
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
   
  // If a new Due Diligence date is entered
  if (dueDiligenceDate){ 
    
    // If previous due diligence date exists, find event and delete
    if (dueDiligenceOldDate){
      var dueDiligenceID = getIdFromName('' + buyerName + ' - Due Diligence Deadline', dueDiligenceOldDate, email)
      
      if (calendar.getEventById(dueDiligenceID)){
        calendar.getEventById(dueDiligenceID).deleteEvent()
      }
    }
    
    // Create new event with new date
    eventName = '' + buyerName + ' - Due Diligence Deadline'
    newEvent = calendar.createAllDayEvent(eventName, new Date(dueDiligenceDate),{location: ''})
    ss.getRange('W' + rowNum + '').setValue(dueDiligenceDate).setNumberFormat('m"/"d"/"yy')
  }
  
  // If a new F&A date is entered
  if (financingDate){ 
    
    // If previous F&A exists, find event and delete
    if (financingOldDate){
      var financingID = getIdFromName('' + buyerName + ' - F&A Deadline', financingOldDate, email)
      if (calendar.getEventById(financingID)){
        calendar.getEventById(financingID).deleteEvent()
      }
    }
      
    // Create new event with new date
    eventName = '' + buyerName + ' - F&A Deadline'
    newEvent = calendar.createAllDayEvent(eventName, new Date(financingDate),{location: ''})
    ss.getRange('X' + rowNum + '').setValue(financingDate).setNumberFormat('m"/"d"/"yy')
  }
  
  // If a new Settlement date is entered
  if (settlementDate){ 
    
    // If previous Settlement exists, find event and delete
    if (settlementOldDate){
      var settlementID = getIdFromName('' + buyerName + ' - Settlement & Closing Deadline', settlementOldDate, email)
      if (calendar.getEventById(settlementID)){
        calendar.getEventById(settlementID).deleteEvent()
      }
    }
    
    // Create new event with new date
    eventName = '' + buyerName + ' - Settlement & Closing Deadline'
    newEvent = calendar.createAllDayEvent(eventName, new Date(settlementDate),{location: ''})
    ss.getRange('Y' + rowNum + '').setValue(settlementDate).setNumberFormat('m"/"d"/"yy')
  }
  
  // Send out UC emails if there aren't any existing deadlines and it's not the 2nd time through this function
  if (!dueDiligenceOldDate && !financingOldDate && !settlementOldDate && email !== 'homie.com_1cs8eji9ahpmol4rvqllcq8bco@group.calendar.google.com'){
    sendUCEmails()
  }
}

function getIdFromName(name, date, email){
  var ss = SpreadsheetApp.getActive()
  var calendar = CalendarApp.getCalendarById(email)
  var events = calendar.getEventsForDay(new Date(date))
  var title = ''
  
  for (var i = 0; i < events.length; i++){
    title = events[i].getTitle()
    if (title === name){
      return events[i].getId()
    }
  }
  return ''
}

function sendUCEmails(){
  MailApp.sendEmail('mike.degroot@homie.com', '#1', "You're under contract!")
  MailApp.sendEmail('mdegroot09@gmail.com', '#2', "You're under contract!")
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

function deleteEvents(){
  var ss = SpreadsheetApp.getActive()
  var buyerName = ss.getRange('V2').getValue()
  
  // If no buyer is selected, throw an error
  if (!buyerName){
    return makeBuyerRed()
  }
  
  // Get row number of buyer being changed
  var rowNum = ss.getRange('AA1').setFormula('=IFERROR(MATCH(V2,A:A,0),"")').getValue()
  ss.getRange('W' + rowNum + ':Y' + rowNum + '').setNumberFormat('m"/"d"/"yy')
  
  // Capture old dates
  var dueDiligenceOldDate = ss.getRange('W' + rowNum + '').setNumberFormat('mmmm" "d", "yyyy').getValue()
  var financingOldDate = ss.getRange('X' + rowNum + '').setNumberFormat('mmmm" "d", "yyyy').getValue()
  var settlementOldDate = ss.getRange('Y' + rowNum + '').setNumberFormat('mmmm" "d", "yyyy').getValue()
  
  // Delete events from agent's calendar
  var email = ss.getSheetByName('Dashboard').getRange('B6').getValue()
  var calendar = CalendarApp.getCalendarById(email)
    
  // If previous due diligence date exists, find event and delete
  if (dueDiligenceOldDate){
    var dueDiligenceID = getIdFromName('' + buyerName + ' - Due Diligence Deadline', dueDiligenceOldDate, email)
    
    if (calendar.getEventById(dueDiligenceID)){
      calendar.getEventById(dueDiligenceID).deleteEvent()
    }
  }
    
  // If previous F&A exists, find event and delete
  if (financingOldDate){
    var financingID = getIdFromName('' + buyerName + ' - F&A Deadline', financingOldDate, email)
    if (calendar.getEventById(financingID)){
      calendar.getEventById(financingID).deleteEvent()
    }
  }
    
  // If previous Settlement exists, find event and delete
  if (settlementOldDate){
    var settlementID = getIdFromName('' + buyerName + ' - Settlement & Closing Deadline', settlementOldDate, email)
    if (calendar.getEventById(settlementID)){
      calendar.getEventById(settlementID).deleteEvent()
    }
  }
  
  // Delete events from shared group calendar
  email = 'homie.com_1cs8eji9ahpmol4rvqllcq8bco@group.calendar.google.com'
  calendar = CalendarApp.getCalendarById(email)
    
  // If previous due diligence date exists, find event and delete
  if (dueDiligenceOldDate){
    var dueDiligenceID = getIdFromName('' + buyerName + ' - Due Diligence Deadline', dueDiligenceOldDate, email)
    
    if (calendar.getEventById(dueDiligenceID)){
      calendar.getEventById(dueDiligenceID).deleteEvent()
    }
  }
    
  // If previous F&A exists, find event and delete
  if (financingOldDate){
    var financingID = getIdFromName('' + buyerName + ' - F&A Deadline', financingOldDate, email)
    if (calendar.getEventById(financingID)){
      calendar.getEventById(financingID).deleteEvent()
    }
  }
    
  // If previous Settlement exists, find event and delete
  if (settlementOldDate){
    var settlementID = getIdFromName('' + buyerName + ' - Settlement & Closing Deadline', settlementOldDate, email)
    if (calendar.getEventById(settlementID)){
      calendar.getEventById(settlementID).deleteEvent()
    }
  }
  
  // Clear previous dates from buyer row
  ss.getRange('W' + rowNum + ':Y' + rowNum + '').clear({contentsOnly: true})
  
  // Reset the formatting for the date inputs 
  redoFormatting()
}