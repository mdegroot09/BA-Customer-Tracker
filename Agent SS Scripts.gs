function onEdit(e){
  var range = e.range;
  var columnEdited = range.getColumn();
  var rowEdited = range.getRow();
  var ss = SpreadsheetApp.getActive();
  var sheetName = ss.getActiveSheet().getName()
  
  // Check for stage or status and update Last Changed column
  if (columnEdited === 7 || columnEdited === 12){
    updateLastChange(rowEdited)
  }
  
  // Check for stage or status change in New/Warm Leads tab 
  if (sheetName === 'New/Warm Leads' && (columnEdited === 7 || columnEdited === 12)){
    var cell = range.getA1Notation()
    var val = ss.getRange(cell).getValue()
    
    // Moving from New/Warm Lead to Opportunity 
    if ((val === 'Searching' || val === 'Touring' || val === 'Offering' || val === 'UC') && columnEdited === 7){
      return moveToOpp(rowEdited)
    }
    
    // Moving from New/Warm Lead to Cold Lead 
    else if (val === 'Cold Lead' && columnEdited === 7){
      return moveToCold(rowEdited)
    }
    
    // Moving from New/Warm Lead to Archive 
    else if ((val === 'Closed' && columnEdited === 7) || ((val === 'Lost' || val === 'Abandoned') && columnEdited === 12)){
      if (val === 'Closed'){
        var closedDateExists = checkClosedDate(rowEdited)
        if (!closedDateExists) {
          return
        }
        else {
          return archive(rowEdited)
        }
      }
      else {
        return archive(rowEdited)
      }
    }
  }
  
  // Check for stage or status change in Cold Leads tab
  else if (sheetName === 'Cold Leads' && (columnEdited === 7 || columnEdited === 12)){
    var cell = range.getA1Notation()
    var val = ss.getRange(cell).getValue()
    
    // Moving from Cold Lead to Opportunity 
    if ((val === 'Searching' || val === 'Touring' || val === 'Offering' || val === 'UC') && columnEdited === 7){
      return moveToOpp(rowEdited)
    }
    
    // Moving from Cold Lead to New/Warm Lead 
    else if ((val === 'Warm Lead' || val === 'New Lead') && columnEdited === 7){
      return moveToWarm(rowEdited)
    }
    
    // Moving from Cold Lead to Archive  
    else if ((val === 'Closed' && columnEdited === 7) || ((val === 'Lost' || val === 'Abandoned') && columnEdited === 12)){
      if (val === 'Closed'){
        var closedDateExists = checkClosedDate(rowEdited)
        if (!closedDateExists) {
          return
        }
        else {
          return archive(rowEdited)
        }
      }
      else {
        return archive(rowEdited)
      }
    }
  }
  
  // Check for stage or status change in Opportunity tab
  else if (sheetName === 'Opportunities' && (columnEdited === 7 || columnEdited === 12)){
    var cell = range.getA1Notation()
    var val = ss.getRange(cell).getValue()
    
    // Moving from Opportunity to Cold Lead 
    if (val === 'Cold Lead' && columnEdited === 7){
      return moveToCold(rowEdited)
    }
    
    // Moving from Opportunity to New/Warm Lead 
    else if ((val === 'Warm Lead' || val === 'New Lead') && columnEdited === 7){
      return moveToWarm(rowEdited)
    }
    
    // Moving from Opportunity to Archive 
    else if ((val === 'Closed' && columnEdited === 7) || ((val === 'Lost' || val === 'Abandoned') && columnEdited === 12)){
      if (val === 'Closed'){
        var closedDateExists = checkClosedDate(rowEdited)
        if (!closedDateExists) {
          return
        }
        else {
          return archive(rowEdited)
        }
      }
      else {
        return archive(rowEdited)
      }
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
        return moveToOpp(rowEdited)
      }
      else if (val === 'Warm Lead' || val === 'New Lead'){
        ss.getRange(cell).setBackground(null)
        return moveToWarm(rowEdited)
      }
      else if (val === 'Cold Lead'){
        ss.getRange(cell).setBackground(null)
        return moveToCold(rowEdited)
      }
    }
    
    // Moving from Archive by changing status
    else if (val === 'Open' && columnEdited === 12){
      
      var stageCell = "G"+rowEdited+""
      var stage = ss.getRange(stageCell).getValue()
      ss.getRange("L"+rowEdited+"").setBackground(null);
      
      // Moving from Archive to New/Warm Lead 
      if (stage === 'New Lead' || stage === 'Warm Lead'){
        return moveToWarm(rowEdited)
      }
      
      // Moving from Archive to Cold Lead 
      else if (stage === 'Cold Lead'){
        return moveToCold(rowEdited)
      }
      
      // Moving from Archive to Opportunity 
      else if (stage === 'Searching' || stage === 'Touring' || stage === 'Offering' || stage === 'UC'){
        return moveToOpp(rowEdited)
      }
    }
  }

  // Alert if dates are manually added to row.
  else if (sheetName === 'Opportunities' && (columnEdited === 23 || columnEdited === 24 || columnEdited === 25) && rowEdited > 2){
    alertUser('Use the form above to create/delete contract deadlines.')
    ss.getRange('W' + rowEdited + '').setValue(ss.getRange('AC' + rowEdited + '').getValue())
    ss.getRange('X' + rowEdited + '').setValue(ss.getRange('AD' + rowEdited + '').getValue())
    return ss.getRange('Y' + rowEdited + '').setValue(ss.getRange('AE' + rowEdited + '').getValue())
  }
  
  // Archive closed buyers who are already
  else if ((sheetName === 'Opportunities' || sheetName === 'New/Warm Leads' || sheetName === 'Cold Leads') && columnEdited === 26 && val !== '' && ss.getRange('G' + rowEdited).getValue() === 'Closed'){
    archive(rowEdited)
  }
}

function moveToOpp(rowEdited){
  var ss = SpreadsheetApp.getActive();
  ss.getSheetByName('Opportunities').insertRowsBefore(4,1)
  var range = "A"+rowEdited+":"+rowEdited+""
  ss.getRange(range).copyTo(ss.getSheetByName('Opportunities').getRange('A4:4'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false)
  ss.deleteRows(rowEdited, 1)
  
  ss.getSheetByName('Opportunities').getRange('V2')
  .setDataValidation(
    SpreadsheetApp.newDataValidation().setAllowInvalid(true).requireValueInRange(ss.getRange('Opportunities!$A$4:$A'), true).build()
  )
}

function moveToWarm(rowEdited){
  var ss = SpreadsheetApp.getActive();
  ss.getSheetByName('New/Warm Leads').insertRowsBefore(4,1)
  var range = "A"+rowEdited+":"+rowEdited+""
  ss.getRange(range).copyTo(ss.getSheetByName('New/Warm Leads').getRange('A4:4'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false)
  ss.deleteRows(rowEdited, 1)
}

function moveToCold(rowEdited){
  var ss = SpreadsheetApp.getActive();
  ss.getSheetByName('Cold Leads').insertRowsBefore(4,1)
  var range = "A"+rowEdited+":"+rowEdited+""
  ss.getRange(range).copyTo(ss.getSheetByName('Cold Leads').getRange('A4:4'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false)
  ss.deleteRows(rowEdited, 1)
}

function archive(rowEdited){
  var ss = SpreadsheetApp.getActive();
  ss.getSheetByName('Archive').insertRowsBefore(4,1)
  var range = "A"+rowEdited+":"+rowEdited+""
  ss.getRange(range).copyTo(ss.getSheetByName('Archive').getRange('A4:4'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false)
  ss.deleteRows(rowEdited, 1)
}



function addBuyer(){
  var ss = SpreadsheetApp.getActive();
  ss.insertRowsBefore(4,1)
  
  ss.getRange('L4').setValue('Open')
  ss.getRange('O4').setFormula('=IF(B4="","",VLOOKUP(B4,Setting!A:B,2,false))')
  ss.getRange('Q4').setFormula('=IF(J4="","",IFS(J4="TBD","TBD",MONTH(J4)=1,"January",MONTH(J4)=2,"February",MONTH(J4)=3,"March",MONTH(J4)=4,"April",MONTH(J4)=5,"May",MONTH(J4)=6,"June",MONTH(J4)=7,"July",MONTH(J4)=8,"August",MONTH(J4)=9,"September",MONTH(J4)=10,"October",MONTH(J4)=11,"November",MONTH(J4)=12,"December"))');
  ss.getRange('R4').setFormula('=IF(J4="","",IF(J4="TBD","TBD",year(J4)))');
  ss.getRange('S4').setFormula('=IFS(N4="TBD","TBD",N4="","",N4>0,O4&" "&N4)');
  ss.getRange('AA4').setNumberFormat('m"/"d" "h":"mma/p')
  ss.getRange('AA4').setValue('=NOW()')
  var date = ss.getRange('AA4').getValue()
  ss.getRange('AA4').setValue(date)
  
  ss.getSheetByName('Opportunities').getRange('V2')
  .setDataValidation(
    SpreadsheetApp.newDataValidation().setAllowInvalid(true).requireValueInRange(ss.getRange('Opportunities!$A$4:$A'), true).build()
  )
}

function updateLastChange(rowEdited){
  var ss = SpreadsheetApp.getActive()
  ss.getRange('AF' + rowEdited).setValue("=NOW()").setNumberFormat('m"/"d"/"yy')
  var date = ss.getRange('AF' + rowEdited).getValue()
  ss.getRange('AF' + rowEdited).setValue(date)
  
  // Hide last columns
  ss.getActiveSheet().hideColumn(ss.getRange('AC:AF'));
}

function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu('UC Menu')
  .addItem('Convert to UC', 'convertUC')
  .addToUi();
  
  SpreadsheetApp.getUi()
  .createMenu('Enter Deadline')
  .addItem('Enter Deadline', 'enterDeadline')
  .addToUi();
}

function convertUC() {
  var ui = SpreadsheetApp.getUi(); 
  var ss = SpreadsheetApp.getActive()
  
  var range = ss.getRange("A:A")
  var columnAValues = range.getValues()

  var result = ui.prompt(
      'Convert to Under Contract',
      "Please enter the Buyer's name:",
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var name = result.getResponseText();
  
  // If user clicked "OK"
  if (button == ui.Button.OK) {
    
    var rowNum = findBuyerName(name, columnAValues)
    
    // if rowNum is not an empty string
    if (rowNum){
      
      // If cancelled or closed, quit macro
      var dueDiligenceDate = enterDeadline('Due Diligence', rowNum)
      if (dueDiligenceDate === 'error'){
        return 
      } 
      
      // If cancelled or closed, quit macro
      var financingDate = enterDeadline('Financing & Appraisal', rowNum)
      if (financingDate === 'error'){
        return 
      }
      
      // If cancelled or closed, quit macro
      var settlementDate = enterDeadline('Settlement', rowNum)
      if (settlementDate === 'error'){
        return 
      }
      
      ui.alert('dueDiligenceDate: ' + dueDiligenceDate + ', financingDate: ' + financingDate + ', settlementDate: ' + settlementDate)
      return updateCalendar(dueDiligenceDate, financingDate, settlementDate, rowNum)
    }
    
    // if rowNum is an empty string
    else {
      ui.alert('"' + name + '" not found in Opportunities. Please check the spelling and try again.')
    }
  } 
}

function findBuyerName(name, columnAValues){
  var ss = SpreadsheetApp.getActive()
  var ui = SpreadsheetApp.getUi()
  
  // Look for buyer name entered in column A
  var j = 0
  for (var i = 4; i < columnAValues.length; i++){
    if (ss.getRange('A' + i).getValue() && name === ss.getRange('A' + i).getValue()){
      return i
    }
    if (!ss.getRange('A' + i).getValue()){
      j++
    }
    if (j === 3){
      return ''
    }
  }
}

function enterDeadline(deadline, rowNum){
  var ui = SpreadsheetApp.getUi()
  
  var response = ui.prompt(
    deadline,
    "Please enter the " + deadline + " Deadline:",
    ui.ButtonSet.OK_CANCEL);
  
  var button = response.getSelectedButton()
  var date = response.getResponseText()
  
  // Attempt to convert to a date with the input
  if (!isNaN(new Date(date))){
    return date
  }
  
  // If button clicked is "OK", continue
  if (button == ui.Button.OK){
  
    // keep asking for date until valid date is entered
    while (true){
      response = ui.prompt(
        deadline,
        'Please enter a valid ' + deadline + ' Deadline (e.g. "' + new Date() + '"):',
        ui.ButtonSet.OK_CANCEL)
        
      // Get button clicked and response
      button = response.getSelectedButton()
      date = response.getResponseText()
        
      // If button clicked isn't "OK", return error
      if (button != ui.Button.OK){
        return 'error'
      }
      
      // Attempt to convert to a date with the input
      if (!isNaN(new Date(date))){
        return date
      }
    }   
  }
  
  // If button clicked isn't "OK", return error
  else {
    return 'error'
  }
}

function updateCalendar(dueDiligenceDate, financingDate, settlementDate, rowNum){
  var ss = SpreadsheetApp.getActive()
  var buyerName = ss.getRange('A' + rowNum).getValue()
  
  SpreadsheetApp.getUi().alert('dueDiligenceDate:' + dueDiligenceDate + ', financingDate: ' + financingDate + ', settlementDate: ' + settlementDate + ', rowNum: ' + rowNum)
  
  var dueDiligenceOldDate = ''
  var financingOldDate = ''
  var settlementOldDate = ''
  
  // Define OldDate variables if rowNum is valid
  if (rowNum) {
    // Capture old dates
    dueDiligenceOldDate = ss.getRange('AC' + rowNum + '').setNumberFormat('mmmm" "d", "yyyy').getValue()
    financingOldDate = ss.getRange('AD' + rowNum + '').setNumberFormat('mmmm" "d", "yyyy').getValue()
    settlementOldDate = ss.getRange('AE' + rowNum + '').setNumberFormat('mmmm" "d", "yyyy').getValue()
  }
  
  // Set format of old dates
  ss.getRange('AC' + rowNum + ':AE' + rowNum + '').setNumberFormat('m"/"d"/"yy')
  
  // Capture new dates and change format
  dueDiligenceDate = ss.getRange('W' + rowNum).setValue(dueDiligenceDate).setNumberFormat('mmmm" "d", "yyyy').getValue()
  financingDate = ss.getRange('X' + rowNum).setValue(financingDate).setNumberFormat('mmmm" "d", "yyyy').getValue()
  settlementDate = ss.getRange('Y' + rowNum).setValue(settlementDate).setNumberFormat('mmmm" "d", "yyyy').getValue()

  var email = ss.getSheetByName('Dashboard').getRange('B6').getValue()
  deleteCreateEvents(email, rowNum, dueDiligenceOldDate, financingOldDate, settlementOldDate, dueDiligenceDate, financingDate, settlementDate)
  
  email = 'homie.com_1cs8eji9ahpmol4rvqllcq8bco@group.calendar.google.com'
  deleteCreateEvents(email, rowNum, dueDiligenceOldDate, financingOldDate, settlementOldDate, dueDiligenceDate, financingDate, settlementDate)
  
  // Change status to 'UC'
  ss.getRange('G' + rowNum).setValue('UC')
  
  return alertUser('Success! Events have been added to your calendar.')
}

function deleteCreateEvents(email, rowNum, dueDiligenceOldDate, financingOldDate, settlementOldDate, dueDiligenceDate, financingDate, settlementDate){
  var calendar = CalendarApp.getCalendarById(email)
  var ss = SpreadsheetApp.getActive()
  var buyerName = ss.getRange('A' + rowNum).getValue()
  var eventName = ''
  var newEvent = ''
   
  // If a new Due Diligence date is entered
  if (dueDiligenceDate && dueDiligenceDate !== 'N/A'){ 
    
    // If previous due diligence date exists, find event and delete
    if (dueDiligenceOldDate && dueDiligenceOldDate !== 'N/A'){
      var dueDiligenceID = getIdFromName('' + buyerName + ' - Due Diligence Deadline', dueDiligenceOldDate, email)
      
      if (calendar.getEventById(dueDiligenceID)){
        calendar.getEventById(dueDiligenceID).deleteEvent()
      }
    }
    
    // Create new event with new date
    eventName = '' + buyerName + ' - Due Diligence Deadline'
    newEvent = calendar.createAllDayEvent(eventName, new Date(dueDiligenceDate),{location: ''})
    ss.getRange('W' + rowNum + '').setValue(dueDiligenceDate).setNumberFormat('m"/"d"/"yy')
    ss.getRange('AC' + rowNum + '').setValue(dueDiligenceDate).setNumberFormat('m"/"d"/"yy')
  }
  
  // Set Due Diligence date to N/A
  else if (dueDiligenceDate === 'N/A'){
    ss.getRange('W' + rowNum + '').setValue('N/A')
    ss.getRange('AC' + rowNum + '').setValue('N/A')
  }
  
  // If a new F&A date is entered
  if (financingDate && financingDate !== 'N/A'){ 
    
    // If previous F&A exists, find event and delete
    if (financingOldDate && financingOldDate !== 'N/A'){
      var financingID = getIdFromName('' + buyerName + ' - F&A Deadline', financingOldDate, email)
      if (calendar.getEventById(financingID)){
        calendar.getEventById(financingID).deleteEvent()
      }
    }
      
    // Create new event with new date
    eventName = '' + buyerName + ' - F&A Deadline'
    newEvent = calendar.createAllDayEvent(eventName, new Date(financingDate),{location: ''})
    ss.getRange('X' + rowNum + '').setValue(financingDate).setNumberFormat('m"/"d"/"yy')
    ss.getRange('AD' + rowNum + '').setValue(financingDate).setNumberFormat('m"/"d"/"yy')
  }
  
  // Set Due Diligence date to N/A
  else if (financingDate === 'N/A'){
    ss.getRange('X' + rowNum + '').setValue('N/A')
    ss.getRange('AD' + rowNum + '').setValue('N/A')
  }
  
  // If a new Settlement date is entered
  if (settlementDate && settlementDate !== 'N/A'){ 
    
    // If previous Settlement exists, find event and delete
    if (settlementOldDate & settlementOldDate !== 'N/A'){
      var settlementID = getIdFromName('' + buyerName + ' - Settlement & Closing Deadline', settlementOldDate, email)
      if (calendar.getEventById(settlementID)){
        calendar.getEventById(settlementID).deleteEvent()
      }
    }
    
    // Create new event with new date
    eventName = '' + buyerName + ' - Settlement & Closing Deadline'
    newEvent = calendar.createAllDayEvent(eventName, new Date(settlementDate),{location: ''})
    ss.getRange('Y' + rowNum + '').setValue(settlementDate).setNumberFormat('m"/"d"/"yy')
    ss.getRange('AE' + rowNum + '').setValue(settlementDate).setNumberFormat('m"/"d"/"yy')
  }
  
  // Set Due Diligence date to N/A
  else if (settlementDate === 'N/A'){
    ss.getRange('Y' + rowNum + '').setValue('N/A')
    ss.getRange('AE' + rowNum + '').setValue('N/A')
  }
  
  // Send out UC emails if there aren't any existing deadlines and it's not the 2nd time through this function
  if (!dueDiligenceOldDate && !financingOldDate && !settlementOldDate && email !== 'homie.com_1cs8eji9ahpmol4rvqllcq8bco@group.calendar.google.com'){
    //    sendUCEmails(email)
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

function sendUCEmails(type, agentName, emailAddress, toolsLink, buyerName){
  var ss = SpreadsheetApp.getActive()
  
  // If converting to UC
  if (type === 'UC'){
    
    var toolsHTML = ''
    
    // check for toolsLink
    if (toolsLink){
      toolsHTML = "Here is the link to the Offer in Tools: " + toolsLink + "<br><br>"
    }
    
    MailApp.sendEmail({
      // to: email + "," + 'mdegroot09@gmail.com',
      to: emailAddress,
      subject: buyerName + " Under Contract", 
      htmlBody: 
      buyerName + " is now under contract<br><br>" +
      toolsHTML + 
      "Thanks,<br>" +
      agentName
    })
  }
  
  else if (type === 'Cancelled'){
    MailApp.sendEmail({
      to: emailAddress,
      subject: "Contract Cancelled", 
      htmlBody: 
      "You're under contract!<br><br>" +
      "Thanks,<br>" +
      "<img src='https://simplejoys.s3.us-east-2.amazonaws.com/email%20signature-1576377050955.png'>"
    })
  }
}

function alertUser(text){
  var ui = SpreadsheetApp.getUi()
  ui.alert(text)
}

function resetOldDateFormats(rowNum){
  var ss = SpreadsheetApp.getActive()
  ss.getRange('W' + rowNum + '').setNumberFormat('m"/"d"/"yy').getValue()
  ss.getRange('X' + rowNum + '').setNumberFormat('m"/"d"/"yy').getValue()
  ss.getRange('Y' + rowNum + '').setNumberFormat('m"/"d"/"yy').getValue()
}

function checkClosedDate(rowEdited){
  var ss = SpreadsheetApp.getActive()
  var closedDate = ss.getRange('Z' + rowEdited).getValue()
  if (!closedDate){
    var ui = SpreadsheetApp.getUi();
    ui.alert('Enter a closed date into cell: Z' + rowEdited);
    return false
  }
  else {
    return true
  }
}

function cancelContract(){
  var ss = SpreadsheetApp.getActive()
  var ui = SpreadsheetApp.getUi()
  
  var range = ss.getRange("A:A")
  var columnAValues = range.getValues()
  
  var result = ui.prompt(
    'Cancel Contract',
    "Please enter the Buyer's name:",
    ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  var buyerName = result.getResponseText();
  
  // If user clicked "OK"
  if (button == ui.Button.OK) {
    
    var rowNum = findBuyerName(buyerName, columnAValues)
    
    // if rowNum is an empty string
    if (!rowNum){
      return ui.alert('"' + buyerName + '" not found in Opportunities. Please check the spelling and try again.')
    }
  }
  
  // If user cancelled or closed the box
  else {
    return
  }
  
  // Get row number of buyer being changed
  ss.getRange('W' + rowNum + ':Y' + rowNum + '').setNumberFormat('m"/"d"/"yy')
  
  // Capture old dates
  var dueDiligenceOldDate = ss.getRange('W' + rowNum + '').setNumberFormat('mmmm" "d", "yyyy').getValue()
  var financingOldDate = ss.getRange('X' + rowNum + '').setNumberFormat('mmmm" "d", "yyyy').getValue()
  var settlementOldDate = ss.getRange('Y' + rowNum + '').setNumberFormat('mmmm" "d", "yyyy').getValue()
  
  // Delete events from agent's calendar
  var email = ss.getSheetByName('Dashboard').getRange('B6').getValue()
  var calendar = CalendarApp.getCalendarById(email)
    
  // If previous due diligence date exists, find event and delete
  if (dueDiligenceOldDate && dueDiligenceOldDate !== 'N/A'){
    var dueDiligenceID = getIdFromName('' + buyerName + ' - Due Diligence Deadline', dueDiligenceOldDate, email)
    
    if (calendar.getEventById(dueDiligenceID)){
      calendar.getEventById(dueDiligenceID).deleteEvent()
    }
  }
    
  // If previous F&A exists, find event and delete
  if (financingOldDate && financingOldDate !== 'N/A'){
    var financingID = getIdFromName('' + buyerName + ' - F&A Deadline', financingOldDate, email)
    if (calendar.getEventById(financingID)){
      calendar.getEventById(financingID).deleteEvent()
    }
  }
    
  // If previous Settlement exists, find event and delete
  if (settlementOldDate && settlementOldDate !== 'N/A'){
    var settlementID = getIdFromName('' + buyerName + ' - Settlement & Closing Deadline', settlementOldDate, email)
    if (calendar.getEventById(settlementID)){
      calendar.getEventById(settlementID).deleteEvent()
    }
  }
  
  // Delete events from shared group calendar
  email = 'homie.com_1cs8eji9ahpmol4rvqllcq8bco@group.calendar.google.com'
  calendar = CalendarApp.getCalendarById(email)
    
  // If previous due diligence date exists, find event and delete
  if (dueDiligenceOldDate && dueDiligenceOldDate !== 'N/A'){
    var dueDiligenceID = getIdFromName('' + buyerName + ' - Due Diligence Deadline', dueDiligenceOldDate, email)
    
    if (calendar.getEventById(dueDiligenceID)){
      calendar.getEventById(dueDiligenceID).deleteEvent()
    }
  }

  // If previous F&A exists, find event and delete
  if (financingOldDate && financingOldDate !== 'N/A'){
    var financingID = getIdFromName('' + buyerName + ' - F&A Deadline', financingOldDate, email)
    if (calendar.getEventById(financingID)){
      calendar.getEventById(financingID).deleteEvent()
    }
  }
    
  // If previous Settlement exists, find event and delete
  if (settlementOldDate && settlementOldDate !== 'N/A'){
    var settlementID = getIdFromName('' + buyerName + ' - Settlement & Closing Deadline', settlementOldDate, email)
    if (calendar.getEventById(settlementID)){
      calendar.getEventById(settlementID).deleteEvent()
    }
  }
  
  // Clear previous dates from buyer row
  ss.getRange('W' + rowNum + ':Y' + rowNum + '').clear({contentsOnly: true})
  ss.getRange('AC' + rowNum + ':AE' + rowNum + '').clear({contentsOnly: true})
  
  // Change stage to Cancelled
  ss.getRange('G' + rowNum).setValue('Cancelled')
  
  // Success alert
  alertUser('Events were successfully deleted from your calendar.')
}