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
    
    // Converting from New/Warm Lead to Opportunity 
    if ((val === 'Searching' || val === 'Touring' || val === 'Offering' || val === 'UC') && columnEdited === 7){
      convertToOpp(rowEdited)
    }
    
    // Converting from New/Warm Lead to Cold Lead 
    else if (val === 'Cold Lead' && columnEdited === 7){
      convertToCold(rowEdited)
    }
    
    // Converting from New/Warm Lead to Archive 
    else if ((val === 'Closed' && columnEdited === 7) || ((val === 'Lost' || val === 'Abandoned') && columnEdited === 12)){
      archive(rowEdited)
    }
  }
  
  // Check for stage or status change in Cold Leads tab
  else if (sheetName === 'Cold Leads' && (columnEdited === 7 || columnEdited === 12)){
    var cell = range.getA1Notation()
    var val = ss.getRange(cell).getValue()
    
    // Converting from Cold Lead to Opportunity 
    if ((val === 'Searching' || val === 'Touring' || val === 'Offering' || val === 'UC') && columnEdited === 7){
      convertToOpp(rowEdited)
    }
    
    // Converting from Cold Lead to New/Warm Lead 
    else if ((val === 'Warm Lead' || val === 'New Lead') && columnEdited === 7){
      convertToWarm(rowEdited)
    }
    
    // Converting from Cold Lead to Archive  
    else if ((val === 'Closed' && columnEdited === 7) || ((val === 'Lost' || val === 'Abandoned') && columnEdited === 12)){
      archive(rowEdited)
    }
  }
  
  // Check for stage or status change in Opportunity tab
  else if (sheetName === 'Opportunities' && (columnEdited === 7 || columnEdited === 12)){
    var cell = range.getA1Notation()
    var val = ss.getRange(cell).getValue()
    
    // Converting from Opportunity to Cold Lead 
    if (val === 'Cold Lead' && columnEdited === 7){
      convertToCold(rowEdited)
    }
    
    // Converting from Opportunity to New/Warm Lead 
    else if ((val === 'Warm Lead' || val === 'New Lead') && columnEdited === 7){
      convertToWarm(rowEdited)
    }
    
    // Converting from Opportunity to Archive 
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
        convertToOpp(rowEdited)
      }
      else if (val === 'Warm Lead' || val === 'New Lead'){
        ss.getRange(cell).setBackground(null)
        convertToWarm(rowEdited)
      }
      else if (val === 'Cold Lead'){
        ss.getRange(cell).setBackground(null)
        convertToCold(rowEdited)
      }
    }
    
    // Moving from Archive by changing status
    else if (val === 'Open' && columnEdited === 12){
      
      var stageCell = "G"+rowEdited+""
      var stage = ss.getRange(stageCell).getValue()
      ss.getRange("L"+rowEdited+"").setBackground(null);
      
      // Converting from Archive to New/Warm Lead 
      if (stage === 'New Lead' || stage === 'Warm Lead'){
        convertToWarm(rowEdited)
      }
      
      // Converting from Archive to Cold Lead 
      else if (stage === 'Cold Lead'){
        convertToCold(rowEdited)
      }
      
      // Converting from Archive to Opportunity 
      else if (stage === 'Searching' || stage === 'Touring' || stage === 'Offering' || stage === 'UC'){
        convertToOpp(rowEdited)
      }
    }
  }
}

function convertToOpp(rowEdited){
  var ss = SpreadsheetApp.getActive();
  ss.getSheetByName('Opportunities').insertRowsBefore(4,1)
  var range = "A"+rowEdited+":AB"+rowEdited+""
  ss.getRange(range).copyTo(ss.getSheetByName('Opportunities').getRange('A4:AB4'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false)
  ss.deleteRows(rowEdited, 1)
}

function convertToWarm(rowEdited){
  var ss = SpreadsheetApp.getActive();
  ss.getSheetByName('New/Warm Leads').insertRowsBefore(4,1)
  var range = "A"+rowEdited+":AB"+rowEdited+""
  ss.getRange(range).copyTo(ss.getSheetByName('New/Warm Leads').getRange('A4:AB4'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false)
  ss.deleteRows(rowEdited, 1)
}

function convertToCold(rowEdited){
  var ss = SpreadsheetApp.getActive();
  ss.getSheetByName('Cold Leads').insertRowsBefore(4,1)
  var range = "A"+rowEdited+":AB"+rowEdited+""
  ss.getRange(range).copyTo(ss.getSheetByName('Cold Leads').getRange('A4:AB4'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false)
  ss.deleteRows(rowEdited, 1)
}

function archive(rowEdited){
  var ss = SpreadsheetApp.getActive();
  ss.getSheetByName('Archive').insertRowsBefore(4,1)
  var range = "A"+rowEdited+":AB"+rowEdited+""
  ss.getRange(range).copyTo(ss.getSheetByName('Archive').getRange('A4:AB4'), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false)
  ss.deleteRows(rowEdited, 1)
}