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
    
    // Converting to Opportunity from New/Warm Lead
    if ((val === 'Searching' || val === 'Touring' || val === 'Offering' || val === 'UC') && columnEdited === 7){
      convertToOpp(rowEdited)
    }
    
    // Converting to Cold Lead from New/Warm Lead
    else if (val === 'Cold Lead' && columnEdited === 7){
      convertToCold(rowEdited)
    }
    
    // Converting Closed and Lost to Archive from New/Warm Lead 
    else if ((val === 'Closed' && columnEdited === 7) || ((val === 'Lost' || val === 'Lost') && columnEdited === 12)){
      archive(rowEdited)
    }
  }
  
  // Check for stage or status change in Cold Leads tab
  else if (sheetName === 'Cold Leads' && (columnEdited === 7 || columnEdited === 12)){
    var cell = range.getA1Notation()
    var val = ss.getRange(cell).getValue()
    
    // Converting to Opportunity from Cold Lead
    if ((val === 'Searching' || val === 'Touring' || val === 'Offering' || val === 'UC') && columnEdited === 7){
      convertToOpp(rowEdited)
    }
    
    // Converting to New/Warm Lead from Cold Lead
    else if ((val === 'Warm Lead' || val === 'New Lead') && columnEdited === 7){
      convertToWarm(rowEdited)
    }
    
    // Converting Closed and Lost to Archive from Cold Lead 
    else if ((val === 'Closed' && columnEdited === 7) || ((val === 'Lost' || val === 'Lost') && columnEdited === 12)){
      archive(rowEdited)
    }
  }
  
  // Check for stage or status change in Opportunity tab
  else if (sheetName === 'Opportunities' && (columnEdited === 7 || columnEdited === 12)){
    var cell = range.getA1Notation()
    var val = ss.getRange(cell).getValue()
    
    // Converting to Cold Lead from Opportunity
    if (val === 'Cold Lead' && columnEdited === 7){
      convertToCold(rowEdited)
    }
    
    // Converting to New/Warm Lead from Opportunity
    else if ((val === 'Warm Lead' || val === 'New Lead') && columnEdited === 7){
      convertToWarm(rowEdited)
    }
    
    // Converting Closed and Lost to Archive from Opportunity
    else if ((val === 'Closed' && columnEdited === 7) || ((val === 'Lost' || val === 'Lost') && columnEdited === 12)){
      archive(rowEdited)
    }
  }
  
  // Check for status change in Archive tab
  else if (sheetName === 'Archive' && columnEdited === 12){
    var cell = range.getA1Notation()
    var val = ss.getRange(cell).getValue()
    
    // Moving from Archive
    if ((val === 'Open') && columnEdited === 12){
      
      var stageCell = "G"+rowEdited+""
      var stage = ss.getRange(stageCell).getValue()
      ss.getRange("L"+rowEdited+"").setBackground(null);
      
      // Converting to New/Warm Lead from Archive
      if (stage === 'New Lead' || stage === 'Warm Lead'){
        convertToWarm(rowEdited)
      }
      
      // Converting to Cold Lead from Archive
      else if (stage === 'Cold Lead'){
        convertToCold(rowEdited)
      }
      
      // Converting to Opportunity from Archive
      else if (stage === 'Searching' || stage === 'Touring' || stage === 'Offering' || stage === 'UC'){
        convertToOpp(rowEdited)
      }
    }
    
    // Converting Closed and Lost to Archive from Opportunity
    else if ((val === 'Closed' && columnEdited === 7) || (val === 'Lost' && columnEdited === 12)){
      archive(rowEdited)
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
