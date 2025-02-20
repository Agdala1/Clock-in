function doGet(e) {
    if (e.parameter && e.parameter.admin) {
        return HtmlService.createHtmlOutputFromFile('admin')
            .setWidth(600)
            .setHeight(400);
    } else {
        return HtmlService.createHtmlOutputFromFile('TimeTracker')
            .setWidth(600)
            .setHeight(400);
    }
}

function getEmployees() { 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var employeeSheet = ss.getSheetByName("EMPLOYEES"); 
  var getLastRow = employeeSheet.getLastRow();  
  return employeeSheet.getRange(2, 1, getLastRow - 1, 1).getValues();  
}

function clockIn(employee, clockTime) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("MAIN");
  var lastRow = mainSheet.getLastRow();
  
  var return_date = '';
  var error = 'SUCCESS';
  var return_array = [];

  // Verificar si el empleado ya está clocked in
  for (var j = 2; j <= lastRow; j++) {
    if (employee == mainSheet.getRange(j, 1).getValue() && mainSheet.getRange(j, 3).getValue() == '') {
      error = 'Need to Clock Out before Clocking In';
      return_array.push([error, return_date, employee]);
      return return_array;
    }
  }
  
  // Crear una fecha a partir de clockTime
  var new_date = clockTime ? new Date(clockTime) : new Date();
  return_date = getDate(new_date);
  
  mainSheet.getRange(lastRow + 1, 1).setValue(employee).setFontSize(12);
  mainSheet.getRange(lastRow + 1, 2).setValue(new_date).setNumberFormat("MM/dd/yyyy hh:mm:ss A/P").setHorizontalAlignment("left").setFontSize(12);
  
  return_array.push([error, return_date, employee]);
  return return_array;
}

function clockOut(employee, clockTime) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("MAIN");
  var lastRow = mainSheet.getLastRow();
  
  var foundRecord = false;
  var new_date = clockTime ? new Date(clockTime) : new Date();
  var return_date = getDate(new_date);
  var error = 'SUCCESS';
  var return_array = [];
  
  for (var j = 2; j <= lastRow; j++) {
    if (employee == mainSheet.getRange(j, 1).getValue() && mainSheet.getRange(j, 3).getValue() == '') {
      mainSheet.getRange(j, 3).setValue(new_date).setNumberFormat("MM/dd/yyyy hh:mm:ss A/P").setHorizontalAlignment("left").setFontSize(12);
      var totalTime = (mainSheet.getRange(j, 3).getValue() - mainSheet.getRange(j, 2).getValue()) / (60 * 60 * 1000);
      mainSheet.getRange(j, 4).setValue(totalTime.toFixed(2)).setNumberFormat("#0.00").setHorizontalAlignment("left").setFontSize(12);
      foundRecord = true;     
    }
  }
  
  if (foundRecord == false) {
    return_array.push(['Need to Clock In First', '', employee]);
    return return_array; 
  }
  
  TotalHours();
  
  return_array.push([error, return_date, employee]);
  return return_array;
}

function getActiveEmployees() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("MAIN");
  var lastRow = mainSheet.getLastRow();
  var activeEmployees = [];
  
  for (var j = 2; j <= lastRow; j++) {
    var employeeName = mainSheet.getRange(j, 1).getValue();
    var clockInTime = mainSheet.getRange(j, 2).getValue();
    var clockOutTime = mainSheet.getRange(j, 3).getValue();
    
    // Solo agregar si está clocked in (no tiene hora de salida)
    if (clockOutTime === '') {
      activeEmployees.push({
        name: employeeName,
        clockInTime: clockInTime
      });
    }
  }
  
  return activeEmployees; // Devuelve la lista de empleados activos
}

function TotalHours() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("MAIN");
  var lastRow = mainSheet.getLastRow();
  var totals = [];
  
  for (var j = 2; j <= lastRow; j++) {
    var rate = mainSheet.getRange(j, 4).getValue();
    var name = mainSheet.getRange(j, 1).getValue();
    var foundRecord = false;
    
    for (var i = 0; i < totals.length; i++) {
       if (name == totals[i][0] && rate != '') {         
         totals[i][1] = totals[i][1] + rate;
         foundRecord = true;
       }
    }
    
    if (foundRecord == false && rate != '') {
      totals.push([name, rate]);
    }
  }
  
  mainSheet.getRange("F5:G1000").clear();
  
  for (var i = 0; i < totals.length; i++) {
    mainSheet.getRange(2 + i, 6).setValue(totals[i][0]).setFontSize(12);
    mainSheet.getRange(2 + i, 7).setValue(totals[i][1]).setFontSize(12);  
  } 
}

function addZero(i) {
  if (i < 10) {
    i = "0" + i;
  }
  return i;
}

function getDate(date_in) {
  var currentDate = date_in;
  var currentMonth = currentDate.getMonth() + 1;
  var currentYear = currentDate.getFullYear();
  var currentHours = (addZero(currentDate.getHours()) > 12) ? addZero(currentDate.getHours()) - 12 : addZero(currentDate.getHours());
  var currentMinutes = addZero(currentDate.getMinutes());
  var currentSeconds = addZero(currentDate.getSeconds());
  var suffix = (addZero(currentDate.getHours()) >= 12) ? 'PM' : 'AM';
  var date = currentMonth.toString() + '/' + currentDate.getDate().toString() + '/' + 
             currentYear.toString() + ' ' + currentHours.toString() + ':' +
             currentMinutes.toString() + ':' + currentSeconds.toString() + ' ' + suffix;
  
  return date;
}
