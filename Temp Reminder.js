var nameRow = 2;
var emailRow = 3;
var firstDateRow = 4;
var reportHour = 16;
var noWorkDay = "STAT";

var timeZone = "Canada/Eastern";
var sheetName = "Daily Temperature Checks";
var sheetLink = "https://docs.google.com/spreadsheets/d/1HZCjDJNtgJ7egpnmx24MBCMHcNC8GifVOKrbZrmXKWQ/edit#gid=0/";

var emailSubject = "! Temperature Reminder !";
var emailMessageStart = "Please measure your temperature and\nreport it or enter it directly at:";
var emailMessageEnd = "Thank you and stay safe!"

var emailReports = ["artiomtz@gmail.com"]; // add recipients with commas
var emailReportSubject = "Temperature End Of Day Report";
var emailReportMessage = "Employees which did not report their temperature for today:";


function TempReminder()
{
  try
  { 
    var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName)
    var data = sheet.getDataRange().getValues();
  }
  catch(e)
  {
    return;
  }  
  var todayRow = todaysRow(data);
  if (todayRow)
    CheckToday(sheet, todayRow);
}


function CheckToday(sheet, todayRow)
{
  var firstEntry = sheet.getRange(todayRow, 2).getValues().toString().toUpperCase();
  if (firstEntry == noWorkDay)
    return;
  
  var numColumns = sheet.getLastColumn();
  var TimeHourNow = Number(Utilities.formatDate(new Date(), timeZone, "HH"));
  var didntTakeTemp = [];
  
  for(i=2; i<=numColumns; ++i)
  {
    if (isCellEmpty(sheet, todayRow, i))
    {
      if (TimeHourNow < reportHour)
        sendNotification(sheet, i, todayRow);
      else 
        didntTakeTemp.push(i);
    }
  }
  if (didntTakeTemp.length)
    endOfDayReport(sheet, didntTakeTemp);
}


function sendNotification(sheet, column, todayRow)
{
  var name = sheet.getRange(nameRow, column).getValues().toString();
  var email = sheet.getRange(emailRow, column).getValues().toString();
  var subject = emailSubject;
  var message = "Hi " + name + "\n\n" + emailMessageStart + "\n" + sheetLink + "&range=A" + todayRow + " \n\n" + emailMessageEnd;
  try
  {
    MailApp.sendEmail(email, subject, message);
  }
  catch(e) {}
}


function endOfDayReport(sheet, didntTakeTemp)
{
  var message = emailReportMessage + "\n\n";  
  for(i=0; i<didntTakeTemp.length; ++i)
  {
    var employee = sheet.getRange(nameRow, didntTakeTemp[i]).getValues().toString();
    if (employee != "")
      message += employee + "\n";
    else
      message += sheet.getRange(emailRow, didntTakeTemp[i]).getValues().toString() + "\n";
  }
  message += "\nThank you."
  try
  {
    for (var reportEmail in emailReports)
      MailApp.sendEmail(emailReports[reportEmail], emailReportSubject, message);
  }
  catch(e) {}
}


function todaysRow(data)
{
  var day = Utilities.formatDate(new Date(), timeZone, "E");
  if (day == "Sat" || day == "Sun")
    return 0;
  
  var today = Utilities.formatDate(new Date(), timeZone, "dd/MM/yyyy");
  for(i=firstDateRow-1; i<data.length; ++i)
  {
    try
    {
      var workDay = Utilities.formatDate(data[i][0], timeZone, "dd/MM/yyyy");
      if (today == workDay)
        return i + 1;      
    }
    catch(e)
    {
      break;
    }
  }
  return 0;
}


function isCellEmpty(sheet, row, column)
{
  var value = sheet.getRange(row, column).getValues();
  
  if (value[0].toString() == "")
    return true;
  else
    return false;
}
