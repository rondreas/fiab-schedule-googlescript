function schedules(){
  var spreadsheet = SpreadsheetApp.openByUrl(
     'https://docs.google.com/spreadsheets/d/1GiEqoj8pRmQ8Co1hynNg786tLlFPCRxHGdyhRo9r40w/edit');  
  var sheets = spreadsheet.getSheets();
  data = [];
  sheets.forEach(function(sheet){
    data.push(parseSheet(sheet));
  });
  return JSON.stringify(data);
}

function doGet() {
  var template = HtmlService.createTemplateFromFile("ui");
  return template.evaluate();
}

function main() {  
  var spreadsheet = SpreadsheetApp.openByUrl(
     'https://docs.google.com/spreadsheets/d/1GiEqoj8pRmQ8Co1hynNg786tLlFPCRxHGdyhRo9r40w/edit');  
  var sheets = spreadsheet.getSheets();
  var schedule = parseSheet(sheets[0]);
  Logger.log(JSON.stringify(schedule.contents.Andreas, null, 2));
}

/**
 * Get the last date in the month.
 * @return {Number} - a number 28-31 for the last date of objects month.
 */
Date.prototype.getLastDateInMonth = function() {
  var tmp_date = new Date(this.valueOf());
  tmp_date.setMonth(this.getMonth() + 1);
  tmp_date.setDate(0);
  return tmp_date.getDate();
}

/**
 * Create a Date object from Sheet name ex "March 2017" -> 2017-03-01T00:00:00
 * @param {Sheet}
 * @return {Date}
 */
function sheetToDate(sheet) {
  var name = sheet.getName();
  var date = new Date(
    parseInt(name.split(' ')[1]),
    monthToNum(name.split(' ')[0]),
    1
  );
  return date;
}

/**
 * Take a string formatted month and return the month number.
 * @param {String} - Month as a string
 * @return {Number} - Month number as according to a Date object
 */
function monthToNum(month) {
  var months = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December"
  ];  
  return months.indexOf(month);
}

/**
 * Given a Date object, create an array matching the date structure in our Sheet.
 * @param {Date}
 * @return {Array<string>} - Array of dates of given month, starting monday, blank if not in month.
 */
function getDates(date) {
  var result = [];
  
  // Date object of first date in month.
  var start = new Date(date.getYear(), date.getMonth(), 1);
  
  // Date object of last date in month.
  var end = new Date(date.getYear(), date.getMonth(), date.getLastDateInMonth());
  
  // Push empty strings for each missing day starting from Monday
  for (var i = 0; i < (start.getDay() + 6) % 7; i++ ) {
    result.push('');
  }
  
  // Push dates in month
  for (var i = start.getDate(); i <= end.getDate(); i++ ) {
    result.push(i.toString());
  }
  
  // Push empty strings for each missing day to Sunday
  for (var i = (end.getDay() + 6) % 7; i < 6; i++ ) {
    result.push('');
  }
  
  return result;
}

/**
 * Scan first column in sheet, for every non-empty cell push to array names, until a duplicate is found.
 * @param {Sheet} sheet
 * @return {Array.<string>} - Array of unique names in sheet.
 */
function getNamesInSheet(sheet) {
  var result = [];
  var range = sheet.getRange(1, 1, sheet.getLastRow());  
  var values = range.getValues();
  
  // Loop all values in column
  for (var i = 0; i < values.length; i++){
    // If value found in our array of names, break, else push to array.
    if (result.indexOf(values[i].toString()) > -1){
      break;
    } else if (values[i].toString() != ''){
      result.push(values[i].toString());
    }
  }
  return result;
}

/**
 * @param {Sheet}
 * @return {Object}
 */
function parseSheet(sheet) {
  var json = {
    name: sheet.getName(),
    id: sheet.getSheetId(),
    contents: {}
  };
  
  var rooms = {
    'B': 'Bank', 
    'X': 'Bunker', 
    'Z': 'Zombie', 
    'U': 'Upstairs'
  };
  
  var dates = getDates(sheetToDate(sheet));
  var names = getNamesInSheet(sheet);  
  var week = 0;
  
  // Get values for cells
  var range = sheet.getRange(1, 1, sheet.getLastRow(), 15);
  var values = range.getValues();
  
  // For each row
  for (var row = 0; row < values.length; row++) {
    var curName = values[row][0];
    if (curName != ''){
      
      // Slice our list of dates on the schedule to only focus on current week...
      var curWeek = dates.slice(week*7, week*7 + 7);
      
      // For each shift in the week ...
      for (var column = 1; column < values[row].length; column++) {
        var room = values[row][column];
        if (room != '') {
          
          // Every day will mean two shifts, so using floor division we can now get the date for the
          // currently found shift.
          var day = (( column - 1 ) / 2 >> 0 );
          
          // Create and set date for current column
          var curDate = sheetToDate(sheet);
          curDate.setDate(curWeek[day]);
          
          var curTimes = scheduleTimes(column);
          
          var dt_start = new Date(curDate.getFullYear(), curDate.getMonth(), curDate.getDate(), curTimes[0][0], curTimes[0][1], 0, 0);
          var dt_end = new Date(curDate.getFullYear(), curDate.getMonth(), curDate.getDate(), curTimes[1][0], curTimes[1][1], 0, 0);
          
          var shift = {room: rooms[room], start: dt_start.valueOf(), end: dt_end.valueOf()};
          
          if (json.contents[curName]){
            // If user already have shifts, push
            json.contents[curName].push(shift);
          }
          else {
            // Set the first shift
            json.contents[curName] = [shift];
          }
        }
      }
      // If we reached the last unique name, increment week counter.
      if (curName == names.slice(-1)[0]){
        week++;
      }
    }
  }
  return json;
}

/**
 * Accept a column number and return array with hour and minute
 * @param {Number} col - index of column value in row
 * @returns {Array<Number>} - [[hh, mm], [hh, mm]]
 */
function scheduleTimes(col){
  // For week days ...
  if (col < 9) {
    // First or Second shift ...
    if (col % 2 == 1) {
      return [[12, 30], [17, 0]];
    } else {
      return [[17, 0], [23, 0]];
    }
    // For weekend ...
  } else {
    // First or Second shift ...
    if (col % 2 == 1) {
      return [[9, 30], [15, 30]];
    } else {
      return [[15, 30], [23, 0]];
    }
  }
}

/** 
 * Create or update calendar with events specified in json
 * @param {Array.<Object>} schedule - json array of all shifts for a given employee
 */
function addCalendar(schedule) {
  // Hardcoded name for our calendar
  var name = "Fox in a Box"
  if (checkForCalendar(name)){
    var cal = CalendarApp.getCalendarsByName(name)[0];
    Logger.log("Calendar named %s already exist, will attempt update existing one...", name);
  } else {
    var cal = CalendarApp.createCalendar(name, {color: "#F78A1D"});
    Logger.log("Created calendar named: '%s' with the id: '%s'", cal.getName(), cal.getId());
  }
  // Go through shift in our schedule, if it already exists, do nothing. Else create the event.
  schedule.forEach(function(shift){
    var events = cal.getEvents(new Date(parseInt(shift.start)), new Date(parseInt(shift.end)));
    if (events.length > 0) {
      Logger.log("Pre-existing event; %s starting %s", shift.room, new Date(parseInt(shift.start)));
    }  else {
      Logger.log("Could not find event during that period.");
      Logger.log("Create new event; %s starting %s", shift.room, new Date(parseInt(shift.start)));
      cal.createEvent(shift.room, new Date(parseInt(shift.start)), new Date(parseInt(shift.end)));
    }
  });
}

/**
 * Search owned calendars for one matching name.
 * @param {String} caledarName - name to search for
 * @return {Boolean} - true if found, else false
 */
function checkForCalendar(calendarName){
  var calendars = CalendarApp.getAllOwnedCalendars();
  for (var i = 0; i < calendars.length; i++){
    if (calendarName == calendars[i].getName()){return true};
  }
  return false;
}

/**
 * Get all calendar events for month of date.
 * @param {Calendar} calendar
 * @param {Date} date
 * @return {Array.<CalendarEvent>}
 */
function calendarGetMonthEvents(calendar, date) {
  var events = calendar.getEvents(
    new Date(date.getYear(), date.getMonth(), 1),
    new Date(date.getYear(), date.getMonth(), date.getLastDateInMonth() + 1)
  );
  return events;
}
