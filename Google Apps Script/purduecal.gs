/**
 * @OnlyCurrentDoc  Limits the script to only accessing the current document.
 */
function onOpen(e) {
  DocumentApp.getUi()
      .createAddonMenu()
      .addItem('Export to Google Calendar', 'exportSchedule')
      .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function parseDayOfWeek(day) {
  var table = {
    M: CalendarApp.Weekday.MONDAY,
    T: CalendarApp.Weekday.TUESDAY,
    W: CalendarApp.Weekday.WEDNESDAY,
    R: CalendarApp.Weekday.THURSDAY,
    F: CalendarApp.Weekday.FRIDAY
  };
  return table[day] || "TBA";
}

// "MWF" -> [CalendarApp.Weekday.MONDAY, CalendarApp.Weekday.WEDNESDAY, CalendarApp.Weekday.FRIDAY]
// "TBA" -> []
function parseDays(s) {
  if (s === "TBA") {
    return [];
  } else {
    var days = [];
    for (var i = 0; i < s.length; i++) {
      days.push(parseDayOfWeek(s.charAt(i)));
    }
    return days;
  }
}

function getCalendar() {
  var name = "My Classes";
  var calendars = CalendarApp.getCalendarsByName(name);
  if (calendars.length > 0) {
    return calendars[0];
  } else {
    return CalendarApp.createCalendar(name);
  }
}

function createDateTime(date, time) {
  return new Date(date + " " + time);
}

// via http://stackoverflow.com/a/3639412
function getNextWeekDay(date, num) {
  return new Date(date.getTime() + ((num - date.getDay() + 7) % 7 + 1) * 86400000);
}

function getClassesForWeek(date, days) {
  var table = ["M", "T", "W", "R", "F"];
  var res = [];
  for (var i = 0; i < days.length; i++) {
    var first_day = days.charAt(i);
    var week_day = table.indexOf(first_day);
    if (week_day == 0) res.push(date);
    else res.push(getNextWeekDay(date, week_day));
  }
  return res;
}

function exportSchedule() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var text = body.getText();
  var data = JSON.parse(text);
  
  for (var i = 0; i < data.length; i++) {
    // Parse the JSON
    var e = data[i];
    var crn = e["CRN"];
    var course = e["Course"];
    var title = e["Title"];
    var campus = e["Campus"];
    var credits = e["Credits"];
    var level = e["Level"];
    var location = e["Location"];
    var instructor = e["Instructor"];
    var days = parseDays(e["Days"]);
    var combined_time = e["Time"];
    if (combined_time != "TBA") {
      var start_end_split = combined_time.split(" - ");
      
      // Start & End Times
      var start_time = parseTime(start_end_split[0]);
      var end_time = parseTime(start_end_split[1]);
      
      var cal_start_date_time = getClassesForWeek(createDateTime(e["Start Date"], start_time), e["Days"]);
      var cal_end_date_time = getClassesForWeek(createDateTime(e["Start Date"], end_time), e["Days"]);
                        
      // Add events to calendar
      var calendar = getCalendar();      
      calendar.setTimeZone("America/New_York");
      
      for (var j = 0; j < cal_start_date_time.length; j++) {
        var start_date_time = cal_start_date_time[j];
        var end_date_time = cal_end_date_time[j];
        
        var recurrence = CalendarApp.newRecurrence().addWeeklyRule()
          .onlyOnWeekday(days[j])
          .until(createDateTime(e["End Date"], end_time));
        var series = calendar.createEventSeries(course, start_date_time, end_date_time, recurrence);
      
        var description = title + "\n" +
          "\nInstructor: " + instructor +
          "\nCampus: " + campus +
          "\nCRN: " + crn +
          "\nCredits: " + credits +
          "\nLevel: " + level;
        series.setDescription(description);
        series.setLocation(location);
      }
    }
  }
  openDialog();
}

function parseTime(time) {
  var am_pm_split = time.split(" ");
  var hr_min_split = am_pm_split[0].split(":");
  var hours = parseInt(hr_min_split[0]);
  var minutes = parseInt(hr_min_split[1]);
  var am_pm = am_pm_split[1];
  // via http://stackoverflow.com/a/15083891
  if (am_pm == "pm" && hours < 12) hours = hours + 12;
  if (am_pm == "am" && hours == 12) hours = hours - 12;
  var sHours = hours.toString();
  var sMinutes = minutes.toString();
  if (hours < 10) sHours = "0" + sHours;
  if (minutes < 10) sMinutes = "0" + sMinutes;
  return sHours + ":" + sMinutes;
}

function openDialog() {
  var html = HtmlService.createHtmlOutputFromFile('index')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  DocumentApp.getUi()
     .showModalDialog(html, 'Success!');
}