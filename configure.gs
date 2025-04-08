// FUNCTIONALITY
// This script takes all contact in "MyContacts" (i.e. not hidden ones) which have a birthday added and
// creates a recurring event in the calendar. The script can run scheduled (ex. every night) to update 
// changes from the contacts. Rather than an update, all exisitng entries will be deleted and the creation
// is done for each event again.

// KNOWN LIMITATIONS
// * The script is NOT optimized for performance and resources or big amounts of contacts 
//   (tests go as far as around 1.200 contacts).
// * There are limitations as to how many events can be created, how long scripts can run and others, that
//   vary depending on the account type and maybe other factors.


// Create events for "significant dates" of the contact (other than birthdays)
let CREATE_SPECIAL_EVENTS = true;

// Include all contacts (skipping hidden and other contacts)
let INCLUDE_ALL_CONTACTS = false;

// Show the birth year of the contact in the calendar entry 
let SHOW_BIRTH_YEAR = true;

// Create reminder for the events x minutes before the event, 0 means at the time of the event -1 does not set a reminder
let REMINDER_EMAIL = false;
let REMINDER_POPUP = true;

// Number of events for each birthday, the first being in the year of the birth:
let MAX_AGE = 200;
let MAX_EVENT_AGE = 200;

// Minimum event or birth year
let MIN_YEAR = 1900;
let DEFAULT_BIRTH_YEAR = 2000
let DEFAYLT_EVENT_YEAR = 1900 

// Name of the calendar in which the birthdays should be added.
// Please use this calendar for the birthdays only: During the upgrade, all events in this calendar will be deleted!
var CALENDAR_NAME = "BirthdaysCal"; 

// Output log entries (true/false)
var LOGGING_ENABLED = true;

// Max retries count
let MAX_RETRIES_COUNT = 3

var calendar;

function updateBirthdays() {
  if (CalendarApp.getCalendarsByName(CALENDAR_NAME)[0] == null) {
    if (LOGGING_ENABLED) Logger.log("Creating calendar " + CALENDAR_NAME);
    CalendarApp.createCalendar(CALENDAR_NAME);
  }
  calendar = CalendarApp.getCalendarsByName(CALENDAR_NAME)[0];

// Standard is 200 years of appointments, so this 4 calls should cover all generated events:
  clearCalendar(1899);
  clearCalendar(1999);
  clearCalendar(2099);
  clearCalendar(2199);
  createBirthdayEvents_v2();
}

// Function required for Web App Deployement
function doGet() {
  updateBirthdays();
  return HtmlService.createHtmlOutput().createHtmlOutput('<b>All done, thanks for using Cal-to-Con!</b>');
}
