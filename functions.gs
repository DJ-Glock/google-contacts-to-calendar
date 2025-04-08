function createBirthdayEvents() {
  // Reading all contacts of the users Google Contacts.
  // The request is paginated, as the maximum number of contacts is 1.000 otherwise.
  // So far it has been tested with around 1.200 contacts, which is working fine at
  // least once per day (too many calls will lead to failures due to limitations).

  var people = People.People;
  var connections = people.Connections.list('people/me', {
    'pageSize': 1000, // Fetch up to 1000 contacts at a time
    'personFields': 'names,birthdays,memberships,phoneNumbers,emailAddresses,events'
  });
  var contacts = connections.connections;
  var nextPageToken = connections.nextPageToken;

  while (nextPageToken) {
    var nextPage = people.Connections.list('people/me', {
      'pageToken': nextPageToken,
      'pageSize': 1000,
      'personFields': 'names,birthdays,memberships,phoneNumbers,emailAddresses,events'
    });
    contacts = contacts.concat(nextPage.connections);
    nextPageToken = nextPage.nextPageToken;
  }

  if (!contacts) {
    if (LOGGING_ENABLED) Logger.log("No contacts found, exiting.");
    return;
  }

  if (LOGGING_ENABLED) Logger.log("Number of contacts found: " + contacts.length);

  // Loop through each contact, checking if they have a birthday and are stored in "MyContacts"
  for (var i = 0; i < contacts.length; i++) {
    var person = contacts[i];
    var isMyContact = INCLUDE_ALL_CONTACTS;

    // Check if the person is in the group "MyContacts" (i.e. not a hidden contact)
    for (var j = 0; j < person.memberships.length; j++) {
      var membership = person.memberships[j];
      if (membership.contactGroupMembership.contactGroupId.includes("myContacts")) {
        isMyContact = true;
        break; // Exit the inner loop once found
      }
    }

    // Events are only created if:
    // 1. the contacts is in MyContacts or
    // 2. INCLUDE_ALL_CONTACTS in the config section is set to "true"
    if (isMyContact) {
      if (person.birthdays) {
        let tryNumber = 1
        while(true) {
          try {
            addBirthday(person)
            break
          } catch (e) {
            console.log("Failed to add birthday for " + person.names[0].displayName + "(" + tryNumber + "): " + e);
            if (tryNumber >= MAX_RETRIES_COUNT) {
              console.log("Retries exhausted")
              throw(e)
              return
            } else {
              tryNumber++
            }
          }
        }
      }

      // Creating event series for other dates of the contact (ex. wedding anniversaries)
      if (person.events && CREATE_SPECIAL_EVENTS) {
        let tryNumber = 1
        while(true) {
          try {
            addSpecialEvent(person)
            break
          } catch (e) {
            console.log("Failed to add special event for " + person.names[0].displayName + "(" + tryNumber + "): " + e);
            if (tryNumber >= MAX_RETRIES_COUNT) {
              console.log("Retries exhausted")
              throw(e)
              return
            } else {
              tryNumber++
            }
          }
        }
      }
    }
  }
}

function addBirthday(person) {
  var birthdayRaw = person.birthdays[0].date;
        var birthdayYear = DEFAULT_BIRTH_YEAR;
        let isValidBirthYear = birthdayRaw.year > MIN_YEAR;

        if (isValidBirthYear) birthdayYear = birthdayRaw.year;
        var birthdayDate = new Date(birthdayYear + '/' + birthdayRaw.month + '/' + birthdayRaw.day);

        let startTime = new Date();
        startTime.setFullYear(birthdayYear);
        startTime.setMonth(birthdayDate.getMonth());
        startTime.setDate(birthdayDate.getDate());
        startTime.setHours(10);
        startTime.setMinutes(0);
        startTime.setSeconds(0);

        let endTime = new Date();
        endTime.setFullYear(birthdayYear);
        endTime.setMonth(birthdayDate.getMonth());
        endTime.setDate(birthdayDate.getDate());
        endTime.setDate(birthdayDate.getDate());
        endTime.setHours(11);
        endTime.setMinutes(0);
        endTime.setSeconds(0);

        var yearOrigin = "";
        if (SHOW_BIRTH_YEAR && isValidBirthYear) yearOrigin = " (" + birthdayYear + ")";
        var title = "🎂 " + person.names[0].displayName + yearOrigin

        // Create a recurring event
        var event = calendar.createEventSeries(
          title,
          startTime,
          endTime,
          CalendarApp.newRecurrence().addYearlyRule().times(MAX_AGE)
        );

        if (REMINDER_EMAIL) event.addEmailReminder(0);
        if (REMINDER_POPUP) event.addPopupReminder(0);
        if (LOGGING_ENABLED) Logger.log("Birthday added for " + person.names[0].displayName);
}

function addSpecialEvent(person) {
  for (var k = 0; k < person.events.length; k++) {
    var specialEvent = person.events[k];
    let isValidEventYear = specialEvent.date.year > MIN_YEAR;

    var eventYear = DEFAYLT_EVENT_YEAR;
    if (isValidEventYear) eventYear = specialEvent.date.year;

    var displayedEventYear = ""
    if (isValidEventYear) displayedEventYear = " (" + eventYear + ")"

    var specialEventDate = new Date(eventYear + '/' + specialEvent.date.month + '/' + specialEvent.date.day);
    var specialEventName = specialEvent.formattedType;

    let title = "🥳 " + person.names[0].displayName + " (" + specialEventName + ")" + displayedEventYear

    let startTime = new Date();
    startTime.setFullYear(eventYear);
    startTime.setMonth(specialEventDate.getMonth());
    startTime.setDate(specialEventDate.getDate());
    startTime.setHours(10);
    startTime.setMinutes(0);
    startTime.setSeconds(0);

    let endTime = new Date();
    endTime.setFullYear(eventYear);
    endTime.setMonth(specialEventDate.getMonth());
    endTime.setDate(specialEventDate.getDate());
    endTime.setDate(specialEventDate.getDate());
    endTime.setHours(11);
    endTime.setMinutes(0);
    endTime.setSeconds(0);

    // Create a recurring event
    var event = calendar.createEventSeries(
      title,
      startTime,
      endTime,
      CalendarApp.newRecurrence().addYearlyRule().times(MAX_EVENT_AGE)
    );

    if (REMINDER_EMAIL) event.addEmailReminder(0);
    if (REMINDER_POPUP) event.addPopupReminder(0);
    if (LOGGING_ENABLED) Logger.log("Special event " + specialEventName + " added for " + person.names[0].displayName);
  }
}

function clearCalendar(baseYear) {
  var events = calendar.getEvents(new Date(baseYear + "/01/01"), new Date((baseYear + 1) + "/01/01"));

  for (var i in events) {
    var event = events[i];

    // Check if the event is recurring
    if (event.isRecurringEvent()) {
      // Get the ID of the recurring event
      var recurringEventId = event.getId().split('_')[0];

      // Get the recurring event by ID
      var recurringEvent = calendar.getEventSeriesById(recurringEventId);

      // Delete the entire recurring event series
      let tryNumber = 1
      while(true) {
        try {
          recurringEvent.deleteEventSeries()
          if (LOGGING_ENABLED) Logger.log("Removed " + event.getTitle());
          break
        } catch (e) {
          console.log("Failed to remove event series" + event.getTitle() + "(" + tryNumber + "): " + e);
          if (tryNumber >= MAX_RETRIES_COUNT) {
            console.log("Retries exhausted")
            throw(e)
            return
          } else {
            tryNumber++
          }
        }
      }
    } else {
      let tryNumber = 1
      while(true) {
        try {
          event.deleteEvent();
          if (LOGGING_ENABLED) Logger.log("Removed " + event.getTitle());
          break
        } catch (e) {
          console.log("Failed to remove event" + event.getTitle() + "(" + tryNumber + ")")
          if (tryNumber >= MAX_RETRIES_COUNT) {
            console.log("Retries exhausted")
            throw(e)
            return
          } else {
            tryNumber++
          }
        }
      }
    }
  }
}
