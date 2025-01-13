// SOURCE: https://github.com/qoomon/GoogleContactsEventsToGoogleCalendarSync

// # INSTRUCTION Initial setup...
// 1) Add People and Calendar service by clicking on the "+"" icon next to "Services" at the left pannel.

// # INSTRUCTION Sync birthdays and special events from your Google Contacts into any Google Calendar...
// 1) Adjust the "getContactsEventLocalization" funtion and the "const Config = {...}" object below before you proceed.
// 2) Click "Save project to Drive" above afterwards
// 3) Run this script for the first time...
//   1) Select "run_syncEvents" in the dropdown menu above, then click "Run"
//.  2) Click "Advanced" during warnings to proceed and grant permissions to this script to access your contacts and calanders.
// 4) Create a "Trigger" to run "syncEvents" daily...
//   1) Click "Triggers" in the most left pannel
//   2) Click "Add Trigger" Button at the bottom right
//.  3) Select "run_syncEvents" for "Choose which function to run"
//.  3) Select "Day Timer" for "Select type of time based trigger"
//.  4) Click "Save" Button

// # INSTRUCTION Remove all synced events...
// 1) Select "run_removeEvents" in the dropdown menu above, then click "Run"

const userLocale = getUserLocale();
const ContactsEventLocalization = getContactsEventLocalization(userLocale);
console.info("User Locale:", userLocale + ":", ContactsEventLocalization);

const Config = {
  // --- Google Contacts ---
  contacts: {
    // If undefined all contacts are synced
    // If set only contacts with that label are synced
    //   To get the contactsLabelId...
    //   - Open https://contacts.google.com/
    //   - Click on any contact label on the left pannel,
    //     the last part of the url address is the contactsLabelId (https://contacts.google.com/label/[contactsLabelId]?...)
    labelId: "CHANGE_ME",
    // Only those contact event types are synced. Add custom labels if needed.
    annualEventTypes: [
      ContactsEventLocalization.birthday,
      ContactsEventLocalization.anniversary,
    ],
  },
  // --- Google Calendar ---
  calendar: {
    // Target calendar for contact events. Set to "primary" for the default calendar.
    //   To get the calendarId for a calendar...
    //   - Open https://calendar.google.com/
    //   - Hover over any of your calenders you have write premissions
    //   - Click on the 3 dot menu and then click on "Settings and sharing"
    //   - Sroll down to "Integrate calendar" > "Calendar ID"
    id: "CHANGE_ME@group.calendar.google.com",
    eventSummaryPrefix: "⌘ ", // ⌘, ❖, ✱, 
  },
};
Config.calendar.eventSummaryPrefix+= " "; // Always add a Thin Space (U+2009) for design purpose;

function getContactsEventLocalization(userLocale) {
  const localization =  {
    "en": { birthday: "Birthday", anniversary: "Anniversary" },
    "de": { birthday: "Geburtstag", anniversary: "Jahrestag" },
  }[userLocale];
  if(!localization) {
    throw new Error(`Unsupported localization '${userLocale}'.` +
      `\nAdd localization entry for '${userLocale}' at \`function getContactsEventLocalization\` .`);
  };
}

// --- main methods START ---

const CALENDAR_CONTACTS_EVENTS_SOURCE = "contacts";

function run_debug(event) {
  console.log("event:", JSON.stringify(event, null ,2));
}

function run_syncEvents() {
  try {
    const contactsEvents = getContactsEvents({
      types: Config.contacts.annualEventTypes,
      labelId: Config.contacts.labelId,
    });
    console.info("Contacts events count: " + contactsEvents.length);

    const calendarContactsEvents = getCalendarContactsEvents({
      calendarId: Config.calendar.id,
    });
    // TODO handle event.status !== "cancelled" (single instance events from recurring events, after edit single event)
    //      update status to "confirmed" or delete parent event
    console.info("Calendar Contact events count: " + calendarContactsEvents.length);

    // --- remove legacy calendar events ---
    const contactsEventIdSet = new Set(contactsEvents.map((event) => event.id));
    const calendarContactsEventsToDelete = calendarContactsEvents.filter((event) => !contactsEventIdSet.has(event.extendedProperties.private.contactEventId));
    console.info("Calendar Contact events to delete: " + calendarContactsEventsToDelete.length);
    calendarContactsEventsToDelete.forEach((calendarEvent) => {
      removeCalendarEvent(Config.calendar.id, calendarEvent);
    });

    // --- create or update calendar events ---
    contactsEvents.forEach((contactEvent) => {
      createOrUpdateCalendarEventFromContactEvent(Config.calendar.id, contactEvent);
    });
  } catch (error) {
    console.error("ERROR", error.stack);
  }
}

function run_removeEvents() {
  const events = getCalendarContactsEvents({
    calendarId: Config.calendar.id,
  });
  console.info("events: ", events.length);

  events.forEach((event) => {
    removeCalendarEvent(Config.calendar.id, event);
  });
}

// --- main methods END ---

function getCalendarContactsEvents({ calendarId, privateExtendedProperties }) {
  const result = [];

  let nextPageToken = null;
  do {
    const response = Calendar.Events.list(calendarId, {
      privateExtendedProperty: Object.entries(Object.assign({
          source: CALENDAR_CONTACTS_EVENTS_SOURCE,
        }, privateExtendedProperties ?? {}))
        .map(([key, value]) => `${key}=${value}`),
      pageToken: nextPageToken,
    });

    nextPageToken = response.nextPageToken;

    result.push(...response.items);
  } while (nextPageToken);

  return result;
}

function getContactsConections({ labelId }) {
  const result = [];

  let nextPageToken = null;
  do {
    const response = People.People.Connections.list("people/me", {
      personFields: "names,birthdays,events,memberships",
      pageToken: nextPageToken,
    });
    nextPageToken = response.nextPageToken;

    const connections = labelId
      ? response.connections?.filter((connection) =>
          connection.memberships?.some((membership) => membership.contactGroupMembership?.contactGroupId === labelId))
      : response.connections;
    result.push(...connections);
  } while (nextPageToken);

  return result;
}

function getContactsEvents({ labelId, types }) {
  types = types.map((type) => type.toLowerCase());
  return getContactsConections({ labelId })
    .flatMap(getContactEvents)
    .filter((event) => types.includes(event.type.toLowerCase()));
}

function getContactEvents(connection) {
  const contact = {
    resourceName: connection.resourceName,
    name: connection.names?.[0].displayName,
  };
  if (!contact.name) {
    console.warn("skip connection without name");
    return [];
  }

  const contactEventTypes = new Set();
  const events = [];

  const birthday = connection.birthdays?.[0];
  if (birthday) {
    if(connection.birthdays?.length > 1){
      console.warn(`Ambigous birthday from ${contact.name}`);
    }
    events.push({
      type: ContactsEventLocalization.birthday,
      date: birthday.date,
    });
  }

  // Special Events
  connection.events?.forEach((connectionEvent) => {
    const eventLabel = connectionEvent.formattedType;
    if (!eventLabel) {
      console.warn(`skip event without label from ${contact.name}`);
      return;
    }

    if(contactEventTypes.has(eventLabel)) {
      console.warn(`skip ambigous ${eventLabel} from ${contact.name}`);
      return;
    }
    contactEventTypes.add(eventLabel);

    events.push({
      type: eventLabel,
      date: connectionEvent.date,
    });
  });

  // enrich event
  events.forEach((event) => {
    event.contact = contact;
    event.id = buildContactEventId(event.type, event.contact.resourceName);

    event.summary = `${event.contact.name}'s ${event.type}`;
    if(event.date.year){
      event.summary += ` (${event.date.year})`;
    }
  });

  return events;

  function buildContactEventId(type, resourceName) {
    const value = `${resourceName}-${type}`;
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value);
    return Utilities.base64EncodeWebSafe(digest).replace(/=+$/,"");
  }
}

function createOrUpdateCalendarEventFromContactEvent(calendarId, contactEvent) {
  // NOTE as of now (2025-01-01) there is no way to determine the creation date of the contact, therefore we use 1970 as the event start date
  const contactEventDate = new Date([
    contactEvent.date.year ?? 1970,
    String(contactEvent.date.month).padStart(2, "0"),
    String(contactEvent.date.day).padStart(2, "0"),
  ].join("-"));

  const calendarEvent = {
    eventType: "default",
    summary: `${Config.calendar.eventSummaryPrefix ?? "" }${contactEvent.summary}`,
    start: { date: contactEventDate.toISOString().split("T")[0] },
    end: { date: nextDay(contactEventDate).toISOString().split("T")[0] },
    recurrence: [(contactEvent.date.month === 2 && contactEvent.date.day === 29)
      ? "RRULE:FREQ=YEARLY;BYMONTH=2;BYMONTHDAY=-1" // Exception for Feb 29th!
      : "RRULE:FREQ=YEARLY"],
    description: `<a href="https://contacts.google.com/person/${contactEvent.contact.resourceName.replace(/^people\//,'')}"><b>Google Contacts</b></a>`,
    transparency: "transparent", // The event does not block time on the calendar.
    visibility: "private",
    extendedProperties: {
      private: {
        source: CALENDAR_CONTACTS_EVENTS_SOURCE,
        contactEventId: contactEvent.id,
      }
    }
  };

  const existingEvents = getCalendarContactsEvents({
    calendarId,
    privateExtendedProperties: {
      contactEventId: contactEvent.id,
    },
  });
  if(existingEvents.length > 1) {
    throw new Error(`Ambiguous ${contactEvent.type} from ${contactEvent.contact.name} on ${contactEventDate.toISOString().split("T")[0]}`);
  }

  const existingEvent = existingEvents[0];
  if (!existingEvent) {
    console.info(`Create ${contactEvent.type} from ${contactEvent.contact.name} on ${contactEventDate.toISOString().split("T")[0]}`);
    Calendar.Events.insert(calendarEvent, calendarId);
  } else {
    const eventHasUpdates = calendarEvent.summary !== existingEvent.summary ||
      calendarEvent.start.date !== existingEvent.start.date ||
      calendarEvent.end.date !== existingEvent.end.date ||
      calendarEvent.recurrence[0] !== existingEvent.recurrence[0] ||
      calendarEvent.description !== existingEvent.description;
    if(eventHasUpdates) {
      console.info(`Update ${contactEvent.type} from ${contactEvent.contact.name} on ${contactEventDate.toISOString().split("T")[0]}`);
      Calendar.Events.update(calendarEvent, calendarId, existingEvent.id);
    } else {
      console.log(`No changes for ${contactEvent.type} from ${contactEvent.contact.name} on ${contactEventDate.toISOString().split("T")[0]}`);
    }
  }
}

function removeCalendarEvent(calendarId, event) {
  console.info(`Remove '${event.summary}' on ${event.start.date}`);
  Calendar.Events.remove(calendarId, event.id);
}

function getUserLocale() {
  let userLocale = Session.getActiveUserLocale();
  if(userLocale) {
    console.info(`Store user locale '${userLocale}'.`);
    PropertiesService.getUserProperties().setProperty("userLocale", userLocale);
  } else {
    userLocale = PropertiesService.getUserProperties().getProperty("userLocale");
    if(!userLocale) {
      throw new Error("Could not determine user locale. Run this script manually once.");
    }
  }
  return userLocale;
}

function nextDay(date) {
  const nextDayDate = new Date(date);
  nextDayDate.setDate(date.getDate() + 1);
  return nextDayDate;
}
