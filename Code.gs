// # INSTRUCTION to sync birthdays and special events from your Google Contacts into any Google Calendar...
// 1) Copy of this project
// 2) Adjust the "config" object below before you proceed, click "Save project to Drive" above afterwards
// 3) Run this script for the first time...
//   1) Select "run_syncEvents" in the dropdown menu above, then click "Run"
//.  2) Click 'Advanced' during warnings to proceed and grant permissions to this script to access your contacts and calanders.
// 4) Create a "Trigger" to run "syncEvents" daily...
//   1) Click "Triggers" in the most left pannel
//   2) Click "Add Trigger" Button at the bottom right
//.  3) Select "run_syncEvents" for "Choose which function to run"
//.  3) Select "Day Timer" for "Select type of time based trigger"
//.  4) Click "Save" Button

// # INSTRUCTION to remove all synced events...
// 1) Select "run_removeEvents" in the dropdown menu above, then click "Run"  

const config = {
  // If undefined all contacts are synced
  // If set only contacts with that label are synced
  //   To get the contactLabelId...
  //   - Open https://contacts.google.com/ 
  //.  - Click on any contact label on the left pannel,
  //     the last part of the url address is the contactLabelId (https://contacts.google.com/label/[contactLabelId]?...)
  contactLabelId: undefined,
  // Only those contact event types are synced. Add Custom ones if needed.
  contactAnnualEventTypes: ['Birthday', 'Anniversary'],

  // Target calendar for contact events. Set to 'primary' for the default calendar.
  //   To get the calendarId for a calendar...
  //.  - Open https://calendar.google.com/
  //   - Hover over any of your calenders you have write premissions
  //   - Click on the 3 dot menu and then click on "Settings and sharing"
  //   - Sroll down to "Integrate calendar" > "Calendar ID"
  calendarId: "primary",
};

// --- main methods START ---

function run_syncEvents() {
  try {
    const contactsEvents = getContactsEvents({
      types: config.contactAnnualEventTypes,
      labelId: config.contactLabelId,
    });
    console.info("contactsEvents: " + contactsEvents.length);

    const calendarContactsEvents = getCalendarContactsEvents({
      calendarId: config.calendarId,
    });
    console.info("calendarContactsEvents: " + calendarContactsEvents.length);

    // --- remove legacy calendar events ---
    const contactsEventIdSet = new Set(contactsEvents.map((event) => event.id));
    const calendarContactsEventsToDelete = calendarContactsEvents.filter((event) => !contactsEventIdSet.has(event.extendedProperties.private.contactEventId));
    console.info("calendarContactsEventsToDelete: " + calendarContactsEventsToDelete.length);
    calendarContactsEventsToDelete.forEach((calendarEvent) => {
      removeCalendarEvent(config.calendarId, calendarEvent);
    });

    // --- create or update calendar events ---
    contactsEvents.forEach((contactEvent) => {
      createOrUpdateCalendarEventFromContactEvent(config.calendarId, contactEvent);
    });
  } catch (error) {
    console.error("ERROR", error.stack);
  }
}

function run_removeEvents() {
  const events = getCalendarContactsEvents({
    calendarId: config.calendarId,
  });
  console.info("events: ", events.length);

  events.forEach((event) => {
    removeCalendarEvent(config.calendarId, event);
  });
}

// --- main methods END ---

function getCalendarContactsEvents({ calendarId, privateExtendedProperties }) {
  return Calendar.Events.list(calendarId, {
    privateExtendedProperty: Object.entries(Object.assign({
        source: 'contacts',
      }, privateExtendedProperties ?? {}))
      .map(([key, value]) => `${key}=${value}`),
  }).items;
}

function getContactsConections({ labelId }) {
  const result = [];

  const pageSize = 100;
  let nextPageToken = null;
  do {
    const response = People.People.Connections.list("people/me", {
      personFields: "names,birthdays,events,memberships",
      pageSize,
      pageToken: nextPageToken,
    });
    nextPageToken = response.nextPageToken;

    const connections = labelId
      ? response.connections?.filter((connection) =>
          connection.memberships?.some((membership) => membership.contactGroupMembership?.contactGroupId === labelId),
        )
      : response.connections;
    result.push(...connections);
  } while (nextPageToken);

  return result;
}

function getContactsEvents({ labelId, types }) {
  types = types.map((type) => type.toLowerCase())
  return getContactsConections({ labelId })
    .flatMap(getContactEvents)
    .filter((event) => types.includes(event.type.toLowerCase()));
}

function getContactEvents(connection) {
  const contactName = connection.names?.[0].displayName;
  if (!contactName) {
    console.warn("skip connection without name");
    return [];
  }

  const contact = {
    resourceName: connection.resourceName,
    name: contactName,
  };

  const events = [];

  // Gather Birthday
  {
    if(connection.birthdays?.length > 1){
      console.error(`Ambigous birthday from ${contactName}`)
    }

    const birthday = connection.birthdays?.[0]
    if (birthday) {
      let summary = `${contactName}'s Birthday`;
      if (birthday.date.year) {
        summary += ` (${birthday.date.year})`;
      }
      const event = {
        type: "birthday",
        summary,
        date: birthday.date,
        contact,
      };
      event.id = buildContactEventId(event);
      events.push(event);
    }
  }

  // Gather Special Events
  {
    const eventTypes = new Set();
    connection.events?.forEach((connectionEvent) => {
      const eventLabel = connectionEvent.formattedType;
      if (!eventLabel) {
        console.warn(`skip event without label from ${contactName}`);
        return;
      }

      if(eventTypes.has(eventLabel)) {
        console.warn(`skip ambigous ${eventLabel} from ${contactName}`);
        return;
      }
      eventTypes.add(eventLabel);

      let summary = `${contactName}'s ${eventLabel}`;
      if(connectionEvent.date.year){
        summary += ` (${connectionEvent.date.year})`;
      }
      const event = {
        type: eventLabel,
        summary,
        date: connectionEvent.date,
        contact,
      };
      event.id = buildContactEventId(event);
      events.push(event);
    });
  }

  return events;

  function buildContactEventId({contact, type, date}) {
    const value = `${contact.resourceName}-${type}-${date.year ?? '0000'}-${String(date.month).padStart(2, '0')}-${String(date.day).padStart(2, '0')}`;
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value);
    return Utilities.base64EncodeWebSafe(digest).replace(/=+$/,'')
  }
}

function createOrUpdateCalendarEventFromContactEvent(calendarId, contactEvent) {
  // TODO handle no year, use contact creation date
  const contactEventDate = new Date(`${contactEvent.date.year ?? 1970}-${String(contactEvent.date.month).padStart(2, '0')}-${String(contactEvent.date.day).padStart(2, '0')}`);

  const calendarEvent = {
    eventType: 'default',
    summary: `★ ${contactEvent.summary}`,
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
        source: 'contacts',
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

function nextDay(date) {
  const nextDayDate = new Date(date);
  nextDayDate.setDate(date.getDate() + 1);
  return nextDayDate;
}
