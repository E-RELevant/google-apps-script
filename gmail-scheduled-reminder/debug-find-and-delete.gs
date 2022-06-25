function DebugFindAndDelete() {
events = CalendarApp.getDefaultCalendar().getEvents( new Date('January 01, 1900 00:00:00 UTC'), new Date('December 31, 2100 00:00:00 UTC'))
         .filter(e => e.getTag('ScheduledEmail'));
  console.log(`Number of events: ${events.length}`);
  for (i in events) {events[i].deleteEvent()}
}