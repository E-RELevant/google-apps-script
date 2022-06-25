function createTimeDrivenTriggers() {
  // Trigger every 1 hour.
  ScriptApp.newTrigger('main')
      .timeBased()
      .everyHours(1)
      .create();
}