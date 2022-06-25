//#region    Global Variables
globalThis.scheduledLabel = 'Calendar Reminder 📅⏳';
globalThis.targetLabel    = GmailApp.getUserLabelByName(scheduledLabel);

// Configuration Variables
globalThis.remindBeforeInHours = 3
globalThis.reminderPrefix      = 'Scheduled email'
globalThis.reminderDescSubject = 'Subject'
globalThis.reminderDescTo      = 'To'
//#endregion Global Variables

function main() {
  try {
    verifyTargetLabelExistence();

    const newScheduledThreads = getScheduled();

    for (const thread of newScheduledThreads) {
      createReminder(thread);

      // assign "targetLabel" to the scheduled email
      thread.addLabel(targetLabel);
    }
  } catch (err) {
    console.log(err);
  }
}

//#region    Functions
/**
 * Verifies that the "targetLabel" exists, otherwise creates it.
 */
function verifyTargetLabelExistence() {

  if (targetLabel == null) {
    // create the `scheduledLabel`
    GmailApp.createLabel(scheduledLabel);
    targetLabel = GmailApp.getUserLabelByName(scheduledLabel);
  }
}

/**
 * Returns now's date formatted to 'yyyy-MMM-dd HH:mm:ss'.
 */
function formatDate() {
  const d    = new Date();
  const yyyy = d.getFullYear();
  let   MMM  = d.toLocaleString('en-US', {month: 'short'});
  let   dd   = d.getDate();
  let   HH   = d.getHours();
  let   mm   = d.getMinutes();
  let   ss   = d.getSeconds();

  if (dd < 10) dd = '0' + dd;
  if (HH < 10) mm = '0' + HH;
  if (mm < 10) mm = '0' + mm;
  if (ss < 10) mm = '0' + ss;

  return `${yyyy}-${MMM}-${dd} ${HH}:${mm}:${ss}`;
}

/**
 * Returns an array with all new scheduled emails.
 * 'new' = without "targetLabel".
 */
function getScheduled() {
  const threads = GmailApp.search('is:scheduled');
  var newScheduledThreads = new Array();
  if (threads.length > 0) {
    for (const thread of threads) {
      // check if marked with the target label
      labels = thread.getLabels();
      if (labels.length > 0) {
        labelNames = [];
        for (let label in labels) {
          // add label name to labelNames array
          labelNames.push(labels[label].getName());
        }
        isNew = !labelNames.includes(scheduledLabel);
        if (!isNew) {
          continue;
        }
      }
      Logger.log(
        `Subject: ${thread.getFirstMessageSubject()}
        Scheduled for: ${thread.getMessages()[thread.getMessageCount()-1].getDate()}`
      )
      // otherwise, add
      newScheduledThreads.push(thread);
    }
  } else {
    console.log("There are no new scheduled emails upcoming.")
  }

  return newScheduledThreads;
}

/**
 * Creates a calander event related to a GMail thread.
 *
 * @param {thread} The GMail thread.
 */
function createReminder(thread) {
  let thSubject = thread.getFirstMessageSubject();
  let thReminderDate = thread.getLastMessageDate();
  let thRecipients = thread.getMessages()[0].getTo();

  // event creation
  var event = CalendarApp.getDefaultCalendar().createEvent(
    `${reminderPrefix}: "${thSubject}"`,
    new Date(thReminderDate),
    new Date(thReminderDate),
  );

  // set description
  event.setDescription(
    `<strong>${reminderDescSubject}</strong>: ${thSubject}` +
    `<br><strong>${reminderDescTo}</strong>: ${thRecipients}` +
    '<br><a href="https://mail.google.com/mail/u/0/#scheduled">Browser</a>'
    + ' | ' + `<a href="${thread.getPermalink()}">Permalink</a>`
  )

  // set reminders
  event.addEmailReminder(60 * remindBeforeInHours);
  event.addPopupReminder(60 * remindBeforeInHours);

  let daysUntilSent = (thReminderDate - new Date()) / (1000 * 3600 * 24);
  if (daysUntilSent > 1) {
    event.addEmailReminder(60 * 24);
    event.addPopupReminder(60 * 24);
  }

  // set custom tag
  event.setTag('ScheduledEmail', true);
  Logger.log(
    `Event ID: ${event.getId()}
     ScheduledEmail[tag]: ${event.getTag('ScheduledEmail')}`
  )
}
//#endregion Functions