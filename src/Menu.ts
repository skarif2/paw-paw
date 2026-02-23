/**
 * Creates a custom menu in Google Sheets on open — owner only 🐾
 * Non-owners open the sheet normally with no menu visible.
 */
function onOpen(): void {
  const owner = SpreadsheetApp.getActiveSpreadsheet().getOwner()?.getEmail();
  const me = Session.getActiveUser().getEmail();

  if (!owner || me !== owner) return; // 😾 Not the owner — slink away silently

  SpreadsheetApp.getUi()
    .createMenu('Paw-Paw 🐱')
    .addItem('Sync Roster with Slack', 'syncAllActiveSheets')
    .addItem('Send Daily Slack Briefing', 'sendDailySlackBriefing')
    .addSeparator()
    .addItem('Refresh Holiday Colors 🎌', 'refreshHolidayFormatting')
    .addSeparator()
    .addItem('Create Next Month Sheet', 'createNextMonthSheet')
    .addItem('Force Tomorrow Headcount', 'sendTomorrowHeadcount')
    .addToUi();
}
