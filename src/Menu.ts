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
    .addItem('🔄 Update Slack Message...', 'promptUpdateSlackBriefing')
    .addSeparator()
    .addItem('Refresh Holiday Colors 🎌', 'refreshHolidayFormatting')
    .addSeparator()
    .addItem('Create Next Month Sheet', 'createNextMonthSheet')
    .addItem('Force Tomorrow Headcount', 'sendTomorrowHeadcount')
    .addToUi();
}

/**
 * Prompts the owner for a date and triggers a silent Slack message update. 🐾
 */
function promptUpdateSlackBriefing(): void {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "Update Slack Roll Call 🐾", 
    "Enter the date you want to update in yyyy-MM-dd format (e.g. 2026-02-25):", 
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const inputStr = response.getResponseText().trim();
    // Validate format
    if (!/^\d{4}-\d{2}-\d{2}$/.test(inputStr)) {
      ui.alert("🙀 Hiss!", "Invalid format. Dates must be exactly yyyy-MM-dd.", ui.ButtonSet.OK);
      return;
    }
    
    ui.alert("😸 On it!", `Pawing through the sheet for ${inputStr} and updating Slack...`, ui.ButtonSet.OK);
    updateHistoricSlackBriefing(inputStr);
  } else {
    ui.alert("😴 Cancelled. Going back to sleep.");
  }
}
