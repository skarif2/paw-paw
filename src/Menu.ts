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
    .addItem('📅 Create Sheet', 'promptCreateSheetForMonth')
    .addItem('🛠️ Sync With Slack', 'promptSyncAllActiveSheets')
    .addItem('⚙️ Refresh Holidays', 'promptRefreshHolidayFormatting')
    .addSeparator()
    .addItem('📢 Daily Briefing', 'promptSendDailySlackBriefing')
    .addItem('🪄 Update Briefing', 'promptUpdateSlackBriefing')
    .addSeparator()
    .addItem('🍽️ Update Headcount', 'promptSendTomorrowHeadcount')
    .addToUi();
}

/**
 * Prompts the owner before syncing the roster. 🐾
 */
function promptSyncAllActiveSheets(): void {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Sync Roster with Slack 🔄",
    "Are you sure you want to sync the sheet with the live Slack roster?\n\nThis will add new joiners and remove leavers (based on EXCLUDED_USERS).",
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    ui.alert("😸 On it!", "Pawing through Slack and updating the sheets...", ui.ButtonSet.OK);
    syncAllActiveSheets();
  } else {
    ui.alert("😴 Cancelled. Going back to sleep.");
  }
}

/**
 * Prompts the owner before manually sending the daily Slack briefing. 🐾
 */
function promptSendDailySlackBriefing(): void {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Send Daily Slack Briefing 📢",
    "Are you sure you want to broadcast today's summary to Slack?\n\nThis will post a new message to the channel even if one was already sent today.",
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    ui.alert("😸 On it!", "Posting today's roll call to Slack...", ui.ButtonSet.OK);
    sendDailySlackBriefing();
  } else {
    ui.alert("😴 Cancelled. The channel is safe.");
  }
}

/**
 * Prompts the owner before manually sending tomorrow's headcount. 🐾
 */
function promptSendTomorrowHeadcount(): void {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Force Tomorrow Headcount 🍽️",
    "Are you sure you want to manually trigger the evening headcount?\n\nThis will calculate tomorrow's headcount, post/update it in Discord, and lock tomorrow's row.",
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    ui.alert("😸 On it!", "Crunching the numbers and updating Discord...", ui.ButtonSet.OK);
    sendTomorrowHeadcount();
  } else {
    ui.alert("😴 Cancelled. No numbers were crunched.");
  }
}

/**
 * Prompts the owner before refreshing holiday formatting across all sheets. 🐾
 */
function promptRefreshHolidayFormatting(): void {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Refresh Holiday Colors 🎌",
    "Are you sure you want to refresh holiday formatting?\n\nThis will re-apply colors, data validation, and row locks across all active and future sheets based on the current 'Holidays' tab.",
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    ui.alert("😸 On it!", "Repainting the spreadsheet with the latest holiday colors...", ui.ButtonSet.OK);
    refreshHolidayFormatting();
  } else {
    ui.alert("😴 Cancelled. The colors remain as they are.");
  }
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

/**
 * Prompts the owner for a month and triggers sheet creation. 🐾
 */
function promptCreateSheetForMonth(): void {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "Create Roster Sheet 📅",
    "Enter the month you want to create in yyyy-MM format (e.g. 2026-01):",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const inputStr = response.getResponseText().trim();

    if (!inputStr) {
      ui.alert("🙀 Hiss!", "No month provided. Going back to sleep.", ui.ButtonSet.OK);
      return;
    }

    if (!/^\d{4}-\d{2}$/.test(inputStr)) {
      ui.alert("🙀 Hiss!", "Invalid format. Month must be exactly yyyy-MM.", ui.ButtonSet.OK);
      return;
    }

    ui.alert("😸 On it!", `Pawing together the roster for ${inputStr}...`, ui.ButtonSet.OK);
    createSheetForMonth(inputStr);
  } else {
    ui.alert("😴 Cancelled. The future can wait.");
  }
}
