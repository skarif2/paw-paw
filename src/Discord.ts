/**
 * Builds the URL for a specific Discord webhook message.
 * @param {string} webhook - Base Discord webhook URL
 * @param {string} messageId - Target message ID
 * @returns {string} URL pointing to that specific message
 */
function getWebhookMessageUrl(webhook: string, messageId: string): string {
  return webhook.replace(
    /\/webhooks\/([^/]+)\/([^/]+)$/,
    `/webhooks/$1/$2/messages/${messageId}`,
  );
}

/**
 * Locks the row matching the given date so no one can edit it.
 * @param {string} date - The date to lock in `yyyy-MM-dd` format
 */
function lockRowByDate(date: string): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dateToLock = new Date(date);

  // Find the current month's sheet
  const sheetName = Utilities.formatDate(
    dateToLock,
    CONFIG.TIMEZONE,
    'yyyy-MM',
  );
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  const dateStr = Utilities.formatDate(
    dateToLock,
    CONFIG.TIMEZONE,
    'yyyy-MM-dd',
  );

  // Skip index 0 (header) and index 1 (hidden ID row) — data rows start at index 2 🐾
  for (let i = 2; i < data.length; i++) {
    const rowDate = data[i][0];

    if (rowDate instanceof Date) {
      const rowDateStr = Utilities.formatDate(
        rowDate,
        CONFIG.TIMEZONE,
        'yyyy-MM-dd',
      );

      if (rowDateStr === dateStr) {
        const rowNum = i + 1; // +1 because arrays are 0-indexed but sheet rows are 1-indexed
        const protection = sheet
          .getRange(rowNum, 1, 1, data[i].length)
          .protect();
        protection.setDescription(`Locked Past Date: ${rowDateStr}`);
        protection.removeEditors(protection.getEditors());
        if (protection.canDomainEdit()) {
          protection.setDomainEdit(false);
        }
      }
    }
  }
  console.log('🔒 Row locked — no edits allowed past this point.');
}

/**
 * Finalizes tomorrow's meal headcount, posts it to Discord, and locks the row.
 * No further changes are allowed by the team after this runs. 😼
 */
function sendTomorrowHeadcount(): void {
  try {
    const now = new Date();
    const tomorrow = new Date(now);
    tomorrow.setDate(now.getDate() + 1);

    const tomorrowStr = Utilities.formatDate(
      tomorrow,
      CONFIG.TIMEZONE,
      'yyyy-MM-dd',
    );
    const tomorrowSheetName = Utilities.formatDate(
      tomorrow,
      CONFIG.TIMEZONE,
      'yyyy-MM',
    );

    const { offdays, noFoodDays } = getDateConfig();

    const isWeekend = tomorrow.getDay() === 0 || tomorrow.getDay() === 6;
    const isOffday = offdays.some((h) => h.date === tomorrowStr);
    const isNoFoodDay = noFoodDays.some((d) => d.date === tomorrowStr);

    if (isWeekend) {
      console.log(
        `😴 Napping — ${tomorrowStr} is a non-working day, no headcount needed.`,
      );
      return;
    }

    if (isOffday || isNoFoodDay) {
      console.log(
        `🐟🚫 No food arranged for ${tomorrowStr} — skipping the Discord meal order. Slack still purrs on! 😸`,
      );
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(tomorrowSheetName);

    if (!sheet) {
      console.error(
        `🙀 Sheet ${tomorrowSheetName} has vanished! Cannot send headcount.`,
      );
      return;
    }

    const data = sheet.getDataRange().getValues();
    let yesCount: number | string = 0;

    // Skip index 0 (header) and index 1 (hidden ID row) — data rows start at index 2 🐾
    for (let i = 2; i < data.length; i++) {
      const rowDate = data[i][0];
      if (rowDate instanceof Date) {
        const rowDateStr = Utilities.formatDate(
          rowDate,
          CONFIG.TIMEZONE,
          'yyyy-MM-dd',
        );

        if (rowDateStr === tomorrowStr) {
          yesCount = data[i][data[i].length - 1];
          break;
        }
      }
    }

    // 1. Send the final count to Discord
    console.log(
      `🐟 Placing the lunch order for ${tomorrowStr}: ${yesCount} hungry humans.`,
    );
    sendOrUpdateDiscordMessage(tomorrowStr, Number(yesCount));

    // 2. Lock the row for tomorrow — no take-backs 🔒
    lockRowByDate(tomorrowStr);

    // 3. Monthly admin check: create next-next month's sheet if we're on the 25th
    if (typeof checkAndCreateFutureSheet === 'function') {
      checkAndCreateFutureSheet();
    }

    console.log(
      `😸 Purr-fect! Order placed and roster locked for ${tomorrowStr}.`,
    );

    const summary = `Headcount sent for ${tomorrowStr}. Today's row is locked.`;
    sendOwnerReport(true, 'sendTomorrowHeadcount', summary);
  } catch (error) {
    console.error(error as Error);
    // Pass the full error object so we capture the stack trace in the DM
    sendOwnerReport(false, 'sendTomorrowHeadcount', error);
  }
}

/**
 * Sends a new Discord message or updates the existing one for a given date.
 * @param {string} dateStr - Target date in `yyyy-MM-dd` format
 * @param {number} yesCount - Number of people eating in the office
 */
function sendOrUpdateDiscordMessage(dateStr: string, yesCount: number): void {
  const props = PropertiesService.getScriptProperties();
  const lastDate = props.getProperty(PROPERTY_KEYS.LAST_DATE);
  const lastMessageId = props.getProperty(
    PROPERTY_KEYS.LAST_DISCORD_MESSAGE_ID,
  );

  if (lastDate === dateStr && lastMessageId) {
    updateExistingDiscordMessage(lastMessageId, yesCount, dateStr);
  } else {
    sendNewDiscordMessage(dateStr, yesCount);
  }
}

/**
 * Generates the Discord message content based on whether it is Ramadan or not.
 * @param {string} dateStr - The date to check in `yyyy-MM-dd` format
 * @param {number} yesCount - The headcount for that day
 * @returns {string} The formatted content string to post
 */
function getMealMessageContent(dateStr: string, yesCount: number): string {
  // const { permittedHomeOffice } = getDateConfig();

  // 🐾 Check if dateStr falls broadly into any of the Permitted HO ranges
  // Assuming Permitted HO maps to Ramadan rules for meals.
  // const isPermittedHO = permittedHomeOffice.some(
  //   (pHO) => dateStr >= pHO.start && dateStr <= pHO.end,
  // );

  // if (isPermittedHO) {
  //   return `Lunch: **0**, Iftar: **${yesCount}**`;
  // }

  return `Lunch: **${yesCount}**`;
}

/**
 * Builds the URLFetch options for a Discord webhook request.
 * Centralizes the bot identity (username, avatar) and JSON content type.
 * @param {GoogleAppsScript.URL_Fetch.HttpMethod} method - HTTP method (`post`, `patch`, `delete`)
 * @param {string} content - The message content string
 * @returns {GoogleAppsScript.URL_Fetch.URLFetchRequestOptions}
 */
function buildDiscordOptions(
  method: GoogleAppsScript.URL_Fetch.HttpMethod,
  content: string,
): GoogleAppsScript.URL_Fetch.URLFetchRequestOptions {
  return {
    method,
    contentType: 'application/json',
    payload: JSON.stringify({
      username: CONFIG.DISCORD.USERNAME,
      avatar_url: CONFIG.DISCORD.AVATAR_URL,
      content,
    }),
  };
}

/**
 * Updates an existing Discord webhook message by ID.
 * Falls back to sending a new message if the original is not found (404).
 * @param {string} messageId - The Discord message ID to patch
 * @param {number} yesCount - Meal count
 * @param {string} dateStr - Date string for content generation
 */
function updateExistingDiscordMessage(
  messageId: string,
  yesCount: number,
  dateStr: string,
): void {
  const { DISCORD_WEBHOOK } = getProperties();
  const mealContent = getMealMessageContent(dateStr, yesCount);

  const patchUrl = getWebhookMessageUrl(DISCORD_WEBHOOK, messageId);
  const result = makeHttpRequest(
    `${patchUrl}?wait=true`,
    buildDiscordOptions('patch', mealContent),
  );

  if (!result.success && result.responseCode === 404) {
    console.log('🐾 Old message gone — pawing a fresh one to Discord.');
    sendNewDiscordMessage(dateStr, yesCount);
  }
}

/**
 * Posts a new meal headcount message to Discord via webhook.
 * Saves the resulting message ID to script properties for future updates.
 * @param {string} dateStr - Date string for content generation
 * @param {number} yesCount - Meal count
 */
function sendNewDiscordMessage(dateStr: string, yesCount: number): void {
  const { DISCORD_WEBHOOK } = getProperties();
  const mealContent = getMealMessageContent(dateStr, yesCount);

  const result = makeHttpRequest(
    `${DISCORD_WEBHOOK}?wait=true`,
    buildDiscordOptions('post', mealContent),
  );

  if (result.success && result.data.id) {
    const props = PropertiesService.getScriptProperties();
    props.setProperties({
      [PROPERTY_KEYS.LAST_DATE]: dateStr,
      [PROPERTY_KEYS.LAST_DISCORD_MESSAGE_ID]: result.data.id,
    });
  }
}

/**
 * Deletes a specific Discord webhook message by ID.
 * Useful for manual cleanup of stale or test messages.
 * @param {string} messageId - The Discord message ID to delete
 */
function deleteDiscordMessage(messageId: string): void {
  const { DISCORD_WEBHOOK } = getProperties();

  const deleteUrl = getWebhookMessageUrl(DISCORD_WEBHOOK, messageId);

  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: 'delete',
  };

  const result = makeHttpRequest(`${deleteUrl}?wait=true`, options);
  console.log(`🗑️ Message deleted — response: ${result.responseCode}`);
}
