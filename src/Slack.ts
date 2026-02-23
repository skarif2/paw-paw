/**
 * Fetches Slack user info for a list of user IDs in batches.
 * @param {string[]} userIds - Array of Slack user IDs to look up
 * @returns {Record<string, SlackUser>} Map of user ID → SlackUser
 */
function getUserInfoBatch(userIds: string[]): Record<string, SlackUser> {
  const { SLACK_TOKEN } = getProperties();
  const userMap: Record<string, SlackUser> = {};

  // 🐱 Process in small batches — even cats don't chase all mice at once
  const batchSize = 10;
  for (let i = 0; i < userIds.length; i += batchSize) {
    const batch = userIds.slice(i, i + batchSize);

    batch.forEach((userId) => {
      try {
        const url = `${CONFIG.SLACK_API_BASE}/users.info?user=${userId}`;
        const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
          headers: { Authorization: `Bearer ${SLACK_TOKEN}` },
        };

        const result = makeHttpRequest(url, options);

        if (result.success && result.data.ok) {
          const user = result.data.user;

          // Filter out bots and deleted accounts — only real humans eat lunch 🐟
          if (!user.is_bot && !user.deleted) {
            userMap[userId] = {
              name: user.profile.display_name_normalized || user.profile.real_name_normalized || user.real_name || user.name,
              email: user.profile.email,
              image: user.profile.image_original
            }
          }
        }
      } catch (error) {
        console.warn(`😿 Couldn't sniff out user ${userId}:`, (error as Error).toString());
      }
    });

    // Brief nap between batches to stay on Slack's good side 😴
    if (i + batchSize < userIds.length) {
      Utilities.sleep(100);
    }
  }

  return userMap;
}


/**
 * Fetches all members of the configured Slack channel, paginating as needed.
 * @returns {Record<string, SlackUser>} Map of user ID → SlackUser
 */
function getChannelUsers(): Record<string, SlackUser> {
  const { SLACK_TOKEN, SLACK_CHANNEL_ID } = getProperties();
  const userMap: Record<string, SlackUser> = {};
  let cursor = '';
  const MAX_PAGES = 5; // 5 × 100 = 500 members max, purr-fectly safe 🐾
  let page = 0;

  do {
    const url = `${CONFIG.SLACK_API_BASE
      }/conversations.members?channel=${SLACK_CHANNEL_ID}&limit=100${cursor ? '&cursor=' + cursor : ''
      }`;
    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
      headers: { Authorization: `Bearer ${SLACK_TOKEN}` },
    };

    const result = makeHttpRequest(url, options);

    if (!result.success || !result.data.ok) {
      throw new Error(`Failed to get channel members: ${result.data?.error || result.error}`);
    }

    const userIds: string[] = result.data.members;
    const userInfoBatch = getUserInfoBatch(userIds);

    Object.assign(userMap, userInfoBatch);

    cursor = result.data.response_metadata?.next_cursor || '';
    page++;
  } while (cursor && page < MAX_PAGES);

  console.log('🐾 Channel roster fetched:', userMap);
  return userMap;
}

/**
 * Sends the daily Paw-Paw briefing to the configured Slack channel.
 * Posts a status summary message and kicks off a thread for the day's chat.
 */
function sendDailySlackBriefing(): void {
  const { SLACK_TOKEN, SLACK_CHANNEL_ID } = getProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = new Date();
  const todayStr = Utilities.formatDate(today, CONFIG.TIMEZONE, "yyyy-MM-dd");

  // 1. QUICK EXIT: Skip weekends, holidays, and off-days — the cat is napping 😴
  const dayOfWeek = today.getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    console.log("😴 It's the weekend — this cat is napping. No briefing today.");
    return;
  }

  if (getDateConfig().offdays.includes(todayStr)) {
    console.log("😴 It's an off-day — paws up. No briefing today.");
    return;
  }

  // 2. Get the current month's sheet
  const sheetName = Utilities.formatDate(today, CONFIG.TIMEZONE, "yyyy-MM");
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    console.warn(`🙀 No sheet found for ${sheetName} — where did it go?!`);
    return;
  }

  // 3. Find today's data row
  const data = sheet.getDataRange().getValues();
  let todayRow: any[] | null = null;
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    const rowDate = data[i][0];
    if (rowDate instanceof Date) {
      const rowDateStr = Utilities.formatDate(rowDate, CONFIG.TIMEZONE, "yyyy-MM-dd");
      if (rowDateStr === todayStr) {
        todayRow = data[i];
        break;
      }
    }
  }

  // Fallback: if the sheet itself marks this as a non-working day ("-"), bail out
  if (!todayRow || todayRow[1] === "-") {
    console.log("😿 No valid data found for today. Staying quiet.");
    return;
  }

  // 4. Group members by attendance status
  const groups: Record<string, string[]> = { "Office": [], "WFH": [], "Leave": [] };

  for (let col = 1; col < headers.length - 1; col++) {
    const status = todayRow[col];
    const memberName = headers[col];

    if (groups[status]) {
      groups[status].push(memberName);
    }
  }

  // 5. Construct the message
  const dateHeading = Utilities.formatDate(today, CONFIG.TIMEZONE, "EEEE, MMM d");

  const officeList = groups["Office"].length > 0 ? groups["Office"].join(", ") : "_None_";
  const wfhList = groups["WFH"].length > 0 ? groups["WFH"].join(", ") : "_None_";
  const leaveList = groups["Leave"].length > 0 ? groups["Leave"].join(", ") : "_None_";

  const messageText = `<!here> | *${dateHeading}*\n\n` +
    `*🏢 ON-SITE:* ${officeList}\n` +
    `*🏠 WFH:* ${wfhList}\n` +
    `*🌴 MIA:* ${leaveList}\n\n` +
    `Please post your *status* (Starting, AFK, Back, etc.) in the thread and feel free to *chitchat!*`

  // 6. Send the main message to Slack
  const url = `${CONFIG.SLACK_API_BASE}/chat.postMessage`;
  const mainPayload = {
    channel: SLACK_CHANNEL_ID,
    text: messageText
  };

  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${SLACK_TOKEN}`
    },
    payload: JSON.stringify(mainPayload)
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());

    if (!result.ok) {
      console.error("🙀 Slack hissed back at us (main message):", result.error);
      return;
    }

    console.log("😸 Daily briefing delivered — purrfect!");

    // 7. Reply in the thread to kick off the day's conversation
    const threadTimestamp = result.ts;

    const threadPayload = {
      channel: SLACK_CHANNEL_ID,
      text: "Good morning!",
      thread_ts: threadTimestamp
    };

    options.payload = JSON.stringify(threadPayload);
    const threadResponse = UrlFetchApp.fetch(url, options);
    const threadResult = JSON.parse(threadResponse.getContentText());

    if (!threadResult.ok) throw new Error(`Slack Thread Failed: ${threadResult.error}`);

    sendOwnerReport(true, "sendDailySlackBriefing", "Daily briefing and thread posted.");
  } catch (error) {
    console.error(error as Error);
    sendOwnerReport(false, "sendDailySlackBriefing", error);
  }
}

/**
 * Digs into threads and deletes all bot-sent messages and replies.
 * Useful for a hard reset after testing or a bad run. 🙀
 */
function deepCleanupBotMessages(): void {
  const { SLACK_TOKEN, SLACK_CHANNEL_ID } = getProperties();

  const historyUrl = `https://slack.com/api/conversations.history?channel=${SLACK_CHANNEL_ID}&limit=50`;
  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: "get",
    headers: { Authorization: `Bearer ${SLACK_TOKEN}` }
  };

  try {
    const response = UrlFetchApp.fetch(historyUrl, options);
    const result = JSON.parse(response.getContentText());

    if (!result.ok) {
      console.error("🙀 Couldn't fetch channel history:", result.error);
      return;
    }

    const messages = result.messages;
    let deletedCount = 0;

    const deleteMsg = (ts: string) => {
      const deleteUrl = "https://slack.com/api/chat.delete";
      const deleteOptions: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
        method: "post",
        contentType: "application/json",
        headers: { Authorization: `Bearer ${SLACK_TOKEN}` },
        payload: JSON.stringify({ channel: SLACK_CHANNEL_ID, ts: ts })
      };
      const res = UrlFetchApp.fetch(deleteUrl, deleteOptions);
      if (JSON.parse(res.getContentText()).ok) deletedCount++;
      Utilities.sleep(1200); // Prevent Slack rate limiting
    };

    // Loop through the main messages
    for (let i = 0; i < messages.length; i++) {
      const msg = messages[i];

      // 1. If this message has a thread, dive in and delete bot replies first
      if (msg.thread_ts) {
        const repliesUrl = `https://slack.com/api/conversations.replies?channel=${SLACK_CHANNEL_ID}&ts=${msg.thread_ts}`;
        const repliesResponse = UrlFetchApp.fetch(repliesUrl, options);
        const repliesResult = JSON.parse(repliesResponse.getContentText());

        if (repliesResult.ok) {
          for (let j = 0; j < repliesResult.messages.length; j++) {
            const reply = repliesResult.messages[j];
            if (reply.bot_id || reply.app_id) {
              console.log(`🐾 Swatting thread reply: ${reply.ts}`);
              deleteMsg(reply.ts);
            }
          }
        }
      }

      // 2. Delete the main parent message if the bot sent it
      if (msg.bot_id || msg.app_id) {
        console.log(`🐾 Swatting main message: ${msg.ts}`);
        deleteMsg(msg.ts);
      }
    }

    console.log(`😸 All clean! Knocked ${deletedCount} message(s) off the shelf.`);

  } catch (e) {
    console.error("🙀 Hissed at a fetch error:", (e as Error).toString());
  }
}

/**
 * Sends a private DM report to the owner on success or failure.
 * On failure, includes the stack trace if an `Error` object is provided.
 * @param {boolean} isSuccess - Whether the operation succeeded
 * @param {string} functionName - Name of the function that was running
 * @param {Error|string} detail - Success summary string, or an `Error` object with a stack trace
 */
function sendOwnerReport(isSuccess: boolean, functionName: string, detail: any): void {
  const { SLACK_TOKEN, SLACK_OWNER_ID } = getProperties();

  let messageText = "";

  if (isSuccess) {
    messageText = `✅ *Paw-Paw Success*: \`${functionName}\` completed successfully.\n> ${detail}`;
  } else {
    // Extract message and stack trace if detail is an Error object
    const errorMsg = detail.message || detail;
    const stackTrace = detail.stack ? `\n\`\`\`${detail.stack}\`\`\`` : "";

    messageText = `🚨 *Paw-Paw System Alert*\n` +
      `*Function:* \`${functionName}\`\n` +
      `*Error:* ${errorMsg}${stackTrace}`;
  }

  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: `Bearer ${SLACK_TOKEN}` },
    payload: JSON.stringify({
      channel: SLACK_OWNER_ID,
      text: messageText
    })
  };

  UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", options);
}
