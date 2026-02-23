# Paw-Paw 🐱

> *"Purr-fectly automated, so the cat can nap." 😸*

**Paw-Paw** is an internal **Roster & Meal Management System** built for the Craftsmen Saga team in Dhaka. It lives inside a **Google Spreadsheet** as a Google Apps Script project and automates three main workflows:

1. Keeping the attendance roster in sync with the Slack channel membership.
2. Sending a daily attendance briefing to Slack every morning.
3. Calculating tomorrow's office headcount and posting it to a Discord meal-ordering channel every evening.

Everything is automated via time-based triggers. The script also handles edge cases (weekends, public holidays, team off-days, Ramadan, member joins/leaves) without any manual intervention.

---

## Table of Contents

1. [Why This Exists](#1-why-this-exists)
2. [How It Works: Big Picture](#2-how-it-works-big-picture)
3. [Tech Stack](#3-tech-stack)
4. [Directory Structure](#4-directory-structure)
5. [The Spreadsheet Layout](#5-the-spreadsheet-layout)
   - [Attendance Sheets (`yyyy-MM`)](#attendance-sheets-yyyy-mm)
   - [Holidays Sheet](#holidays-sheet)
6. [Source Files & Core Functions](#6-source-files--core-functions)
   - [Config.ts](#configts)
   - [Slack.ts](#slackts)
   - [Spreadsheet.ts](#spreadsheetts)
   - [Discord.ts](#discordts)
   - [Http.ts](#httpts)
   - [Menu.ts](#menuts)
   - [types.d.ts](#typesdts)
7. [Triggers & When They Run](#7-triggers--when-they-run)
   - [Automated (Time-Based) Triggers](#automated-time-based-triggers)
   - [Manual Triggers (Spreadsheet Menu)](#manual-triggers-spreadsheet-menu)
8. [Configuration Variables](#8-configuration-variables)
   - [How to Set Script Properties](#how-to-set-script-properties)
   - [SLACK\_TOKEN](#slack_token)
   - [SLACK\_CHANNEL\_ID](#slack_channel_id)
   - [SLACK\_OWNER\_ID](#slack_owner_id)
   - [DISCORD\_WEBHOOK](#discord_webhook)
   - [GOOGLE\_SHEET\_ID](#google_sheet_id)
   - [Auto-Managed Properties](#auto-managed-properties)
9. [Development & Deployment](#9-development--deployment)
   - [Prerequisites](#prerequisites)
   - [First-Time Setup](#first-time-setup)
   - [Day-to-Day Workflow](#day-to-day-workflow)
   - [Clasp & Google Apps Script Setup](#clasp--google-apps-script-setup)
10. [Error Reporting & Observability](#10-error-reporting--observability)
11. [Persona & Tone](#11-persona--tone)
12. [Frequently Asked Questions](#12-frequently-asked-questions)

---

## 1. Why This Exists

The Saga team in Dhaka orders office lunch every day. The caterer needs a headcount the evening before. Before Paw-Paw, someone had to manually count who would be in the office tomorrow, DM the count to a Discord channel, and hope no one forgot to update their status.

Paw-Paw automates all of this:

- It reads each person's attendance status from a Google Sheet (which they maintain themselves via a dropdown).
- Every evening it counts tomorrow's "Office" entries, posts the number to a Discord channel, and **locks the row** so no one can change their status after the deadline.
- Every morning it posts a Slack briefing showing who is in/WFH/on leave, with a thread for status updates.
- When someone joins or leaves the Slack channel, the roster is automatically updated in the spreadsheet.

---

## 2. How It Works: Big Picture

``` plaintext
                      ┌─────────────────────────────────┐
                      │   Google Spreadsheet            │
                      │                                 │
                      │  ┌──────────┐   ┌───────────┐   │
                      │  │ 2025-01  │   │ Holidays  │   │
                      │  │ 2025-02  │   │  (config) │   │
                      │  │ …        │   └───────────┘   │
                      │  └──────────┘                   │
                      └─────────────┬─────────────────┬─┘
                                    │                 │
               ┌────────────────────▼────┐    ┌───────▼─────────────────┐
               │  Morning Trigger        │    │  Evening Trigger        │
               │  sendDailySlackBriefing │    │  sendTomorrowHeadcount  │
               └────────────┬────────────┘    └──────────┬──────────────┘
                            │                            │
               ┌────────────▼────────────┐    ┌──────────▼──────────────┐
               │  Slack Channel          │    │  Discord Channel        │
               │  Posts today's          │    │  Posts tomorrow's       │
               │  attendance summary     │    │  meal headcount         │
               └─────────────────────────┘    └─────────────────────────┘
```

**Daily Flow:**

| Time | What Happens |
| ------ | ------------- |
| Morning (~9 AM) | `sendDailySlackBriefing` runs → posts today's Office/WFH/Leave summary to Slack with a thread |
| Throughout day | Team members update their own status cells in the Sheet (dropdowns, protected columns) |
| Evening (~4:30 PM) | `sendTomorrowHeadcount` runs → reads tomorrow's Office count → posts to Discord → locks tomorrow's row |
| 25th of the month | `sendTomorrowHeadcount` also calls `checkAndCreateFutureSheet()` → provisions the next-next month's sheet |

---

## 3. Tech Stack

| Layer | Technology | Notes |
| --- | --- | --- |
| **Runtime** | Google Apps Script (GAS) V8 | Runs server-side inside Google's infrastructure |
| **Language** | TypeScript 5.7 | Compiled to ES2019 `.js` for GAS |
| **Architecture** | Global Namespace (no modules) | No `import`/`export`. All files share one global scope, required by GAS |
| **Compiler** | `tsc` | Outputs to `dist/`, configured via `tsconfig.json` |
| **Deployment** | `clasp` (Command Line Apps Script) v3 | Pushes compiled `.js` files to Google Apps Script project |
| **Package Manager** | `pnpm` | Manages dev dependencies locally |
| **Storage** | Google Apps Script `PropertiesService` | Stores API keys and runtime state (never hardcoded) |
| **HTTP** | `UrlFetchApp` (GAS built-in) | Wrapped in `makeHttpRequest` with retry + exponential backoff |
| **Scheduling** | GAS Time-based Triggers | Set up in the GAS dashboard (not code-driven) |
| **Logging** | Google Cloud Stackdriver / `console.log` | Visible in the GAS dashboard → Executions |

### Key Dependencies (`devDependencies` only; nothing ships to GAS except compiled JS)

```json
{
  "@google/clasp": "^3.2.0",
  "@types/google-apps-script": "^1.0.83",
  "typescript": "^5.7.0"
}
```

**Why no runtime dependencies?** Google Apps Script does not support Node.js module imports. Everything available at runtime is either compiled TypeScript or native GAS globals (`SpreadsheetApp`, `UrlFetchApp`, `PropertiesService`, etc.).

### `tsconfig.json` highlights

```json
{
  "target": "ES2019",
  "module": "None",
  "types": ["google-apps-script"],
  "strict": true,
  "rootDir": "./src",
  "outDir": "./dist"
}
```

- `"module": "None"`: Critical. Prevents TypeScript from emitting any `require()` or `import` statements that GAS cannot handle.
- `"types": ["google-apps-script"]`: Gives full IntelliSense for all GAS globals.

---

## 4. Directory Structure

``` plaintext
paw-paw/
├── src/                        ← Source TypeScript files (edit these)
│   ├── Config.ts               ← Global constants, property cache, date config
│   ├── Slack.ts                ← Slack API: roster sync, briefings, owner DMs
│   ├── Spreadsheet.ts          ← Sheet creation, sync, formatting, protections
│   ├── Discord.ts              ← Meal headcount, Discord webhook, row locking
│   ├── Http.ts                 ← HTTP wrapper with retry + exponential backoff
│   ├── Menu.ts                 ← onOpen() (owner-only custom menu)
│   ├── types.d.ts              ← Ambient TypeScript interfaces (NOT emitted)
│   └── appsscript.json         ← GAS manifest (timezone, runtimeVersion)
│
├── dist/                       ← Auto-generated, DO NOT EDIT (git-ignored)
│   ├── Config.js
│   ├── Slack.js
│   ├── Spreadsheet.js
│   ├── Discord.js
│   ├── Http.js
│   ├── Menu.js
│   └── appsscript.json
│
├── .clasp.json                 ← Links local project to remote GAS project (scriptId)
├── tsconfig.json               ← TypeScript compiler config
├── package.json                ← pnpm scripts & devDependencies
└── GEMINI.md                   ← AI assistant context for this project
```

> **Rule:** Only ever edit files in `src/`. The `dist/` folder is regenerated on every `pnpm run push`.

---

## 5. The Spreadsheet Layout

### Attendance Sheets (`yyyy-MM`)

Each month has its own sheet tab named in `yyyy-MM` format (e.g., `2025-01`, `2025-02`). The script identifies these sheets using the pattern `/^\d{4}-\d{2}$/`.

**Structure:**

| Column A | Column B | Column C | … | Last Column |
| --- | --- | --- | --- | --- |
| **Date** | **Alice** | **Bob** | … | **Total** |
| 2025-01-01 | Leave | Leave | … | — |
| 2025-01-02 | Office | WFH | … | 3 |
| … | … | … | … | … |

- **Column A (Date):** Formatted `yyyy-MM-dd`. Locked; no one can edit it.
- **Member Columns (B to second-to-last):** Each member has their own column. The header is their Slack display name. Each member can only edit their own column (enforced via Google Sheets column protections, with their Google email address added as the editor).
- **Last Column (Total):** A `COUNTIF` formula that counts "Office" entries for that row. Locked.

**Row States:**

| Row Type | Cell Value | Validation | Protection |
| --- | --- | --- | --- |
| Weekday | `Office` / `WFH` / `Leave` | Dropdown: Office, WFH, Leave | Per-member column only |
| Weekend | `—` | None | Full row locked (no editors) |
| Holiday | `Leave` | Dropdown: WFH, Leave | Full row locked (no editors) |
| Team Offday | `Leave` | Dropdown: WFH, Leave | Full row locked (no editors) |
| Past date (nightly) | *(unchanged)* | n/a | Warning-only lock via `setWarningOnly(true)` |

**Summary Section (below data rows):**

Three rows appearing 3 rows below the last date row (i.e., `daysInMonth + 4`):

| | Alice | Bob | … |
| --- | --- | --- | --- |
| **Office** | `=COUNTIF(B2:B32,"Office")` | … | … |
| **WFH** | `=COUNTIF(B2:B32,"WFH")` | … | … |
| **Leave** | `=COUNTIF(B2:B32,"Leave")` | … | … |

**Conditional Formatting (applied automatically):**

| Rule | Background | Font | Condition |
| --- | --- | --- | --- |
| Weekend | `#EFEFEF` (grey) | `#999999` | `WEEKDAY($A2, 2) > 5` |
| Holiday | `#D9EAD3` (light green) | `#274E13` (dark green) | Date matches a holiday in the Holidays sheet |
| Offday | `#FCE5CD` (light orange) | `#B45F06` (dark orange) | Date matches an offday in the Holidays sheet |

Weekly borders (thick bottom) are drawn after each Sunday to visually group weeks.

---

### Holidays Sheet

A special sheet tab named exactly **`Holidays`** (case-sensitive). This is the single source of truth for all off-days and holidays. It is read once per script execution and cached in memory.

**Layout:**

| Column A | Column B | Column C | Column E | Column F |
| --- | --- | --- | --- | --- |
| **Date** | **Name** | **Type** | **Key** | **Value** |
| 2025-02-21 | Ekushey February | Holiday | Start | 2025-03-01 |
| 2025-03-17 | Sheikh Mujib Birthday | Holiday | End | 2025-03-30 |
| 2025-05-01 | May Day | Holiday | | |
| 2025-03-14 | Team Offday | Offday | | |

- **Col A:** Date (Date cell type; GAS will parse it).
- **Col B:** Human-readable name (for reference only).
- **Col C:** Type. Must be exactly `Holiday` or `Offday` (case-sensitive).
  - `Holiday` → green row, no office option, row locked.
  - `Offday` → orange row, no office option, row locked.
- **Col E2 / F2:** Label `Start` / Ramadan start date.
- **Col E3 / F3:** Label `End` / Ramadan end date.

> **Important:** Row 1 is the header row and is skipped. Ramadan dates go in rows 2 and 3 of columns E/F specifically.

When a Ramadan date range is active, the Discord meal message changes format from `**N**` to `Lunch: **0**, Iftar: **N**`.

---

## 6. Source Files & Core Functions

### `Config.ts`

The foundation that everything else reads from. Loaded first in GAS's execution order (files load alphabetically).

#### `CONFIG` (constant)

```typescript
const CONFIG: AppConfig = {
  SLACK_API_BASE: 'https://slack.com/api',
  DISCORD: {
    USERNAME: 'Saga Check-in 🍽️',
    AVATAR_URL: '...',    // Avatar shown on Discord messages
  },
  TIMEZONE: 'GMT+6',
};
```

Hardcoded non-sensitive configuration. To change the Discord bot name or avatar, edit these values here.

#### `PROPERTY_KEYS` (constant)

A map of string keys used to read/write `PropertiesService`. Centralizes key names to prevent typos across files.

#### `getProperties(): ScriptProperties`

**Purpose:** Returns all sensitive script properties (API keys, IDs) as a typed object.

**How it works:** On first call per execution, it fetches all properties from `PropertiesService.getScriptProperties()` and caches them in the module-level `_propertyCache` variable. All subsequent calls in the same execution return the cache with zero extra API calls.

**Why cache?** GAS charges execution time per API call. Fetching properties once per run is efficient.

#### `getDateConfig(): DateConfig`

**Purpose:** Reads the `Holidays` sheet and returns `{ holidays: string[], offdays: string[], ramadan: { start, end } }`.

**How it works:**

1. Checks `_dateConfig` cache. If populated, returns immediately.
2. Finds the `Holidays` sheet by name.
3. Reads the entire data range in a single `getValues()` call.
4. Iterates rows (skipping header), classifies each date as `Holiday` or `Offday`, formats as `yyyy-MM-dd` strings.
5. Reads Ramadan dates from E2/F2 and E3/F3.
6. Caches result and returns.

**Fallback:** If no `Holidays` sheet exists, logs a warning and returns empty arrays (graceful degradation). The app continues running without holiday logic.

---

### `Slack.ts`

All communication with the Slack API lives here.

#### `getChannelUsers(): Record<string, SlackUser>`

**Purpose:** Fetches all human, non-bot, non-deleted members of the configured Slack channel.

**How it works:**

1. Calls `conversations.members` with `limit=100`, paginating via cursor up to 5 pages (500 members max).
2. For each page of member IDs, calls `getUserInfoBatch()`.
3. Returns a map of `userId → SlackUser { name, email, image }`.

**What counts as a valid member?** Only users where `is_bot === false` AND `deleted === false`. Slackbots, deleted users, and deactivated accounts are excluded.

#### `getUserInfoBatch(userIds: string[]): Record<string, SlackUser>`

**Purpose:** Resolves Slack user IDs to human-readable names and emails.

**How it works:**

- Processes IDs in batches of 10 using `users.info` API.
- 100ms sleep between batches to stay within Slack's rate limits.
- Name resolution priority: `display_name_normalized` → `real_name_normalized` → `real_name` → `name`.

#### `sendDailySlackBriefing(): void`

**Purpose:** Posts the morning attendance summary to the Slack channel and starts a reply thread.

**How it works:**

1. **Skip checks:** Returns early if today is a weekend or appears in the `offdays` list.
2. Finds the current month's sheet (e.g., `2025-01`).
3. Locates today's row by matching dates.
4. **Early exit:** If the row's second cell is `"-"` (sheet-level non-working day), skips.
5. Groups member columns into `Office`, `WFH`, and `Leave` lists.
6. Constructs a Slack block message with `<!here>`, today's date, and the three groups.
7. Posts the main message via `chat.postMessage`.
8. Saves the returned `ts` timestamp and posts a `"Good morning!"` reply in the thread.
9. Calls `sendOwnerReport(true, ...)` on success or `sendOwnerReport(false, ...)` on failure.

**Slack message format:**

``` plaintext
<!here> | *Monday, Jan 6*

*🏢 ON-SITE:* Alice, Bob, Carol
*🏠 WFH:* Dave
*🌴 MIA:* Eve

Please post your *status* (Starting, AFK, Back, etc.) in the thread and feel free to *chitchat!*
```

#### `deepCleanupBotMessages(): void`

**Purpose:** Emergency cleanup tool. Deletes all bot messages (and their thread replies) from the last 50 channel messages. Useful after testing or a bad run.

**How it works:**

1. Fetches the last 50 messages via `conversations.history`.
2. For each message with a thread, fetches replies via `conversations.replies` and deletes any `bot_id`/`app_id` replies.
3. Deletes the parent message if it was bot-sent.
4. 1200ms sleep between deletions to stay within Slack rate limits.

> **Note:** This is a manual-only, owner-run function. It is not exposed in the menu currently. Call it from the GAS script editor directly if needed.

#### `sendOwnerReport(isSuccess, functionName, detail): void`

**Purpose:** Sends a private DM to the owner's Slack account reporting the outcome of any automated function.

**Success format:**

``` plaintext
✅ *Paw-Paw Success*: `sendTomorrowHeadcount` completed successfully.
> Headcount sent for 2025-01-06. Today's row is locked.
```

**Failure format:**

``` plaintext
🚨 *Paw-Paw System Alert*
*Function:* `sendTomorrowHeadcount`
*Error:* Something went wrong
```stack trace here```
```

---

### `Spreadsheet.ts`

Everything related to creating, populating, and maintaining the Google Sheets lives here.

#### Helper Functions

**`getSheetInfo(sheet)`**: Extracts `{ year, month, daysInMonth }` from a `yyyy-MM` sheet name. Used everywhere to avoid recalculating date math.

**`columnToLetter(column)`**: Converts 1-based column index to Excel letter notation (`1 → A`, `27 → AA`). Used for building COUNTIF formula strings.

**`writeSummarySection(sheet, daysInMonth, memberCount, totalCols)`**: Writes the Office/WFH/Leave summary rows below the data. Sets bold labels in column A and `COUNTIF` formulas for each member column. Adds a thick top border.

**`addFormattingRules(sheet, totalRows, totalCols)`**: Applies all three conditional format rules (weekend grey, holiday green, offday orange) and weekly bottom borders (after each Sunday). This completely replaces any existing conditional format rules on the sheet.

---

#### `createNextMonthSheet(): void`

**Purpose:** Creates the *next* month's attendance sheet from scratch, fully populated and protected.

**Trigger:** Called manually from the menu, or automatically by `checkAndCreateFutureSheet()` on the 25th.

**How it works (in order):**

1. **Find the latest existing dated sheet.** Scans all sheet tabs for `yyyy-MM` names, picks the latest. If none exist, defaults to the current month.
2. **Calculate target month.** Adds 1 month to the latest sheet's date.
3. **Guard:** If the sheet already exists, exits immediately (idempotent).
4. **Fetch Slack roster.** Calls `getChannelUsers()` to get current members, sorted alphabetically by name.
5. **Pre-compute day metadata.** For each day in the month, calculates: `isWeekend`, `isHoliday`, `isOffday`. Done once to avoid repeated date comparisons.
6. **Generate rows.** One row per calendar day:
   - Weekends: all cells set to `"—"`.
   - Holidays/Offdays: all member cells set to `"Leave"`.
   - Working days: all member cells set to `"Office"`.
   - Total column: COUNTIF formula counting "Office" in the row.
7. **Write to sheet.** Batch-writes header row + all data rows in two `setValues()` calls.
8. **Data Validation.** Per row:
   - Weekends: no validation (they're locked anyway).
   - Holidays/Offdays: dropdown `[WFH, Leave]`.
   - Working days: dropdown `[Office, WFH, Leave]`.
9. **Row Protections.** Weekend rows and holiday/offday rows are fully locked (no editors).
10. **Column Protections.** Each member column is protected with only that member's email as editor. This is the key mechanism that prevents members from editing each other's status.
11. **Structural Protections.** The header row, date column (col A), and totals column (last col) are locked.
12. **Summary Section.** Writes the Office/WFH/Leave COUNTIF summary rows.
13. **Styling.** Date format, font size (11pt), vertical/horizontal alignment, row height (32px), column widths, frozen first row and first column.

---

#### `syncAllActiveSheets(): void`

**Purpose:** Syncs the current month and all future month sheets with the live Slack channel roster. Handles both joiners and leavers.

**Trigger:** Called manually from the "Sync Roster with Slack" menu item, or can be set as a periodic trigger.

**How it works:**

1. Fetches the live Slack roster.
2. **Gatekeeper check:** Compares the current month's sheet headers against the live roster. If they match exactly, logs "no sync needed" and exits early to avoid unnecessary writes.
3. If out of sync, iterates all `yyyy-MM` sheets from the current month onward and calls `processSheetSync()` on each.

#### `processSheetSync(sheet, currentSlackUsers, currentSlackNames, isFutureSheet): void`

**Purpose:** Applies the member diff to a single sheet, handling leavers and joiners.

**Leavers (members in sheet but not in Slack):**

- **Future sheets:** Delete their column entirely.
- **Current month:** Fill remaining days (today onward) with `"-"` and strip their column protection. Past days are preserved.

**Joiners (members in Slack but not in sheet):**

- Insert a new column before the Totals column.
- Set header cell with bold style.
- Fill values: past days and weekends get `"-"`, holidays get `"Leave"`, working days get `"Office"`.
- Apply data validation rules.
- Add column protection with the member's email.

After processing, recalculates Total column formulas and rewrites the summary section.

---

#### `refreshHolidayFormatting(): void`

**Purpose:** Re-applies holiday/offday formatting across all active sheets (current month onward). Run this manually after editing the `Holidays` tab mid-month.

**How it works:** Iterates all `yyyy-MM` sheets from the current month onward, calls `refreshSheetHolidayFormatting()` on each.

#### `refreshSheetHolidayFormatting(sheet, isFutureSheet): void`

**Purpose:** Refreshes a single sheet's holiday/offday state.

**What it does:**

- **Future sheets:** Updates cell values (sets holidays/offdays back to `"Leave"`, restores `"Office"` if a date was accidentally un-marked), updates data validation dropdowns, re-applies row protections.
- **Current month:** Only re-applies conditional formatting colors (does NOT touch cell values, preserving manually entered attendance).
- Weekend rows are always skipped.

#### `checkAndCreateFutureSheet(): void`

**Purpose:** Thin wrapper that calls `createNextMonthSheet()` only if today is on or after the 25th. Called by `sendTomorrowHeadcount` every evening.

---

### `Discord.ts`

Handles all meal headcount logic and Discord webhook communication.

#### `sendTomorrowHeadcount(): void`

**Purpose:** The main nightly function. Calculates tomorrow's meal count, posts it to Discord, locks the row, and on the 25th creates the next-next-month sheet.

**How it works:**

1. Calculate tomorrow's date.
2. **Skip check:** If tomorrow is a weekend or a team offday, logs and exits.
3. Finds tomorrow's sheet (e.g., if today is Jan 31, tomorrow is Feb 1 → opens `2025-02`).
4. Reads tomorrow's row, picks the value in the last column (the Total COUNTIF result).
5. Calls `sendOrUpdateDiscordMessage(tomorrowStr, yesCount)`.
6. Calls `lockRowByDate(tomorrowStr)` to lock tomorrow's row.
7. Calls `checkAndCreateFutureSheet()` for the monthly admin check.
8. Sends owner report (success or failure with stack trace).

#### `sendOrUpdateDiscordMessage(dateStr, yesCount): void`

**Purpose:** Smart send. Updates the existing message if one was already sent today, otherwise posts a new one.

**How it works:**

- Reads `LAST_DATE` and `LAST_DISCORD_MESSAGE_ID` from script properties.
- If `LAST_DATE === dateStr` and a message ID exists → calls `updateExistingDiscordMessage()` (PATCH).
- Otherwise → calls `sendNewDiscordMessage()` (POST).

This means if the trigger runs multiple times in a day (e.g., manual retry), it edits the same Discord message instead of spamming new ones.

#### `getMealMessageContent(dateStr, yesCount): string`

**Purpose:** Formats the Discord message content.

- **During Ramadan:** `Lunch: **0**, Iftar: **N**`
- **Regular days:** `**N**`

Ramadan detection: checks if `dateStr` falls within the `[ramadan.start, ramadan.end]` range from `getDateConfig()`. Uses simple string comparison (ISO date strings sort correctly lexicographically).

#### `sendNewDiscordMessage(dateStr, yesCount): void`

Posts a new message via `POST https://discord.com/api/webhooks/.../...?wait=true`.

`?wait=true` is important because it makes Discord return the created message object (including `id`), which is then saved to `LAST_DISCORD_MESSAGE_ID` for future updates.

#### `updateExistingDiscordMessage(messageId, yesCount, dateStr): void`

Patches the existing message via `PATCH .../messages/{id}?wait=true`.

If Discord returns 404 (message was deleted), falls back to `sendNewDiscordMessage()` automatically.

#### `lockRowByDate(date): void`

**Purpose:** Locks the given date's row in its month's sheet so no further edits are allowed after the headcount is submitted.

**How it works:**

1. Finds the correct month sheet.
2. Scans rows to find the matching date.
3. Applies a `protection` with `setWarningOnly(true)`. Editors see a warning popup but cannot be fully blocked (a Google API limitation for non-owner editors); the intent is clearly communicated.

#### `deleteDiscordMessage(messageId): void`

**Purpose:** Utility for manually deleting a specific Discord webhook message by ID. Used for cleanup. Called directly from the GAS script editor.

#### `getWebhookMessageUrl(webhook, messageId): string`

**Purpose:** Transforms `https://discord.com/api/webhooks/123/abc` into `https://discord.com/api/webhooks/123/abc/messages/{messageId}` using a regex substitution.

---

### `Http.ts`

#### `makeHttpRequest(url, options, maxRetries = 3): HttpRequestResult`

**Purpose:** A robust HTTP wrapper around GAS's `UrlFetchApp.fetch()` with automatic retries.

**Behavior:**

- Always sets `muteHttpExceptions: true` so GAS doesn't throw on non-2xx responses; the function handles them explicitly.
- On `2xx`: returns `{ success: true, data: parsedJSON, responseCode }`.
- On `429 (Rate Limited)`: sleeps with **exponential backoff** (capped at 10 seconds) and retries.
- On other non-2xx: returns `{ success: false, data: null, error: string, responseCode }` immediately (no retry).
- On network exception: waits `1000ms × attempt` and retries. After all retries exhausted, **throws** (so the caller's try/catch can capture it and call `sendOwnerReport`).

**Retry strategy:**

| Attempt | Behavior |
| --------- | --------- |
| 1 | Immediate |
| 2 (on 429) | Sleep 2s |
| 3 (on 429) | Sleep 4s (capped at 10s) |
| After max retries | Throws `Error` |

---

### `Menu.ts`

#### `onOpen(): void`

**Purpose:** GAS's reserved trigger function. Runs automatically whenever anyone opens the spreadsheet.

**Access control:** Compares the active user's email to the spreadsheet owner's email. If they don't match, the function returns silently with no menu created and no error shown. Non-owners open the spreadsheet normally.

**Menu structure (owner only):**

``` plaintext
Paw-Paw 🐱
├── Sync Roster with Slack          → syncAllActiveSheets()
├── Send Daily Slack Briefing       → sendDailySlackBriefing()
├── ─────────────────────
├── Refresh Holiday Colors 🎌       → refreshHolidayFormatting()
├── ─────────────────────
├── Create Next Month Sheet         → createNextMonthSheet()
└── Force Tomorrow Headcount        → sendTomorrowHeadcount()
```

---

### `types.d.ts`

Ambient TypeScript interfaces. This is a type-only file that is **not emitted** to JavaScript (GAS never sees it).

| Interface | Purpose |
| --------- | --------- |
| `DateConfig` | Return type of `getDateConfig()` |
| `AppConfig` | Type of the `CONFIG` constant |
| `ScriptProperties` | Return type of `getProperties()` |
| `SlackUser` | User object with `name`, `email`, `image?` |
| `HttpRequestResult` | Return type of `makeHttpRequest()` |

---

## 7. Triggers & When They Run

### Automated (Time-Based) Triggers

These are set up **manually** in the Google Apps Script dashboard. They do not configure themselves.

> **How to set up:** In the GAS editor → click the ⏰ **Triggers** icon (left sidebar) → **+ Add Trigger** → select function, deployment, event type (Time-driven), and schedule.

| Function | Recommended Time | What It Does |
| --------- | ----------------- | -------------- |
| `sendDailySlackBriefing` | Every day, 9:00–10:00 AM (Bangladesh time) | Posts today's attendance summary to Slack. Auto-skips weekends and offdays. |
| `sendTomorrowHeadcount` | Every day, 4:00–5:00 PM (Bangladesh time) | Reads tomorrow's "Office" count → posts to Discord → locks row → on 25th creates next-next month sheet. |

**Important notes about GAS triggers:**

- GAS triggers run in UTC by default. The `appsscript.json` sets `"timeZone": "Asia/Dhaka"` for the script, but trigger times in the dashboard are displayed in the timezone of your Google Account. Set them accordingly.
- Triggers fire within a **1-hour window** (not exact). `"9:00–10:00 AM"` means it will run sometime in that hour.
- If the trigger fails (throws an unhandled exception), GAS will email the script owner after consecutive failures.

### Manual Triggers (Spreadsheet Menu)

The **Paw-Paw 🐱** custom menu is visible only to the spreadsheet owner. Access it by opening the spreadsheet and clicking the menu.

| Menu Item | Function Called | When to Use |
| --------- | ---------------- | ------------- |
| **Sync Roster with Slack** | `syncAllActiveSheets()` | When someone joins/leaves the Slack channel and you want to update the sheet immediately, without waiting for the next automated run. Also updates all future month sheets. |
| **Send Daily Slack Briefing** | `sendDailySlackBriefing()` | If the morning trigger fails or you want to manually post the briefing (e.g., for testing). Note: this will post a new message even if one was already sent today. |
| **Refresh Holiday Colors 🎌** | `refreshHolidayFormatting()` | After updating the `Holidays` sheet tab (adding/removing/changing dates), run this to re-apply colors, data validation, and row protections to all current and future month sheets. |
| **Create Next Month Sheet** | `createNextMonthSheet()` | To manually provision next month's sheet ahead of the automated 25th-of-month check. Safe to run multiple times; it exits early if the sheet already exists. |
| **Force Tomorrow Headcount** | `sendTomorrowHeadcount()` | If the evening trigger failed and you need to manually post the headcount to Discord and lock the row. |

---

## 8. Configuration Variables

All sensitive configuration is stored in **Google Apps Script Script Properties** and never hardcoded. These are encrypted key-value pairs accessible only within the script.

### How to Set Script Properties

1. Open the Google Apps Script editor for the project.
2. Click ⚙️ **Project Settings** (gear icon in left sidebar).
3. Scroll down to **Script Properties**.
4. Click **Add script property** and add each key-value pair below.

---

### `SLACK_TOKEN`

- **What it is:** The Slack Bot User OAuth Token. Used to authenticate all Slack API calls (fetching channel members, user info, posting messages, deleting messages).
- **Format:** Starts with `xoxb-`
- **Scopes required by the Slack app:**
  - `channels:read`: list channel members
  - `users:read`: fetch user info
  - `users:read.email`: get member email addresses (needed for sheet column protections)
  - `chat:write`: post messages in channels
  - `chat:write.public`: post to channels without joining
  - `im:write`: send DMs (used for owner reports)

**How to get it:**

1. Go to [https://api.slack.com/apps](https://api.slack.com/apps).
2. Select your app (or create one: **Create New App → From scratch**).
3. In the left sidebar, go to **OAuth & Permissions**.
4. Under **Scopes → Bot Token Scopes**, add all the scopes listed above.
5. Scroll up and click **Install to Workspace** (or **Reinstall** if already installed).
6. Copy the **Bot User OAuth Token** shown after installation.

---

### `SLACK_CHANNEL_ID`

- **What it is:** The unique ID of the Slack channel where the daily briefing is posted and from which the member roster is synced.
- **Format:** Starts with `C` (e.g., `C01AB23CD`)

**How to get it:**

1. Open Slack in a browser (not the desktop app).
2. Navigate to the target channel.
3. The Channel ID is the last part of the URL: `https://app.slack.com/client/T.../C01AB23CD`
4. **Alternative:** Right-click the channel name in the sidebar → **View channel details** → scroll to the bottom → copy the Channel ID.

> **Important:** The bot must be **invited to the channel** before it can read members or post. Send `/invite @YourBotName` in the channel.

---

### `SLACK_OWNER_ID`

- **What it is:** The Slack Member ID of the person who should receive private DM reports on success/failure of automated runs.
- **Format:** Starts with `U` (e.g., `U01XY789Z`)

**How to get it:**

1. Open Slack.
2. Click your profile picture (or any team member's profile you want to set as owner).
3. Click **View full profile**.
4. Click the `⋯` (More) button.
5. Click **Copy member ID**.

---

### `DISCORD_WEBHOOK`

- **What it is:** A Discord webhook URL pointing to the channel where meal headcounts are posted.
- **Format:** `https://discord.com/api/webhooks/{webhook.id}/{webhook.token}`

**How to get it:**

1. Open Discord and go to the target server.
2. Click ⚙️ on the channel where headcounts should be posted (**Edit Channel**).
3. Go to **Integrations → Webhooks**.
4. Click **New Webhook** (or select an existing one).
5. Give it a name (this is overridden by `CONFIG.DISCORD.USERNAME` in the script).
6. Click **Copy Webhook URL**.

> **Security note:** Treat the webhook URL like a password. Anyone with this URL can post to your Discord channel.

---

### `GOOGLE_SHEET_ID`

- **What it is:** The unique identifier of the Google Spreadsheet that Paw-Paw manages.
- **Format:** A long alphanumeric string

**How to get it:**

1. Open the Google Spreadsheet in a browser.
2. Look at the URL: `https://docs.google.com/spreadsheets/d/`**`THIS_IS_THE_ID`**`/edit`
3. Copy the ID portion.

> **Note:** Currently, this property is stored but the script primarily uses `SpreadsheetApp.getActiveSpreadsheet()` since it runs in the bound GAS context. This property exists for potential future standalone script usage.

---

### Auto-Managed Properties

These are written and read by the script internally. **Do not set or modify these manually** unless you know what you're doing.

| Key | Managed By | Purpose |
| --------- | ----------- | --------- |
| `LAST_DATE` | `sendNewDiscordMessage` | The date string of the last Discord message sent (e.g., `2025-01-06`). Used to decide whether to POST a new message or PATCH the existing one. |
| `LAST_DISCORD_MESSAGE_ID` | `sendNewDiscordMessage` | Discord's message ID of the last headcount message. Used to PATCH instead of creating duplicate messages. |
| `LAST_SLACK_MESSAGE_TS` | Reserved | Slack message timestamp, reserved for potential thread-management features. |

---

## 9. Development & Deployment

### Prerequisites

| Tool | Version | Install |
| --------- | --------- | --------- |
| Node.js | 18+ | [nodejs.org](https://nodejs.org) |
| pnpm | Latest | `npm install -g pnpm` |
| @google/clasp | Installed via pnpm | included in `devDependencies` |
| Google Account | n/a | Must have access to the GAS project |

### First-Time Setup

**1. Clone and install dependencies:**

```bash
git clone <repo-url>
cd paw-paw
pnpm install
```

**2. Login to clasp with your Google Account:**

```bash
pnpm run login
# Opens a browser window. Log in with the same Google account
# that owns the GAS project
```

**3. Verify `.clasp.json`:**

```json
{
  "scriptId": "<YOUR_GAS_SCRIPT_ID>",
  "rootDir": "dist/"
}
```

The `scriptId` must match the remote GAS project. Find it in the GAS editor URL: `https://script.google.com/home/projects/`**`THIS_IS_THE_SCRIPT_ID`**`/edit`.

**4. Set Script Properties:**
Go to the GAS project → Project Settings → Script Properties and add all the keys described in [Section 8](#8-configuration-variables).

**5. Set up Triggers:**
Go to the GAS project → Triggers (⏰) and configure the two time-based triggers per [Section 7](#7-triggers--when-they-run).

---

### Day-to-Day Workflow

**Making changes:**

1. Edit `.ts` files in `src/`.
2. Build and push in one command:

```bash
pnpm run push
```

This runs:

```bash
tsc                                  # Compiles src/*.ts → dist/*.js
cp src/appsscript.json dist/         # Copies the GAS manifest
clasp push                           # Uploads dist/* to Google Apps Script
```

**Build only (no push):**

```bash
pnpm run build
```

**Login (if session expires):**

```bash
pnpm run login
```

---

### Clasp & Google Apps Script Setup

If you need to create a brand new GAS project from scratch and link it:

1. Create a new Google Spreadsheet.
2. Go to **Extensions → Apps Script** (this creates a bound GAS project).
3. In the GAS editor, note the script ID from the URL.
4. Update `.clasp.json` with the new `scriptId`.
5. Run `pnpm run push` to upload the code.
6. Set all Script Properties.
7. Set up the triggers.

**Cloning from an existing project** (if you have the scriptId):

```bash
npx clasp clone <scriptId>
```

---

## 10. Error Reporting & Observability

### Owner DMs

Every automated function that can fail (`sendDailySlackBriefing`, `sendTomorrowHeadcount`) wraps its logic in `try/catch` and calls `sendOwnerReport()`. This sends a Slack DM to the `SLACK_OWNER_ID` with:

- On success: which function ran and a short summary.
- On failure: which function failed, the error message, and the full stack trace (formatted as a Slack code block).

This means you don't need to check the GAS dashboard daily. Any failure will appear in your Slack DMs.

### GAS Execution Logs

Every function uses `console.log()` and `console.warn()` throughout. These are visible in:

- GAS Editor → **Executions** (left sidebar icon): shows all recent runs with logs and errors.
- Google Cloud Console → **Stackdriver Logging** (because `appsscript.json` sets `"exceptionLogging": "STACKDRIVER"`).

All log messages include cat-themed context that clearly identifies the operation:

- `😸` = success
- `😴` = skipped (non-working day, no-op)
- `😿` / `🙀` = warning or error
- `🐾` = informational progress

### GAS Email Notifications

By default, GAS sends an email to the project owner if a triggered function fails consecutively. This is a GAS platform feature. Configure it under **Triggers → (your trigger) → Notifications**.

---

## 11. Persona & Tone

Paw-Paw runs with the persona of a **playful, curious, and slightly aloof cat**. This applies to:

- **Code comments:** Written with cat-appropriate observations (`// even cats don't chase all mice at once`).
- **Log messages:** Use cat emojis and language (`😸 Purr-fect!`, `😴 Napping now`, `🙀 Hissing at errors`).
- **Slack messages:** The daily briefing is friendly and inviting.
- **Discord messages:** Concise and to the point. The cat delivers the count, nothing more.
- **Error reports:** Dramatic. The cat is alarmed. But the stack trace is always included for practicality.
- **Menu labels:** Clear function names with subtle cat branding (`Paw-Paw 🐱`, `Refresh Holiday Colors 🎌`).

When adding new features or modifying existing log messages/comments, maintain this tone. It makes the system more enjoyable to maintain and easier to distinguish Paw-Paw's messages from generic system noise.

---

## 12. Frequently Asked Questions

**Q: The morning briefing wasn't sent today. What happened?**

A: Check in order:

1. Is today a weekend or in the `Holidays` sheet as an offday? → Expected, no briefing.
2. Check your Slack DMs from Paw-Paw. Did you get a 🚨 alert?
3. Go to GAS Editor → Executions and check the log for `sendDailySlackBriefing`.
4. Verify the trigger is still active in the Triggers tab.

---

**Q: A new person joined our Slack channel but they're not in the sheet.**

A: Open the spreadsheet → **Paw-Paw 🐱 → Sync Roster with Slack**. This will add the new member's column to all current and future month sheets with the correct default values.

---

**Q: Someone left the team. Their column still shows in the sheet.**

A: Run **Sync Roster with Slack**. For the current month, their remaining days will be filled with `"—"` and their column protection removed. For future month sheets, the column is deleted entirely.

---

**Q: I added dates to the Holidays sheet but the colors didn't update.**

A: Open the spreadsheet → **Paw-Paw 🐱 → Refresh Holiday Colors 🎌**. This re-reads the Holidays tab and re-applies all formatting, validation, and protections to the current and future sheets.

---

**Q: The Discord count was wrong / sent twice / not sent.**

A:

- **Wrong count:** Check the Total column in tomorrow's sheet row. The COUNTIF formula is `=COUNTIF(B{row}:{lastCol}{row}, "Office")`. It depends on correct "Office" values in the row.
- **Sent twice:** The `LAST_DATE` and `LAST_DISCORD_MESSAGE_ID` properties should prevent this. Check Script Properties. If `LAST_DATE` matches today and `LAST_DISCORD_MESSAGE_ID` is set, it will patch instead of post.
- **Not sent:** Check if tomorrow is a weekend or in the offdays list. Check your owner DM for error reports.

---

**Q: I need to override the headcount for a day manually.**

A: You can edit tomorrow's Total cell directly in the sheet before the evening trigger runs. However, the cell is formula-driven, so you'd need to replace it with a number. After the trigger fires and locks the row, you can use **Force Tomorrow Headcount** from the menu to re-send, but you'd need to unlock the row first (via the GAS Protections tab in the sheet editor).

---

**Q: What happens on the 25th of the month?**

A: When `sendTomorrowHeadcount` runs on the evening of the 25th, it calls `checkAndCreateFutureSheet()`, which calls `createNextMonthSheet()`. This provisions the sheet for *next* month (e.g., if it's January 25th, it creates March's sheet). February's sheet should already exist (having been created on December 25th). This ensures there's always a future month sheet ready before the current month ends.

---

**Q: Why does the column protection restrict only the member's own column?**

A: Each member's column is protected with `removeEditors()` (removes all editors including inheritors) and then `addEditor(member.email)`. This means:

- Only that team member can edit their own attendance column.
- The spreadsheet owner can still edit all columns (owners bypass protections in Google Sheets).
- If a member needs to edit someone else's status, the owner must do it.

---

**Q: Can I run this as a standalone GAS project (not bound to a spreadsheet)?**

A: Currently no. The script uses `SpreadsheetApp.getActiveSpreadsheet()` throughout, which only works in a spreadsheet-bound context. To convert to standalone, you'd need to use `SpreadsheetApp.openById(GOOGLE_SHEET_ID)` instead.

---

*Project Paw-Paw. Built with 🐾 for the Craftsmen Saga team, Dhaka.*
