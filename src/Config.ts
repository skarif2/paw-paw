/** Global configuration constants for Paw-Paw 🐾 */
const CONFIG: AppConfig = {
  SLACK_API_BASE: 'https://slack.com/api',
  DISCORD: {
    USERNAME: 'Saga Check-in 🍽️',
    AVATAR_URL: 'https://saganews.com/hs-fs/hubfs/saga-icon.png?width=120&height=120&name=saga-icon.png',
  },
  TIMEZONE: 'GMT+6', // e.g. 'Asia/Dhaka', 'UTC', 'America/New_York'
};

const PROPERTY_KEYS = {
  EXCLUDED_USERS: 'EXCLUDED_USERS',
  DISCORD_WEBHOOK: 'DISCORD_WEBHOOK',
  SLACK_TOKEN: 'SLACK_TOKEN',
  SLACK_CHANNEL_ID: 'SLACK_CHANNEL_ID',
  SLACK_OWNER_ID: 'SLACK_OWNER_ID',
  GOOGLE_SHEET_ID: 'GOOGLE_SHEET_ID',
  LAST_SLACK_MESSAGE_TS: 'LAST_SLACK_MESSAGE_TS',
  LAST_DATE: 'LAST_DATE',
  LAST_DISCORD_MESSAGE_ID: 'LAST_DISCORD_MESSAGE_ID',
};

/** Cached script properties — fetched once per execution to reduce API calls 🐾 */
let _propertyCache: ScriptProperties | null = null;

function getProperties(): ScriptProperties {
  if (!_propertyCache) {
    const props = PropertiesService.getScriptProperties();
    _propertyCache = {
      DISCORD_WEBHOOK: props.getProperty(PROPERTY_KEYS.DISCORD_WEBHOOK) || '',
      SLACK_TOKEN: props.getProperty(PROPERTY_KEYS.SLACK_TOKEN) || '',
      SLACK_CHANNEL_ID: props.getProperty(PROPERTY_KEYS.SLACK_CHANNEL_ID) || '',
      SLACK_OWNER_ID: props.getProperty(PROPERTY_KEYS.SLACK_OWNER_ID) || '',
      GOOGLE_SHEET_ID: props.getProperty(PROPERTY_KEYS.GOOGLE_SHEET_ID) || '',
      EXCLUDED_USERS: (props.getProperty(PROPERTY_KEYS.EXCLUDED_USERS) || '').split(',').map(s => s.trim()).filter(Boolean),
    };
  }
  return _propertyCache;
}

/** Cached date config — read once per execution from the Holidays sheet tab 🐾 */
let _dateConfig: DateConfig | null = null;

/**
 * Reads holidays, offdays, and Ramadan dates from the `Holidays` sheet tab.
 * Result is cached for the lifetime of the execution — single `getValues()` call.
 *
 * @remarks
 * Sheet layout expected:
 * - Col A: Date  |  Col B: Name  |  Col C: Type (`Holiday` or `Offday`)
 * - E2: `Start`, F2: Ramadan start date
 * - E3: `End`,   F3: Ramadan end date
 */
function getDateConfig(): DateConfig {
  if (_dateConfig) return _dateConfig;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Holidays');

  if (!sheet) {
    // 🙀 Hissing — no Holidays tab found, falling back to empty lists
    console.warn('🙀 Holidays sheet not found — date config will be empty!');
    _dateConfig = { holidays: [], offdays: [], workingHolidays: [], permittedHomeOffice: [] };
    return _dateConfig;
  }

  const data = sheet.getDataRange().getValues();
  const holidays: { date: string; name: string }[] = [];
  const offdays: { date: string; name: string }[] = [];
  const workingHolidays: { date: string; name: string }[] = [];
  const permittedHomeOffice: { name: string; start: string; end: string }[] = [];

  // Row 0 is the header — start at row 1 🐾
  for (let i = 1; i < data.length; i++) {
    const rawDate = data[i][0];
    const holName = String(data[i][1] || '').trim();
    const type = String(data[i][2] || '').trim();

    if (rawDate && type) {
      const dateStr = Utilities.formatDate(new Date(rawDate), CONFIG.TIMEZONE, 'yyyy-MM-dd');
      if (type === 'Holiday') holidays.push({ date: dateStr, name: holName });
      else if (type === 'Offday') offdays.push({ date: dateStr, name: holName });
      else if (type === 'W. Holiday') workingHolidays.push({ date: dateStr, name: holName });
    }

    // Parse Permitted Home Office ranges from cols E, F, G (indices 4, 5, 6)
    const phName = String(data[i][4] || '').trim();
    const phStart = data[i][5];
    const phEnd = data[i][6];

    if (phName && phStart && phEnd) {
      const toDateStr = (val: any) => (val instanceof Date ? Utilities.formatDate(val, CONFIG.TIMEZONE, 'yyyy-MM-dd') : String(val).trim());
      permittedHomeOffice.push({
        name: phName,
        start: toDateStr(phStart),
        end: toDateStr(phEnd)
      });
    }
  }

  _dateConfig = { holidays, offdays, workingHolidays, permittedHomeOffice };
  console.log(`😸 Date config loaded — ${holidays.length} holiday(s), ${offdays.length} offday(s), ${workingHolidays.length} working holiday(s), ${permittedHomeOffice.length} perm HO range(s).`);
  return _dateConfig;
}
