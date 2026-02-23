/** Global configuration constants for Paw-Paw 🐾 */
const CONFIG: AppConfig = {
  SLACK_API_BASE: 'https://slack.com/api',
  DISCORD: {
    USERNAME: 'Saga Check-in 🍽️',
    AVATAR_URL: 'https://cdn.jsdelivr.net/gh/twitter/twemoji@14.0.2/assets/72x72/1f37d.png',
  },
  TIMEZONE: 'GMT+6', // e.g. 'Asia/Dhaka', 'UTC', 'America/New_York'
};

const PROPERTY_KEYS = {
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
    _dateConfig = { holidays: [], offdays: [], ramadan: { start: '', end: '' } };
    return _dateConfig;
  }

  const data = sheet.getDataRange().getValues();
  const holidays: string[] = [];
  const offdays: string[] = [];

  // Row 0 is the header — start at row 1 🐾
  for (let i = 1; i < data.length; i++) {
    const raw = data[i][0];
    const type = String(data[i][2]).trim();

    if (!raw || !type) continue;

    const dateStr = Utilities.formatDate(new Date(raw), CONFIG.TIMEZONE, 'yyyy-MM-dd');
    if (type === 'Holiday') holidays.push(dateStr);
    else if (type === 'Offday') offdays.push(dateStr);
  }

  // Ramadan: E2 = 'Start' / F2 = date; E3 = 'End' / F3 = date
  const toDateStr = (val: any): string => {
    if (!val) return '';
    if (val instanceof Date) return Utilities.formatDate(val, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    return String(val).trim();
  };

  const ramadan = {
    start: data.length > 1 ? toDateStr(data[1][5]) : '',
    end:   data.length > 2 ? toDateStr(data[2][5]) : '',
  };

  _dateConfig = { holidays, offdays, ramadan };
  console.log(`😸 Date config loaded — ${holidays.length} holiday(s), ${offdays.length} offday(s).`);
  return _dateConfig;
}
