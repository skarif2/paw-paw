interface DateConfig {
  holidays: string[];
  offdays: string[];
  ramadan: { start: string; end: string };
}

interface AppConfig {
  SLACK_API_BASE: string;
  DISCORD: { USERNAME: string; AVATAR_URL: string };
  TIMEZONE: string;
}

interface ScriptProperties {
  DISCORD_WEBHOOK: string;
  SLACK_TOKEN: string;
  SLACK_CHANNEL_ID: string;
  SLACK_OWNER_ID: string;
  GOOGLE_SHEET_ID: string;
  LAST_SLACK_MESSAGE_TS?: string;
  LAST_DATE?: string;
  LAST_DISCORD_MESSAGE_ID?: string;
}

interface SlackUser {
  id: string;
  name: string;
  email: string;
  image?: string;
}

interface HttpRequestResult {
  success: boolean;
  data: any;
  responseCode: number;
  error?: string;
}

// Declare Google Apps Script global explicitly for IDE
declare namespace GoogleAppsScript {
  export namespace Spreadsheet {
    export interface Sheet {}
  }
  export namespace URL_Fetch {
    export interface URLFetchRequestOptions {}
  }
}
