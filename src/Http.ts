/**
 * Makes an HTTP request with error handling and exponential-backoff retries.
 * @param {string} url - The request URL
 * @param {Object} options - URLFetch options (method, headers, payload, etc.)
 * @param {number} maxRetries - Maximum number of retry attempts (default: 3)
 * @returns {HttpRequestResult} Parsed response or throws after all retries are exhausted
 */
function makeHttpRequest(url: string, options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {}, maxRetries: number = 3): HttpRequestResult {
  let lastError: any;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const response = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        ...options,
      });

      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();

      if (responseCode >= 200 && responseCode < 300) {
        return {
          success: true,
          data: responseText ? JSON.parse(responseText) : null,
          responseCode,
        };
      }

      // Handle rate limiting with exponential backoff, capped at 10 seconds
      if (responseCode === 429) {
        const delay = Math.min(1000 * Math.pow(2, attempt), 10000);
        console.warn(`😾 Rate limited — napping for ${delay}ms then trying again (attempt ${attempt}).`);
        Utilities.sleep(delay);
        continue;
      }

      return {
        success: false,
        data: null,
        error: `HTTP ${responseCode}: ${responseText}`,
        responseCode,
      };
    } catch (error) {
      lastError = error;
      if (attempt < maxRetries) {
        const delay = 1000 * attempt;
        console.warn(
          `😿 Request failed — waiting ${delay}ms before another attempt (${attempt}):`,
          (error as any).toString(),
        );
        Utilities.sleep(delay);
      }
    }
  }

  throw new Error(`Request failed after ${maxRetries} attempts: ${lastError?.toString()}`);
}
