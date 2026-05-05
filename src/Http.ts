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

      // Handle rate limiting with explicit header parsing and safety caps
      if (responseCode === 429) {
        let delayMs = 0;
        lastError = new Error(`HTTP 429: Rate Limited (${responseText || 'No Body'})`);

        // 1. Try checking the HTTP headers first
        const headers = response.getAllHeaders();
        const retryAfterKey = Object.keys(headers).find(k => k.toLowerCase() === 'retry-after');
        if (retryAfterKey) {
          const headerVal = parseFloat(headers[retryAfterKey] as string);
          if (!isNaN(headerVal)) delayMs = headerVal * 1000;
        }

        // 2. Override with JSON body if available (usually more precise for Discord webhooks)
        try {
          if (responseText) {
            const errorBody = JSON.parse(responseText);
            if (errorBody && typeof errorBody.retry_after === 'number') {
              delayMs = errorBody.retry_after * 1000;
            }
          }
        } catch (e) {
          // Ignore parse errors (e.g. Cloudflare HTML pages)
        }

        // 3. Fallback & Caps
        if (delayMs <= 0) {
          const fallbacks = [20000, 60000, 120000];
          delayMs = fallbacks[attempt - 1] || 120000;
        }

        // 4. Safety abort to prevent Google Apps Script timeouts
        if (delayMs > 150000) {
          throw new Error(`Discord demanded a sleep of ${delayMs / 1000}s, which exceeds the 150s safety cap. Aborting!`);
        }

        console.warn(`😾 Rate limited — napping for ${delayMs}ms then trying again (attempt ${attempt}).`);
        Utilities.sleep(delayMs);
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
