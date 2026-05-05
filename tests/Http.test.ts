describe('Http Module', () => {
  let mockFetch: jest.Mock;
  let mockSleep: jest.Mock;

  beforeEach(() => {
    mockFetch = (global as any).UrlFetchApp.fetch as jest.Mock;
    mockSleep = (global as any).Utilities.sleep as jest.Mock;
    mockFetch.mockReset();
    mockSleep.mockReset();
    
    // Silence console lines during expected log/warn tests
    jest.spyOn(console, 'warn').mockImplementation(() => {});
  });

  afterEach(() => {
    jest.restoreAllMocks();
  });

  it('should return successful response data for 2xx HTTP codes', () => {
    mockFetch.mockReturnValue({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({ ok: true })
    });

    const result = makeHttpRequest('https://example.com');
    expect(result.success).toBe(true);
    expect(result.data).toEqual({ ok: true });
    expect(mockFetch).toHaveBeenCalledTimes(1);
  });

  it('should not parse data if response string is empty', () => {
    mockFetch.mockReturnValue({
      getResponseCode: () => 204,
      getContentText: () => ''
    });

    const result = makeHttpRequest('https://example.com');
    expect(result.success).toBe(true);
    expect(result.data).toBe(null);
  });

  it('should retry on 429 rate limit with the updated fallback', () => {
    mockFetch
      .mockReturnValueOnce({ getResponseCode: () => 429, getContentText: () => 'Rate Limit', getAllHeaders: () => ({}) })
      .mockReturnValueOnce({ getResponseCode: () => 200, getContentText: () => '{"ok": true}' });

    const result = makeHttpRequest('https://example.com');
    
    expect(result.success).toBe(true);
    expect(mockFetch).toHaveBeenCalledTimes(2);
    expect(mockSleep).toHaveBeenCalledWith(20000); // attempt 1 fallback -> 20000
  });

  it('should retry on 429 by precisely parsing the Retry-After header', () => {
    mockFetch
      .mockReturnValueOnce({ 
        getResponseCode: () => 429, 
        getContentText: () => 'Rate Limit', 
        getAllHeaders: () => ({ 'Retry-After': '6.5' }) 
      })
      .mockReturnValueOnce({ getResponseCode: () => 200, getContentText: () => '{"ok": true}' });

    const result = makeHttpRequest('https://example.com');
    
    expect(result.success).toBe(true);
    expect(mockSleep).toHaveBeenCalledWith(6500); // 6.5s * 1000 = 6500ms
  });

  it('should handle runtime exceptions and retry', () => {
    mockFetch
      .mockImplementationOnce(() => { throw new Error('Network timeout'); })
      .mockReturnValueOnce({ getResponseCode: () => 200, getContentText: () => '{"ok": true}' });

    const result = makeHttpRequest('https://example.com');
    
    expect(result.success).toBe(true);
    expect(mockFetch).toHaveBeenCalledTimes(2);
    expect(mockSleep).toHaveBeenCalledWith(1000); // 1000 * attempt (1) for general exceptions
  });

  it('should exhaust maxRetries and throw an error', () => {
    mockFetch.mockImplementation(() => { throw new Error('Fatal generic Error'); });

    expect(() => makeHttpRequest('https://example.com', {}, 3)).toThrow('Request failed after 3 attempts');
    expect(mockFetch).toHaveBeenCalledTimes(3);
    // 1000ms on first failure, 2000ms on second
    expect(mockSleep).toHaveBeenNthCalledWith(1, 1000);
    expect(mockSleep).toHaveBeenNthCalledWith(2, 2000);
  });

  it('should return success = false when valid HTTP response is mapped outside 2xx', () => {
    mockFetch.mockReturnValue({
      getResponseCode: () => 404,
      getContentText: () => '{"error": "Not Found"}'
    });

    // Valid response code but outside 2xx throws success: false, it doesn't retry on non-429 client errors usually
    const result = makeHttpRequest('https://example.com');
    
    expect(result.success).toBe(false);
    expect(result.responseCode).toBe(404);
    expect(result.error).toContain('HTTP 404');
    // Only calls once, does not retry 400s or 500s unless explicit exception or 429
    expect(mockFetch).toHaveBeenCalledTimes(1); 
  });
});
