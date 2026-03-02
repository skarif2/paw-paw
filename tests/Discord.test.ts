describe('Discord Module', () => {
  beforeEach(() => {
    // Reset internal caches and mocks
    (global as any)._propertyCache = null;
    (global as any)._dateConfig = null;
    
    (global as any).PropertiesService.getScriptProperties.mockReturnValue({
      getProperty: jest.fn((key) => {
        if (key === 'DISCORD_WEBHOOK') return 'https://discord.com/api/webhooks/123/token';
        return null;
      }),
      setProperty: jest.fn(),
      setProperties: jest.fn(),
    });
    ((global as any).UrlFetchApp.fetch as jest.Mock).mockClear();
    
    // Silence console
    jest.spyOn(console, 'log').mockImplementation(() => {});
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    jest.restoreAllMocks();
  });

  describe('getWebhookMessageUrl', () => {
    it('should format message URL correctly', () => {
      const webhookUrl = 'https://discord.com/api/webhooks/12345/token67890';
      const expected = 'https://discord.com/api/webhooks/12345/token67890/messages/msg111';
      expect(getWebhookMessageUrl(webhookUrl, 'msg111')).toBe(expected);
    });
  });

  describe('getMealMessageContent', () => {
    it('should format lunch only when not permitted HO', () => {
      // Mock _dateConfig
      (global as any)._dateConfig = {
        permittedHomeOffice: [],
        holidays: [], offdays: [], workingHolidays: []
      };
      
      const content = getMealMessageContent('2026-03-05', 15);
      expect(content).toBe('Lunch: **15**');
    });

    it('should format Iftar only when within Permitted HO date range', () => {
      (global as any)._dateConfig = {
        permittedHomeOffice: [{ start: '2026-02-18', end: '2026-03-20', name: 'Ramadan' }],
        holidays: [], offdays: [], workingHolidays: []
      };
      
      const content = getMealMessageContent('2026-03-05', 10);
      expect(content).toBe('Lunch: **0**, Iftar: **10**');
    });
  });

  describe('sendNewDiscordMessage', () => {
    it('should POST message and save message ID to script properties', () => {
      (global as any).UrlFetchApp.fetch.mockReturnValue({
        getResponseCode: () => 200,
        getContentText: () => JSON.stringify({ id: 'msg999' })
      });

      const mockSetProperties = jest.fn();
      (global as any).PropertiesService.getScriptProperties.mockReturnValue({
        getProperty: jest.fn(),
        setProperties: mockSetProperties
      });
      // Ensure date config passes
      (global as any)._dateConfig = { permittedHomeOffice: [] };

      sendNewDiscordMessage('2026-03-05', 12);

      expect((global as any).UrlFetchApp.fetch).toHaveBeenCalledWith(
        expect.stringContaining('?wait=true'),
        expect.objectContaining({ method: 'post' })
      );
      expect(mockSetProperties).toHaveBeenCalledWith({
        'LAST_DATE': '2026-03-05',
        'LAST_DISCORD_MESSAGE_ID': 'msg999'
      });
    });
  });

  describe('updateExistingDiscordMessage', () => {
    it('should PATCH if message exists', () => {
      (global as any).UrlFetchApp.fetch.mockReturnValue({
        getResponseCode: () => 200,
        getContentText: () => JSON.stringify({ id: 'msg999' })
      });
      (global as any)._dateConfig = { permittedHomeOffice: [] };

      updateExistingDiscordMessage('msg999', 15, '2026-03-05');

      expect((global as any).UrlFetchApp.fetch).toHaveBeenCalledWith(
        expect.stringContaining('/messages/msg999?wait=true'),
        expect.objectContaining({ method: 'patch' })
      );
    });
  });

  describe('lockRowByDate', () => {
    it('should lock the correct row in the spreadsheet', () => {
      const mockProtection = {
        setDescription: jest.fn(),
        removeEditors: jest.fn(),
        canDomainEdit: jest.fn(() => true),
        setDomainEdit: jest.fn(),
        getEditors: jest.fn(() => []),
      };

      const mockGetRange = jest.fn(() => ({ protect: () => mockProtection }));
      const mockGetSheetByName = jest.fn(() => ({
        getDataRange: () => ({
          getValues: () => [
            ['Header'],
            ['IdRow'],
            [new Date('2026-03-04T00:00:00Z'), 'Office'],
            [new Date('2026-03-05T00:00:00Z'), 'WFH']
          ]
        }),
        getRange: mockGetRange
      }));

      (global as any).SpreadsheetApp.getActiveSpreadsheet.mockReturnValue({
        getSheetByName: mockGetSheetByName
      });
      (global as any).Utilities.formatDate.mockImplementation((d: Date, tz: string, fmt: string) => {
        if (fmt === 'yyyy-MM-dd') return d.toISOString().split('T')[0];
        if (fmt === 'yyyy-MM') return d.toISOString().split('T')[0].substring(0, 7);
        return '';
      });

      lockRowByDate('2026-03-05');
      // March 5th is index 3 -> row 4
      expect(mockGetRange).toHaveBeenCalledWith(4, 1, 1, 2);
      expect(mockProtection.setDescription).toHaveBeenCalledWith('Locked Past Date: 2026-03-05');
    });
  });
});
