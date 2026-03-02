describe('Config Module', () => {
  beforeEach(() => {
    // Reset property caches before each test using the global setter injected by setup.ts
    // This allows us to clear out state between test runs.
    if (typeof _propertyCache !== 'undefined') {
      (global as any)._propertyCache = null;
    }
    if (typeof _dateConfig !== 'undefined') {
      (global as any)._dateConfig = null;
    }
  });

  describe('Global Constants', () => {
    it('should expose CONFIG and PROPERTY_KEYS globally', () => {
      expect(CONFIG).toBeDefined();
      expect(CONFIG.SLACK_API_BASE).toBe('https://slack.com/api');
      expect(CONFIG.TIMEZONE).toBe('GMT+6');
      expect(PROPERTY_KEYS).toBeDefined();
    });
  });

  describe('getProperties()', () => {
    it('should fetch from PropertiesService when cache is empty', () => {
      const mockProps = {
        DISCORD_WEBHOOK: 'https://webhook.test',
        EXCLUDED_USERS: 'U123, U456',
      };
      
      const mockGetProperty = jest.fn((key) => mockProps[key as keyof typeof mockProps] || null);
      (global as any).PropertiesService.getScriptProperties.mockReturnValue({
        getProperty: mockGetProperty,
      });

      const props = getProperties();
      
      expect(mockGetProperty).toHaveBeenCalled();
      expect(props.DISCORD_WEBHOOK).toBe('https://webhook.test');
      expect(props.EXCLUDED_USERS).toEqual(['U123', 'U456']);
    });

    it('should return cached properties on subsequent calls', () => {
      const mockGetProperty = jest.fn((key) => 'test');
      (global as any).PropertiesService.getScriptProperties.mockReturnValue({
        getProperty: mockGetProperty,
      });

      // First call fetches API
      getProperties();
      // Second call uses cache
      getProperties();

      expect(mockGetProperty).toHaveBeenCalledTimes(6); // 6 keys mapped in getProperties()
    });
  });

  describe('getDateConfig()', () => {
    it('should return empty config and warn if Holidays sheet is missing', () => {
      const mockWarn = jest.spyOn(console, 'warn').mockImplementation(() => {});
      (global as any).SpreadsheetApp.getActiveSpreadsheet.mockReturnValue({
        getSheetByName: jest.fn(() => null),
      });

      const dateConfig = getDateConfig();
      expect(dateConfig.holidays).toEqual([]);
      expect(dateConfig.workingHolidays).toEqual([]);
      expect(mockWarn).toHaveBeenCalledWith('🙀 Holidays sheet not found — date config will be empty!');
      mockWarn.mockRestore();
    });

    it('should parse holidays, offdays, and working holidays from sheet data', () => {
      const mockData = [
        ['Date', 'Name', 'Type', '', 'Permitted HO', 'Start', 'End'], // Header
        ['2026-03-26', 'Independence Day', 'Holiday', '', '', '', ''],
        ['2026-04-14', 'Pohela Boishakh', 'Offday', '', '', '', ''],
        ['2026-05-01', 'May Day', 'W. Holiday', '', 'Ramadan', '2026-02-18', '2026-03-20'],
      ];

      (global as any).SpreadsheetApp.getActiveSpreadsheet.mockReturnValue({
        getSheetByName: jest.fn(() => ({
          getDataRange: jest.fn(() => ({
            getValues: jest.fn(() => mockData),
          })),
        })),
      });
      (global as any).Utilities.formatDate.mockImplementation((d: Date) => d.toISOString().split('T')[0]);

      const dateConfig = getDateConfig();
      
      expect(dateConfig.holidays).toHaveLength(1);
      expect(dateConfig.holidays[0].name).toBe('Independence Day');
      
      expect(dateConfig.offdays).toHaveLength(1);
      expect(dateConfig.offdays[0].name).toBe('Pohela Boishakh');

      expect(dateConfig.workingHolidays).toHaveLength(1);
      
      expect(dateConfig.permittedHomeOffice).toHaveLength(1);
      expect(dateConfig.permittedHomeOffice[0].name).toBe('Ramadan');
    });
  });
});
