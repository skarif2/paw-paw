describe('Spreadsheet Module', () => {
  describe('getSheetInfo', () => {
    it('should parse yyyy-MM sheet names', () => {
      const mockSheet = { getName: () => '2026-03' };
      const info = getSheetInfo(mockSheet as any);
      expect(info.year).toBe(2026);
      expect(info.month).toBe(3);
      expect(info.daysInMonth).toBe(31);
    });
  });

  describe('columnToLetter', () => {
    it('should convert numbers to Excel columns', () => {
      expect(columnToLetter(1)).toBe('A');
      expect(columnToLetter(26)).toBe('Z');
      expect(columnToLetter(27)).toBe('AA');
    });
  });

  describe('createSheetForMonth', () => {
    it('should reject invalid format', () => {
      const mockError = jest.spyOn(console, 'error').mockImplementation(() => {});
      createSheetForMonth('invalid-date');
      expect(mockError).toHaveBeenCalled();
      mockError.mockRestore();
    });
    
    it('should abort if sheet already exists', () => {
      const mockGetSheetByName = jest.fn(() => ({})); // Returns a truthy object (sheet exists)
      (global as any).SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => ({
        getSheetByName: mockGetSheetByName
      }));
      (global as any).Utilities.formatDate = jest.fn(() => '2026-03');

      createSheetForMonth('2026-03');
      expect(mockGetSheetByName).toHaveBeenCalled();
      // Test will not throw because it aborted safely
    });
  });
});
