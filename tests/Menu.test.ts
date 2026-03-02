describe('Menu Module', () => {
  describe('onOpen', () => {
    it('should create menu if user is owner', () => {
      const mockCreateMenu = jest.fn().mockReturnThis();
      const mockAddItem = jest.fn().mockReturnThis();
      const mockAddSeparator = jest.fn().mockReturnThis();
      const mockAddToUi = jest.fn().mockReturnThis();

      (global as any).SpreadsheetApp.getUi = jest.fn(() => ({
        createMenu: mockCreateMenu,
        addItem: mockAddItem,
        addSeparator: mockAddSeparator,
        addToUi: mockAddToUi
      }));

      (global as any).SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => ({
        getOwner: () => ({ getEmail: () => 'owner@example.com' })
      }));
      (global as any).Session.getActiveUser = jest.fn(() => ({
        getEmail: () => 'owner@example.com'
      }));

      onOpen();

      expect(mockCreateMenu).toHaveBeenCalledWith('Paw-Paw 🐱');
      expect(mockAddItem).toHaveBeenCalledWith('📅 Create Sheet', 'promptCreateSheetForMonth');
      expect(mockAddToUi).toHaveBeenCalled();
    });

    it('should silently return if user is not owner', () => {
      const mockCreateMenu = jest.fn();
      (global as any).SpreadsheetApp.getUi = jest.fn(() => ({
        createMenu: mockCreateMenu
      }));

      (global as any).SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => ({
        getOwner: () => ({ getEmail: () => 'owner@example.com' })
      }));
      // 😼 Hacker trying to open the UI
      (global as any).Session.getActiveUser = jest.fn(() => ({
        getEmail: () => 'hacker@example.com'
      }));

      onOpen();

      expect(mockCreateMenu).not.toHaveBeenCalled();
    });
  });
});
