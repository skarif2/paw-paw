describe('Slack Module', () => {
  beforeEach(() => {
    (global as any)._propertyCache = null;
    (global as any)._dateConfig = null;
    ((global as any).UrlFetchApp.fetch as jest.Mock).mockClear();
    (global as any).PropertiesService.getScriptProperties.mockReturnValue({
      getProperty: jest.fn((k) => {
        if (k === 'SLACK_TOKEN') return 'xoxb-test';
        if (k === 'SLACK_CHANNEL_ID') return 'C123';
        if (k === 'EXCLUDED_USERS') return 'U000';
        return '';
      })
    });
    jest.spyOn(console, 'log').mockImplementation(() => {});
    jest.spyOn(console, 'warn').mockImplementation(() => {});
  });

  afterEach(() => {
    jest.restoreAllMocks();
  });

  describe('getUserInfoBatch', () => {
    it('should map valid users and skip bots/deleted', () => {
      (global as any).UrlFetchApp.fetch.mockReturnValue({
        getResponseCode: () => 200,
        getContentText: () => JSON.stringify({
          ok: true,
          user: {
            is_bot: false, deleted: false,
            profile: { display_name_normalized: 'Omen', email: 'omen@example.com' }
          }
        })
      });

      const users = getUserInfoBatch(['U111']);
      expect(users['U111']).toBeDefined();
      expect(users['U111'].name).toBe('Omen');
      expect((global as any).UrlFetchApp.fetch).toHaveBeenCalledTimes(1);
    });

    it('should sleep when paging over batch size', () => {
      (global as any).UrlFetchApp.fetch.mockReturnValue({
        getResponseCode: () => 200,
        getContentText: () => JSON.stringify({ ok: true, user: { is_bot: true } }) // All bots, skipped from map
      });
      
      const mockSleep = (global as any).Utilities.sleep as jest.Mock;
      mockSleep.mockClear();

      const userIds = Array.from({ length: 15 }, (_, i) => `U${i}`);
      getUserInfoBatch(userIds); // 1.5 batches

      expect(mockSleep).toHaveBeenCalledWith(100);
      expect((global as any).UrlFetchApp.fetch).toHaveBeenCalledTimes(15);
    });
  });

  describe('getChannelUsers', () => {
    it('should paginate and exclude users defined in properties', () => {
      (global as any).UrlFetchApp.fetch
        .mockReturnValueOnce({  // conversations.members
          getResponseCode: () => 200,
          getContentText: () => JSON.stringify({
            ok: true,
            members: ['U000', 'U111']
          })
        })
        .mockReturnValueOnce({ // users.info (U000 is excluded, so only U111 is requested)
          getResponseCode: () => 200,
          getContentText: () => JSON.stringify({
            ok: true,
            user: { is_bot: false, profile: { display_name_normalized: 'John' } }
          })
        });

      // Clear the property cache from setup.ts to hit mock again if needed
      if (typeof _propertyCache !== 'undefined') (global as any)._propertyCache = null;

      const userMap = getChannelUsers();
      
      expect(Object.keys(userMap)).toEqual(['U111']);
      expect(userMap['U111'].name).toBe('John');
    });
  });

  describe('buildDailyBriefingBlocks', () => {
    it('should correctly format blocks for given attendance groups', () => {
      const groups = { "Office": ["U111"], "WFH": ["U222", "U333"], "Leave": [] };
      const blocks = buildDailyBriefingBlocks("Monday", groups);

      expect(blocks).toHaveLength(6);
      
      const richTextObj: any = blocks.find((b: any) => b.type === 'rich_text');
      expect(richTextObj).toBeDefined();
      expect(richTextObj.elements).toHaveLength(3); // Office, WFH, Leave
      
      // Office element
      expect(richTextObj.elements[0].elements).toHaveLength(2); // "ON-SITE:", U111
      
      // Leave element -- empty
      expect(richTextObj.elements[2].elements).toHaveLength(2); // "Leave:", "None" (fallback text)
      expect(richTextObj.elements[2].elements[1].text).toBe('None');
    });
  });
});
