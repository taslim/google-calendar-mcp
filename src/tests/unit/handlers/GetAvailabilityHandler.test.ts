import { describe, it, expect, vi, beforeEach } from 'vitest';
import { GetAvailabilityHandler } from '../../../handlers/core/GetAvailabilityHandler.js';
import { OAuth2Client } from 'google-auth-library';
import { CalendarRegistry } from '../../../services/CalendarRegistry.js';

// Mock the googleapis module
vi.mock('googleapis', () => ({
  google: {
    calendar: vi.fn(() => ({
      events: {
        list: vi.fn()
      },
      calendarList: {
        get: vi.fn()
      }
    }))
  },
  calendar_v3: {}
}));

// Mock datetime utils
vi.mock('../../../utils/datetime.js', () => ({
  hasTimezoneInDatetime: vi.fn((datetime: string) =>
    /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(Z|[+-]\d{2}:\d{2})$/.test(datetime)
  ),
  convertToRFC3339: vi.fn((datetime: string) => {
    if (/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(Z|[+-]\d{2}:\d{2})$/.test(datetime)) {
      return datetime;
    }
    return `${datetime}Z`;
  }),
  createTimeObject: vi.fn((datetime: string, timezone: string) => ({
    dateTime: datetime,
    timeZone: timezone
  }))
}));

// Mock event filter
vi.mock('../../../filters/event-filter.js', () => ({
  applyEventFilters: vi.fn((events: unknown[], filters: unknown[] | undefined) => ({
    events,
    eventsFiltered: 0
  }))
}));

describe('GetAvailabilityHandler', () => {
  let handler: GetAvailabilityHandler;
  let mockOAuth2Client: OAuth2Client;
  let mockAccounts: Map<string, OAuth2Client>;
  let mockCalendar: { events: { list: ReturnType<typeof vi.fn> } };

  beforeEach(() => {
    CalendarRegistry.resetInstance();

    handler = new GetAvailabilityHandler();
    mockOAuth2Client = new OAuth2Client();
    mockAccounts = new Map([['test', mockOAuth2Client]]);

    mockCalendar = {
      events: {
        list: vi.fn()
      }
    };

    vi.spyOn(handler as any, 'getCalendar').mockReturnValue(mockCalendar);
    vi.spyOn(handler as any, 'getCalendarTimezone').mockResolvedValue('America/Los_Angeles');
  });

  describe('Basic Availability Query', () => {
    it('should compute busy and free blocks from events', async () => {
      mockCalendar.events.list.mockResolvedValue({
        data: {
          items: [
            {
              id: 'evt1',
              status: 'confirmed',
              start: { dateTime: '2025-01-15T10:00:00Z' },
              end: { dateTime: '2025-01-15T11:00:00Z' },
              transparency: 'opaque',
              attendees: []
            },
            {
              id: 'evt2',
              status: 'confirmed',
              start: { dateTime: '2025-01-15T14:00:00Z' },
              end: { dateTime: '2025-01-15T15:00:00Z' },
              transparency: 'opaque',
              attendees: []
            }
          ],
          nextPageToken: undefined
        }
      });

      const args = {
        timeMin: '2025-01-15T00:00:00Z',
        timeMax: '2025-01-15T23:59:59Z',
        calendarId: 'primary'
      };

      const result = await handler.runTool(args, mockAccounts);
      const response = JSON.parse(result.content[0].text);

      expect(response.busy).toHaveLength(2);
      expect(response.free).toHaveLength(3); // before, between, after
      expect(response.timezone).toBe('America/Los_Angeles');
      expect(response.calendars_checked).toContain('primary');
      expect(response.window.start).toBeDefined();
      expect(response.window.end).toBeDefined();
    });

    it('should skip transparent events', async () => {
      mockCalendar.events.list.mockResolvedValue({
        data: {
          items: [
            {
              id: 'evt1',
              status: 'confirmed',
              start: { dateTime: '2025-01-15T10:00:00Z' },
              end: { dateTime: '2025-01-15T11:00:00Z' },
              transparency: 'transparent',
              attendees: []
            }
          ],
          nextPageToken: undefined
        }
      });

      const args = {
        timeMin: '2025-01-15T00:00:00Z',
        timeMax: '2025-01-15T23:59:59Z',
        calendarId: 'primary'
      };

      const result = await handler.runTool(args, mockAccounts);
      const response = JSON.parse(result.content[0].text);

      expect(response.busy).toHaveLength(0);
      expect(response.free).toHaveLength(1); // entire window is free
    });

    it('should skip declined events', async () => {
      mockCalendar.events.list.mockResolvedValue({
        data: {
          items: [
            {
              id: 'evt1',
              status: 'confirmed',
              start: { dateTime: '2025-01-15T10:00:00Z' },
              end: { dateTime: '2025-01-15T11:00:00Z' },
              attendees: [{ self: true, responseStatus: 'declined' }]
            }
          ],
          nextPageToken: undefined
        }
      });

      const args = {
        timeMin: '2025-01-15T00:00:00Z',
        timeMax: '2025-01-15T23:59:59Z',
        calendarId: 'primary'
      };

      const result = await handler.runTool(args, mockAccounts);
      const response = JSON.parse(result.content[0].text);

      expect(response.busy).toHaveLength(0);
    });

    it('should skip cancelled events', async () => {
      mockCalendar.events.list.mockResolvedValue({
        data: {
          items: [
            {
              id: 'evt1',
              status: 'cancelled',
              start: { dateTime: '2025-01-15T10:00:00Z' },
              end: { dateTime: '2025-01-15T11:00:00Z' },
              attendees: []
            }
          ],
          nextPageToken: undefined
        }
      });

      const args = {
        timeMin: '2025-01-15T00:00:00Z',
        timeMax: '2025-01-15T23:59:59Z',
        calendarId: 'primary'
      };

      const result = await handler.runTool(args, mockAccounts);
      const response = JSON.parse(result.content[0].text);

      expect(response.busy).toHaveLength(0);
    });
  });

  describe('No Events', () => {
    it('should return full window as free when no events', async () => {
      mockCalendar.events.list.mockResolvedValue({
        data: { items: [], nextPageToken: undefined }
      });

      const args = {
        timeMin: '2025-01-15T00:00:00Z',
        timeMax: '2025-01-15T23:59:59Z',
        calendarId: 'primary'
      };

      const result = await handler.runTool(args, mockAccounts);
      const response = JSON.parse(result.content[0].text);

      expect(response.busy).toHaveLength(0);
      expect(response.free).toHaveLength(1);
    });
  });

  describe('Timezone Handling', () => {
    it('should use explicit timezone when provided', async () => {
      mockCalendar.events.list.mockResolvedValue({
        data: { items: [], nextPageToken: undefined }
      });

      const args = {
        timeMin: '2025-01-15T00:00:00Z',
        timeMax: '2025-01-15T23:59:59Z',
        calendarId: 'primary',
        timezone: 'Europe/London'
      };

      const result = await handler.runTool(args, mockAccounts);
      const response = JSON.parse(result.content[0].text);

      expect(response.timezone).toBe('Europe/London');
    });

    it('should fall back to primary calendar timezone', async () => {
      mockCalendar.events.list.mockResolvedValue({
        data: { items: [], nextPageToken: undefined }
      });

      const args = {
        timeMin: '2025-01-15T00:00:00Z',
        timeMax: '2025-01-15T23:59:59Z',
        calendarId: 'primary'
      };

      const result = await handler.runTool(args, mockAccounts);
      const response = JSON.parse(result.content[0].text);

      expect(response.timezone).toBe('America/Los_Angeles');
    });
  });

  describe('Response Format', () => {
    it('should include all required response fields', async () => {
      mockCalendar.events.list.mockResolvedValue({
        data: { items: [], nextPageToken: undefined }
      });

      const args = {
        timeMin: '2025-01-15T00:00:00Z',
        timeMax: '2025-01-15T23:59:59Z',
        calendarId: 'primary'
      };

      const result = await handler.runTool(args, mockAccounts);
      const response = JSON.parse(result.content[0].text);

      expect(response).toHaveProperty('timezone');
      expect(response).toHaveProperty('window');
      expect(response.window).toHaveProperty('start');
      expect(response.window).toHaveProperty('end');
      expect(response).toHaveProperty('busy');
      expect(response).toHaveProperty('free');
      expect(response).toHaveProperty('calendars_checked');
      expect(response).toHaveProperty('events_filtered');
    });

    it('should include day and label on busy blocks', async () => {
      mockCalendar.events.list.mockResolvedValue({
        data: {
          items: [
            {
              id: 'evt1',
              status: 'confirmed',
              start: { dateTime: '2025-01-15T10:00:00Z' },
              end: { dateTime: '2025-01-15T11:00:00Z' },
              attendees: []
            }
          ],
          nextPageToken: undefined
        }
      });

      const args = {
        timeMin: '2025-01-15T00:00:00Z',
        timeMax: '2025-01-15T23:59:59Z',
        calendarId: 'primary'
      };

      const result = await handler.runTool(args, mockAccounts);
      const response = JSON.parse(result.content[0].text);

      expect(response.busy[0]).toHaveProperty('day');
      expect(response.busy[0]).toHaveProperty('label');
      expect(response.busy[0]).toHaveProperty('calendarId');
      expect(response.busy[0]).toHaveProperty('isAllDay');
    });

    it('should include day and label on free blocks', async () => {
      mockCalendar.events.list.mockResolvedValue({
        data: { items: [], nextPageToken: undefined }
      });

      const args = {
        timeMin: '2025-01-15T00:00:00Z',
        timeMax: '2025-01-15T23:59:59Z',
        calendarId: 'primary'
      };

      const result = await handler.runTool(args, mockAccounts);
      const response = JSON.parse(result.content[0].text);

      expect(response.free[0]).toHaveProperty('day');
      expect(response.free[0]).toHaveProperty('label');
    });
  });

  describe('Error Handling', () => {
    it('should report errors for failing calendars without crashing', async () => {
      mockCalendar.events.list.mockRejectedValue(new Error('Calendar not found'));

      const args = {
        timeMin: '2025-01-15T00:00:00Z',
        timeMax: '2025-01-15T23:59:59Z',
        calendarId: 'nonexistent@calendar.google.com'
      };

      const result = await handler.runTool(args, mockAccounts);
      const response = JSON.parse(result.content[0].text);

      expect(response.errors).toBeDefined();
      expect(response.errors).toHaveLength(1);
      expect(response.errors[0].calendarId).toBe('nonexistent@calendar.google.com');
    });
  });

  describe('Multi-Account Handling', () => {
    it('should use all accounts when none specified', async () => {
      const spy = vi.spyOn(handler as any, 'getClientsForAccounts');
      mockCalendar.events.list.mockResolvedValue({
        data: { items: [], nextPageToken: undefined }
      });

      const args = {
        timeMin: '2025-01-15T00:00:00Z',
        timeMax: '2025-01-15T23:59:59Z',
        calendarId: 'primary'
      };

      await handler.runTool(args, mockAccounts);

      expect(spy).toHaveBeenCalledWith(undefined, mockAccounts);
    });

    it('should use specified account when provided', async () => {
      const spy = vi.spyOn(handler as any, 'getClientsForAccounts');
      mockCalendar.events.list.mockResolvedValue({
        data: { items: [], nextPageToken: undefined }
      });

      const args = {
        timeMin: '2025-01-15T00:00:00Z',
        timeMax: '2025-01-15T23:59:59Z',
        calendarId: 'primary',
        account: 'test'
      };

      await handler.runTool(args, mockAccounts);

      expect(spy).toHaveBeenCalledWith('test', mockAccounts);
    });
  });
});
