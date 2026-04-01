import { CallToolResult } from "@modelcontextprotocol/sdk/types.js";
import { OAuth2Client } from "google-auth-library";
import { BaseToolHandler } from "./BaseToolHandler.js";
import { calendar_v3 } from 'googleapis';
import { applyEventFilters } from "../../filters/event-filter.js";
import { createStructuredResponse } from "../../utils/response-builder.js";
import { convertToRFC3339, convertLocalTimeToUTC } from "../../utils/datetime.js";
import {
    GetAvailabilityResponse,
    BusyBlock,
    FreeBlock,
} from "../../types/structured-responses.js";

interface GetAvailabilityArgs {
    timeMin: string;
    timeMax: string;
    calendarId?: string | string[];
    timezone?: string;
    account?: string | string[];
}

/**
 * Slim representation of a calendar event for availability computation.
 * Only includes fields needed to determine busy/free status.
 */
interface SlimEvent {
    id: string;
    recurringEventId?: string;
    calendarId: string;
    start: string;
    end: string;
    isAllDay: boolean;
    transparent: boolean;
    declined: boolean;
    extendedProperties?: calendar_v3.Schema$Event['extendedProperties'];
}

/** Minimal field mask for slim event fetching. */
const SLIM_FIELDS = 'items(id,recurringEventId,status,start,end,transparency,attendees(self,responseStatus),extendedProperties),nextPageToken';

/**
 * Checks whether the authenticated user has declined an event.
 */
function isDeclined(event: calendar_v3.Schema$Event): boolean {
    if (!event.attendees) return false;
    const self = event.attendees.find(a => a.self);
    return self?.responseStatus === 'declined';
}

/** Cached formatters keyed by timezone (tz is constant within a single runTool call). */
/**
 * Returns the full day-of-week name for an ISO datetime string.
 */
function dayName(isoDate: string, tz: string): string {
    try {
        const date = new Date(isoDate);
        if (isNaN(date.getTime())) return '';
        return new Intl.DateTimeFormat('en-US', { weekday: 'long', timeZone: tz }).format(date);
    } catch {
        return '';
    }
}

/**
 * Returns a human-readable date label like "Monday, March 30".
 */
function dayLabel(isoDate: string, tz: string): string {
    try {
        const date = new Date(isoDate);
        if (isNaN(date.getTime())) return '';
        return new Intl.DateTimeFormat('en-US', {
            weekday: 'long',
            month: 'long',
            day: 'numeric',
            timeZone: tz,
        }).format(date);
    } catch {
        return '';
    }
}

/**
 * Converts an all-day event date range (YYYY-MM-DD) to full ISO datetime range
 * anchored at the start-of-day in the given timezone.
 */
function allDayToRange(startDate: string, endDate: string, tz: string): { start: string; end: string } {
    // All-day events use date strings like "2025-01-01" where end is exclusive.
    // Convert to the start-of-day in the target timezone expressed as UTC ISO.
    const toStartOfDay = (dateStr: string): string => {
        const [year, month, day] = dateStr.split('-').map(Number);
        return convertLocalTimeToUTC(year, month - 1, day, 0, 0, 0, tz)
            .toISOString();
    };

    return { start: toStartOfDay(startDate), end: toStartOfDay(endDate) };
}

/**
 * Converts an ISO datetime string to epoch milliseconds.
 */
function toEpoch(iso: string): number {
    return new Date(iso).getTime();
}

/**
 * Computes free slots from busy blocks within a time window.
 * Expects `busy` to be pre-sorted by start time (ascending).
 * Busy blocks are merged (overlapping/adjacent blocks coalesced) before gap detection.
 */
function computeFreeSlots(
    busy: Array<{ start: string; end: string }>,
    windowStart: string,
    windowEnd: string,
    tz: string,
): FreeBlock[] {
    if (busy.length === 0) {
        return [{
            start: windowStart,
            end: windowEnd,
            day: dayName(windowStart, tz),
            label: dayLabel(windowStart, tz),
        }];
    }

    const merged: Array<{ start: string; end: string }> = [{ ...busy[0] }];

    for (let i = 1; i < busy.length; i++) {
        const last = merged[merged.length - 1];
        if (toEpoch(busy[i].start) <= toEpoch(last.end)) {
            if (toEpoch(busy[i].end) > toEpoch(last.end)) last.end = busy[i].end;
        } else {
            merged.push({ ...busy[i] });
        }
    }

    const free: FreeBlock[] = [];
    let cursor = windowStart;

    for (const block of merged) {
        if (toEpoch(block.start) > toEpoch(cursor)) {
            free.push({
                start: cursor,
                end: block.start,
                day: dayName(cursor, tz),
                label: dayLabel(cursor, tz),
            });
        }
        if (toEpoch(block.end) > toEpoch(cursor)) cursor = block.end;
    }

    if (toEpoch(cursor) < toEpoch(windowEnd)) {
        free.push({
            start: cursor,
            end: windowEnd,
            day: dayName(cursor, tz),
            label: dayLabel(cursor, tz),
        });
    }

    return free;
}

export class GetAvailabilityHandler extends BaseToolHandler {
    async runTool(args: GetAvailabilityArgs, accounts: Map<string, OAuth2Client>): Promise<CallToolResult> {
        const selectedAccounts = this.getClientsForAccounts(args.account, accounts);

        // Determine timezone: explicit arg, or primary calendar's timezone, or UTC
        const tz = args.timezone ?? await this.resolveTimezone(selectedAccounts);

        const timeMin = convertToRFC3339(args.timeMin, tz);
        const timeMax = convertToRFC3339(args.timeMax, tz);

        const calendarIds = this.normalizeCalendarIds(args.calendarId);
        let accountCalendarMap: Map<string, string[]>;
        const resolutionWarnings: string[] = [];

        if (selectedAccounts.size > 1 || (calendarIds.length > 0 && calendarIds.some(id => id !== 'primary' && !id.includes('@')))) {
            if (calendarIds.length === 0) {
                // No calendar IDs specified -- query primary on every account
                accountCalendarMap = new Map<string, string[]>();
                for (const accountId of selectedAccounts.keys()) {
                    accountCalendarMap.set(accountId, ['primary']);
                }
            } else if (selectedAccounts.size > 1) {
                const { resolved, warnings } = await this.calendarRegistry.resolveCalendarsToAccounts(
                    calendarIds,
                    selectedAccounts,
                );
                accountCalendarMap = resolved;
                resolutionWarnings.push(...warnings);

                if (accountCalendarMap.size === 0) {
                    await this.throwNoCalendarsFoundError(calendarIds, selectedAccounts);
                }
            } else {
                // Single account with name-based calendars -- resolve names
                const [accountId] = selectedAccounts.keys();
                const client = selectedAccounts.get(accountId)!;
                const resolvedIds = await this.resolveCalendarIds(client, calendarIds);
                accountCalendarMap = new Map([[accountId, resolvedIds]]);
            }
        } else {
            // Simple case: single account, IDs already look like IDs (or nothing specified)
            const [accountId] = selectedAccounts.keys();
            accountCalendarMap = new Map([[accountId, calendarIds.length > 0 ? calendarIds : ['primary']]]);
        }

        // Fetch slim events from all accounts/calendars
        const allSlimEvents: SlimEvent[] = [];
        const allCalendarsChecked: string[] = [];
        const errors: Array<{ calendarId: string; error: string }> = [];

        await Promise.all(
            Array.from(accountCalendarMap.entries()).map(async ([accountId, calendarsForAccount]) => {
                const client = selectedAccounts.get(accountId)!;
                const results = await Promise.all(calendarsForAccount.map(async (calendarId) => {
                    try {
                        return { calendarId, events: await this.fetchEventsSlim(client, calendarId, timeMin, timeMax, tz) };
                    } catch (err) {
                        const message = err instanceof Error ? err.message : String(err);
                        process.stderr.write(`get-availability: failed to fetch from ${calendarId} (${accountId}): ${message}\n`);
                        return { calendarId, events: [] as SlimEvent[], error: message };
                    }
                }));
                for (const r of results) {
                    allSlimEvents.push(...r.events);
                    allCalendarsChecked.push(r.calendarId);
                    if (r.error) errors.push({ calendarId: r.calendarId, error: r.error });
                }
            }),
        );

        const { events: filteredEvents, eventsFiltered } = applyEventFilters(allSlimEvents, this.eventFilters);

        const busy: BusyBlock[] = [];
        for (const event of filteredEvents) {
            if (event.transparent || event.declined) continue;
            busy.push({
                start: event.start,
                end: event.end,
                day: dayName(event.start, tz),
                label: dayLabel(event.start, tz),
                calendarId: event.calendarId,
                isAllDay: event.isAllDay,
            });
        }

        busy.sort((a, b) => toEpoch(a.start) - toEpoch(b.start));

        const busyForGaps = busy.map(b => {
            if (b.isAllDay) return allDayToRange(b.start, b.end, tz);
            return { start: b.start, end: b.end };
        });
        const free = computeFreeSlots(busyForGaps, timeMin, timeMax, tz);

        const response: GetAvailabilityResponse = {
            timezone: tz,
            window: { start: timeMin, end: timeMax },
            busy,
            free,
            calendars_checked: allCalendarsChecked,
            events_filtered: eventsFiltered,
            ...(errors.length > 0 && { errors }),
            ...(resolutionWarnings.length > 0 && { warnings: resolutionWarnings }),
        };

        return createStructuredResponse(response);
    }

    /**
     * Fetch events with minimal fields for availability computation.
     */
    private async fetchEventsSlim(
        client: OAuth2Client,
        calendarId: string,
        timeMin: string,
        timeMax: string,
        tz: string,
    ): Promise<SlimEvent[]> {
        const calendar = this.getCalendar(client);
        const results: SlimEvent[] = [];
        let pageToken: string | undefined;

        do {
            const res = await calendar.events.list({
                calendarId,
                timeMin,
                timeMax,
                singleEvents: true,
                orderBy: 'startTime',
                timeZone: tz,
                maxResults: 2500,
                pageToken,
                fields: SLIM_FIELDS,
            });

            for (const event of res.data.items ?? []) {
                if (event.status === 'cancelled') continue;
                if (!event.id || !event.start || !event.end) continue;

                const isAllDay = !!event.start.date;
                const start = isAllDay ? event.start.date : event.start.dateTime;
                const end = isAllDay ? event.end.date : event.end.dateTime;
                if (!start || !end) continue;

                results.push({
                    id: event.id,
                    ...(event.recurringEventId && { recurringEventId: event.recurringEventId }),
                    calendarId,
                    start,
                    end,
                    isAllDay,
                    transparent: event.transparency === 'transparent',
                    declined: isDeclined(event),
                    extendedProperties: event.extendedProperties,
                });
            }

            pageToken = res.data.nextPageToken ?? undefined;
        } while (pageToken);

        return results;
    }

    /**
     * Resolve timezone from the first available account's primary calendar.
     */
    private async resolveTimezone(accounts: Map<string, OAuth2Client>): Promise<string> {
        const sortedAccountIds = Array.from(accounts.keys()).sort();
        for (const accountId of sortedAccountIds) {
            const client = accounts.get(accountId);
            if (!client) continue;
            try {
                return await this.getCalendarTimezone(client, 'primary');
            } catch {
                // Continue to next account
            }
        }
        return 'UTC';
    }

    /**
     * Normalise the calendarId argument to an array of strings.
     */
    private normalizeCalendarIds(calendarId: string | string[] | undefined): string[] {
        if (!calendarId) return [];
        if (Array.isArray(calendarId)) return calendarId;
        return [calendarId];
    }
}
