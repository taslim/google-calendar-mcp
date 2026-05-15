/**
 * Activated when `PRINCIPAL_EMAILS` is set. Under service-account + DWD
 * impersonation, Google's `"primary"` keyword resolves to the impersonated
 * EA's own calendar — almost never the intended one. This module redirects
 * that default to the principal's calendars instead.
 */

import { McpError, ErrorCode } from "@modelcontextprotocol/sdk/types.js";

export const PRIMARY_ALIAS = 'primary';

let cached: string[] | null = null;

function load(): string[] {
    if (cached !== null) return cached;
    const raw = process.env.PRINCIPAL_EMAILS ?? '';
    cached = raw
        .split(',')
        .map((s) => s.trim())
        .filter(Boolean);
    return cached;
}

export function getPrincipalCalendars(): string[] {
    return load();
}

export function isPrincipalMode(): boolean {
    return load().length > 0;
}

/** Test hook — clear cache between tests that vary PRINCIPAL_EMAILS. */
export function resetPrincipalCalendarsCache(): void {
    cached = null;
}

/**
 * Substitute `"primary"` with the first principal calendar before any Google
 * API call. Honors the convention "first email in PRINCIPAL_EMAILS is the
 * principal's primary." No-op when principal mode is off (Google resolves
 * `"primary"` against the impersonated identity instead).
 */
export function resolvePrimaryAlias(calendarId: string): string {
    if (!isPrincipalMode()) return calendarId;
    if (calendarId !== PRIMARY_ALIAS) return calendarId;
    const principals = getPrincipalCalendars();
    return principals[0] ?? calendarId;
}

/**
 * Reject undefined / empty `calendarId` for write operations. There's no safe
 * default for an irreversible action like create / update / delete. `"primary"`
 * is allowed — {@link resolvePrimaryAlias} handles substitution.
 */
export function assertWriteCalendarId(calendarId: string | undefined | null): void {
    if (!isPrincipalMode()) return;
    if (!calendarId) {
        const ids = getPrincipalCalendars().join(', ');
        throw new McpError(
            ErrorCode.InvalidRequest,
            `calendarId is required. Use 'primary' for the principal's main calendar, ` +
            `or pass a specific one (${ids}).`,
        );
    }
}
