import { readFileSync, existsSync } from 'fs';

/**
 * Configurable event filter rule.
 *
 * Filters are applied post-fetch against in-memory event objects.
 * Two matching modes:
 *   - `equals` -- static value match on the property at `where`
 *   - `inValuesOf` -- cross-event dedup: exclude/include events whose `where`
 *     property value appears in the set of values collected from `inValuesOf`
 *     across ALL events in the batch (Phase 1 pre-scan).
 *
 * `calendarIds` scopes a filter to specific calendars; events from other
 * calendars pass through unaffected.
 */
export interface EventFilter {
  action: 'exclude' | 'includeOnly';
  /** Dot-notation path to the event property to test (e.g. "id", "extendedProperties.private.workHoldId"). */
  where: string;
  /** Static value to compare against. Mutually exclusive with `inValuesOf`. */
  equals?: string;
  /** Dot-notation path whose values across all events form the match set. Mutually exclusive with `equals`. */
  inValuesOf?: string;
  /** Restrict this filter to events from these calendar IDs. Omit to apply globally. */
  calendarIds?: string[];
}

/** Result of applying filters to an event list. */
export interface FilterResult<T> {
  events: T[];
  eventsFiltered: number;
}

/**
 * Walk a dot-notation property path on an arbitrary object.
 * Returns `undefined` if any segment is missing or the base is not an object.
 *
 * Example: getProp(event, "extendedProperties.private.workHoldId")
 */
export function getProp(obj: unknown, path: string): unknown {
  let current: unknown = obj;
  for (const key of path.split('.')) {
    if (current == null || typeof current !== 'object') return undefined;
    current = (current as Record<string, unknown>)[key];
  }
  return current;
}

/**
 * Two-phase event filter.
 *
 * Phase 1 (pre-scan): For cross-event filters (`inValuesOf` without `equals`),
 * scans all events to build the set of values that the `inValuesOf` path
 * resolves to. This enables dedup-style filters such as "exclude events whose
 * `id` appears in the set of `extendedProperties.private.workHoldId` values".
 *
 * Phase 2 (per-event): Evaluates every filter as a predicate on each event.
 * `exclude` + match => drop. `includeOnly` + no match => drop. First
 * disqualifying filter short-circuits.
 */
export function applyEventFilters<T>(events: T[], filters: EventFilter[] | undefined): FilterResult<T> {
  if (!filters || filters.length === 0) return { events, eventsFiltered: 0 };

  // Phase 1: pre-scan for cross-event filters -- build match sets
  const matchSets = new Map<number, Set<string>>();
  for (let i = 0; i < filters.length; i++) {
    const filter = filters[i];
    if (!filter.equals && filter.inValuesOf) {
      const values = new Set<string>();
      for (const event of events) {
        const val = getProp(event, filter.inValuesOf);
        if (typeof val === 'string' && val.length > 0) values.add(val);
      }
      matchSets.set(i, values);
    }
  }

  // Phase 2: apply all filters as per-event predicates
  let eventsFiltered = 0;
  const result: T[] = [];

  for (const event of events) {
    let keep = true;

    for (let i = 0; i < filters.length; i++) {
      const filter = filters[i];
      let matches: boolean;

      // Skip this filter for events outside its calendar scope
      if (filter.calendarIds) {
        const eventCalId = getProp(event, 'calendarId');
        if (typeof eventCalId !== 'string' || !filter.calendarIds.includes(eventCalId)) continue;
      }

      if (filter.equals !== undefined) {
        matches = getProp(event, filter.where) === filter.equals;
      } else if (filter.inValuesOf) {
        const propValue = getProp(event, filter.where);
        matches = typeof propValue === 'string' && (matchSets.get(i)?.has(propValue) ?? false);
      } else {
        continue;
      }

      if (filter.action === 'exclude' && matches) { keep = false; break; }
      if (filter.action === 'includeOnly' && !matches) { keep = false; break; }
    }

    if (keep) result.push(event);
    else eventsFiltered++;
  }

  return { events: result, eventsFiltered };
}

/**
 * Expected shape of the JSON config file pointed to by EVENT_FILTER_CONFIG.
 */
interface EventFilterConfig {
  eventFilters?: EventFilter[];
}

/**
 * Load event filters from a JSON config file.
 *
 * Returns an empty array (not undefined) when the file is missing, empty,
 * or malformed -- the server should start gracefully regardless of filter
 * config issues. Parse errors are logged to stderr.
 */
export function loadEventFilters(configPath: string | undefined): EventFilter[] {
  if (!configPath) return [];
  if (!existsSync(configPath)) {
    process.stderr.write(`event-filter: config file not found: ${configPath}\n`);
    return [];
  }

  try {
    const raw = JSON.parse(readFileSync(configPath, 'utf-8')) as Partial<EventFilterConfig>;
    if (!Array.isArray(raw.eventFilters)) return [];

    // Basic structural validation -- reject obviously malformed entries
    const valid: EventFilter[] = [];
    for (const entry of raw.eventFilters) {
      if (
        typeof entry !== 'object' || entry == null ||
        (entry.action !== 'exclude' && entry.action !== 'includeOnly') ||
        typeof entry.where !== 'string' || entry.where.length === 0
      ) {
        process.stderr.write(`event-filter: skipping malformed filter entry: ${JSON.stringify(entry)}\n`);
        continue;
      }
      if (entry.equals === undefined && !entry.inValuesOf) {
        process.stderr.write(`event-filter: skipping filter without equals or inValuesOf: ${JSON.stringify(entry)}\n`);
        continue;
      }

      valid.push({
        action: entry.action,
        where: entry.where,
        ...(entry.equals !== undefined && { equals: entry.equals }),
        ...(entry.inValuesOf && { inValuesOf: entry.inValuesOf }),
        ...(Array.isArray(entry.calendarIds) && { calendarIds: entry.calendarIds }),
      });
    }

    if (valid.length > 0) {
      process.stderr.write(`event-filter: loaded ${valid.length} filter(s) from ${configPath}\n`);
    }
    return valid;
  } catch (err) {
    process.stderr.write(`event-filter: failed to parse config: ${err}\n`);
    return [];
  }
}
