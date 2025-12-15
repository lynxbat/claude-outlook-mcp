# Feature Request: Improved Calendar Query Capabilities

## Summary

The current calendar operations cannot reliably query events more than ~30 days in the future. The `search` operation fails to find events by date, and `upcoming` is limited to a `days` parameter that doesn't extend far enough for scheduling use cases.

---

## Problem Description

### Use Case: Scheduling a Meeting 3+ Weeks Out

User needs to check availability for Wednesday, January 7, 2026 (23 days from today, December 15, 2025).

**Attempted queries:**

```javascript
// Attempt 1: Search by date
{ operation: "search", searchTerm: "January 7 2026" }
// Result: "No events found"

// Attempt 2: Search by date variation
{ operation: "search", searchTerm: "January 7" }
// Result: "No events found"

// Attempt 3: Search by meeting name
{ operation: "search", searchTerm: "Jeremy Nick weekly" }
// Result: "No events found"

// Attempt 4: Upcoming with 30 days
{ operation: "upcoming", days: 30 }
// Result: Returns events through mid-January but misses target date
```

**Visual confirmation:** User's Outlook calendar screenshot clearly shows events on January 7, 2026 (Jeremy/Nick weekly, IT Daily, etc.)

### The Gap

- User can SEE events in Outlook
- MCP cannot QUERY those same events
- Forces user to manually check calendar instead of AI-assisted scheduling

---

## Proposed Solutions

### Option 1: Date Range Query (Recommended)

Add a new operation or parameters for querying a specific date range:

```javascript
// Query specific date
{
  operation: "events_on_date",
  date: "2026-01-07"
}
// Returns: All events on January 7, 2026

// Query date range
{
  operation: "events_in_range",
  startDate: "2026-01-06",
  endDate: "2026-01-10"
}
// Returns: All events from Jan 6-10, 2026
```

### Option 2: Extended Upcoming Range

Increase the `days` limit for `upcoming` operation:

```javascript
{
  operation: "upcoming",
  days: 90  // Allow up to 90 days (current limit appears to be ~30)
}
```

### Option 3: Free/Busy Query

Add operation to check availability for a specific time slot:

```javascript
{
  operation: "check_availability",
  date: "2026-01-07",
  startTime: "13:00",
  endTime: "14:00"
}
// Returns: { available: true } or { available: false, conflict: "Jeremy / Nick - weekly" }
```

### Option 4: Enhanced Search

Make `search` work with date patterns:

```javascript
{
  operation: "search",
  searchTerm: "2026-01-07"  // Search by ISO date
}
// Returns: All events on that date

{
  operation: "search",
  searchTerm: "weekly",
  dateRange: { start: "2026-01-01", end: "2026-01-31" }
}
// Returns: Events matching "weekly" in January 2026
```

---

## AppleScript Implementation Notes

The current implementation likely uses AppleScript to query Outlook. Date-based queries can be done with:

```applescript
tell application "Microsoft Outlook"
  set targetDate to date "Wednesday, January 7, 2026"
  set dayStart to targetDate
  set dayEnd to targetDate + (1 * days)

  set dayEvents to every calendar event whose start time ≥ dayStart and start time < dayEnd

  repeat with evt in dayEvents
    -- return event details
  end repeat
end tell
```

---

## Use Cases This Would Enable

1. **Scheduling assistance** - "Am I free on Jan 7 at 1pm?"
2. **Meeting coordination** - "Find a slot next month for a 30-min call"
3. **Conflict detection** - "Does this proposed meeting conflict with anything?"
4. **Weekly planning** - "What does my week of Jan 6-10 look like?"

---

## Current Workarounds

1. User manually checks Outlook calendar
2. User takes screenshot and shares with AI
3. AI interprets screenshot (error-prone, can't see all details)

None of these are ideal for an AI-assisted workflow.

---

## Priority

**Medium-High** - Calendar availability is a core scheduling need. Current limitations force manual workarounds that break the AI-assisted workflow.

---

*Reported: 2025-12-15*
*Context: Attempting to schedule World of Books call for January 2026*
