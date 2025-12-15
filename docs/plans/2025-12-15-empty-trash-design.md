# Design: empty_trash Operation

## Overview

Two-phase operation to permanently empty the Deleted Items folder with preview before execution.

## API

### Phase 1: Preview

```javascript
{ operation: "empty_trash", preview: true }

// Response
{
  "preview": true,
  "itemCount": 847,
  "oldestItem": "2024-01-15",
  "newestItem": "2025-12-14",
  "totalSizeMB": 156.4
}
```

### Phase 2: Execute

```javascript
{ operation: "empty_trash", confirm: true }

// Response
{
  "deleted": 847,
  "message": "Permanently deleted 847 items from Deleted Items"
}
```

## Validation Rules

- Must provide either `preview: true` OR `confirm: true`
- If neither provided → error with usage hint
- If both provided → error (ambiguous intent)
- If `confirm: true` with 0 items → success with "Deleted Items folder is already empty"

## AppleScript Implementation

### Preview
```applescript
tell application "Microsoft Outlook"
  set deletedFolder to deleted items
  set msgs to messages of deletedFolder
  set itemCount to count of msgs
  -- Iterate to get dates and sizes
end tell
```

### Execute
```applescript
tell application "Microsoft Outlook"
  set deletedFolder to deleted items
  repeat with msg in (messages of deletedFolder)
    permanently delete msg
  end repeat
end tell
```

## Error Handling

| Scenario | Response |
|----------|----------|
| No `preview` or `confirm` param | Error: "empty_trash requires either preview: true or confirm: true" |
| Both `preview` and `confirm` | Error: "Cannot use both preview and confirm - use one at a time" |
| Outlook not running | Error: "Microsoft Outlook is not running" |
| No Deleted Items access | Error: "Could not access Deleted Items folder" |
| Deletion fails mid-way | Return partial count: "Deleted 423 of 847 items before error: {reason}" |
| Empty folder on confirm | Success: "Deleted Items folder is already empty" (deleted: 0) |

## Notes

- Outlook's `permanently delete` command bypasses Deleted Items - items are truly gone
- No retry logic - if deletion fails, user can run again to clean up remaining items
- Large folders may take several seconds for preview/execute
