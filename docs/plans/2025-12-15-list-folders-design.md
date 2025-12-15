# Design: `list_folders` Operation

## Overview

New `list_folders` operation returns folder metadata with full hierarchical paths. The existing `folders` operation remains unchanged for backward compatibility.

## Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `includeCounts` | boolean | `false` | Include `count` and `unreadCount` fields (slower) |
| `excludeDeleted` | boolean | `true` | Hide folders under Deleted Items |
| `account` | string | `null` | Filter to specific account email address |

## Response Format

Each folder returns an object:

```json
{
  "path": ["People", "Julia Ward"],
  "account": "nick@work.com",
  "specialFolder": null
}
```

With `includeCounts: true`:

```json
{
  "path": ["People", "Julia Ward"],
  "account": "nick@work.com",
  "specialFolder": null,
  "count": 127,
  "unreadCount": 0
}
```

**`specialFolder` values:** `inbox`, `sent`, `drafts`, `trash`, `junk`, `archive`, or `null` for custom folders.

**Full response example:**

```json
[
  {"path": ["Inbox"], "account": "nick@work.com", "specialFolder": "inbox"},
  {"path": ["Deleted Items"], "account": "nick@work.com", "specialFolder": "trash"},
  {"path": ["People"], "account": "nick@work.com", "specialFolder": null},
  {"path": ["People", "Julia Ward"], "account": "nick@work.com", "specialFolder": null},
  {"path": ["Supply Chain", "OMS", "Infios"], "account": "nick@work.com", "specialFolder": null}
]
```

## Implementation Approach

### AppleScript Strategy

1. Get all accounts, iterate each account's `mail folders`
2. For each folder, recursively traverse `mail folders` (subfolders)
3. Build path array as we recurse
4. Check folder properties for special folder detection
5. Optionally count messages if `includeCounts` requested

### Key AppleScript Properties

- `mail folders of account` - Top-level folders
- `mail folders of folder` - Subfolders (recursive)
- `name of folder` - Folder name
- `count of messages` - Email count
- `unread count` - Unread count

### Filtering Logic

- `excludeDeleted`: Skip recursion when hitting "Deleted Items" folder
- `account`: Only iterate matching account

## Display Formatting

When the MCP returns results, the output text joins paths with `/` for readability:

```
Found 24 folders:

Inbox (nick@work.com) [inbox] - 38 emails, 5 unread
Sent Items (nick@work.com) [sent]
People (nick@work.com)
People/Julia Ward (nick@work.com) - 127 emails
People/Greg Kellerman (nick@work.com) - 43 emails
Supply Chain/OMS/Infios (nick@work.com) - 12 emails
```

Format: `path (account) [specialFolder] - counts if requested`

The raw JSON stays as arrays for programmatic use, but the human-readable output joins with `/`.

## Decisions Summary

| Decision | Choice |
|----------|--------|
| Operation name | `list_folders` |
| Backward compat | New operation, keep `folders` unchanged |
| Path format | Array of strings `["Parent", "Child"]` |
| Counts | Opt-in via `includeCounts: true` |
| Deleted folders | Hidden by default, `excludeDeleted: false` to show |
| Account filter | Return all by default, optional `account` param |
