# Folder Management Design

**Date:** 2025-12-04
**Status:** Approved

## Overview

Add full folder management to the Outlook MCP plugin: create, move, rename, and delete folders. Extends the existing `outlook_mail` tool.

## Decisions

| Decision | Choice |
|----------|--------|
| Operations | Create, move, rename, delete (full management) |
| Tool approach | Extend existing `outlook_mail` tool |
| Email identification | Message ID (unique, unambiguous) |
| Nested folders | Supported via path syntax (e.g., "Programs/Bloomreach") |
| Testing | Dedicated `_MCP_Test` folder, cleaned up after tests |

## New Operations

| Operation | Parameters | Description |
|-----------|------------|-------------|
| `create_folder` | `name`, `parent` (optional) | Create folder, optionally nested under parent |
| `move_email` | `messageId`, `targetFolder` | Move email by ID to target folder |
| `rename_folder` | `folder`, `newName` | Rename existing folder |
| `delete_folder` | `folder` | Delete folder (must be empty) |

## AppleScript Implementation

### Create folder
```applescript
tell application "Microsoft Outlook"
  set newFolder to make new mail folder with properties {name:"Bloomreach"}
  -- For nested: make new mail folder at mail folder "Programs" with properties {name:"Bloomreach"}
end tell
```

### Move email
```applescript
tell application "Microsoft Outlook"
  set theMsg to message id 12345
  move theMsg to mail folder "Bloomreach"
end tell
```

### Rename folder
```applescript
tell application "Microsoft Outlook"
  set name of mail folder "OldName" to "NewName"
end tell
```

### Delete folder
```applescript
tell application "Microsoft Outlook"
  delete mail folder "FolderName"
end tell
```

## Schema Changes

### Updated operation enum
```typescript
enum: ["unread", "search", "send", "folders", "read",
       "create_folder", "move_email", "rename_folder", "delete_folder"]
```

### New parameters
```typescript
name: { type: "string", description: "Folder name to create" }
parent: { type: "string", description: "Parent folder path (optional, for nesting)" }
messageId: { type: "string", description: "Email message ID to move" }
targetFolder: { type: "string", description: "Destination folder path" }
newName: { type: "string", description: "New folder name" }
```

### Updated read/search output
Add `messageId` to email output:
```
--- Email 1 ---
ID: 12345
From: sender@example.com
Date: Thursday, December 4, 2025
Subject: Example Subject
```

## Helper Function Updates

Extend `buildFolderRef` to handle nested paths:
- Input: `"Programs/Bloomreach"`
- Split on `/` and traverse: `mail folder "Bloomreach" of mail folder "Programs"`

Add `buildNestedFolderRef(path: string): string` to helpers.ts.

## Testing Strategy

### Unit tests (tests/unit/helpers.test.ts)
- `buildNestedFolderRef("Programs/Bloomreach")` returns correct AppleScript path
- `buildNestedFolderRef("Inbox")` returns `inbox` (special case)
- `parseEmailOutputWithId()` extracts message ID from output

### Integration tests (tests/integration/folders.test.ts)
```typescript
describe("folder management", () => {
  const TEST_FOLDER = "_MCP_Test";

  beforeAll(() => createFolder(TEST_FOLDER));
  afterAll(() => deleteFolder(TEST_FOLDER));

  it("creates folder");
  it("creates nested folder");
  it("renames folder");
  it("moves email to folder");
  it("deletes empty folder");
  it("fails to delete non-empty folder");
});
```

## Error Handling

| Operation | Error Condition | Response |
|-----------|----------------|----------|
| `create_folder` | Folder already exists | Return error message, don't throw |
| `create_folder` | Parent folder not found | Return error with suggestion |
| `move_email` | Message ID not found | Return error message |
| `move_email` | Target folder not found | Return error message |
| `rename_folder` | Folder not found | Return error message |
| `rename_folder` | New name already exists | Return error message |
| `delete_folder` | Folder not found | Return error message |
| `delete_folder` | Folder not empty | Return error, list email count |

All errors return `isError: false` with descriptive text (not exceptions), allowing conversation to continue gracefully.

## Implementation Order

1. Update helpers.ts with `buildNestedFolderRef`
2. Update read/search to include message ID in output
3. Add `createFolder` function
4. Add `moveEmail` function
5. Add `renameFolder` function
6. Add `deleteFolder` function
7. Update tool schema with new operations
8. Add MCP handler cases
9. Export new functions
10. Add unit tests
11. Add integration tests
