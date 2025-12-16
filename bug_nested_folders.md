# Bug Report: Nested Folder Paths Not Working for Read/Count/Search Operations

**STATUS: FIXED** (2025-12-16)
**Reported:** 2025-12-16
**Severity:** High - Blocks access to emails in subfolders

## Summary

The `read`, `count`, and `search` operations fail when using nested folder paths like `"Analytics/Daily Sales"`, even though:
1. The `list_folders` operation correctly returns these paths
2. The `move_email` operation correctly handles these paths

## Root Cause

Inconsistent use of folder reference helpers:

| Operation | Helper Used | Nested Folder Support |
|-----------|-------------|----------------------|
| `moveEmail` | `buildNestedFolderRef` | ✓ Works |
| `createFolder` | `buildNestedFolderRef` | ✓ Works |
| `renameFolder` | `buildNestedFolderRef` | ✓ Works |
| `deleteFolder` | `buildNestedFolderRef` | ✓ Works |
| `readEmails` | `buildFolderRef` | ✗ **Broken** |
| `countEmails` | `buildFolderRef` | ✗ **Broken** |
| `searchEmails` | `buildFolderRef` | ✗ **Broken** |

## Technical Details

**`buildFolderRef("Analytics/Daily Sales")`** produces:
```applescript
mail folder "Analytics/Daily Sales"
```
This treats the entire path as a single folder name, which doesn't exist.

**`buildNestedFolderRef("Analytics/Daily Sales")`** produces:
```applescript
mail folder "Daily Sales" of mail folder "Analytics"
```
This correctly navigates the folder hierarchy.

## Reproduction Steps

```javascript
// 1. This works - move email to nested folder
{ operation: "move_email", messageId: "7752", targetFolder: "Analytics/Daily Sales" }
// Result: Success

// 2. This fails - read from the same nested folder
{ operation: "read", folder: "Analytics/Daily Sales", limit: 5 }
// Result: Error: Microsoft Outlook got an error: Can't get mail folder "Analytics/Daily Sales"

// 3. This also fails
{ operation: "count", folder: "Analytics/Daily Sales" }
// Result: Error: Microsoft Outlook got an error: Can't get mail folder "Analytics/Daily Sales"
```

## Fix Required

In `index.ts`, change these three lines:

**Line ~488 (searchEmails):**
```typescript
// FROM:
const folderRef = buildFolderRef(folder);
// TO:
const folderRef = buildNestedFolderRef(folder);
```

**Line ~1816 (countEmails):**
```typescript
// FROM:
const folderRef = buildFolderRef(folder);
// TO:
const folderRef = buildNestedFolderRef(folder);
```

**Line ~1889 (readEmails):**
```typescript
// FROM:
const folderRef = buildFolderRef(folder);
// TO:
const folderRef = buildNestedFolderRef(folder);
```

## Impact

- Cannot read emails from any subfolder (only top-level folders work)
- Cannot count emails in subfolders
- Cannot search within subfolders
- This breaks workflows that organize emails into hierarchical folder structures
