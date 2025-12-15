# list_folders Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add a new `list_folders` operation that returns folder metadata with full hierarchical paths.

**Architecture:** New AppleScript function recursively traverses folder hierarchy, building path arrays. Returns JSON objects with path, account, specialFolder, and optional counts.

**Tech Stack:** TypeScript, AppleScript, Bun test runner

---

## Task 1: Add list_folders to Tool Schema

**Files:**
- Modify: `index.ts:28-29` (operation enum)
- Modify: `index.ts:28` (operation description)
- Modify: `index.ts:119-129` (add new parameters)

**Step 1: Update operation enum**

In `index.ts`, find the `operation` property in `OUTLOOK_MAIL_TOOL.inputSchema.properties` and add `list_folders`:

```typescript
operation: {
  type: "string",
  description: "Operation to perform: 'unread', 'search', 'send', 'draft', 'reply', 'forward', 'folders', 'read', 'create_folder', 'move_email', 'rename_folder', 'delete_folder', 'count', 'save_attachments', or 'list_folders'",
  enum: ["unread", "search", "send", "draft", "reply", "forward", "folders", "read", "create_folder", "move_email", "rename_folder", "delete_folder", "count", "save_attachments", "list_folders"]
},
```

**Step 2: Add new parameters for list_folders**

Add these properties after `destinationFolder` (around line 129):

```typescript
includeCounts: {
  type: "boolean",
  description: "Include email count and unread count for each folder (slower, default: false)"
},
excludeDeleted: {
  type: "boolean",
  description: "Exclude folders under Deleted Items (default: true)"
},
account: {
  type: "string",
  description: "Filter to specific account email address (optional, returns all accounts if not specified)"
}
```

**Step 3: Verify changes compile**

Run: `bun run index.ts 2>&1 | head -5`
Expected: Server starts without errors

---

## Task 2: Define FolderInfo Type

**Files:**
- Modify: `index.ts` (add after line 14, before Tool Definitions section)

**Step 1: Add the type definition**

Add this interface after the imports, before the Tool Definitions section:

```typescript
// Folder info type for list_folders operation
interface FolderInfo {
  path: string[];
  account: string;
  specialFolder: string | null;
  count?: number;
  unreadCount?: number;
}
```

**Step 2: Verify changes compile**

Run: `bun run index.ts 2>&1 | head -5`
Expected: Server starts without errors

---

## Task 3: Write Integration Test (Failing)

**Files:**
- Modify: `tests/integration/folders.test.ts`

**Step 1: Add test imports and describe block**

Add after the existing `describe("folder management"` block (end of file):

```typescript
describe("listFolders", () => {
  beforeAll(async () => {
    await ensureOutlookRunning();
  });

  it("returns array of folder objects", async () => {
    const folders = await listFolders();

    expect(Array.isArray(folders)).toBe(true);
    expect(folders.length).toBeGreaterThan(0);

    // Each folder should have required properties
    const folder = folders[0];
    expect(Array.isArray(folder.path)).toBe(true);
    expect(typeof folder.account).toBe("string");
    expect(folder.specialFolder === null || typeof folder.specialFolder === "string").toBe(true);
  }, TEST_TIMEOUT);

  it("includes Inbox with specialFolder set", async () => {
    const folders = await listFolders();

    const inbox = folders.find(f =>
      f.path.length === 1 &&
      f.path[0].toLowerCase() === "inbox"
    );

    expect(inbox).toBeDefined();
    expect(inbox?.specialFolder).toBe("inbox");
  }, TEST_TIMEOUT);

  it("returns nested folder paths as arrays", async () => {
    // Create a nested test folder
    await createFolder("_ListTest");
    await createFolder("_ListTestSub", "_ListTest");

    const folders = await listFolders({ excludeDeleted: false });

    const nestedFolder = folders.find(f =>
      f.path.length === 2 &&
      f.path[0] === "_ListTest" &&
      f.path[1] === "_ListTestSub"
    );

    expect(nestedFolder).toBeDefined();

    // Cleanup
    await deleteFolder("_ListTest/_ListTestSub");
    await deleteFolder("_ListTest");
  }, TEST_TIMEOUT);

  it("excludes deleted folders by default", async () => {
    const folders = await listFolders();

    // Should not have any folder paths starting with Deleted Items
    const deletedFolders = folders.filter(f =>
      f.path[0]?.toLowerCase().includes("deleted")
    );

    // Deleted Items itself should still appear, but not its children
    const deletedChildren = folders.filter(f =>
      f.path.length > 1 &&
      f.path[0]?.toLowerCase().includes("deleted")
    );

    expect(deletedChildren.length).toBe(0);
  }, TEST_TIMEOUT);

  it("includes deleted folders when excludeDeleted is false", async () => {
    const folders = await listFolders({ excludeDeleted: false });

    // This just verifies the parameter is accepted
    expect(Array.isArray(folders)).toBe(true);
  }, TEST_TIMEOUT);

  it("includes counts when includeCounts is true", async () => {
    const folders = await listFolders({ includeCounts: true });

    expect(folders.length).toBeGreaterThan(0);

    const inbox = folders.find(f =>
      f.path.length === 1 &&
      f.path[0].toLowerCase() === "inbox"
    );

    expect(inbox).toBeDefined();
    expect(typeof inbox?.count).toBe("number");
    expect(typeof inbox?.unreadCount).toBe("number");
  }, TEST_TIMEOUT);
});
```

**Step 2: Update import to include listFolders**

Update the import at top of file:

```typescript
import { getMailFolders, createFolder, renameFolder, deleteFolder, listFolders } from "../../index";
```

**Step 3: Run test to verify it fails**

Run: `bun test tests/integration/folders.test.ts -t "listFolders"`
Expected: FAIL - `listFolders` is not exported

---

## Task 4: Implement listFolders Function

**Files:**
- Modify: `index.ts` (add after `getMailFolders` function, around line 906)

**Step 1: Add the listFolders function**

Add after `getMailFolders` function:

```typescript
// Function to list folders with full paths and metadata
async function listFolders(options: {
  includeCounts?: boolean;
  excludeDeleted?: boolean;
  account?: string;
} = {}): Promise<FolderInfo[]> {
  const { includeCounts = false, excludeDeleted = true, account } = options;
  console.error(`[listFolders] Getting folders with options: includeCounts=${includeCounts}, excludeDeleted=${excludeDeleted}, account=${account || 'all'}`);
  await checkOutlookAccess();

  // Build the counts script portion
  const countsScript = includeCounts ? `
            set folderCount to count of messages of theFolder
            set unreadCount to 0
            repeat with msg in messages of theFolder
              if is read of msg is false then
                set unreadCount to unreadCount + 1
              end if
            end repeat
            set countInfo to "," & folderCount & "," & unreadCount` : `
            set countInfo to ""`;

  const script = `
    tell application "Microsoft Outlook"
      set folderList to {}
      set accountFilter to "${account || ""}"
      set excludeDeletedItems to ${excludeDeleted}

      repeat with theAccount in exchange accounts
        set accountEmail to email address of theAccount

        -- Skip if filtering by account and this isn't it
        if accountFilter is not "" and accountEmail is not accountFilter then
          -- skip this account
        else
          -- Process each top-level folder
          repeat with theFolder in mail folders of theAccount
            my processFolder(theFolder, {}, accountEmail, excludeDeletedItems, folderList)
          end repeat
        end if
      end repeat

      return folderList
    end tell

    on processFolder(theFolder, parentPath, accountEmail, excludeDeletedItems, folderList)
      tell application "Microsoft Outlook"
        set folderName to name of theFolder
        set currentPath to parentPath & {folderName}

        -- Check if this is Deleted Items and we should skip children
        set isDeletedItems to folderName is "Deleted Items" or folderName is "Trash"

        -- Determine special folder type
        set specialType to "null"
        if folderName is "Inbox" then
          set specialType to "inbox"
        else if folderName is "Sent Items" or folderName is "Sent" then
          set specialType to "sent"
        else if folderName is "Drafts" then
          set specialType to "drafts"
        else if folderName is "Deleted Items" or folderName is "Trash" then
          set specialType to "trash"
        else if folderName is "Junk Email" or folderName is "Junk" then
          set specialType to "junk"
        else if folderName is "Archive" then
          set specialType to "archive"
        end if

        -- Get counts if requested
        ${countsScript}

        -- Build path string as JSON array
        set pathJSON to "["
        repeat with i from 1 to count of currentPath
          if i > 1 then set pathJSON to pathJSON & ","
          set pathJSON to pathJSON & "\\"" & item i of currentPath & "\\""
        end repeat
        set pathJSON to pathJSON & "]"

        -- Add folder info as JSON-like string
        set folderInfo to pathJSON & "|" & accountEmail & "|" & specialType & countInfo
        set end of folderList to folderInfo

        -- Process subfolders unless this is Deleted Items and we're excluding
        if not (isDeletedItems and excludeDeletedItems) then
          repeat with subFolder in mail folders of theFolder
            my processFolder(subFolder, currentPath, accountEmail, excludeDeletedItems, folderList)
          end repeat
        end if
      end tell
    end processFolder
  `;

  try {
    const result = await runAppleScript(script);
    console.error(`[listFolders] Raw result length: ${result.length}`);

    // Parse the result - each folder is separated by ", "
    const folderStrings = result.split(", ");
    const folders: FolderInfo[] = [];

    for (const folderStr of folderStrings) {
      if (!folderStr.trim()) continue;

      // Parse: pathJSON|account|specialType[,count,unreadCount]
      const parts = folderStr.split("|");
      if (parts.length < 3) continue;

      try {
        const path = JSON.parse(parts[0]);
        const accountEmail = parts[1];
        const specialPart = parts[2];

        // Parse special folder and optional counts
        const specialParts = specialPart.split(",");
        const specialFolder = specialParts[0] === "null" ? null : specialParts[0];

        const folderInfo: FolderInfo = {
          path,
          account: accountEmail,
          specialFolder
        };

        if (includeCounts && specialParts.length >= 3) {
          folderInfo.count = parseInt(specialParts[1], 10) || 0;
          folderInfo.unreadCount = parseInt(specialParts[2], 10) || 0;
        }

        folders.push(folderInfo);
      } catch (parseError) {
        console.error(`[listFolders] Failed to parse folder: ${folderStr}`, parseError);
      }
    }

    console.error(`[listFolders] Parsed ${folders.length} folders`);
    return folders;
  } catch (error) {
    console.error("[listFolders] Error:", error);
    throw error;
  }
}
```

**Step 2: Run tests to check progress**

Run: `bun test tests/integration/folders.test.ts -t "returns array"`
Expected: Still fails - function not exported yet

---

## Task 5: Add Case Handler for list_folders

**Files:**
- Modify: `index.ts` (add case in switch statement, around line 2680)

**Step 1: Add the case handler**

Add after `case "save_attachments"` block (around line 2679):

```typescript
          case "list_folders": {
            const folders = await listFolders({
              includeCounts: args.includeCounts,
              excludeDeleted: args.excludeDeleted !== false, // default true
              account: args.account
            });

            // Format for display
            const formatPath = (path: string[]) => path.join("/");
            const lines = folders.map(f => {
              let line = formatPath(f.path);
              line += ` (${f.account})`;
              if (f.specialFolder) line += ` [${f.specialFolder}]`;
              if (f.count !== undefined) {
                line += ` - ${f.count} emails`;
                if (f.unreadCount && f.unreadCount > 0) {
                  line += `, ${f.unreadCount} unread`;
                }
              }
              return line;
            });

            return {
              content: [{
                type: "text",
                text: folders.length > 0 ?
                  `Found ${folders.length} folders:\n\n${lines.join("\n")}` :
                  "No folders found."
              }],
              isError: false
            };
          }
```

**Step 2: Verify it compiles**

Run: `bun run index.ts 2>&1 | head -5`
Expected: Server starts

---

## Task 6: Update Type Guard

**Files:**
- Modify: `index.ts:2282-2283` (operation type)
- Modify: `index.ts:2313` (operation array)

**Step 1: Update operation type union**

In `isMailArgs`, update the operation type (line 2283):

```typescript
  operation: "unread" | "search" | "send" | "draft" | "reply" | "forward" | "folders" | "read" | "create_folder" | "move_email" | "rename_folder" | "delete_folder" | "count" | "save_attachments" | "list_folders";
```

**Step 2: Add new parameters to type**

Add after `destinationFolder?: string;` (around line 2300):

```typescript
  includeCounts?: boolean;
  excludeDeleted?: boolean;
```

**Step 3: Update operation array in validation**

Update the includes array (line 2313):

```typescript
  if (!operation || !["unread", "search", "send", "draft", "reply", "forward", "folders", "read", "create_folder", "move_email", "rename_folder", "delete_folder", "count", "save_attachments", "list_folders"].includes(operation)) {
```

**Step 4: Verify compilation**

Run: `bun run index.ts 2>&1 | head -5`
Expected: Server starts

---

## Task 7: Export listFolders Function

**Files:**
- Modify: `index.ts:2893-2911` (exports section)

**Step 1: Add listFolders to exports**

Find the export block (around line 2893) and add `listFolders`:

```typescript
export {
  readEmails,
  searchEmails,
  getUnreadEmails,
  sendEmail,
  createDraft,
  replyEmail,
  forwardEmail,
  getMailFolders,
  listFolders,
  createFolder,
  moveEmail,
  renameFolder,
  deleteFolder,
  getTodayEvents,
  getUpcomingEvents,
  searchEvents,
  createEvent,
  listContacts,
  searchContacts
};
```

**Step 2: Run integration tests**

Run: `bun test tests/integration/folders.test.ts -t "listFolders"`
Expected: Tests should now run (may fail on AppleScript specifics)

---

## Task 8: Debug and Fix AppleScript

**Files:**
- Modify: `index.ts` (listFolders function)

**Step 1: Run tests and observe failures**

Run: `bun test tests/integration/folders.test.ts -t "returns array" --timeout 60000`
Observe the actual error output.

**Step 2: Fix AppleScript issues**

Common issues to watch for:
- AppleScript handler scope issues (may need to inline instead of using `on processFolder`)
- String escaping in folder names
- Account type (may need to check `accounts` not just `exchange accounts`)

If the recursive handler doesn't work, use an iterative approach with a queue.

**Step 3: Re-run tests until passing**

Run: `bun test tests/integration/folders.test.ts -t "listFolders" --timeout 60000`
Expected: All listFolders tests pass

---

## Task 9: Commit Implementation

**Step 1: Check what changed**

Run: `git status && git diff --stat`

**Step 2: Stage and commit**

```bash
git add index.ts tests/integration/folders.test.ts
git commit -m "feat: add list_folders operation with hierarchical paths

- New operation returns folder metadata with full path arrays
- Includes account, specialFolder type detection
- Optional includeCounts parameter for email/unread counts
- excludeDeleted parameter (default true) hides Deleted Items children
- account parameter filters to specific email account"
```

---

## Task 10: Update README Documentation

**Files:**
- Modify: `README.md`

**Step 1: Add list_folders to operations table**

Find the mail operations table and add:

```markdown
| `list_folders` | List all folders with full paths and metadata | `includeCounts`, `excludeDeleted`, `account` |
```

**Step 2: Add detailed documentation**

Add a section explaining the new operation with example output.

**Step 3: Commit docs**

```bash
git add README.md
git commit -m "docs: add list_folders operation to README"
```

---

## Summary

| Task | Description |
|------|-------------|
| 1 | Add list_folders to tool schema |
| 2 | Define FolderInfo type |
| 3 | Write failing integration tests |
| 4 | Implement listFolders function |
| 5 | Add case handler |
| 6 | Update type guard |
| 7 | Export function |
| 8 | Debug and fix AppleScript |
| 9 | Commit implementation |
| 10 | Update README |
