# Feature Request: Hierarchical Folder Paths in Folder List

**STATUS: IMPLEMENTED** - The `list_folders` operation now provides hierarchical paths. See usage below.

## Usage

```javascript
// Request
{ operation: "list_folders", excludeDeleted: true }

// Response format
"Supply Chain/WMS (account@email.com)"
"Ecommerce/Platforms/SFCC/Euvic (account@email.com)"
"Inbox (account@email.com) [inbox]"  // special folders marked
```

**Parameters:**
- `excludeDeleted` (default: true) - Hides folders under Deleted Items
- `includeCounts` (default: false) - Include email/unread counts per folder
- `account` - Filter to specific account

---

## Original Problem Statement (Resolved)

The original `folders` operation returns a flat list of folder names without hierarchy information. This created several issues:

1. **Cannot determine folder location** - No way to know if a folder is nested under another folder
2. **Deleted folders appear active** - Folders in "Deleted Items" appear alongside active folders
3. **Duplicate names are ambiguous** - Multiple folders with the same name (e.g., "NRF", "Shopify", "Analytics") cannot be distinguished
4. **Cannot determine folder paths for operations** - When moving emails or creating subfolders, the correct path is unclear

## Current Behavior

```javascript
// Request
{ operation: "folders" }

// Response (flat list)
[
  "Inbox",
  "Deleted Items",
  "__Priority",
  "People",
  "Julia Ward",
  "NRF",
  "Supply Chain",
  "WMS",
  "NRF",  // duplicate - which one?
  ...
]
```

## Proposed Behavior

Return full paths showing folder hierarchy:

```javascript
// Request
{ operation: "folders" }

// Response (hierarchical paths)
[
  "Inbox",
  "Deleted Items",
  "Deleted Items/__Priority",
  "People",
  "People/Julia Ward",
  "People/PJD",
  "Conferences/NRF",
  "Supply Chain",
  "Supply Chain/WMS",
  "Supply Chain/OMS/Infios",
  "Ecommerce/Platforms/NRF",  // now distinguishable from Conferences/NRF
  ...
]
```

## Requirements

### Core Requirements

1. **Full path format**: Return folder paths using `/` as delimiter (e.g., `Parent/Child/Grandchild`)
2. **Include all levels**: Show the complete path from root to leaf for every folder
3. **Consistent ordering**: Parent folders should appear before their children
4. **Handle special characters**: Escape or handle folder names containing `/` if they exist

### Optional Enhancements (Lower Priority)

1. **Email count per folder**: Include count of emails in each folder
   ```javascript
   { path: "Inbox", count: 38 }
   ```

2. **Filter parameter**: `excludeDeleted` (default: true) to hide folders under Deleted Items
   ```javascript
   { operation: "folders", excludeDeleted: true }
   ```

3. **Folder type indicator**: Distinguish local folders ("On My Computer") from Exchange/server folders
   ```javascript
   { path: "Inbox", type: "exchange" }
   { path: "On My Computer/Local Archive", type: "local" }
   ```

## Test Cases

### Test Case 1: Basic Hierarchy
**Setup**: Create folder structure:
```
People/
  Julia Ward/
  Greg Kellerman/
```

**Expected Output** should include:
```
"People"
"People/Julia Ward"
"People/Greg Kellerman"
```

### Test Case 2: Deleted Folder Visibility
**Setup**: Delete a folder named `__Priority`

**Expected Output** should include:
```
"Deleted Items/__Priority"
```
NOT:
```
"__Priority"  // misleading - appears active
```

### Test Case 3: Duplicate Folder Names
**Setup**: Create folders:
```
Conferences/NRF/
Ecommerce/Platforms/NRF/
```

**Expected Output** should include:
```
"Conferences/NRF"
"Ecommerce/Platforms/NRF"
```
NOT:
```
"NRF"
"NRF"  // ambiguous
```

### Test Case 4: Deep Nesting
**Setup**: Create folder structure:
```
Supply Chain/
  OMS/
    Infios/
```

**Expected Output** should include:
```
"Supply Chain"
"Supply Chain/OMS"
"Supply Chain/OMS/Infios"
```

### Test Case 5: Special Characters in Folder Names
**Setup**: Create folder named `Q4 2025 / Planning`

**Expected Output**: Should handle gracefully (escape, encode, or use alternate delimiter)

## Implementation Notes

### AppleScript Considerations

The current implementation likely uses AppleScript to query Outlook. The folder hierarchy can be obtained by:

1. Recursively traversing `mail folders` of each account
2. Building the path string as you traverse
3. For each folder, prepending its parent's path

### Suggested AppleScript Approach

```applescript
on getFolderPath(theFolder)
    set thePath to name of theFolder
    set parentFolder to container of theFolder
    repeat while class of parentFolder is mail folder
        set thePath to (name of parentFolder) & "/" & thePath
        set parentFolder to container of parentFolder
    end repeat
    return thePath
end getFolderPath
```

## Impact

This change would:
- **Eliminate confusion** about folder locations
- **Enable accurate folder operations** (move, create subfolder)
- **Surface deleted folders clearly** so users know to clean them up
- **Disambiguate duplicate names** across different branches

## Backward Compatibility

This is a breaking change to the output format. Options:
1. **New operation**: Add `folders_tree` or `folders_hierarchical` operation
2. **Parameter**: Add `hierarchical: true` parameter (default false for backward compat)
3. **Breaking change**: Update `folders` output format (recommend if MCP is not widely deployed)

---

# Bug Report: Email Reply/Forward Operations

**STATUS: ALL BUGS FIXED**

## Summary

~~Multiple issues with the `reply`, `draft`, and `forward` operations make it impossible to reliably create email replies with modified recipients.~~

## Use Case

User needs to reply to an email thread but change the recipients:
- Move someone from CC to TO line (to direct the action to them)
- Keep original sender and others on CC
- Maintain thread continuity (subject line, In-Reply-To headers, conversation threading)

## Bug 1: `reply` Operation Cannot Modify Recipients

**FIXED:** Added 9 new parameters: `replyTo`, `replyCc`, `replyBcc` (override), `addTo`, `addCc`, `addBcc` (append), `removeTo`, `removeCc`, `removeBcc` (remove).

### Problem
The `reply` operation does not accept `to` or `cc` parameters. It can only reply to the original sender with optional `replyAll: true`.

### Actual Behavior
```javascript
// Request
{
  operation: "reply",
  messageId: "...",
  replyBody: "...",
  to: "leah@example.com",  // IGNORED
  cc: "paul@example.com"   // IGNORED
}

// Result: Replies only to original sender, ignores to/cc parameters
```

### Expected Behavior
```javascript
// Request
{
  operation: "reply",
  messageId: "...",
  replyBody: "...",
  to: "leah@example.com",           // Override TO recipients
  cc: "paul@example.com, steve@...", // Override CC recipients
  replyAll: false                    // Don't auto-include original recipients
}

// Result: Creates reply in thread with specified recipients
```

### Impact
Cannot use `reply` for common workflow: "Reply but redirect the action to a different person"

---

## Bug 2: `draft` Operation Creates New Email, Not Reply

**FIXED:** Added `replyToMessageId` parameter to `draft` operation. Creates threaded reply draft when provided.

### Problem
When trying to work around Bug 1 by using `draft`, it creates a completely new email rather than a reply in the thread.

### Actual Behavior
```javascript
// Request (attempting to draft a reply)
{
  operation: "draft",
  to: "leah@example.com",
  cc: "paul@example.com",
  subject: "Re: Original Subject",
  body: "..."
}

// Result: Opens new email composition window
// - Subject has "Re:" but is NOT a reply
// - No thread association
// - No In-Reply-To header
// - Will appear as new conversation in recipients' inboxes
```

### Expected Behavior
A new operation like `draft_reply` or parameter `replyToMessageId`:
```javascript
{
  operation: "draft",
  replyToMessageId: "original-message-id",
  to: "leah@example.com",
  cc: "paul@example.com",
  body: "..."
}

// Result: Opens draft that IS a reply to the thread
```

---

## Bug 3: `forward` Operation Breaks Threading

**RESOLVED:** Bug 1 fix eliminates need for forward workaround. Also added `forwardBcc` parameter for complete recipient control.

### Problem
Using `forward` as a workaround creates "FW:" prefix instead of "Re:", breaking thread continuity.

### Actual Behavior
```javascript
// Request
{
  operation: "forward",
  messageId: "...",
  forwardTo: "leah@example.com",
  forwardCc: "paul@example.com, steve@example.com",
  forwardComment: "Leah, can you find a time..."
}

// Result:
// - Subject: "FW: Original Subject" (not "Re:")
// - Breaks thread grouping in recipients' mail clients
// - Appears as forward, not reply
```

### Impact
- Recipients see "FW:" and think it's being forwarded outside the group
- Thread is broken - replies to the forward won't group with original thread
- Confusing for all participants

---

## Bug 4: Multiple CC Addresses Malformed

**FIXED:** Extracted `parseRecipients()` helper to properly parse comma-separated addresses into individual recipients. Applied to TO, CC, and BCC fields in all operations.

### Problem
When specifying multiple CC addresses, they are wrapped in a single set of angle brackets as one entity, rather than being parsed as separate recipients.

### Actual Behavior
Screenshot evidence shows CC field displayed as:
```
<paul@example.com, steve@example.com, jessica@example.com>
```

Instead of:
```
paul@example.com; steve@example.com; jessica@example.com
```

### Impact
- Outlook may fail to parse the malformed recipient list
- Email may fail to send
- Recipients may not receive the email

### Root Cause (Suspected)
The code likely wraps the entire CC string in angle brackets:
```javascript
// Buggy
cc: `<${ccAddresses}>`  // "<a@b.com, c@d.com>"

// Should be
cc: ccAddresses.split(',').map(e => e.trim()).join('; ')
```

---

## Bug 5: False Success Messages

**FIXED:** Changed "sent successfully" to "queued for delivery" for send/reply/forward operations.

### Problem
MCP operations return success messages that don't reflect actual Outlook state.

### Actual Behavior
```javascript
// Request
{ operation: "forward", ... }

// Response
{ success: true, message: "Email forwarded successfully" }

// Actual Result: Draft window opened, email NOT sent
```

### Evidence
User screenshot showed:
1. MCP reported "forwarded successfully"
2. Outlook showed draft window open (not sent)
3. Email was never delivered

### Expected Behavior
- If operation opens a draft: Return `{ success: true, status: "draft_created" }`
- If operation sends: Return `{ success: true, status: "sent" }`
- If operation fails: Return `{ success: false, error: "..." }`

Alternatively, add a `sendImmediately: boolean` parameter to control behavior.

---

## Proposed Solutions

### Solution 1: Enhanced `reply` Operation

Add optional `to` and `cc` parameters to `reply`:

```javascript
{
  operation: "reply",
  messageId: "...",
  replyBody: "...",

  // New optional parameters:
  to: "override@example.com",        // Override TO (default: reply to sender)
  cc: "person1@..., person2@...",    // Override CC
  preserveOriginalRecipients: false  // If true, merge with original recipients
}
```

### Solution 2: New `reply_with_recipients` Operation

If modifying `reply` is risky, add a new operation:

```javascript
{
  operation: "reply_with_recipients",
  messageId: "...",
  body: "...",
  to: "...",
  cc: "...",
  bcc: "..."
}
```

### Solution 3: Fix CC Address Parsing

Update CC handling to properly parse multiple addresses:

```javascript
// Parse comma-separated addresses
const ccList = ccAddresses
  .split(',')
  .map(addr => addr.trim())
  .filter(addr => addr.length > 0);

// Format for Outlook
const formattedCc = ccList.join('; ');
```

### Solution 4: Accurate Status Reporting

Return actual operation result:

```javascript
{
  success: true,
  operation: "forward",
  result: "draft_opened",  // or "sent", "queued", "failed"
  draftId: "...",          // if draft was created
  message: "Draft created. Open Outlook to review and send."
}
```

---

## Test Cases for Email Operations

### Test Case 1: Reply with Modified TO
**Setup**: Email thread from Alice, with Bob and Charlie on CC

**Request**:
```javascript
{
  operation: "reply",
  messageId: "...",
  to: "bob@example.com",
  cc: "alice@example.com, charlie@example.com",
  replyBody: "Bob, can you take point on this?"
}
```

**Expected**:
- Subject: "Re: Original Subject" (not "FW:")
- TO: bob@example.com
- CC: alice@example.com, charlie@example.com
- Thread maintained (groups with original in all mail clients)

### Test Case 2: Multiple CC Addresses
**Request**:
```javascript
{
  operation: "reply",
  messageId: "...",
  cc: "a@example.com, b@example.com, c@example.com"
}
```

**Expected**:
- CC field shows three separate recipients
- NOT: `<a@example.com, b@example.com, c@example.com>` as single entity

### Test Case 3: Accurate Success Reporting
**Request**:
```javascript
{ operation: "forward", ..., sendImmediately: false }
```

**Expected Response**:
```javascript
{
  success: true,
  result: "draft_created",
  message: "Draft created in Outlook. Review and send manually."
}
```

**NOT**:
```javascript
{
  success: true,
  message: "Email forwarded successfully"  // Misleading
}
```

---

## Priority

| Bug | Severity | Impact |
|-----|----------|--------|
| Bug 1: reply can't modify recipients | High | Blocks common workflow |
| Bug 2: draft not a reply | High | No workaround for Bug 1 |
| Bug 3: forward breaks threading | Medium | Workaround exists (manual edit) |
| Bug 4: CC address malformed | High | Emails may fail to deliver |
| Bug 5: False success messages | Medium | User confusion, trust issues |

**Recommended Fix Order**: Bug 4 → Bug 1 → Bug 5 → Bug 2 → Bug 3

---

# Feature Request: Auto-Detect HTML in Email Body

**STATUS: IMPLEMENTED** - When `isHtml` is not explicitly provided, the body is automatically scanned for HTML tags.

## Problem

When sending emails, if the body contains HTML tags but `isHtml` is not set to `true`, Outlook renders the literal tags as text:

**Sent:**
```javascript
{
  operation: "send",
  body: "<p>Hello,</p><p>Please review the attached.</p>",
  isHtml: false  // default
}
```

**Displayed in Outlook:**
```
<p>Hello,</p><p>Please review the attached.</p>
```

This is a common mistake that results in malformed emails.

## Proposed Solution: Auto-Detect HTML

Automatically detect HTML content and set the appropriate mode:

```typescript
// In sendEmail(), replyToEmail(), forwardEmail(), createDraft():
const htmlPattern = /<(p|div|br|span|table|ul|ol|li|h[1-6]|a|b|i|strong|em|img|hr)[>\s\/]/i;
const looksLikeHtml = htmlPattern.test(body);
const useHtml = isHtml ?? looksLikeHtml;  // Use explicit isHtml if provided, otherwise auto-detect
```

## Behavior

| `isHtml` param | Body contains HTML | Result |
|----------------|-------------------|--------|
| `true` | any | HTML mode |
| `false` | any | Plain text mode (tags shown literally) |
| not provided | yes | HTML mode (auto-detected) |
| not provided | no | Plain text mode |

## Benefits

- Prevents accidental malformed emails
- Backward compatible (explicit `isHtml` still works)
- Matches user intent automatically

## Priority

**Medium** - Quality of life improvement that prevents common mistakes.

---

# Feature Request: Empty Deleted Items / Permanent Delete

**STATUS: IMPLEMENTED** - The `empty_trash` operation with two-phase safety (preview/confirm).

## Usage

```javascript
// Phase 1: Preview
{ operation: "empty_trash", preview: true }
// Returns: { preview: true, itemCount: 847, oldestItem: "2024-01-15", newestItem: "2025-12-14", totalSizeMB: 156.4 }

// Phase 2: Execute
{ operation: "empty_trash", confirm: true }
// Returns: { deleted: 847, message: "Permanently deleted 847 items from Deleted Items" }
```

---

## Original Use Case

Users need to permanently delete emails to:
- Free up mailbox storage
- Remove sensitive information
- Clean up after bulk triage sessions

## Proposed Operations

### Option 1: `empty_trash` Operation

Empty the entire Deleted Items folder:

```javascript
// Request
{
  operation: "empty_trash",
  confirm: true  // Safety flag required
}

// Response
{
  success: true,
  deleted: 847,
  message: "Permanently deleted 847 items from Deleted Items"
}
```

### Option 2: `permanent_delete` Operation

Permanently delete specific emails by ID:

```javascript
// Request
{
  operation: "permanent_delete",
  messageIds: ["id1", "id2", "id3"],
  confirm: true
}

// Response
{
  success: true,
  deleted: 3
}
```

### Option 3: `purge_folder` Operation

Permanently delete all emails in any folder:

```javascript
// Request
{
  operation: "purge_folder",
  folder: "Deleted Items",  // or any folder path
  olderThan: "30d",         // optional: only items older than 30 days
  confirm: true
}
```

## Safety Considerations

**Required safeguards:**
1. `confirm: true` parameter mandatory - prevents accidental deletion
2. Return count of items to be deleted before acting
3. Consider two-phase: `preview` then `execute`
4. Log all permanent deletions

**Two-phase approach (safest):**
```javascript
// Phase 1: Preview
{ operation: "empty_trash", preview: true }
// Response: { itemCount: 847, oldestItem: "2024-01-15", newestItem: "2025-12-14" }

// Phase 2: Execute (only after user confirms)
{ operation: "empty_trash", confirm: true }
```

## AppleScript Implementation Notes

```applescript
tell application "Microsoft Outlook"
  set deletedFolder to deleted items of exchange account 1
  set msgCount to count of messages of deletedFolder

  -- Permanent delete requires moving to "Permanently Delete" or using delete command
  repeat with msg in messages of deletedFolder
    permanently delete msg
  end repeat
end tell
```

## Priority

Medium - Nice to have for mailbox hygiene, but users can do this manually in Outlook.
