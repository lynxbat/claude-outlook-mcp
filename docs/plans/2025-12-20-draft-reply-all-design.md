# Design: Reply-All Support for Draft Operation

**Date:** 2025-12-20
**Status:** Implemented

## Problem

The `draft` operation with `replyToMessageId` creates a reply draft to the sender only. It ignores recipient parameters and doesn't support reply-all behavior.

Users need to:
- Create reply-all drafts for review before sending
- Modify recipients (add/remove) on reply drafts
- Override recipients entirely on reply drafts

## Solution

Mirror the `reply` operation's recipient control for the `draft` operation when `replyToMessageId` is provided.

## API Design

### New Parameters for Draft Operation

When `replyToMessageId` is provided, these parameters apply:

| Parameter | Type | Description |
|-----------|------|-------------|
| `replyAll` | boolean | Use reply-all (includes all original recipients) |
| `replyTo` | string | Override TO recipients (comma-separated) |
| `replyCc` | string | Override CC recipients (comma-separated) |
| `replyBcc` | string | Override BCC recipients (comma-separated) |
| `addTo` | string | Add to TO recipients (comma-separated) |
| `addCc` | string | Add to CC recipients (comma-separated) |
| `addBcc` | string | Add to BCC recipients (comma-separated) |
| `removeTo` | string | Remove from TO recipients (comma-separated) |
| `removeCc` | string | Remove from CC recipients (comma-separated) |
| `removeBcc` | string | Remove from BCC recipients (comma-separated) |

### Usage Examples

```javascript
// Reply-all draft
{
  operation: "draft",
  replyToMessageId: "10791",
  replyAll: true,
  body: "Thanks everyone."
}

// Reply-all with additional CC
{
  operation: "draft",
  replyToMessageId: "10791",
  replyAll: true,
  addCc: "manager@gnc-hq.com",
  body: "Adding manager for visibility."
}

// Reply with custom recipients (override)
{
  operation: "draft",
  replyToMessageId: "10791",
  replyTo: "specific@gnc-hq.com",
  replyCc: "other@gnc-hq.com",
  body: "Redirecting this thread."
}
```

## Implementation

### 1. Schema Updates

Update descriptions for `replyAll` and recipient parameters to include "for reply or draft operation".

### 2. Function Signature

```typescript
async function createDraft(
  to: string,
  subject: string,
  body: string,
  cc?: string,
  bcc?: string,
  isHtml?: boolean,
  attachments?: string[],
  replyToMessageId?: string,
  replyAll?: boolean,
  recipientOptions?: {
    replyTo?: string;
    replyCc?: string;
    replyBcc?: string;
    addTo?: string;
    addCc?: string;
    addBcc?: string;
    removeTo?: string;
    removeCc?: string;
    removeBcc?: string;
  }
): Promise<string>
```

### 3. AppleScript Changes

In the `if (replyToMessageId)` block:

```applescript
-- Conditionally use reply-all
set replyMsg to reply to theMsg with reply to all without opening window
-- vs
set replyMsg to reply to theMsg without opening window

-- Then apply recipient modifications (same pattern as replyEmail)
```

### 4. Validation

```typescript
// Conflicting parameter checks (same as replyEmail)
if (recipientOptions.replyTo && (recipientOptions.addTo || recipientOptions.removeTo)) {
  return "Error: Cannot use replyTo with addTo or removeTo";
}
// ... same for Cc and Bcc
```

## Test Cases

1. **Basic reply-all draft** - TO: sender, CC: original CCs
2. **Reply-all with addCc** - adds recipient to CC
3. **Reply-all with removeCc** - removes recipient from CC
4. **Reply with replyTo override** - custom TO recipients
5. **Conflicting params** - returns error for replyTo + addTo
