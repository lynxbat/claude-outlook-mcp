# Design: Enhanced Reply Recipients (Bug 1)

## Problem

The `reply` operation cannot modify recipients. It only replies to the original sender (or all with `replyAll`), ignoring any `to`/`cc` parameters.

## Solution

Add 9 new parameters to `reply` operation for full recipient control:

### New Parameters

| Parameter | Type | Description |
|-----------|------|-------------|
| `to` | string | Override all TO recipients |
| `cc` | string | Override all CC recipients |
| `bcc` | string | Override all BCC recipients |
| `addTo` | string | Add to TO recipients |
| `addCc` | string | Add to CC recipients |
| `addBcc` | string | Add to BCC recipients |
| `removeTo` | string | Remove from TO recipients |
| `removeCc` | string | Remove from CC recipients |
| `removeBcc` | string | Remove from BCC recipients |

All parameters accept comma-separated email strings.

### Validation Rules

- Error if both `to` and (`addTo` or `removeTo`) provided
- Error if both `cc` and (`addCc` or `removeCc`) provided
- Error if both `bcc` and (`addBcc` or `removeBcc`) provided

### Implementation Approach

1. Create reply with `reply to theMsg` (preserves threading)
2. Get existing recipients from the reply message
3. Apply modifications:
   - Override: Clear existing recipients, add specified ones
   - Add: Keep existing, append new recipients
   - Remove: Filter out matching email addresses
4. Rebuild recipient lists using `make new to/cc/bcc recipient`

### AppleScript Operations

```applescript
-- Clear existing TO recipients
delete every to recipient of replyMsg

-- Add new recipient
make new to recipient at replyMsg with properties {email address:{address:"x@y.com"}}

-- Get existing recipients (for add/remove operations)
set existingTo to to recipients of replyMsg
```

### Example Usage

**Override recipients:**
```json
{
  "operation": "reply",
  "messageId": "123",
  "replyBody": "Bob, can you take point?",
  "to": "bob@example.com",
  "cc": "alice@example.com, charlie@example.com"
}
```

**Add someone to CC:**
```json
{
  "operation": "reply",
  "messageId": "123",
  "replyBody": "Adding manager to this thread",
  "addCc": "manager@example.com"
}
```

**Remove someone from reply:**
```json
{
  "operation": "reply",
  "messageId": "123",
  "replyBody": "Discussing without Bob",
  "replyAll": true,
  "removeTo": "bob@example.com"
}
```

### Behavior Notes

- Threading preserved (Re: subject, In-Reply-To headers maintained)
- If no recipient params provided, behaves as before (reply to sender or replyAll)
- `replyAll: true` + modifications = start with all original recipients, then apply changes
- `replyAll: false` + modifications = start with sender only, then apply changes
