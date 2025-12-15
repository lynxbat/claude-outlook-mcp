# Outlook MCP Plugin Testing Design

## Overview

Add comprehensive testing to the Outlook MCP plugin to prevent recurring bugs where functions ignore their parameters (e.g., folder parameter being ignored, causing all folders to be scanned).

## Decisions

- **Framework:** Bun's built-in test runner (already available, zero config)
- **Test types:** Unit tests + Integration tests against live Outlook
- **Test data:** Real mailbox with structure/type assertions
- **Scope:** All operations including read and write
- **Write safety:** Test sends go to Nick-Weaver@gnc-hq.com

## Test Structure

```
claude-outlook-mcp/
├── index.ts                 # Main plugin code
├── tests/
│   ├── unit/
│   │   ├── parsing.test.ts  # Test delimiter parsing logic
│   │   └── helpers.test.ts  # Test utility functions
│   └── integration/
│       ├── read.test.ts     # readEmails, searchEmails, getUnreadEmails
│       ├── folders.test.ts  # folder listing, folder targeting
│       ├── send.test.ts     # email sending
│       └── calendar.test.ts # calendar operations
├── tests/setup.ts           # Shared test config & helpers
└── package.json             # Add "test" scripts
```

## Unit Tests

Test the TypeScript parsing logic without needing Outlook.

### `tests/unit/parsing.test.ts`

```typescript
describe("parseEmailOutput", () => {
  it("parses single email correctly", () => {
    const input = "<<<MSG>>>Test Subject<<<FROM>>>sender@test.com<<<DATE>>>Dec 4, 2025<<<CONTENT>>>Hello<<<ENDMSG>>>";
    const result = parseEmailOutput(input);
    expect(result).toHaveLength(1);
    expect(result[0].subject).toBe("Test Subject");
    expect(result[0].sender).toBe("sender@test.com");
  });

  it("parses multiple emails correctly", () => { ... });
  it("handles missing content gracefully", () => { ... });
  it("handles special characters in subject", () => { ... });
  it("returns empty array for empty input", () => { ... });
});
```

### `tests/unit/helpers.test.ts`

- Test `buildFolderRef()` - "Inbox" → `inbox`, other → `mail folder "Name"`
- Test `escapeForAppleScript()` - proper string escaping

## Integration Tests - Read Operations

Test that functions return data from the correct folder.

### `tests/integration/read.test.ts`

```typescript
describe("readEmails", () => {
  it("returns emails from Inbox only", async () => {
    const emails = await readEmails("Inbox", 5);
    expect(emails.length).toBeGreaterThan(0);
    expect(emails.length).toBeLessThanOrEqual(5);
    for (const email of emails) {
      expect(email.subject).toBeDefined();
      expect(email.sender).toBeDefined();
      expect(email.dateSent).toBeDefined();
    }
  });

  it("respects limit parameter", async () => {
    const emails = await readEmails("Inbox", 3);
    expect(emails.length).toBeLessThanOrEqual(3);
  });

  it("returns different results for different folders", async () => {
    const inbox = await readEmails("Inbox", 5);
    const archive = await readEmails("Archive", 5);
    const inboxSubjects = inbox.map(e => e.subject);
    const archiveSubjects = archive.map(e => e.subject);
    expect(inboxSubjects).not.toEqual(archiveSubjects);
  });
});

describe("searchEmails", () => {
  it("searches only specified folder", async () => { ... });
  it("returns empty array when no matches", async () => { ... });
  it("search is case-insensitive", async () => { ... });
});
```

**Key test:** "returns different results for different folders" directly catches the bug we fixed.

## Integration Tests - Write Operations

Test sends go to Nick-Weaver@gnc-hq.com for verification.

### `tests/integration/send.test.ts`

```typescript
const TEST_RECIPIENT = "Nick-Weaver@gnc-hq.com";

describe("sendEmail", () => {
  it("sends email successfully", async () => {
    const subject = `[TEST] Outlook MCP Test - ${Date.now()}`;
    const result = await sendEmail(
      TEST_RECIPIENT,
      subject,
      "This is an automated test email. Safe to delete."
    );
    expect(result).toContain("success");
  });

  it("sends email with CC", async () => { ... });

  it("fails gracefully with invalid recipient", async () => {
    await expect(sendEmail("not-an-email", "Test", "Body"))
      .rejects.toThrow();
  });
});
```

### `tests/integration/calendar.test.ts`

```typescript
describe("createCalendarEvent", () => {
  it("creates event and returns confirmation", async () => {
    const result = await createCalendarEvent(
      "[TEST] MCP Test Event",
      new Date(Date.now() + 86400000),
      new Date(Date.now() + 90000000),
      "Automated test - safe to delete"
    );
    expect(result).toContain("success");
  });
});
```

## Test Setup

### `tests/setup.ts`

```typescript
export const TEST_RECIPIENT = "Nick-Weaver@gnc-hq.com";
export const TEST_TIMEOUT = 30000; // AppleScript can be slow

export async function ensureOutlookRunning(): Promise<void> {
  // Check Outlook is running before integration tests
}
```

## Refactoring Required

Extract inline parsing logic to testable functions:

1. `parseEmailOutput(raw: string): Email[]` - Delimiter parsing
2. `buildFolderRef(folder: string): string` - Folder reference builder
3. `escapeForAppleScript(str: string): string` - String escaping

## Package.json Changes

```json
"scripts": {
  "dev": "bun run index.ts",
  "start": "bun run index.ts",
  "test": "bun test",
  "test:unit": "bun test tests/unit",
  "test:integration": "bun test tests/integration"
}
```

## Running Tests

- `bun test` - Run all tests
- `bun test tests/unit` - Unit tests only (fast, no Outlook needed)
- `bun test tests/integration` - Integration tests only (requires Outlook running)

## When to Run

- Before committing changes to the plugin
- After fixing bugs (add regression test first using TDD)
