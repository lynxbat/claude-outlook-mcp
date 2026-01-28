import { describe, expect, it } from "bun:test";
import { buildFolderRef, buildNestedFolderRef, escapeForAppleScript, parseRecipients, detectHtml, isInboxFolder, buildFolderSearchScript, INBOX_LOCALIZATIONS } from "../../helpers";

describe("buildFolderRef", () => {
  // NOTE: buildFolderRef no longer returns "inbox" for Inbox folder.
  // This is intentional - the AppleScript "inbox" keyword often points to an empty
  // local folder in localized Outlook installations. We always use mail folder references.
  it("returns mail folder reference for Inbox folder (not inbox keyword)", () => {
    expect(buildFolderRef("Inbox")).toBe('mail folder "Inbox"');
  });

  it("returns mail folder reference for other folders", () => {
    expect(buildFolderRef("Archive")).toBe('mail folder "Archive"');
    expect(buildFolderRef("Sent Items")).toBe('mail folder "Sent Items"');
    expect(buildFolderRef("Drafts")).toBe('mail folder "Drafts"');
  });

  it("handles folders with special characters", () => {
    expect(buildFolderRef("My Folder")).toBe('mail folder "My Folder"');
    expect(buildFolderRef("Work/Projects")).toBe('mail folder "Work/Projects"');
  });

  it("treats all inbox variants the same", () => {
    expect(buildFolderRef("inbox")).toBe('mail folder "inbox"');
    expect(buildFolderRef("INBOX")).toBe('mail folder "INBOX"');
  });

  it("handles empty string", () => {
    expect(buildFolderRef("")).toBe('mail folder ""');
  });
});

describe("buildNestedFolderRef", () => {
  // NOTE: buildNestedFolderRef no longer returns "inbox" for Inbox folder.
  // This is intentional to support localized Outlook installations.
  it("returns mail folder reference for Inbox folder (not inbox keyword)", () => {
    expect(buildNestedFolderRef("Inbox")).toBe('mail folder "Inbox"');
  });

  it("returns mail folder reference for flat folders", () => {
    expect(buildNestedFolderRef("Archive")).toBe('mail folder "Archive"');
    expect(buildNestedFolderRef("Reports")).toBe('mail folder "Reports"');
  });

  it("builds nested reference for paths with slashes", () => {
    expect(buildNestedFolderRef("Work/Reports")).toBe('mail folder "Reports" of mail folder "Work"');
  });

  it("handles deeply nested paths", () => {
    expect(buildNestedFolderRef("A/B/C")).toBe('mail folder "C" of mail folder "B" of mail folder "A"');
  });

  it("handles folders with spaces in nested paths", () => {
    expect(buildNestedFolderRef("My Programs/My Project")).toBe('mail folder "My Project" of mail folder "My Programs"');
  });
});

describe("escapeForAppleScript", () => {
  it("escapes double quotes", () => {
    expect(escapeForAppleScript('He said "hello"')).toBe('He said \\"hello\\"');
  });

  it("escapes backslashes", () => {
    expect(escapeForAppleScript("path\\to\\file")).toBe("path\\\\to\\\\file");
  });

  it("escapes both quotes and backslashes", () => {
    expect(escapeForAppleScript('path\\with"quotes')).toBe('path\\\\with\\"quotes');
  });

  it("returns unchanged string when no escaping needed", () => {
    expect(escapeForAppleScript("Hello world")).toBe("Hello world");
    expect(escapeForAppleScript("test@email.com")).toBe("test@email.com");
  });

  it("handles empty string", () => {
    expect(escapeForAppleScript("")).toBe("");
  });

  it("handles multiple consecutive special characters", () => {
    // Input: """\\  (three quotes and two backslashes)
    // Output: \"\"\"\\\\  (escaped quotes and backslashes)
    expect(escapeForAppleScript('"""' + "\\\\" + "")).toBe('\\"\\"\\"\\\\\\\\');
  });
});

describe("parseRecipients", () => {
  it("parses single email address", () => {
    const result = parseRecipients("test@example.com");
    expect(result).toHaveLength(1);
    expect(result[0].address).toBe("test@example.com");
    expect(result[0].name).toBe("test");
  });

  it("parses multiple comma-separated addresses", () => {
    const result = parseRecipients("a@example.com, b@example.com, c@example.com");
    expect(result).toHaveLength(3);
    expect(result[0].address).toBe("a@example.com");
    expect(result[1].address).toBe("b@example.com");
    expect(result[2].address).toBe("c@example.com");
  });

  it("trims whitespace from addresses", () => {
    const result = parseRecipients("  a@example.com  ,   b@example.com   ");
    expect(result).toHaveLength(2);
    expect(result[0].address).toBe("a@example.com");
    expect(result[1].address).toBe("b@example.com");
  });

  it("extracts name from display format", () => {
    const result = parseRecipients("John Doe <john@example.com>");
    expect(result).toHaveLength(1);
    expect(result[0].address).toBe("John Doe <john@example.com>");
    expect(result[0].name).toBe("John Doe");
  });

  it("extracts name from plain email", () => {
    const result = parseRecipients("john.doe@example.com");
    expect(result).toHaveLength(1);
    expect(result[0].name).toBe("john.doe");
  });
});

describe("detectHtml", () => {
  it("detects common HTML tags", () => {
    expect(detectHtml("<p>Hello</p>")).toBe(true);
    expect(detectHtml("<div>Content</div>")).toBe(true);
    expect(detectHtml("<br>")).toBe(true);
    expect(detectHtml("<br/>")).toBe(true);
    expect(detectHtml("<span>text</span>")).toBe(true);
  });

  it("detects HTML tags with attributes", () => {
    expect(detectHtml('<a href="http://example.com">link</a>')).toBe(true);
    expect(detectHtml('<img src="image.png" alt="test">')).toBe(true);
    expect(detectHtml('<div class="container">')).toBe(true);
  });

  it("detects various block elements", () => {
    expect(detectHtml("<table><tr><td>data</td></tr></table>")).toBe(true);
    expect(detectHtml("<ul><li>item</li></ul>")).toBe(true);
    expect(detectHtml("<ol><li>item</li></ol>")).toBe(true);
    expect(detectHtml("<h1>Title</h1>")).toBe(true);
    expect(detectHtml("<h6>Small heading</h6>")).toBe(true);
  });

  it("detects inline formatting tags", () => {
    expect(detectHtml("<b>bold</b>")).toBe(true);
    expect(detectHtml("<i>italic</i>")).toBe(true);
    expect(detectHtml("<strong>strong</strong>")).toBe(true);
    expect(detectHtml("<em>emphasis</em>")).toBe(true);
  });

  it("returns false for plain text", () => {
    expect(detectHtml("Hello world")).toBe(false);
    expect(detectHtml("Just some text")).toBe(false);
    expect(detectHtml("Price is $50")).toBe(false);
  });

  it("returns false for text with angle brackets that arent HTML", () => {
    expect(detectHtml("5 < 10 and 10 > 5")).toBe(false);
    expect(detectHtml("Use -> for arrows")).toBe(false);
    expect(detectHtml("<invalid>not a tag")).toBe(false);
  });

  it("handles case insensitivity", () => {
    expect(detectHtml("<P>Uppercase</P>")).toBe(true);
    expect(detectHtml("<DIV>Mixed</div>")).toBe(true);
    expect(detectHtml("<Br>")).toBe(true);
  });

  it("handles empty string", () => {
    expect(detectHtml("")).toBe(false);
  });
});

describe("isInboxFolder", () => {
  it("recognizes English inbox", () => {
    expect(isInboxFolder("Inbox")).toBe(true);
    expect(isInboxFolder("inbox")).toBe(true);
    expect(isInboxFolder("INBOX")).toBe(true);
  });

  it("recognizes German inbox", () => {
    expect(isInboxFolder("Posteingang")).toBe(true);
  });

  it("recognizes French inbox", () => {
    expect(isInboxFolder("Boîte de réception")).toBe(true);
  });

  it("recognizes Spanish inbox", () => {
    expect(isInboxFolder("Bandeja de entrada")).toBe(true);
  });

  it("returns false for non-inbox folders", () => {
    expect(isInboxFolder("Sent Items")).toBe(false);
    expect(isInboxFolder("Archive")).toBe(false);
    expect(isInboxFolder("Drafts")).toBe(false);
  });
});

describe("buildFolderSearchScript", () => {
  it("generates localization-aware search for Inbox", () => {
    const script = buildFolderSearchScript("theFolder", "Inbox");
    expect(script).toContain("Posteingang");
    expect(script).toContain("Boîte de réception");
    expect(script).toContain("set theFolder to");
  });

  it("generates localization-aware search for localized inbox names", () => {
    const script = buildFolderSearchScript("myFolder", "Posteingang");
    expect(script).toContain("Posteingang");
    expect(script).toContain("set myFolder to");
  });

  it("generates direct reference for non-inbox folders", () => {
    const script = buildFolderSearchScript("theFolder", "Archive");
    expect(script).toBe('set theFolder to mail folder "Archive"');
  });

  it("generates nested reference for path folders", () => {
    const script = buildFolderSearchScript("theFolder", "Work/Projects");
    expect(script).toBe('set theFolder to mail folder "Projects" of mail folder "Work"');
  });
});

describe("INBOX_LOCALIZATIONS", () => {
  it("includes common languages", () => {
    expect(INBOX_LOCALIZATIONS).toContain("Posteingang");  // German
    expect(INBOX_LOCALIZATIONS).toContain("Boîte de réception");  // French
    expect(INBOX_LOCALIZATIONS).toContain("Bandeja de entrada");  // Spanish
    expect(INBOX_LOCALIZATIONS).toContain("Inbox");  // English
  });

  it("has Inbox listed last (to prefer localized folders with messages)", () => {
    expect(INBOX_LOCALIZATIONS[INBOX_LOCALIZATIONS.length - 1]).toBe("Inbox");
  });
});
