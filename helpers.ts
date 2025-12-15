// Shared helper functions for Outlook MCP plugin
// Extracted for testability

export interface ParsedEmail {
  messageId?: string;
  subject: string;
  sender: string;
  dateSent: string;
  content: string;
}

export interface EmailRecipient {
  name: string;
  address: string;
}

/**
 * Extract display name from email address
 * "John Doe <john@example.com>" -> "John Doe"
 * "john@example.com" -> "john"
 */
function extractNameFromEmail(email: string): string {
  const match = email.match(/^([^<]+)</);
  if (match) {
    return match[1].trim();
  }
  return email.split('@')[0];
}

/**
 * Parse comma-separated email addresses into recipient objects
 * "a@b.com, c@d.com" -> [{name: "a", address: "a@b.com"}, {name: "c", address: "c@d.com"}]
 */
export function parseRecipients(emailString: string): EmailRecipient[] {
  return emailString.split(',').map(email => {
    const trimmed = email.trim();
    return { name: extractNameFromEmail(trimmed), address: trimmed };
  });
}

/**
 * Parse email output from AppleScript delimited format
 * Supports two formats:
 * - With ID: <<<MSG>>>subject<<<ID>>>id<<<FROM>>>sender<<<DATE>>>date<<<CONTENT>>>content<<<ENDMSG>>>
 * - Without ID: <<<MSG>>>subject<<<FROM>>>sender<<<DATE>>>date<<<CONTENT>>>content<<<ENDMSG>>>
 */
export function parseEmailOutput(raw: string): ParsedEmail[] {
  if (!raw || raw.trim() === "") {
    return [];
  }

  const emails: ParsedEmail[] = [];
  const messageBlocks = raw.split("<<<MSG>>>").filter(b => b.trim());

  for (const block of messageBlocks) {
    // Check if block contains ID delimiter
    const hasId = block.includes("<<<ID>>>");

    let subjectMatch, idMatch, senderMatch, dateMatch, contentMatch;

    if (hasId) {
      // Format with ID: subject<<<ID>>>id<<<FROM>>>...
      subjectMatch = block.match(/^(.*)<<<ID>>>/s);
      idMatch = block.match(/<<<ID>>>(.*)<<<FROM>>>/s);
      senderMatch = block.match(/<<<FROM>>>(.*)<<<DATE>>>/s);
      dateMatch = block.match(/<<<DATE>>>(.*)<<<CONTENT>>>/s);
      contentMatch = block.match(/<<<CONTENT>>>(.*)<<<ENDMSG>>>/s);
    } else {
      // Format without ID: subject<<<FROM>>>...
      subjectMatch = block.match(/^(.*)<<<FROM>>>/s);
      senderMatch = block.match(/<<<FROM>>>(.*)<<<DATE>>>/s);
      dateMatch = block.match(/<<<DATE>>>(.*)<<<CONTENT>>>/s);
      contentMatch = block.match(/<<<CONTENT>>>(.*)<<<ENDMSG>>>/s);
    }

    if (subjectMatch) {
      const contentText = contentMatch ? contentMatch[1].trim() : "";
      emails.push({
        messageId: hasId && idMatch ? idMatch[1].trim() : undefined,
        subject: subjectMatch[1].trim() || "No subject",
        sender: senderMatch ? senderMatch[1].trim() : "Unknown sender",
        dateSent: dateMatch ? dateMatch[1].trim() : new Date().toString(),
        content: contentText || "[Content not available]"
      });
    }
  }

  return emails;
}

/**
 * Build AppleScript folder reference
 * "Inbox" -> inbox (built-in reference)
 * Other folders -> mail folder "Name" (named reference)
 */
export function buildFolderRef(folder: string): string {
  return folder === "Inbox" ? "inbox" : `mail folder "${folder}"`;
}

/**
 * Build AppleScript folder reference for nested paths
 * "Inbox" -> inbox
 * "Reports" -> mail folder "Reports"
 * "Work/Reports" -> mail folder "Reports" of mail folder "Work"
 */
export function buildNestedFolderRef(path: string): string {
  if (path === "Inbox") {
    return "inbox";
  }

  const parts = path.split("/");

  if (parts.length === 1) {
    return `mail folder "${parts[0]}"`;
  }

  // Build nested reference from innermost to outermost
  // "Work/Reports" -> mail folder "Reports" of mail folder "Work"
  let ref = `mail folder "${parts[parts.length - 1]}"`;
  for (let i = parts.length - 2; i >= 0; i--) {
    ref += ` of mail folder "${parts[i]}"`;
  }
  return ref;
}

/**
 * Escape special characters for AppleScript strings
 */
export function escapeForAppleScript(str: string): string {
  return str
    .replace(/\\/g, "\\\\")
    .replace(/"/g, '\\"');
}

/**
 * Detect if a string contains HTML content
 * Looks for common HTML tags to determine if content should be treated as HTML
 */
export function detectHtml(content: string): boolean {
  const htmlPattern = /<(p|div|br|span|table|ul|ol|li|h[1-6]|a|b|i|strong|em|img|hr)[>\s\/]/i;
  return htmlPattern.test(content);
}
