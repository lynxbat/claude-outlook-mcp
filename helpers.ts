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
 * Common localized names for the Inbox folder across different languages.
 * Used to identify when a user is requesting the inbox regardless of their locale.
 * 
 * NOTE: "Inbox" is listed LAST because in localized Outlook installations,
 * there's often an empty local "Inbox" folder alongside the real localized inbox.
 * By checking localized names first, we find the real inbox with messages.
 */
export const INBOX_LOCALIZATIONS = [
  "Posteingang",     // German
  "Boîte de réception", // French
  "Bandeja de entrada", // Spanish
  "Posta in arrivo", // Italian
  "Postvak IN",      // Dutch
  "Caixa de Entrada", // Portuguese
  "Indbakke",        // Danish
  "Innboks",         // Norwegian
  "Skrzynka odbiorcza", // Polish
  "Входящие",        // Russian
  "受信トレイ",       // Japanese
  "收件箱",          // Chinese
  "Inbox",           // English (checked last - often empty local folder)
];

/**
 * Check if a folder name refers to the inbox (in any supported language)
 */
export function isInboxFolder(folder: string): boolean {
  return INBOX_LOCALIZATIONS.some(
    name => name.toLowerCase() === folder.toLowerCase()
  );
}

/**
 * Build AppleScript folder reference
 * 
 * NOTE: The AppleScript "inbox" keyword often points to an empty local folder
 * instead of the actual Exchange inbox in localized Outlook installations.
 * We always use mail folder references with localization fallback to avoid this issue.
 * 
 * "Inbox" -> searches for inbox by common localized names
 * Other folders -> mail folder "Name" (named reference)
 */
export function buildFolderRef(folder: string): string {
  // Always use named reference - the "inbox" keyword is unreliable
  // with Exchange accounts in localized Outlook installations
  return `mail folder "${folder}"`;
}

/**
 * Build AppleScript folder reference for nested paths
 * "Inbox" -> mail folder "Inbox" (or localized equivalent)
 * "Reports" -> mail folder "Reports"
 * "Work/Reports" -> mail folder "Reports" of mail folder "Work"
 * 
 * NOTE: We no longer use the "inbox" keyword as it's unreliable with Exchange accounts.
 */
export function buildNestedFolderRef(path: string): string {
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
 * Generate AppleScript code to find a folder, with fallback for localized inbox names.
 * This handles the case where "Inbox" might be called "Posteingang", "Boîte de réception", etc.
 * 
 * @param folderVar - The AppleScript variable name to store the found folder
 * @param folderName - The folder name requested by the user
 * @returns AppleScript code that sets folderVar to the correct folder
 */
export function buildFolderSearchScript(folderVar: string, folderName: string): string {
  // Check if user is requesting inbox (in any language)
  if (isInboxFolder(folderName)) {
    // Try each localized inbox name in order (localized names first, "Inbox" last)
    // This ensures we find the real inbox with messages, not the empty local "Inbox"
    const folderSearchBlocks = INBOX_LOCALIZATIONS.map(name => `
        if ${folderVar} is null then
          repeat with aFolder in allFolders
            if name of aFolder is "${name}" then
              -- Only use this folder if it has messages (skip empty local folders)
              try
                if (count of messages of aFolder) > 0 then
                  set ${folderVar} to aFolder
                  exit repeat
                end if
              end try
            end if
          end repeat
        end if`
    ).join("\n");
    
    return `
      -- Find inbox folder (handles localization)
      set ${folderVar} to null
      set allFolders to mail folders
      ${folderSearchBlocks}
      
      -- If still not found, try any folder with an inbox-like name (even if empty)
      if ${folderVar} is null then
        repeat with aFolder in allFolders
          set folderName to name of aFolder
          if folderName is "Inbox" or folderName is "Posteingang" or folderName is "Boîte de réception" then
            set ${folderVar} to aFolder
            exit repeat
          end if
        end repeat
      end if
      
      if ${folderVar} is null then error "Could not find inbox folder"
    `;
  }
  
  // For non-inbox folders, use direct reference
  const nestedRef = buildNestedFolderRef(folderName);
  return `set ${folderVar} to ${nestedRef}`;
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
