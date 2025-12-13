import { describe, expect, it, beforeAll } from "bun:test";
import { ensureOutlookRunning, TEST_TIMEOUT } from "../setup";
import { readEmails } from "../../index";

describe("readEmails date filtering", () => {
  beforeAll(async () => {
    await ensureOutlookRunning();
  });

  it("returns emails without date filter", async () => {
    // Baseline: confirm we can get emails without any filter
    const emails = await readEmails("Inbox", 10);
    console.log(`Without filter: found ${emails.length} emails`);
    expect(emails.length).toBeGreaterThan(0);
  }, TEST_TIMEOUT);

  it("returns emails for a single day range (same start and end date)", async () => {
    // First, get emails without filter to find a valid date
    const allEmails = await readEmails("Inbox", 10);
    expect(allEmails.length).toBeGreaterThan(0);

    // Parse the date from the first email to get a known valid date
    const firstEmail = allEmails[0];
    console.log(`First email date: ${firstEmail.dateSent}`);

    // Parse the date - Outlook returns dates like "Friday, December 5, 2025 at 3:45:00 PM"
    const dateStr = firstEmail.dateSent;
    const dateMatch = dateStr.match(/(\w+),\s+(\w+)\s+(\d+),\s+(\d+)/);

    if (dateMatch) {
      const month = dateMatch[2];
      const day = parseInt(dateMatch[3]);
      const year = parseInt(dateMatch[4]);

      // Convert month name to number
      const months: Record<string, number> = {
        January: 0, February: 1, March: 2, April: 3, May: 4, June: 5,
        July: 6, August: 7, September: 8, October: 9, November: 10, December: 11
      };

      const monthNum = months[month];
      const targetDate = new Date(year, monthNum, day);
      const isoDate = targetDate.toISOString().split('T')[0]; // Format: YYYY-MM-DD

      console.log(`Testing date filter for: ${isoDate}`);

      // Now query with same start and end date (single day)
      const filteredEmails = await readEmails("Inbox", 10, isoDate, isoDate);

      console.log(`With date filter (${isoDate} to ${isoDate}): found ${filteredEmails.length} emails`);

      // Should find at least one email since we know there's an email on this date
      expect(filteredEmails.length).toBeGreaterThan(0);
    } else {
      console.log(`Could not parse date from: ${dateStr}`);
      throw new Error(`Could not parse date format: ${dateStr}`);
    }
  }, TEST_TIMEOUT);

  it("returns emails for a date range spanning multiple days", async () => {
    // Test a week range ending today
    const endDate = new Date();
    endDate.setHours(23, 59, 59, 999);
    const startDate = new Date();
    startDate.setDate(startDate.getDate() - 7);
    startDate.setHours(0, 0, 0, 0);

    const startIso = startDate.toISOString().split('T')[0];
    const endIso = endDate.toISOString().split('T')[0];

    console.log(`Testing week range: ${startIso} to ${endIso}`);

    const filteredEmails = await readEmails("Inbox", 50, startIso, endIso);
    console.log(`Week range found: ${filteredEmails.length} emails`);

    // Should find some emails in the week range
    expect(filteredEmails.length).toBeGreaterThan(0);

    // Verify all returned emails are within the date range (with some tolerance for timezone)
    for (const email of filteredEmails) {
      // Parse Outlook date format: "Friday, December 5, 2025 at 9:19:43 PM"
      const dateStr = email.dateSent;
      const match = dateStr.match(/(\w+),\s+(\w+)\s+(\d+),\s+(\d+)/);
      if (match) {
        const months: Record<string, number> = {
          January: 0, February: 1, March: 2, April: 3, May: 4, June: 5,
          July: 6, August: 7, September: 8, October: 9, November: 10, December: 11
        };
        const month = months[match[2]];
        const day = parseInt(match[3]);
        const year = parseInt(match[4]);
        const emailDate = new Date(year, month, day);

        console.log(`Email date: ${emailDate.toISOString().split('T')[0]}, Range: ${startIso} to ${endIso}`);

        // Compare dates only (ignore time)
        const emailDateOnly = new Date(year, month, day);
        const startDateOnly = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
        const endDateOnly = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());

        expect(emailDateOnly >= startDateOnly).toBe(true);
        expect(emailDateOnly <= endDateOnly).toBe(true);
      }
    }
  }, TEST_TIMEOUT);
});
