import { describe, expect, it } from "bun:test";
import { parseEmailOutput, ParsedEmail } from "../../helpers";

describe("parseEmailOutput", () => {
  it("parses single email correctly", () => {
    const input = "<<<MSG>>>Test Subject<<<FROM>>>sender@test.com<<<DATE>>>Dec 4, 2025<<<CONTENT>>>Hello world<<<ENDMSG>>>";
    const result = parseEmailOutput(input);

    expect(result).toHaveLength(1);
    expect(result[0].subject).toBe("Test Subject");
    expect(result[0].sender).toBe("sender@test.com");
    expect(result[0].dateSent).toBe("Dec 4, 2025");
    expect(result[0].content).toBe("Hello world");
  });

  it("parses multiple emails correctly", () => {
    const input =
      "<<<MSG>>>First Subject<<<FROM>>>first@test.com<<<DATE>>>Dec 4, 2025<<<CONTENT>>>First body<<<ENDMSG>>>" +
      "<<<MSG>>>Second Subject<<<FROM>>>second@test.com<<<DATE>>>Dec 5, 2025<<<CONTENT>>>Second body<<<ENDMSG>>>";
    const result = parseEmailOutput(input);

    expect(result).toHaveLength(2);
    expect(result[0].subject).toBe("First Subject");
    expect(result[1].subject).toBe("Second Subject");
  });

  it("handles missing content gracefully", () => {
    const input = "<<<MSG>>>Test Subject<<<FROM>>>sender@test.com<<<DATE>>>Dec 4, 2025<<<CONTENT>>><<<ENDMSG>>>";
    const result = parseEmailOutput(input);

    expect(result).toHaveLength(1);
    expect(result[0].content).toBe("[Content not available]");
  });

  it("handles special characters in subject", () => {
    const input = "<<<MSG>>>Re: Meeting @ 3pm - Don't forget!<<<FROM>>>sender@test.com<<<DATE>>>Dec 4, 2025<<<CONTENT>>>Test<<<ENDMSG>>>";
    const result = parseEmailOutput(input);

    expect(result).toHaveLength(1);
    expect(result[0].subject).toBe("Re: Meeting @ 3pm - Don't forget!");
  });

  it("returns empty array for empty input", () => {
    expect(parseEmailOutput("")).toEqual([]);
    expect(parseEmailOutput("   ")).toEqual([]);
  });

  it("returns empty array for null/undefined input", () => {
    expect(parseEmailOutput(null as any)).toEqual([]);
    expect(parseEmailOutput(undefined as any)).toEqual([]);
  });

  it("handles multiline content", () => {
    const input = "<<<MSG>>>Subject<<<FROM>>>sender@test.com<<<DATE>>>Dec 4, 2025<<<CONTENT>>>Line 1\nLine 2\nLine 3<<<ENDMSG>>>";
    const result = parseEmailOutput(input);

    expect(result).toHaveLength(1);
    expect(result[0].content).toContain("Line 1");
    expect(result[0].content).toContain("Line 2");
    expect(result[0].content).toContain("Line 3");
  });

  it("handles emails with no subject", () => {
    const input = "<<<MSG>>><<<FROM>>>sender@test.com<<<DATE>>>Dec 4, 2025<<<CONTENT>>>Body<<<ENDMSG>>>";
    const result = parseEmailOutput(input);

    expect(result).toHaveLength(1);
    expect(result[0].subject).toBe("No subject");
  });

  it("trims whitespace from fields", () => {
    const input = "<<<MSG>>>  Spaced Subject  <<<FROM>>>  sender@test.com  <<<DATE>>>  Dec 4, 2025  <<<CONTENT>>>  Body  <<<ENDMSG>>>";
    const result = parseEmailOutput(input);

    expect(result).toHaveLength(1);
    expect(result[0].subject).toBe("Spaced Subject");
    expect(result[0].sender).toBe("sender@test.com");
    expect(result[0].content).toBe("Body");
  });
});
