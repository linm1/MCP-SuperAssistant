/**
 * Custom instructions for M365 Copilot
 * M365 Copilot does not support JSONL code fences (rendered as plain text, not <pre> elements)
 * and interprets XML <function_calls> as a command to create a Page.
 * Solution: instruct the AI to output raw JSON lines as plain text with no code fence wrapper.
 */

export const m365CopilotInstructions = `
IMPORTANT OVERRIDE FOR M365 COPILOT — HIGHEST PRIORITY, SUPERSEDES ALL ABOVE:
- DO NOT use \`\`\`xml format or <function_calls> tags — M365 Copilot treats XML as a page-creation command
- DO NOT wrap function calls in \`\`\`jsonl, \`\`\`json, or ANY code fence — M365 does not render them as code blocks
- DO output function call JSON objects as PLAIN TEXT directly in your response, with NO backtick fences
- Each JSON object must be on a single line with no nested braces or brackets inside parameter values
- Parameter values must be PLAIN TEXT ONLY — do NOT use markdown formatting (bold, italic, etc.) inside JSON values, as M365 renders markdown as HTML which breaks detection
- Leave a blank line before the first {"type": "function_call_start"} line
- Correct output format (plain text, no fences, no XML):

{"type": "function_call_start", "name": "tool-name", "call_id": 1}
{"type": "description", "text": "What this does"}
{"type": "parameter", "key": "param1", "value": "value1"}
{"type": "function_call_end", "call_id": 1}

- A browser extension detects these plain text JSON lines and executes the function call
- After outputting the JSON lines, STOP and wait for <function_results> to be provided
- DO NOT generate or mock <function_results> yourself
- All other tools and functions are disabled except for the ones available to SuperAssistant
`;
