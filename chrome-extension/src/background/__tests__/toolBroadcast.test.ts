/**
 * TDD: Tests for the tool broadcast payload format bug.
 *
 * Root cause: broadcastToolsUpdateToContentScripts() sends
 *   payload: { tools: [...] }
 * but the content script listener at mcp-client.ts:191 checks
 *   Array.isArray(message.payload)
 * which is false for an object, resulting in 0 tools processed.
 */

import { describe, it, expect } from 'vitest';

// ---- Reproduce the BUG (old behaviour) ----
function buildBroadcastPayloadBuggy(tools: any[]) {
  return { tools }; // wrapped in object — this is the bug
}

function receiveToolsBuggy(payload: any): any[] {
  // Current code in mcp-client.ts:191
  return Array.isArray(payload) ? payload : [];
}

// ---- Expected FIX (new behaviour) ----
function buildBroadcastPayloadFixed(tools: any[]) {
  return tools; // flat array
}

function receiveToolsFixed(payload: any): any[] {
  // Defensive: accept both flat array and wrapped { tools } for backwards compat
  if (Array.isArray(payload)) return payload;
  if (Array.isArray(payload?.tools)) return payload.tools;
  return [];
}

const SAMPLE_TOOLS = [
  { name: 'tool_a', description: 'Tool A', input_schema: {} },
  { name: 'tool_b', description: 'Tool B', input_schema: {} },
];

describe('Tool broadcast payload — bug reproduction', () => {
  it('BUG: old sender wraps tools in object, receiver returns 0 tools', () => {
    const payload = buildBroadcastPayloadBuggy(SAMPLE_TOOLS);
    const received = receiveToolsBuggy(payload);
    // This demonstrates the bug: 2 tools sent but 0 received
    expect(received).toHaveLength(0);
  });
});

describe('Tool broadcast payload — fixed behaviour', () => {
  it('FIX: sender sends flat array, receiver correctly returns all tools', () => {
    const payload = buildBroadcastPayloadFixed(SAMPLE_TOOLS);
    const received = receiveToolsFixed(payload);
    expect(received).toHaveLength(2);
    expect(received[0].name).toBe('tool_a');
  });

  it('FIX: receiver handles legacy wrapped { tools } format for backwards compat', () => {
    const payload = { tools: SAMPLE_TOOLS }; // old/legacy format
    const received = receiveToolsFixed(payload);
    expect(received).toHaveLength(2);
  });

  it('FIX: receiver returns empty array for null/undefined payload', () => {
    expect(receiveToolsFixed(null)).toHaveLength(0);
    expect(receiveToolsFixed(undefined)).toHaveLength(0);
    expect(receiveToolsFixed({})).toHaveLength(0);
  });

  it('FIX: receiver returns empty array for empty tools', () => {
    expect(receiveToolsFixed([])).toHaveLength(0);
    expect(receiveToolsFixed({ tools: [] })).toHaveLength(0);
  });
});
