/**
 * TDD: Tests for NoOpJsonSchemaValidator — the CSP-safe replacement for AJV.
 *
 * Root cause: Chrome Manifest V3 blocks new Function() (unsafe-eval CSP).
 * AJV uses new Function() to compile JSON schema validators at runtime.
 * The MCP SDK's Client constructor defaults to AjvJsonSchemaValidator.
 * When listTools() is called, AJV compilation throws EvalError → tools stay empty.
 *
 * Fix: pass a NoOpJsonSchemaValidator to Client({ jsonSchemaValidator: ... })
 * which always passes validation (we don't need output schema validation in this extension).
 */

import { describe, it, expect } from 'vitest';
import { NoOpJsonSchemaValidator } from '../NoOpJsonSchemaValidator.js';

describe('NoOpJsonSchemaValidator', () => {
  it('returns a validator that marks any value as valid', () => {
    const v = new NoOpJsonSchemaValidator();
    const validate = v.getValidator({ type: 'object' });
    const result = validate({ foo: 'bar' });
    expect(result.valid).toBe(true);
    expect(result.errorMessage).toBeUndefined();
  });

  it('passes the input through as data', () => {
    const v = new NoOpJsonSchemaValidator();
    const validate = v.getValidator({ type: 'string' });
    const input = 'hello';
    const result = validate(input);
    expect(result.data).toBe(input);
  });

  it('marks null as valid (never throws)', () => {
    const v = new NoOpJsonSchemaValidator();
    const validate = v.getValidator({});
    expect(validate(null).valid).toBe(true);
  });

  it('marks array as valid even for object schema (no-op)', () => {
    const v = new NoOpJsonSchemaValidator();
    const validate = v.getValidator({ type: 'object', required: ['id'] });
    expect(validate([]).valid).toBe(true);
  });

  it('creates a fresh validator per getValidator call (no shared state)', () => {
    const v = new NoOpJsonSchemaValidator();
    const v1 = v.getValidator({ type: 'string' });
    const v2 = v.getValidator({ type: 'number' });
    expect(v1('foo').valid).toBe(true);
    expect(v2(42).valid).toBe(true);
  });
});
