import type { jsonSchemaValidator, JsonSchemaType, JsonSchemaValidator } from '@modelcontextprotocol/sdk/validation/types.js';

/**
 * CSP-safe no-op JSON Schema validator for use in Chrome Manifest V3 extensions.
 *
 * Chrome MV3 blocks `new Function()` (unsafe-eval CSP). The MCP SDK's default
 * `AjvJsonSchemaValidator` uses AJV which calls `new Function()` at runtime to
 * compile schema validators. This throws an EvalError inside `listTools()`,
 * causing the entire tool-fetch to fail silently and return 0 tools.
 *
 * This validator satisfies the `jsonSchemaValidator` interface required by the
 * MCP SDK `Client` constructor option, but skips actual schema validation.
 * The extension only lists and calls tools — it does not need output schema
 * validation for security purposes.
 *
 * Usage:
 *   new Client(info, { capabilities: {}, jsonSchemaValidator: new NoOpJsonSchemaValidator() })
 */
export class NoOpJsonSchemaValidator implements jsonSchemaValidator {
  getValidator<T>(schema: JsonSchemaType): JsonSchemaValidator<T> {
    return (input: unknown) => ({
      valid: true as const,
      data: input as T,
      errorMessage: undefined,
    });
  }
}
