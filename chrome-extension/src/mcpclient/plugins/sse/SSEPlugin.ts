import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { SSEClientTransport } from '@modelcontextprotocol/sdk/client/sse.js';
import type { Transport } from '@modelcontextprotocol/sdk/shared/transport.js';
import { LoggingMessageNotificationSchema } from '@modelcontextprotocol/sdk/types.js';

import type { ITransportPlugin, PluginMetadata, PluginConfig } from '../../types/plugin.js';
import type { SSEPluginConfig } from '../../types/config.js';
import { createLogger } from '@extension/shared/lib/logger';


const logger = createLogger('SSEPlugin');

export class SSEPlugin implements ITransportPlugin {
  readonly metadata: PluginMetadata = {
    name: 'SSE Transport Plugin',
    version: '1.0.0',
    transportType: 'sse',
    description: 'Server-Sent Events transport for MCP protocol',
    author: 'MCP SuperAssistant',
  };

  private config: SSEPluginConfig = {};
  private transport: Transport | null = null;
  private isConnectedFlag: boolean = false;
  private connectionPromise: Promise<Transport> | null = null;

  async initialize(config: PluginConfig): Promise<void> {
    this.config = {
      keepAlive: true,
      connectionTimeout: 5000,
      readTimeout: 30000,
      ...config,
    } as SSEPluginConfig;

    logger.debug(`Initialized with config:`, this.config);
  }

  async connect(uri: string): Promise<Transport> {
    logger.debug(`[SSEPlugin] Creating transport for: ${uri}`);
    logger.debug(`[SSEPlugin] Environment check - fetch available: ${typeof fetch !== 'undefined'}, EventSource available: ${typeof EventSource !== 'undefined'}`);

    try {
      const transport = await this.createConnection(uri);
      this.transport = transport;
      logger.debug('[SSEPlugin] Transport created successfully');
      return transport;
    } catch (error) {
      logger.error('[SSEPlugin] Transport creation failed:', {
        errorType: error instanceof Error ? error.constructor.name : typeof error,
        errorMessage: error instanceof Error ? error.message : String(error),
        uri,
      });
      throw error;
    }
  }

  private async createConnection(uri: string): Promise<Transport> {
    try {
      // Validate and parse URI
      const url = new URL(uri);
      logger.debug(`[SSEPlugin] Creating SSE transport for: ${url.toString()}`);

      // Pre-flight check: verify server is reachable before creating transport
      try {
        const preflight = await fetch(url.toString(), {
          method: 'GET',
          headers: { 'Accept': 'text/event-stream' },
          signal: AbortSignal.timeout(5000),
        });
        logger.debug(`[SSEPlugin] Pre-flight check: status=${preflight.status}, content-type=${preflight.headers.get('content-type')}`);
        // Abort the preflight response body to free resources
        await preflight.body?.cancel();
      } catch (preflightError) {
        logger.error(`[SSEPlugin] Pre-flight fetch failed:`, {
          errorType: preflightError instanceof Error ? preflightError.constructor.name : typeof preflightError,
          errorMessage: preflightError instanceof Error ? preflightError.message : String(preflightError),
        });
        throw new Error(`SSE Plugin: Server unreachable at ${uri}. Pre-flight check failed: ${preflightError instanceof Error ? preflightError.message : String(preflightError)}`);
      }

      // Create SSE transport
      const transport = new SSEClientTransport(url);

      // Return the transport without testing
      // The main client will handle the connection test
      logger.debug('[SSEPlugin] SSE transport created successfully');
      return transport;
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : String(error);

      // Enhanced error messages for SSE-specific issues
      let enhancedError = errorMessage;
      if (errorMessage.includes('404')) {
        enhancedError = 'SSE endpoint not found (404). Verify the server URL and SSE endpoint path.';
      } else if (errorMessage.includes('timeout')) {
        enhancedError = 'SSE connection timeout. The server may be slow or the endpoint may not support SSE.';
      } else if (errorMessage.includes('Failed to fetch')) {
        enhancedError = 'SSE connection failed. Check if the server is running and accessible.';
      }

      throw new Error(`SSE Plugin: ${enhancedError}`);
    }
  }

  async disconnect(): Promise<void> {
    logger.debug('[SSEPlugin] Disconnecting...');

    if (this.transport) {
      try {
        // SSE transport may not have explicit close method, but we clean up references
        if ('close' in this.transport && typeof this.transport.close === 'function') {
          await (this.transport as any).close();
        }
      } catch (error) {
        logger.warn('[SSEPlugin] Error during transport cleanup:', error);
      }
    }

    this.transport = null;
    this.isConnectedFlag = false;
    this.connectionPromise = null;

    logger.debug('[SSEPlugin] Disconnected');
  }

  isConnected(): boolean {
    // The plugin creates transports but doesn't manage connection state
    // Connection state is managed by the main client
    return this.transport !== null;
  }

  isSupported(uri: string): boolean {
    try {
      const url = new URL(uri);
      return url.protocol === 'http:' || url.protocol === 'https:';
    } catch {
      return false;
    }
  }

  getDefaultConfig(): PluginConfig {
    return {
      keepAlive: true,
      connectionTimeout: 5000,
      readTimeout: 30000,
      headers: {
        Accept: 'text/event-stream',
        'Cache-Control': 'no-cache',
      },
    };
  }

  async isHealthy(): Promise<boolean> {
    if (!this.isConnected() || !this.transport) {
      return false;
    }

    try {
      // For SSE, we can't easily ping the server, so we assume healthy if connected
      // In a real implementation, you might want to track last received event timestamp
      return true;
    } catch (error) {
      logger.warn('[SSEPlugin] Health check failed:', error);
      return false;
    }
  }

  async callTool(client: Client, toolName: string, args: any): Promise<any> {
    if (!this.isConnected()) {
      throw new Error('SSE Plugin: Not connected');
    }

    logger.debug(`Calling tool: ${toolName}`);

    try {
      const result = await client.callTool({ name: toolName, arguments: args });
      logger.debug(`Tool call completed: ${toolName}`);
      return result;
    } catch (error) {
      logger.error(`Tool call failed: ${toolName}`, error);
      throw error;
    }
  }

  async getPrimitives(client: Client): Promise<any[]> {
    if (!this.isConnected()) {
      throw new Error('SSE Plugin: Not connected');
    }

    logger.debug('[SSEPlugin] Getting primitives...');

    try {
      const capabilities = client.getServerCapabilities();
      logger.debug('[SSEPlugin] Server capabilities:', JSON.stringify(capabilities));

      const primitives: any[] = [];

      const promises: Promise<void>[] = [
        client.listTools()
          .then(({ tools }) => { tools.forEach(item => primitives.push({ type: 'tool', value: item })); })
          .catch(error => { logger.warn('[SSEPlugin] listTools() failed:', error instanceof Error ? error.message : String(error)); }),

        client.listResources()
          .then(({ resources }) => { resources.forEach(item => primitives.push({ type: 'resource', value: item })); })
          .catch(error => { logger.warn('[SSEPlugin] listResources() failed:', error instanceof Error ? error.message : String(error)); }),

        client.listPrompts()
          .then(({ prompts }) => { prompts.forEach(item => primitives.push({ type: 'prompt', value: item })); })
          .catch(error => { logger.warn('[SSEPlugin] listPrompts() failed:', error instanceof Error ? error.message : String(error)); }),
      ];

      await Promise.all(promises);
      logger.debug(`Retrieved ${primitives.length} primitives`);
      return primitives;
    } catch (error) {
      logger.error('[SSEPlugin] Failed to get primitives:', error);
      throw error;
    }
  }
}
