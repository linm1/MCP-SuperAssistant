import { BaseAdapterPlugin } from './base.adapter';
import type { AdapterCapability, PluginContext } from '../plugin-types';
import { createLogger } from '@extension/shared/lib/logger';

/**
 * Microsoft 365 Copilot Chat Adapter
 *
 * Supports:
 *   - m365.cloud.microsoft/chat  (direct access)
 *   - m365.cloud.microsoft.mcas.ms/chat  (MCAS proxy, e.g. Parexel enterprise)
 *
 * Injects the MCP popover button next to the "工具/Tools" button in the Fluent AI
 * chat composer toolbar. Text insertion uses document.execCommand on the Lexical editor.
 */

const logger = createLogger('M365CopilotAdapter');

export class M365CopilotAdapter extends BaseAdapterPlugin {
  readonly name = 'M365CopilotAdapter';
  readonly version = '1.0.0';
  readonly hostnames: (string | RegExp)[] = [
    'm365.cloud.microsoft',
    /m365\.cloud\.microsoft(\.mcas\.ms)?/,
  ];
  readonly capabilities: AdapterCapability[] = [
    'text-insertion',
    'form-submission',
    'dom-manipulation',
    'file-attachment',
  ];

  // Stable selectors using data-testid and fai- prefixed class names
  // These are Fluent AI design-system stable identifiers, unlike hashed Griffel classes
  private readonly selectors = {
    // Lexical contenteditable editor
    CHAT_INPUT: '#m365-chat-editor-target-element, [aria-label="傳送訊息給 Copilot"][contenteditable="true"], [aria-label="Send a message to Copilot"][contenteditable="true"], .fai-EditorInput__input[contenteditable="true"]',
    // Send/submit button (only visible when there is text in the editor)
    SUBMIT_BUTTON: '.fai-SendButton, .fai-ChatInput__send, button[type="submit"][aria-label="傳送"], button[type="submit"][aria-label="Send"]',
    // Container holding the + and Tools buttons (primary insertion target)
    BUTTON_INSERTION_CONTAINER: '.fai-ChatInput__attachments',
    // Tools (工具) button — insert MCP button directly after this
    TOOLS_BUTTON: '[data-testid="capability-picker-trigger-btn"]',
    // Plus (+) button
    PLUS_BUTTON: '[data-testid="PlusMenuButton"]',
    // The whole chat input wrapper
    CHAT_INPUT_WRAPPER: '.fai-ChatInput, #m365-chat-input-shared-wrapper',
    // Always-present hidden file input (id confirmed stable via DevTools)
    FILE_INPUT: '#upload-file-button',
  };

  // URL tracking for SPA navigation
  private lastUrl: string = '';
  private urlCheckInterval: NodeJS.Timeout | null = null;

  // State management
  private mcpPopoverContainer: HTMLElement | null = null;
  private mutationObserver: MutationObserver | null = null;
  private popoverCheckInterval: NodeJS.Timeout | null = null;

  // Setup state tracking (prevent duplicate setup)
  private storeEventListenersSetup: boolean = false;
  private domObserversSetup: boolean = false;
  private uiIntegrationSetup: boolean = false;

  // Instance tracking for debugging
  private static instanceCount = 0;
  private instanceId: number;

  // Style injection tracking
  private adapterStylesInjected: boolean = false;

  constructor() {
    super();
    M365CopilotAdapter.instanceCount++;
    this.instanceId = M365CopilotAdapter.instanceCount;
    logger.debug(`Instance #${this.instanceId} created. Total instances: ${M365CopilotAdapter.instanceCount}`);
  }

  async initialize(context: PluginContext): Promise<void> {
    if (this.currentStatus === 'initializing' || this.currentStatus === 'active') {
      this.context?.logger.warn(`M365 Copilot adapter instance #${this.instanceId} already initialized, skipping`);
      return;
    }

    await super.initialize(context);
    this.context.logger.debug(`Initializing M365 Copilot adapter instance #${this.instanceId}...`);

    this.lastUrl = window.location.href;
    this.setupUrlTracking();
    this.setupStoreEventListeners();
    this.injectM365ButtonStyles();
  }

  async activate(): Promise<void> {
    if (this.currentStatus === 'active') {
      this.context?.logger.warn(`M365 Copilot adapter instance #${this.instanceId} already active, skipping`);
      return;
    }

    await super.activate();
    this.context.logger.debug(`Activating M365 Copilot adapter instance #${this.instanceId}...`);

    this.injectM365ButtonStyles();
    this.setupDOMObservers();
    this.setupUIIntegration();

    this.context.eventBus.emit('adapter:activated', {
      pluginName: this.name,
      timestamp: Date.now(),
    });
  }

  async deactivate(): Promise<void> {
    if (this.currentStatus === 'inactive' || this.currentStatus === 'disabled') {
      this.context?.logger.warn('M365 Copilot adapter already inactive, skipping deactivation');
      return;
    }

    await super.deactivate();
    this.context.logger.debug('Deactivating M365 Copilot adapter...');

    this.cleanupUIIntegration();
    this.cleanupDOMObservers();

    this.storeEventListenersSetup = false;
    this.domObserversSetup = false;
    this.uiIntegrationSetup = false;

    this.context.eventBus.emit('adapter:deactivated', {
      pluginName: this.name,
      timestamp: Date.now(),
    });
  }

  async cleanup(): Promise<void> {
    await super.cleanup();
    this.context.logger.debug('Cleaning up M365 Copilot adapter...');

    if (this.urlCheckInterval) {
      clearInterval(this.urlCheckInterval);
      this.urlCheckInterval = null;
    }

    if (this.popoverCheckInterval) {
      clearInterval(this.popoverCheckInterval);
      this.popoverCheckInterval = null;
    }

    const styleElement = document.getElementById('mcp-m365-copilot-button-styles');
    if (styleElement) {
      styleElement.remove();
      this.adapterStylesInjected = false;
    }

    this.cleanupUIIntegration();
    this.cleanupDOMObservers();

    this.storeEventListenersSetup = false;
    this.domObserversSetup = false;
    this.uiIntegrationSetup = false;
  }

  /**
   * Insert text into the M365 Copilot chat input (Lexical editor)
   * Uses execCommand('insertText') which is confirmed working on this editor.
   */
  async insertText(text: string, options?: { targetElement?: HTMLElement }): Promise<boolean> {
    this.context.logger.debug(`Attempting to insert text: ${text.substring(0, 50)}${text.length > 50 ? '...' : ''}`);

    let targetElement: HTMLElement | null = options?.targetElement ?? null;

    if (!targetElement) {
      const selectors = this.selectors.CHAT_INPUT.split(', ');
      for (const selector of selectors) {
        targetElement = document.querySelector(selector.trim()) as HTMLElement;
        if (targetElement) {
          this.context.logger.debug(`Found chat input via: ${selector.trim()}`);
          break;
        }
      }
    }

    if (!targetElement) {
      this.context.logger.error('Could not find M365 Copilot chat input element');
      this.emitExecutionFailed('insertText', 'Chat input element not found');
      return false;
    }

    try {
      targetElement.focus();

      // Move cursor to end of content
      const selection = window.getSelection();
      if (selection) {
        const range = document.createRange();
        range.selectNodeContents(targetElement);
        range.collapse(false);
        selection.removeAllRanges();
        selection.addRange(range);
      }

      // Prepend newline if there is existing content
      const existingText = targetElement.textContent ?? '';
      const insertContent = existingText.trim() ? '\n' + text : text;

      const success = document.execCommand('insertText', false, insertContent);

      // Also fire an InputEvent to ensure Lexical's internal state updates
      targetElement.dispatchEvent(
        new InputEvent('input', {
          inputType: 'insertText',
          data: insertContent,
          bubbles: true,
          cancelable: true,
        })
      );

      this.emitExecutionCompleted('insertText', { text }, {
        success,
        existingLength: existingText.length,
        insertedLength: insertContent.length,
      });

      this.context.logger.debug(`Text inserted successfully (execCommand returned: ${success})`);
      return success;
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      this.context.logger.error(`Error inserting text: ${msg}`);
      this.emitExecutionFailed('insertText', msg);
      return false;
    }
  }

  /**
   * Submit the M365 Copilot chat form.
   * Prefers clicking the send button; falls back to Enter key.
   */
  async submitForm(options?: { formElement?: HTMLFormElement }): Promise<boolean> {
    this.context.logger.debug('Attempting to submit M365 Copilot chat input');

    const selectors = this.selectors.SUBMIT_BUTTON.split(', ');
    let submitButton: HTMLButtonElement | null = null;

    for (const selector of selectors) {
      submitButton = document.querySelector(selector.trim()) as HTMLButtonElement;
      if (submitButton) {
        this.context.logger.debug(`Found submit button via: ${selector.trim()}`);
        break;
      }
    }

    if (!submitButton) {
      // Fallback: dispatch Enter key on the editor
      this.context.logger.warn('Submit button not found, trying Enter key fallback');
      const editor = document.querySelector(this.selectors.CHAT_INPUT) as HTMLElement | null;
      if (editor) {
        editor.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', bubbles: true }));
        editor.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', code: 'Enter', bubbles: true }));
        this.emitExecutionCompleted('submitForm', {}, { success: true, method: 'enter-key' });
        return true;
      }
      this.emitExecutionFailed('submitForm', 'Submit button and editor not found');
      return false;
    }

    try {
      if (submitButton.disabled) {
        this.context.logger.warn('Submit button is disabled');
        this.emitExecutionFailed('submitForm', 'Submit button is disabled');
        return false;
      }

      const rect = submitButton.getBoundingClientRect();
      if (rect.width === 0 || rect.height === 0) {
        this.context.logger.warn('Submit button is not visible');
        this.emitExecutionFailed('submitForm', 'Submit button is not visible');
        return false;
      }

      submitButton.click();
      this.emitExecutionCompleted('submitForm', {}, { success: true, method: 'button-click' });
      this.context.logger.debug('M365 Copilot chat submitted successfully');
      return true;
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      this.context.logger.error(`Error submitting: ${msg}`);
      this.emitExecutionFailed('submitForm', msg);
      return false;
    }
  }

  /**
   * Attach a file to the M365 Copilot chat input.
   *
   * M365 keeps a hidden <input type="file" id="upload-file-button"> always in the DOM.
   * Confirmed via DevTools: setting .files via DataTransfer + firing 'change' triggers
   * the upload and shows the file attachment chip ("已完成上傳").
   *
   * Fallback: drag-drop simulation onto the chat input wrapper.
   */
  async attachFile(file: File, options?: { inputElement?: HTMLInputElement }): Promise<boolean> {
    this.context.logger.debug(`M365: attachFile called for ${file.name} (${file.size} bytes)`);

    if (!file || file.size === 0) {
      this.emitExecutionFailed('attachFile', 'Invalid file: empty or null');
      return false;
    }

    // Primary: inject directly into the always-present hidden file input
    const fileInput = (options?.inputElement ??
      document.querySelector(this.selectors.FILE_INPUT)) as HTMLInputElement | null;

    if (fileInput) {
      try {
        const dt = new DataTransfer();
        dt.items.add(file);
        fileInput.files = dt.files;
        fileInput.dispatchEvent(new Event('change', { bubbles: true }));
        fileInput.dispatchEvent(new Event('input', { bubbles: true }));
        this.context.logger.debug('M365: file injected via #upload-file-button');
        this.emitExecutionCompleted('attachFile', { fileName: file.name }, { success: true, method: 'file-input' });
        return true;
      } catch (e) {
        this.context.logger.warn('M365: direct file input failed, trying drag-drop fallback', e);
      }
    }

    // Fallback: drag-drop simulation onto the chat input area
    try {
      const dropZoneSelectors = this.selectors.CHAT_INPUT_WRAPPER.split(', ');
      let dropZone: HTMLElement | null = null;
      for (const sel of dropZoneSelectors) {
        dropZone = document.querySelector(sel.trim()) as HTMLElement | null;
        if (dropZone) break;
      }
      if (dropZone) {
        const dt = new DataTransfer();
        dt.items.add(file);
        dropZone.dispatchEvent(new DragEvent('dragenter', { bubbles: true, dataTransfer: dt }));
        dropZone.dispatchEvent(new DragEvent('dragover',  { bubbles: true, dataTransfer: dt }));
        dropZone.dispatchEvent(new DragEvent('drop',      { bubbles: true, dataTransfer: dt }));
        this.context.logger.debug('M365: file drop simulated on chat input wrapper');
        this.emitExecutionCompleted('attachFile', { fileName: file.name }, { success: true, method: 'drag-drop' });
        return true;
      }
    } catch (e) {
      this.context.logger.warn('M365: drag-drop fallback failed', e);
    }

    this.emitExecutionFailed('attachFile', 'File attachment failed: input not found and drag-drop failed');
    return false;
  }

  /**
   * Check if the current page is supported by this adapter.
   * Handles both direct and MCAS proxy hostnames.
   */
  isSupported(): boolean {
    const currentHost = window.location.hostname;
    const currentPath = window.location.pathname;

    const isM365Host = this.hostnames.some(pattern => {
      if (typeof pattern === 'string') return currentHost.includes(pattern);
      return (pattern as RegExp).test(currentHost);
    });

    if (!isM365Host) return false;

    // Only activate on the /chat path
    const isChatPage = currentPath.startsWith('/chat');
    this.context?.logger.debug(`M365 adapter support check: host=${currentHost}, path=${currentPath}, supported=${isChatPage}`);
    return isChatPage;
  }

  // ── Private helpers ───────────────────────────────────────────────────────

  private setupUrlTracking(): void {
    if (!this.urlCheckInterval) {
      this.urlCheckInterval = setInterval(() => {
        const currentUrl = window.location.href;
        if (currentUrl !== this.lastUrl) {
          this.context.logger.debug(`URL changed: ${this.lastUrl} → ${currentUrl}`);
          if (this.onPageChanged) {
            this.onPageChanged(currentUrl, this.lastUrl);
          }
          this.lastUrl = currentUrl;
        }
      }, 1000);
    }
  }

  private setupStoreEventListeners(): void {
    if (this.storeEventListenersSetup) return;

    this.context.eventBus.on('tool:execution-completed', (data) => {
      this.context.logger.debug('Tool execution completed:', data);
      this.handleToolExecutionCompleted(data);
    });

    this.context.eventBus.on('ui:sidebar-toggle', (data) => {
      this.context.logger.debug('Sidebar toggled:', data);
    });

    this.storeEventListenersSetup = true;
  }

  private setupDOMObservers(): void {
    if (this.domObserversSetup) return;

    this.context.logger.debug(`Setting up DOM observers for M365 Copilot adapter instance #${this.instanceId}`);

    this.mutationObserver = new MutationObserver((mutations) => {
      let shouldReinject = false;
      for (const mutation of mutations) {
        if (mutation.type === 'childList' && !document.getElementById('mcp-popover-container')) {
          shouldReinject = true;
          break;
        }
      }
      if (shouldReinject) {
        const insertionPoint = this.findButtonInsertionPoint();
        if (insertionPoint) {
          this.context.logger.debug('MCP popover removed, re-injecting');
          this.setupUIIntegration();
        }
      }
    });

    this.mutationObserver.observe(document.body, { childList: true, subtree: true });
    this.domObserversSetup = true;
  }

  private setupUIIntegration(): void {
    if (this.uiIntegrationSetup) {
      this.context.logger.debug(`UI integration already set up for instance #${this.instanceId}, re-injecting`);
    } else {
      this.context.logger.debug(`Setting up UI integration for M365 Copilot adapter instance #${this.instanceId}`);
      this.uiIntegrationSetup = true;
    }

    this.waitForPageReady()
      .then(() => this.injectMCPPopoverWithRetry())
      .catch((error) => this.context.logger.warn('Failed to wait for page ready:', error));
  }

  private async waitForPageReady(): Promise<void> {
    return new Promise((resolve, reject) => {
      let attempts = 0;
      const maxAttempts = 20; // 10 seconds at 500ms

      const checkReady = () => {
        attempts++;
        if (this.findButtonInsertionPoint()) {
          this.context.logger.debug('Page ready for MCP popover injection');
          resolve();
        } else if (attempts >= maxAttempts) {
          reject(new Error('No insertion point found after maximum attempts'));
        } else {
          setTimeout(checkReady, 500);
        }
      };
      setTimeout(checkReady, 200);
    });
  }

  private injectMCPPopoverWithRetry(maxRetries: number = 5): void {
    const attemptInjection = (attempt: number) => {
      this.context.logger.debug(`MCP popover injection attempt ${attempt}/${maxRetries}`);

      if (document.getElementById('mcp-popover-container')) {
        this.context.logger.debug('MCP popover already exists');
        return;
      }

      const insertionPoint = this.findButtonInsertionPoint();
      if (insertionPoint) {
        this.injectMCPPopover(insertionPoint);
      } else if (attempt < maxRetries) {
        this.context.logger.debug(`Insertion point not found, retrying in 1s (attempt ${attempt}/${maxRetries})`);
        setTimeout(() => attemptInjection(attempt + 1), 1000);
      } else {
        this.context.logger.warn('Failed to inject MCP popover after maximum retries');
      }
    };
    attemptInjection(1);
  }

  /**
   * Find where to inject the MCP popover button.
   *
   * Strategy 1 (PRIMARY): Insert after the Tools button
   *   [+] [工具/Tools] [MCP←here] [attachment-list-controls]
   *   Container: .fai-ChatInput__attachments
   *   Anchor:    [data-testid="capability-picker-trigger-btn"]
   *
   * Strategy 2 (FALLBACK): Append to .fai-ChatInput__attachments
   *
   * Strategy 3 (FINAL FALLBACK): Append to .fai-ChatInput wrapper
   */
  private findButtonInsertionPoint(): { container: Element; insertAfter: Element | null } | null {
    this.context.logger.debug('Finding button insertion point for MCP popover');

    // Strategy 1: After the Tools button
    const toolsBtn = document.querySelector(this.selectors.TOOLS_BUTTON);
    if (toolsBtn && toolsBtn.parentElement) {
      this.context.logger.debug('Strategy 1: found Tools button, inserting after it');
      return { container: toolsBtn.parentElement, insertAfter: toolsBtn };
    }

    // Strategy 2: Append to attachments container
    const attachmentsContainer = document.querySelector(this.selectors.BUTTON_INSERTION_CONTAINER);
    if (attachmentsContainer) {
      this.context.logger.debug('Strategy 2: appending to .fai-ChatInput__attachments');
      return { container: attachmentsContainer, insertAfter: null };
    }

    // Strategy 3: Append to root chat input wrapper
    const chatInputRoot = document.querySelector(this.selectors.CHAT_INPUT_WRAPPER);
    if (chatInputRoot) {
      this.context.logger.debug('Strategy 3: appending to .fai-ChatInput root');
      return { container: chatInputRoot, insertAfter: null };
    }

    this.context.logger.debug('Could not find suitable insertion point for MCP popover');
    return null;
  }

  private injectMCPPopover(insertionPoint: { container: Element; insertAfter: Element | null }): void {
    this.context.logger.debug('Injecting MCP popover into M365 Copilot interface');

    try {
      if (document.getElementById('mcp-popover-container')) {
        this.context.logger.debug('MCP popover already exists, skipping injection');
        return;
      }

      const reactContainer = document.createElement('div');
      reactContainer.id = 'mcp-popover-container';
      reactContainer.style.display = 'inline-flex';
      reactContainer.style.alignItems = 'center';
      reactContainer.style.margin = '0 4px';

      const { container, insertAfter } = insertionPoint;
      if (insertAfter && insertAfter.parentNode === container) {
        container.insertBefore(reactContainer, insertAfter.nextSibling);
        this.context.logger.debug('Inserted MCP popover container after Tools button');
      } else {
        container.appendChild(reactContainer);
        this.context.logger.debug('Appended MCP popover container to toolbar');
      }

      this.mcpPopoverContainer = reactContainer;
      this.renderMCPPopover(reactContainer);
      this.context.logger.debug('MCP popover injected and rendered successfully');
    } catch (error) {
      this.context.logger.error('Failed to inject MCP popover:', error);
    }
  }

  private renderMCPPopover(container: HTMLElement): void {
    this.context.logger.debug('Rendering MCP popover for M365 Copilot');

    try {
      import('react').then(React => {
        import('react-dom/client').then(ReactDOM => {
          import('../../components/mcpPopover/mcpPopover').then(({ MCPPopover }) => {
            if (!container || !container.isConnected) {
              this.context.logger.warn('Container disconnected before render');
              return;
            }

            const toggleStateManager = this.createToggleStateManager();

            // Microsoft Fluent Design-compatible button config
            const adapterButtonConfig = {
              className: 'mcp-m365-button-base',
              contentClassName: 'mcp-m365-button-content',
              textClassName: 'mcp-m365-button-text',
              activeClassName: 'mcp-button-active',
            };

            const root = ReactDOM.createRoot(container);
            root.render(
              React.createElement(MCPPopover, {
                toggleStateManager,
                adapterButtonConfig,
                adapterName: this.name,
              })
            );

            this.context.logger.debug('MCP popover rendered successfully with M365 Fluent styling');
          }).catch((error: unknown) => this.context.logger.error('Failed to import MCPPopover:', error));
        }).catch((error: unknown) => this.context.logger.error('Failed to import ReactDOM:', error));
      }).catch((error: unknown) => this.context.logger.error('Failed to import React:', error));
    } catch (error) {
      this.context.logger.error('Failed to render MCP popover:', error);
    }
  }

  private createToggleStateManager() {
    const context = this.context;

    const stateManager = {
      getState: () => {
        try {
          const uiState = context.stores.ui;
          const mcpEnabled = uiState?.mcpEnabled ?? false;
          const autoSubmitEnabled = uiState?.preferences?.autoSubmit ?? false;
          return {
            mcpEnabled,
            autoInsert: autoSubmitEnabled,
            autoSubmit: autoSubmitEnabled,
            autoExecute: false,
          };
        } catch {
          return { mcpEnabled: false, autoInsert: false, autoSubmit: false, autoExecute: false };
        }
      },

      setMCPEnabled: (enabled: boolean) => {
        context.logger.debug(`Setting MCP ${enabled ? 'enabled' : 'disabled'}`);
        try {
          if (context.stores.ui?.setMCPEnabled) {
            context.stores.ui.setMCPEnabled(enabled, 'mcp-popover-toggle');
          } else if (context.stores.ui?.setSidebarVisibility) {
            context.stores.ui.setSidebarVisibility(enabled, 'mcp-popover-toggle-fallback');
          }

          const sidebarManager = (window as any).activeSidebarManager;
          if (sidebarManager) {
            if (enabled) {
              sidebarManager.show().catch((e: any) => context.logger.error('Error showing sidebar:', e));
            } else {
              sidebarManager.hide().catch((e: any) => context.logger.error('Error hiding sidebar:', e));
            }
          }
        } catch (error) {
          context.logger.error('Error in setMCPEnabled:', error);
        }
        stateManager.updateUI();
      },

      setAutoInsert: (enabled: boolean) => {
        context.stores.ui?.updatePreferences?.({ autoSubmit: enabled });
        stateManager.updateUI();
      },

      setAutoSubmit: (enabled: boolean) => {
        context.stores.ui?.updatePreferences?.({ autoSubmit: enabled });
        stateManager.updateUI();
      },

      setAutoExecute: (_enabled: boolean) => {
        stateManager.updateUI();
      },

      updateUI: () => {
        const popoverContainer = document.getElementById('mcp-popover-container');
        if (popoverContainer) {
          const currentState = stateManager.getState();
          popoverContainer.dispatchEvent(
            new CustomEvent('mcp:update-toggle-state', { detail: { toggleState: currentState } })
          );
        }
      },
    };

    return stateManager;
  }

  private cleanupDOMObservers(): void {
    if (this.mutationObserver) {
      this.mutationObserver.disconnect();
      this.mutationObserver = null;
    }
    this.domObserversSetup = false;
  }

  private cleanupUIIntegration(): void {
    const popoverContainer = document.getElementById('mcp-popover-container');
    if (popoverContainer) {
      popoverContainer.remove();
    }
    this.mcpPopoverContainer = null;
  }

  private handleToolExecutionCompleted(data: any): void {
    if (!this.shouldHandleEvents()) return;
    this.context.logger.debug('Tool execution handled by M365 Copilot adapter:', data);
  }

  private emitExecutionCompleted(toolName: string, parameters: any, result: any): void {
    this.context.eventBus.emit('tool:execution-completed', {
      execution: {
        id: this.generateCallId(),
        toolName,
        parameters,
        result,
        timestamp: Date.now(),
        status: 'success',
      },
    });
  }

  private emitExecutionFailed(toolName: string, error: string): void {
    this.context.eventBus.emit('tool:execution-failed', {
      toolName,
      error,
      callId: this.generateCallId(),
    });
  }

  private generateCallId(): string {
    return `m365-copilot-${Date.now()}-${Math.random().toString(36).substring(2, 11)}`;
  }

  /**
   * Inject Microsoft Fluent Design System-compatible styles for the MCP button.
   * Matches the look of existing Fluent UI buttons in the M365 Copilot toolbar.
   */
  private injectM365ButtonStyles(): void {
    if (this.adapterStylesInjected) return;

    try {
      const styleId = 'mcp-m365-copilot-button-styles';
      const existing = document.getElementById(styleId);
      if (existing) existing.remove();

      const styleElement = document.createElement('style');
      styleElement.id = styleId;
      styleElement.textContent = this.getM365ButtonStyles();
      document.head.appendChild(styleElement);

      this.adapterStylesInjected = true;
      this.context.logger.debug('M365 Copilot button styles injected');
    } catch (error) {
      this.context.logger.error('Failed to inject M365 button styles:', error);
    }
  }

  private getM365ButtonStyles(): string {
    return `
/* M365 Copilot MCP Button Styles - Microsoft Fluent Design System */
.mcp-m365-button-base {
  position: relative;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  box-sizing: border-box;
  min-width: 32px;
  height: 32px;
  padding: 0 8px;
  gap: 4px;
  border: none;
  border-radius: 4px;
  background: transparent;
  color: var(--colorNeutralForeground1, #242424);
  font-family: "Segoe UI", -apple-system, BlinkMacSystemFont, Roboto, "Helvetica Neue", sans-serif;
  font-size: 14px;
  font-weight: 400;
  line-height: 20px;
  cursor: pointer;
  user-select: none;
  white-space: nowrap;
  text-decoration: none;
  transition: background 80ms ease, color 80ms ease;
  -webkit-tap-highlight-color: transparent;
  outline: none;
  vertical-align: middle;
}

.mcp-m365-button-base:hover {
  background: var(--colorNeutralBackground1Hover, rgba(0, 0, 0, 0.06));
}

.mcp-m365-button-base:active {
  background: var(--colorNeutralBackground1Pressed, rgba(0, 0, 0, 0.12));
}

.mcp-m365-button-base:focus-visible {
  outline: 2px solid var(--colorBrandStroke1, #0078d4);
  outline-offset: 2px;
}

.mcp-m365-button-base.mcp-button-active {
  color: var(--colorBrandForeground1, #0078d4);
  background: var(--colorBrandBackground2, rgba(0, 120, 212, 0.08));
}

.mcp-m365-button-base.mcp-button-active:hover {
  background: var(--colorBrandBackground2Hover, rgba(0, 120, 212, 0.14));
}

.mcp-m365-button-content {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 4px;
}

.mcp-m365-button-text {
  font-size: 14px;
  font-weight: 400;
  line-height: 20px;
}

/* Dark mode */
@media (prefers-color-scheme: dark) {
  .mcp-m365-button-base {
    color: var(--colorNeutralForeground1, #ffffff);
  }

  .mcp-m365-button-base:hover {
    background: var(--colorNeutralBackground1Hover, rgba(255, 255, 255, 0.08));
  }

  .mcp-m365-button-base:active {
    background: var(--colorNeutralBackground1Pressed, rgba(255, 255, 255, 0.14));
  }

  .mcp-m365-button-base.mcp-button-active {
    color: var(--colorBrandForeground1, #479ef5);
    background: var(--colorBrandBackground2, rgba(71, 158, 245, 0.12));
  }

  .mcp-m365-button-base.mcp-button-active:hover {
    background: var(--colorBrandBackground2Hover, rgba(71, 158, 245, 0.18));
  }
}

/* Ensure proper vertical alignment in Fluent toolbar */
.fai-ChatInput__attachments .mcp-m365-button-base {
  align-self: center;
  margin: 0 2px;
}

/* High contrast mode */
@media (prefers-contrast: high) {
  .mcp-m365-button-base {
    border: 1px solid ButtonText;
  }
  .mcp-m365-button-base:focus-visible {
    outline-width: 3px;
  }
}

/* Reduced motion */
@media (prefers-reduced-motion: reduce) {
  .mcp-m365-button-base {
    transition: none;
  }
}
`;
  }

  // ── Event handlers ────────────────────────────────────────────────────────

  onPageChanged?(url: string, oldUrl?: string): void {
    this.context.logger.debug(`M365 Copilot page changed: ${oldUrl} → ${url}`);
    this.lastUrl = url;

    if (this.isSupported()) {
      setTimeout(() => this.injectM365ButtonStyles(), 300);
      setTimeout(() => this.setupUIIntegration(), 800);
    }

    this.context.eventBus.emit('app:site-changed', {
      site: url,
      hostname: window.location.hostname,
    });
  }

  onHostChanged?(newHost: string, oldHost?: string): void {
    this.context.logger.debug(`M365 Copilot host changed: ${oldHost} → ${newHost}`);
  }

  // ── Public utilities ──────────────────────────────────────────────────────

  public injectMCPPopoverManually(): void {
    this.context.logger.debug('Manual MCP popover injection requested');
    this.injectMCPPopoverWithRetry();
  }

  public isMCPPopoverInjected(): boolean {
    return !!document.getElementById('mcp-popover-container');
  }
}
