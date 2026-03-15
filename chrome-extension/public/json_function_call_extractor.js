/**
 * JSON Function Call Extractor - Converts inline JSON patterns to <pre> elements
 * Monitors page for JSON function call patterns and wraps them in <pre> tags
 */
(function() {
  'use strict';
  
  // Prevent multiple script execution
  if (window.JSONFunctionExtractorExecuted) {
    console.debug('[json-extractor] Script already executed, skipping duplicate execution');
    return;
  }
  window.JSONFunctionExtractorExecuted = true;
  console.debug('[json-extractor] Script execution started');
  
  const CONFIG = {
    debug: true,
    processedAttribute: 'data-json-extracted',
  };
  
  // Track processed elements to avoid re-processing
  const processedElements = new WeakSet();
  let observer = null;
  let scanIntervalId = null;
  
  /**
   * Regex pattern to match JSON function call blocks
   * Matches from ```json to the last closing brace of function_call_end
   */
  const JSON_FUNCTION_PATTERN = /```json\s*\n((?:\{"type":\s*"(?:function_call_start|description|parameter|function_call_end)"[^}]*\}\s*\n?)+)/gi;
  
  /**
   * More lenient pattern for inline JSON without code fences
   * Matches sequences of JSON objects with function call types
   */
  const INLINE_JSON_PATTERN = /(\{"type":\s*"function_call_start"[^}]*\}(?:\s*\n?\s*\{"type":\s*"(?:description|parameter|function_call_end)"[^}]*\})*)/gi;
  
  /**
   * Extract JSON objects from text using regex
   */
  function extractJSONObjects(content) {
    const results = [];
    
    // Pattern to match individual JSON objects with function call types
    const objectPattern = /\{(?:[^{}"]|"(?:\\.|[^"\\])*")*?"type"\s*:\s*"(function_call_start|function_call_end|description|parameter)"(?:[^{}"]|"(?:\\.|[^"\\])*")*?\}/g;
    
    let match;
    while ((match = objectPattern.exec(content)) !== null) {
      const rawJSON = match[0];
      const typeValue = match[1];
      
      try {
        const parsed = JSON.parse(rawJSON);
        results.push({ raw: rawJSON, parsed, type: typeValue });
      } catch (e) {
        if (CONFIG.debug) {
          console.debug('[json-extractor] Failed to parse JSON object:', rawJSON.substring(0, 50));
        }
      }
    }
    
    return results;
  }
  
  /**
   * Validate if extracted JSON is a complete function call
   */
  function isCompleteFunctionCall(objects) {
    let hasStart = false;
    let hasEnd = false;
    
    for (const obj of objects) {
      if (obj.type === 'function_call_start') hasStart = true;
      if (obj.type === 'function_call_end') hasEnd = true;
    }
    
    return hasStart && hasEnd;
  }
  
  /**
   * Create a <pre> element safely without innerHTML
   */
  function createPreElement(jsonContent) {
    const pre = document.createElement('pre');
    pre.className = 'json-function-call';
    pre.setAttribute('data-extracted', 'true');
    
    // Use textContent instead of innerHTML to avoid Trusted Types issues
    // const content = '```json\n' + jsonContent.trim() + '\n```';
    const content = '```\n' + jsonContent.trim() + '\n```';
    pre.textContent = content;
    
    return pre;
  }
  
  /**
   * Process a text node and extract JSON function calls
   */
  function processTextNode(node) {
    if (!node || node.nodeType !== Node.TEXT_NODE) return false;
    
    const text = node.textContent || '';
    if (!text.trim()) return false;
    
    // Quick check for JSON patterns
    if (!text.includes('"type"') || !text.includes('function_call')) {
      return false;
    }
    
    const parent = node.parentNode;
    if (!parent) return false;
    
    let modified = false;
    const replacements = [];
    
    // First, try to match code-fenced JSON blocks
    JSON_FUNCTION_PATTERN.lastIndex = 0;
    let match;
    
    while ((match = JSON_FUNCTION_PATTERN.exec(text)) !== null) {
      const jsonContent = match[1];
      const objects = extractJSONObjects(jsonContent);
      
      if (objects.length > 0 && isCompleteFunctionCall(objects)) {
        replacements.push({
          start: match.index,
          end: match.index + match[0].length,
          jsonContent: jsonContent,
          fullMatch: match[0],
          type: 'code-fenced'
        });
        
        if (CONFIG.debug) {
          console.debug('[json-extractor] Found code-fenced JSON block with', objects.length, 'objects');
        }
      }
    }
    
    // Second, try to match inline JSON blocks (without code fences) if no code-fenced found
    if (replacements.length === 0) {
      INLINE_JSON_PATTERN.lastIndex = 0;
      
      while ((match = INLINE_JSON_PATTERN.exec(text)) !== null) {
        const jsonContent = match[1];
        const objects = extractJSONObjects(jsonContent);
        
        if (objects.length > 0 && isCompleteFunctionCall(objects)) {
          replacements.push({
            start: match.index,
            end: match.index + match[0].length,
            jsonContent: jsonContent,
            fullMatch: match[0],
            type: 'inline'
          });
          
          if (CONFIG.debug) {
            console.debug('[json-extractor] Found inline JSON block with', objects.length, 'objects');
          }
        }
      }
    }
    
    // Process replacements in reverse order to maintain correct indices
    if (replacements.length > 0) {
      replacements.sort((a, b) => b.start - a.start);
      
      // Split the text node at replacement points
      let currentNode = node;
      let currentOffset = 0;
      
      for (const replacement of replacements.reverse()) {
        const beforeText = text.substring(currentOffset, replacement.start);
        const afterOffset = replacement.end;
        
        // Create text node for content before JSON
        if (beforeText) {
          const beforeNode = document.createTextNode(beforeText);
          parent.insertBefore(beforeNode, currentNode);
        }
        
        // Create <pre> element for JSON
        const preElement = createPreElement(replacement.jsonContent);
        parent.insertBefore(preElement, currentNode);
        
        currentOffset = afterOffset;
        modified = true;
        
        if (CONFIG.debug) {
          console.debug('[json-extractor] Inserted <pre> element for', replacement.type, 'JSON block');
        }
      }
      
      // Handle remaining text after last replacement
      const remainingText = text.substring(currentOffset);
      if (remainingText) {
        const remainingNode = document.createTextNode(remainingText);
        parent.insertBefore(remainingNode, currentNode);
      }
      
      // Remove original text node
      parent.removeChild(currentNode);
    }
    
    return modified;
  }
  
  /**
   * Fallback: extract JSON using the element's full textContent.
   * M365 Copilot can render markdown inside JSON parameter values as real HTML
   * (e.g., <em>word</em>), which splits a single logical text node into multiple
   * DOM nodes. Per-text-node processing then fails because no single node contains
   * both function_call_start and function_call_end. This fallback reads the
   * combined textContent (which strips all HTML tags) and replaces the element's
   * content with a clean <pre class="json-function-call"> element.
   */
  function processElementByTextContent(element) {
    const text = element.textContent || '';
    if (!text.includes('"type"') || !text.includes('function_call')) {
      return false;
    }

    INLINE_JSON_PATTERN.lastIndex = 0;
    const match = INLINE_JSON_PATTERN.exec(text);
    if (!match) return false;

    const jsonContent = match[1];
    const objects = extractJSONObjects(jsonContent);
    if (objects.length === 0 || !isCompleteFunctionCall(objects)) return false;

    // Rebuild clean JSON from parsed objects (strips any HTML artifacts in values)
    const cleanJson = objects.map(o => JSON.stringify(o.parsed)).join('\n');
    const pre = createPreElement(cleanJson);

    // Replace element content with the <pre> block
    while (element.firstChild) element.removeChild(element.firstChild);
    element.appendChild(pre);

    element.setAttribute(CONFIG.processedAttribute, 'true');
    processedElements.add(element);

    if (CONFIG.debug) {
      console.debug('[json-extractor] Processed element via textContent fallback:', element.tagName, element.className);
    }

    return true;
  }

  /**
   * Process an element and its text nodes
   */
  function processElement(element) {
    // Skip if already processed
    if (processedElements.has(element)) {
      return false;
    }

    // Skip <pre>, <code>, <script>, <style> elements
    const tagName = element.tagName?.toLowerCase();
    if (['pre', 'code', 'script', 'style'].includes(tagName)) {
      return false;
    }

    // Skip elements that already have extracted JSON
    if (element.hasAttribute(CONFIG.processedAttribute)) {
      return false;
    }

    let modified = false;

    // Get all direct text nodes (not nested)
    const textNodes = [];
    for (let i = 0; i < element.childNodes.length; i++) {
      const child = element.childNodes[i];
      if (child.nodeType === Node.TEXT_NODE) {
        textNodes.push(child);
      } else if (child.nodeType === Node.ELEMENT_NODE) {
        // Recursively process child elements
        const childTag = child.tagName?.toLowerCase();
        if (!['pre', 'code', 'script', 'style'].includes(childTag)) {
          if (processElement(child)) {
            modified = true;
          }
        }
      }
    }

    // Process each text node
    for (const textNode of textNodes) {
      if (processTextNode(textNode)) {
        modified = true;
      }
    }

    // Fallback: if no modification yet, try matching against the full textContent.
    // This handles cases where M365 renders inline HTML (e.g. <em>, <strong>) inside
    // JSON parameter values, which splits the logical text across multiple DOM nodes.
    if (!modified) {
      modified = processElementByTextContent(element);
    }

    if (modified) {
      element.setAttribute(CONFIG.processedAttribute, 'true');
      processedElements.add(element);

      if (CONFIG.debug) {
        console.debug('[json-extractor] Processed element:', element.tagName, element.className);
      }
    }

    return modified;
  }
  
  /**
   * Scan known AI response containers for JSON patterns.
   * Supports multiple site DOM structures:
   *  - <message-content>: custom element used by some chat UIs
   *  - [data-testid="markdown-reply"]: M365 Copilot's AI response container
   */
  function scanDocument() {
    if (CONFIG.debug) {
      console.debug('[json-extractor] Starting document scan');
    }

    let processedCount = 0;

    const messageContentElements = document.querySelectorAll(
      'message-content, [data-testid="markdown-reply"]'
    );

    if (CONFIG.debug) {
      console.debug('[json-extractor] Found', messageContentElements.length, 'AI response container elements');
    }
    
    for (const element of messageContentElements) {
      try {
        if (processElement(element)) {
          processedCount++;
        }
      } catch (error) {
        if (CONFIG.debug) {
          console.warn('[json-extractor] Error processing element:', error);
        }
      }
    }
    
    if (CONFIG.debug && processedCount > 0) {
      console.debug('[json-extractor] Scan complete. Processed', processedCount, 'message-content elements');
    }
  }
  
  /**
   * Setup MutationObserver to watch for new content
   */
  function setupMutationObserver() {
    if (observer) return;
    
    observer = new MutationObserver((mutations) => {
      let shouldScan = false;
      
      for (const mutation of mutations) {
        if (mutation.type === 'childList') {
          for (const node of mutation.addedNodes) {
            if (node.nodeType === Node.ELEMENT_NODE) {
              shouldScan = true;
              break;
            }
          }
        }
        
        if (mutation.type === 'characterData') {
          shouldScan = true;
        }
        
        if (shouldScan) break;
      }
      
      if (shouldScan) {
        requestAnimationFrame(scanDocument);
      }
    });
    
    observer.observe(document.body, {
      childList: true,
      subtree: true,
      characterData: true,
    });
    
    if (CONFIG.debug) {
      console.debug('[json-extractor] MutationObserver started');
    }
  }
  
  /**
   * Start monitoring
   */
  function startMonitoring() {
    console.log('[json-extractor] Starting JSON function call extraction');
    
    // Initial scan
    scanDocument();
    
    // Setup mutation observer
    setupMutationObserver();
    
    console.log('[json-extractor] Monitoring active (mutation observer only)');
  }
  
  /**
   * Stop monitoring
   */
  function stopMonitoring() {
    if (observer) {
      observer.disconnect();
      observer = null;
    }
    
    if (scanIntervalId) {
      clearInterval(scanIntervalId);
      scanIntervalId = null;
    }
    
    console.log('[json-extractor] Monitoring stopped');
  }
  
  /**
   * Force scan now
   */
  function forceScan() {
    console.log('[json-extractor] Force scanning document');
    scanDocument();
  }
  
  /**
   * Clear processed tracking
   */
  function clearProcessed() {
    // Remove all processed attributes
    const processed = document.querySelectorAll(`[${CONFIG.processedAttribute}]`);
    processed.forEach(el => el.removeAttribute(CONFIG.processedAttribute));
    
    console.log('[json-extractor] Cleared', processed.length, 'processed markers');
  }
  
  // Cleanup on page unload
  window.addEventListener('beforeunload', stopMonitoring);
  
  // Expose API
  window.JSONFunctionExtractor = {
    start: startMonitoring,
    stop: stopMonitoring,
    scan: forceScan,
    clearProcessed: clearProcessed,
    config: CONFIG,
    getExtractedElements: function() {
      return Array.from(document.querySelectorAll('pre.json-function-call[data-extracted="true"]'));
    },
  };
  
  // Auto-start
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', startMonitoring);
  } else {
    startMonitoring();
  }
  
  console.log('[json-extractor] API available at window.JSONFunctionExtractor');
  console.log('[json-extractor] Usage: JSONFunctionExtractor.scan() - Force scan');
  console.log('[json-extractor] Usage: JSONFunctionExtractor.stop() - Stop monitoring');
  console.log('[json-extractor] Usage: JSONFunctionExtractor.getExtractedElements() - Get all extracted <pre> elements');
  
})();