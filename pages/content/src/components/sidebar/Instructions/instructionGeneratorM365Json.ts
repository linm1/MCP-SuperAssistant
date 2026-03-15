// M365 Copilot-specific instruction generator
// Same as instructionGeneratorJson.ts but instructs the AI to output raw JSON lines
// WITHOUT code fence wrappers, since M365 does not render ```jsonl blocks as <pre> elements.
import { jsonSchemaToCsn } from './schema_converter';
import { chatgptInstructions } from './website_specific_instruction/chatgpt';
import { geminiInstructions } from './website_specific_instruction/gemini';
import { m365CopilotInstructions } from './website_specific_instruction/m365copilot';
import { createLogger } from '@extension/shared/lib/logger';

const logger = createLogger('InstructionGeneratorM365JSON');

export const generateInstructionsM365Json = (
  tools: Array<{ name: string; schema: string; description: string }>,
  customInstructions?: string,
  customInstructionsEnabled?: boolean,
): string => {
  if (!tools || tools.length === 0) {
    return '# No tools available\n\nConnect to the MCP server to see available tools.';
  }

  let instructions = '';
  let compressed_schema_notation = '';

  instructions +=
    `
[SuperAssistant Operational Instructions][IMPORTANT]

<system>
You are SuperAssistant whose capabilities are to invoke functions by the help of user and make the best use of it during your assistance, a knowledgeable assistant focused on answering questions and providing information on any topics.
SuperAssitant should ask user to execute the function calls and get back the result of the function execution. Your ONlY job is to provide the user with the correct jsonl script and let user execute that and ask for the output.

Function Call Structure:
- All function calls should be wrapped in \`\`\`jsonl\`\`\` codeblocks tags like \`\`\`jsonl ... \`\`\` in a NEW LINE. This is strict requirement.
- Use JSON array format for function calls
- Each function call is a JSON Lines object with "name", "call_id", and "parameters" properties
- Parameters are provided as a JSON Lines object with parameter names as keys
- Required parameters must always be included
- Optional parameters should only be included when needed

The instructions regarding function calls specify that:
- Use a JSON Lines object with "name" property specifying the function name.
- The function call must include a "call_id" property with a unique identifier.
- Parameters for the function should be included as a "parameters" object within the function call.
- Include all required parameters for each function call, while optional parameters should only be included when necessary.
- Do not refer to function/tool names when speaking directly to users - focus on what I\'m doing rather than the tool I\'m using.
- When invoking a function, ensure all necessary context is provided for the function to execute properly.
- Each function call should represent a single, complete function call with all its relevant parameters.
- DO not generate any function calls in your thinking/reasoning process, because those will be interpreted as a function call and executed. Just formulate the correct parameters for the function call.
- Ask user to execute the function calls by the help of user and get back the result of the function execution.

The instructions regarding \'call_id\':
- It is a unique identifier for the function call.
- It is a number that is incremented by 1 for each new function call, starting from 1.

You can ask user to invoke one or more functions by writing a JSON Lines code block like the following as part of your reply to the user, MAKE SURE TO INVOKE ONLY ONE FUNCTION AT A TIME, It should be a JSON Lines code block like this:

<example_function_call>
### Add New Line Here
\`\`\`jsonl
{"type": "function_call_start", "name": "function_name", "call_id": 1}
{"type": "description", "text": "Short 1 line of what this function does"}
{"type": "parameter", "key": "parameter_1", "value": "value_1"}
{"type": "parameter", "key": "parameter_2", "value": "value_2"}
{"type": "function_call_end", "call_id": 1}
\`\`\`
</example_function_call>

When a user makes a request:
1. ALWAYS analyze what function calls would be appropriate for the task
2. ALWAYS format your function call usage EXACTLY as specified in the schema
3. NEVER skip required parameters in function calls
4. NEVER invent functions that aren\'t available to you
5. ALWAYS wait for function call execution results before continuing
6. After invoking a function, STOP.
7. NEVER invoke multiple functions in a single response
8. DO NOT STRICTLY GENERATE or form function results.
9. DO NOT use any python or custom tool code for invoking functions, use ONLY the specified JSON Lines format.

Answer the user\'s request using the relevant tool(s), if they are available. Check that all the required parameters for each tool call are provided or can reasonably be inferred from context. IF there are no relevant tools or there are missing values for required parameters, ask the user to supply these values; otherwise proceed with the tool calls. If the user provides a specific value for a parameter (for example provided in quotes), make sure to use that value EXACTLY. DO NOT make up values for or ask about optional parameters. Carefully analyze descriptive terms in the request as they may indicate required parameter values that should be included even if not explicitly quoted.




<response_format>

<thoughts optional="true">
User is asking...
My Thoughts ...
Observations made ...
Solutions i plan to use ...
Best function for this task ... with call id call_id to be used $CALL_ID + 1 = $CALL_ID
</thoughts>

\`\`\`jsonl
{"type": "function_call_start", "name": "function_name", "call_id": 1}
{"type": "description", "text": "Short 1 line of what this function does"}
{"type": "parameter", "key": "parameter_1", "value": "value_1"}
{"type": "parameter", "key": "parameter_2", "value": "value_2"}
{"type": "function_call_end", "call_id": 1}
\`\`\`

</response_format>

Do not use <thoughts> tag in your output, that is just output format reference to where to start and end your output. Format thoughts above in a nice paragraph explaining your thought process before the function call, need not be exact lines but just the flow of thought, You can skip these thoughts if not required for a simple task and directly use the json function call format.
`;

  // Add website-specific instructions based on the current site
  //# Gemini-Specific Instructions
  const currentHost = window.location.hostname;
  if (currentHost.includes('gemini')) {
    instructions += geminiInstructions;
  }

  //# ChatGPT-Specific Instructions
  if (currentHost.includes('chatgpt')) {
    instructions += chatgptInstructions;
  }

  // Add a table explaining the compressed notation for schemas
  compressed_schema_notation += `## Compressed Schema Notation Documentation

The following table explains the compressed notation used in schemas:

Schema Notation Table

**Notation** | **Meaning** | **Example**
------- | -------- | --------
o | Object | o {p {name:s}}
p {} | Contains the object's properties. |
p {} | Properties block | p {name:s; age:i}
s | String | name:s
i | Integer | age:i
n | Number | score:n
b | Boolean | active:b
a | Array | tags:a[s]
e[values] | Enum | color:e["red", "green", "blue"]
u[types] | Union | value:u[s, n]
lit[value] | Literal | status:lit["active"]
r | Required | name:s r
d=value | Default value | active:b d=true
ap f | Additional properties false | o {p {name:s} ap f}
type(key=value, ...) | Constrained type | name:s(minLength=1)
a[type] | Array with item type | tags:a[s]
o {p {prop:type}} | Nested object | user:o {p {id:i; name:s}}
?type | Optional type | ?s
t[type1, type2, ...] | Tuple | t[s, i]
s[type] | Set | s[i]
d[key, value] | Dictionary | d[s, i]
ClassName | Custom class | User

`;

  // Add available tools section
  instructions += '## AVAILABLE TOOLS FOR SUPERASSISTANT\n\n';

  // Add each tool with its schema
  tools.forEach(tool => {
    instructions += ` - ${tool.name}\n`;

    try {
      const schema = JSON.parse(tool.schema);

      if (tool.description) {
        instructions += `**Description**: ${tool.description}\n`;
      }

      if (schema.properties && Object.keys(schema.properties).length > 0) {
        instructions += '**Parameters**:\n';

        const requiredParams = Array.isArray(schema.required) ? schema.required : [];
        Object.entries(schema.properties).forEach(([paramName, paramDetails]: [string, any]) => {
          const isRequired = requiredParams.includes(paramName);
          instructions += `- \`${paramName}\`: ${paramDetails.description ? paramDetails.description : ''} (${paramDetails.type || 'any'}) (${isRequired ? 'required' : 'optional'})\n`;

          if (paramDetails.type === 'object' && paramDetails.properties) {
            instructions += '  - Properties:\n';
            Object.entries(paramDetails.properties).forEach(([nestedName, nestedDetails]: [string, any]) => {
              instructions += `    - \`${nestedName}\`: ${nestedDetails.description || 'No description'} (${nestedDetails.type || 'any'})\n`;
            });
          }

          if (
            paramDetails.type === 'array' &&
            paramDetails.items &&
            paramDetails.items.type === 'object' &&
            paramDetails.items.properties
          ) {
            instructions += '  - Array items (objects) with properties:\n';
            Object.entries(paramDetails.items.properties).forEach(([itemName, itemDetails]: [string, any]) => {
              instructions += `    - \`${itemName}\`: ${itemDetails.description || 'No description'} (${itemDetails.type || 'any'})\n`;
            });
          }
        });

        instructions += '\n';
      }
    } catch (error) {
      instructions += 'Schema information not available. No Tools Available';
    }
  });

  // Add custom instructions if enabled and available
  if (customInstructionsEnabled && customInstructions && customInstructions.trim()) {
    instructions += '<custom_instructions>\n';
    instructions += customInstructions.trim();
    instructions += '\n</custom_instructions>\n\n';
  }

  instructions += '<\\system>';

  instructions += '\n\n';

  // M365-SPECIFIC: Output plain JSON lines WITHOUT code fence wrappers.
  // M365 Copilot does not render ```jsonl blocks as <pre> elements, and XML triggers a "Page" creation.
  // The browser extension's INLINE_JSON_PATTERN in json_function_call_extractor.js detects
  // raw JSON objects directly from <message-content> elements without needing code fences.
  instructions += 'IMPORTANT: Output function call JSON objects as PLAIN TEXT with NO code fence wrapper.\n\n';
  instructions += 'Example of correct output (plain text, no backtick fences):\n\n';
  instructions += '{"type": "function_call_start", "name": "function_name", "call_id": 1}\n';
  instructions += '{"type": "description", "text": "Short 1 line of what this function does"}\n';
  instructions += '{"type": "parameter", "key": "parameter_1", "value": "value_1"}\n';
  instructions += '{"type": "parameter", "key": "parameter_2", "value": "value_2"}\n';
  instructions += '{"type": "function_call_end", "call_id": 1}\n\n';
  instructions += 'Do NOT wrap these lines in ``` code fences of any kind. Output them as raw plain text.\n';
  instructions += 'Each JSON object must be on a single line with no nested braces in parameter values.\n\n';

  // Append M365-specific override instructions (highest priority)
  instructions += m365CopilotInstructions;

  instructions += '\n\n';
  instructions += 'User Interaction Starts here:';
  instructions += '\n\n\n';
  instructions += '\n\n';
  instructions += '\n\n';
  instructions += '\n\n';
  return instructions;
};
