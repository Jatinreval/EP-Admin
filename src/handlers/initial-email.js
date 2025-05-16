/**
 * Initial Email handler
 * This module handles the AI email response generation for the Initial Email option
 */

import { extractClientInfo, updateStatus } from '../utils/email-utils';

// Function to show the loader
function showLoader() {
  const loaderContainer = document.getElementById('loader-container');
  if (loaderContainer) {
    loaderContainer.style.display = 'flex';
  }
}

// Function to hide the loader
function hideLoader() {
  const loaderContainer = document.getElementById('loader-container');
  if (loaderContainer) {
    loaderContainer.style.display = 'none';
  }
}

// Initialize the module
function initialize() {
  // Add event listeners for buttons
  document.getElementById("openai-button").onclick = sendToAzureOpenAI;
  document.getElementById("reply-button").onclick = replyWithResponse;
  
  // Hide the reply button initially - it will be shown when a response is generated
  document.getElementById("reply-button").style.display = "none";
}

/**
 * Sends email content to Azure OpenAI for processing using fetch API
 */
async function sendToAzureOpenAI() {
  const statusElement = document.getElementById("status");
  const responseContainer = document.getElementById("response-container");
  const copyButton = document.getElementById("copy-button");
  const replyButton = document.getElementById("reply-button");
  
  // Show the loader when starting API call
  showLoader();
  
  updateStatus("Preparing to send to Azure OpenAI...");
  responseContainer.style.display = "none";
  copyButton.style.display = "none";
  replyButton.style.display = "none";

  try {
    // Get client info with email data
    const emailData = await extractClientInfo();
    
    if (!emailData) {
      updateStatus("No email selected");
      hideLoader(); // Hide loader if no email
      return;
    }
    
    try {
      updateStatus("Processing email content...");
      
      // Hard-coded endpoint for testing
      const endpoint = "https://epmfl.openai.azure.com/openai/deployments/gpt-4o/chat/completions?api-version=2024-02-15-preview";
      
      // Try to get API key from various sources
      let apiKey;
      try {
        // First check window.__env from runtime-config.js
        if (window.__env && window.__env.AZURE_OPENAI_API_KEY) {
          apiKey = window.__env.AZURE_OPENAI_API_KEY;
          console.log("Using API key from runtime config");
        } 
        // Fallback to process.env if available
        else if (typeof process !== 'undefined' && process.env && process.env.AZURE_OPENAI_API_KEY) {
          apiKey = process.env.AZURE_OPENAI_API_KEY;
          console.log("Using API key from process.env");
        } 
        // Final fallback
        else {
          apiKey = "your_api_key_here";
          console.log("No API key found in runtime config or process.env");
        }
      } catch (e) {
        console.error("Error accessing API key:", e);
        apiKey = "your_api_key_here";
      }
      
      if (apiKey === "your_api_key_here") {
        updateStatus("Azure OpenAI API key not configured. Please add your key to the .env file.");
        hideLoader(); // Hide loader if API key is missing
        return;
      }
      
      console.log("Endpoint URL:", endpoint);
      updateStatus("Generating AI response...");
      
      // AI PROMPT STRUCTURE FOR INITIAL EMAIL
      // Prepare the request payload for chat completions
      const payload = {
        messages: [
          {
            role: "system",
            content: "You are an AI assistant specialized in analyzing legal correspondence and generating appropriate follow-up emails. Your task is to analyze an email from Bhavesh Mistry (BM) to a client and generate a draft response following specific templates and guidelines. Format your response using markdown for bold text (**bold**) and hyperlinks in the format [link text](URL). Follow the exact structure of the template, including all formatting elements. IMPORTANT: Output ONLY the email content itself without any analysis, explanations, or headers. Do not include any text like 'ANALYSIS:', 'EMAIL:', 'DRAFT EMAIL:', etc. Just start with the actual email content."
          },
          {
            role: "user",
            content: `AI Prompt for Email Analysis and Response Generation

Input Format
You will receive: 
1. Initial Email: The original email sent by BM to the client including the subject line
2. Available Meeting Times: A list of dates and times when BM is available for meetings 
3. Templates: Reference email templates to follow

Here is the email to analyze:
Subject: ${emailData.subject}
From: ${emailData.email}
Date: ${emailData.receivedTime}
Body:
${emailData.body}

Available Meeting Times:
- Thursday, 27 March 2025 at 10:30am
- Thursday, 27 March 2025 at 1:30pm
- Thursday, 27 March 2025 at 2:30pm

Analysis Requirements
Please analyze the initial email to identify:
1.         Client Information:
–          Extract all client names mentioned in the email greeting (e.g., "Hi Barry" → "Barry")
–          Note if multiple clients are addressed (e.g., couples, business partners)
2.         Service Type:
–          Determine if the email is about estate planning or non-estate planning
–          Estate planning indicators include: "estate planning," "asset protection," "will," "enduring power of attorney," "estate plan," "SMSF Trust Deeds," "Family Trust Deeds"
–          Non-estate planning might relate to: divorce, business matters, disputes, etc.
3.         Meeting Format:
–          Identify if the meeting is proposed as:
•           MS Teams meeting (if call then also MS Teams meeting unless noted specifically that call will be on mobile)
•           In-person meeting at the office
•           In-person meeting at another location (specify if mentioned)
–          Default to "MS Teams meeting OR in-person meeting at our office" if unclear

Response Generation Guidelines
Based on your analysis:
1.         Template Selection:
–          Use Template B if the matter involves estate planning
–          Use Template A for all other matters
2.         Email Structure:
–          Begin with "Private and Confidential"
–          Address the client by name
–          Include "Conference" section with available meeting times
–          For estate planning, include "Questionnaire" section with link
–          End with standard closing and "Reply All" instruction
3.         Formatting Requirements:
–          Maintain the exact formatting from the templates
–          Present meeting times as bullet points
–          Preserve all bold formatting for headings

Template Reference
Template A (Non-Estate Planning)
**Private and Confidential**
 
Hi [client name]
 
Further to Bhavesh Mistry's email correspondence today, we look forward to assisting you.
 
**Conference**
 
Please note, Bhavesh Mistry will be available at the following dates and times below for a **[meeting format]** with you for further discussion:
 
- [date/time option 1]; or
- [date/time option 2]; or
- [date/time option 3].
 
We look forward to hearing from you shortly.
 
Please email us ensuring that you select \`Reply All\` to our email so our team can assist you.

Template B (Estate Planning)
**Private and Confidential**
 
Hi [client name]
 
Further to Bhavesh Mistry's email correspondence today, we look forward to assisting you with your estate planning.
 
**Conference**
 
To allow us to understand your intentions, Bhavesh Mistry will be available at the following dates and times below for a **[meeting format]** with you for an initial discussion of your estate planning:
 
- [date/time option 1]; or
- [date/time option 2]; or
- [date/time option 3].
 
Please kindly let us know if any of the above times are suitable and your best contact number. Alternatively, please let us know if there are any other dates and times more suitable for you.
 
**Questionnaire**
 
In preparation for our conference and to assist us in obtaining your initial information and allowing you to start considering your estate plan, please take a few minutes to complete our estate planning questionnaire at the following [link](https://mistryfallahi.com.au/client-asset-protection-enquiry/).
 
We look forward to hearing from you shortly.
 
Please email us ensuring that you select \`Reply All\` to our email so our team can assist you.

Output Format
Your response should include ONLY the Draft Email which is completely formatted following the appropriate template. Do not include headings like "Analysis" or "Draft Email".

Example output structure:
DRAFT EMAIL:
**Subject: Conference - [Include the complete subject here]**
[Complete formatted email]

Important Notes
•           Do not include section headings like "Analysis:" or "Draft Email:" in your response.
•           Do not include placeholders in the final draft email. All fields should be properly populated.
•           Maintain exact formatting from templates including bold text, bullet points, and paragraph spacing.
•           If anything is unclear, default to the most conservative option and note your uncertainty in the analysis.
•           The email subject line MUST have "Conference -" added at the beginning, followed by the complete original subject.
•           For example: "Conference - Estate Planning - Paulina and Alex Pavlova (Our Ref: 10-0204)" 
•           Date and time to be in this format: Thursday, 27 March 2025 at 10:30am, 1:30pm or 2:30pm.
•           Use **bold** markdown formatting for headings and important text.
•           Format the link in the questionnaire section as a proper markdown link [link](URL).`
          }
        ],
        temperature: 0.2,
        max_tokens: 1000,
        top_p: 0.95,
        frequency_penalty: 0,
        presence_penalty: 0,
        stop: null
      };
      
      const response = await fetch(endpoint, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "api-key": apiKey
        },
        body: JSON.stringify(payload)
      });
      
      if (!response.ok) {
        const errorData = await response.json().catch(() => null);
        throw new Error(`API request failed: ${response.status} ${response.statusText}${errorData ? ' - ' + JSON.stringify(errorData) : ''}`);
      }
      
      const data = await response.json();
      const generatedText = data.choices[0].message.content;
      
      // Process the AI response to handle potential formatting issues
      let displayResponse = generatedText;
      let extractedSubject = "";
      
      // First try to get content after "DRAFT EMAIL:" marker
      if (generatedText.includes("DRAFT EMAIL:")) {
        displayResponse = generatedText.split("DRAFT EMAIL:")[1].trim();
      }
      
      // Also handle the case where response contains "### Draft Email:"
      if (displayResponse.includes("### Draft Email:")) {
        displayResponse = displayResponse.split("### Draft Email:")[1].trim();
      }
      
      // Remove any analysis section if present
      if (displayResponse.includes("### Analysis:")) {
        displayResponse = displayResponse.split("### Analysis:")[1].trim();
        // If the response contains both analysis and draft email sections
        if (displayResponse.includes("### Draft Email:")) {
          displayResponse = displayResponse.split("### Draft Email:")[1].trim();
        }
      }
      
      // Extract the subject if present in the format: **Subject: Conference - XXX**
      const subjectMatch = displayResponse.match(/\*\*Subject: (Conference - .+?)\*\*/);
      if (subjectMatch && subjectMatch[1]) {
        extractedSubject = subjectMatch[1];
        // Store the subject in a data attribute on the response container for later use
        responseContainer.setAttribute('data-subject', extractedSubject);
        // Remove the subject line from the display response
        displayResponse = displayResponse.replace(/\*\*Subject: .+?\*\*\s*/g, '');
      }
      
      // Display the processed response
      responseContainer.innerText = displayResponse;
      responseContainer.style.display = "block";
      
      // Show copy and reply buttons
      copyButton.style.display = "inline-block";
      replyButton.style.display = "inline-block";
      
      updateStatus("Response generated!");
      
      // Hide the loader after successful response
      hideLoader();
      
    } catch (apiError) {
      console.error("API error:", apiError);
      updateStatus(`Error calling Azure OpenAI API: ${apiError.message}`);
      hideLoader(); // Hide loader on API error
    }
  } catch (error) {
    updateStatus(`Error: ${error.message}`);
    hideLoader(); // Hide loader on any other error
  }
}

/**
 * Reply to the current email with the generated response
 */
function replyWithResponse() {
  // Show loader during reply creation
  showLoader();
  
  const statusElement = document.getElementById("status");
  const responseContainer = document.getElementById("response-container");
  
  try {
    const responseText = responseContainer.innerText;
    
    if (!responseText) {
      statusElement.innerText = "No response to reply with. Please generate a response first.";
      hideLoader();
      return;
    }
    
    // Format response as HTML with Arial 10 font
    const formattedHtml = formatEmailAsHtml(responseText);
    
    // Use the formatted HTML in the reply all
    Office.context.mailbox.item.displayReplyAllForm(formattedHtml);
    statusElement.innerText = "Reply All created with the generated response!";
    hideLoader();
  } catch (error) {
    statusElement.innerText = `Error creating reply: ${error.message}`;
    hideLoader();
  }
}

/**
 * Formats the response text as HTML with proper styling
 * @param {string} text - The text to format
 * @returns {string} - HTML formatted response
 */
function formatEmailAsHtml(text) {
  // Get the subject from the container if available
  const responseContainer = document.getElementById("response-container");
  const subject = responseContainer.getAttribute('data-subject') || '';
  
  // Remove any markdown headers (lines starting with #)
  let html = text.replace(/^#{1,6}\s+(.*)$/gm, '$1');
  
  // Remove any "DRAFT EMAIL:" or similar prefixes
  html = html.replace(/^DRAFT EMAIL:[\s\n]*/i, '');
  html = html.replace(/^EMAIL:[\s\n]*/i, '');
  html = html.replace(/^RESPONSE:[\s\n]*/i, '');
  html = html.replace(/^ANALYSIS:[\s\n]*/i, '');
  
  // Remove analysis sections (any text before "Private and Confidential")
  const privateConfidentialIndex = html.indexOf('**Private and Confidential**');
  if (privateConfidentialIndex > 0) {
    html = html.substring(privateConfidentialIndex);
  }
  
  // Add the subject line above "Private and Confidential" if available
  if (subject) {
    html = `<strong>Subject:</strong> ${subject}<br><br>` + html;
  }
  
  // Replace ** bold ** with HTML bold
  html = html.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
  
  // Convert line breaks to <br> tags
  html = html.replace(/\n/g, '<br>');
  
  // Convert markdown-style links [text](url) to HTML links
  html = html.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2">$1</a>');
  
  // Convert bullet points (lines starting with "- ") to HTML list items
  const bulletPattern = /<br>- (.*?)(?=<br>|$)/g;
  if (html.match(bulletPattern)) {
    html = html.replace(bulletPattern, '<br>• $1');
  }
  
  // Wrap the content in an HTML structure with Arial 10px font
  return `<div style="font-family: Arial, sans-serif; font-size: 10pt;">${html}</div>`;
}

// Generate email content based on the original email
async function generateEmailFromPrompt(emailContent, clientInfo) {
  try {
    updateStatus('Generating email...');
    showLoader();

    // Get the API key from environment or global variable
    const apiKey = process.env.REACT_APP_OPENAI_API_KEY || window.OPENAI_API_KEY;
    if (!apiKey) {
      throw new Error('OpenAI API key not found');
    }

    const endpoint = "https://epmfl.openai.azure.com/openai/deployments/gpt-4o/chat/completions?api-version=2024-02-15-preview";

    // Extract useful information from the client info object
    const clientInfoPrompt = generateClientInfoPrompt(clientInfo);

    // Structure the prompt for the OpenAI API
    const payload = {
      messages: [
        {
          role: "system",
          content: `You are an assistant helping a law firm generate professional and courteous email responses. 
IMPORTANT: Output ONLY the email content itself - do NOT include any analysis, explanations, or headers/notes before the actual email.
DO NOT include text like "ANALYSIS:", "EMAIL:", or "DRAFT EMAIL:" in your output.
The output should begin with "**Private and Confidential**" and then proceed directly with the email content.`
        },
        {
          role: "user",
          content: `I need to respond to the following email. Please draft a professional reply:

${emailContent}

${clientInfoPrompt}

Response Generation Guidelines:
1. Begin with "**Private and Confidential**" at the top of the email
2. Format the email as follows:
   - Use bold text for emphasis where appropriate using markdown (e.g., **bold text**)
   - Use bullets for lists where appropriate (e.g., "- item")
3. Maintain a professional, courteous tone throughout
4. Address specific questions or concerns raised in the original email
5. Include appropriate next steps or action items if applicable
6. Include a professional sign-off

OUTPUT FORMAT: Generate ONLY the email content itself, starting with "**Private and Confidential**". Do not include any analysis or explanations before the actual email content.`
        }
      ],
      temperature: 0.7,
      max_tokens: 800,
      top_p: 0.95,
      frequency_penalty: 0,
      presence_penalty: 0,
      stop: null
    };

    const response = await fetch(endpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "api-key": apiKey
      },
      body: JSON.stringify(payload)
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => null);
      throw new Error(`API request failed: ${response.status} ${response.statusText}${errorData ? ' - ' + JSON.stringify(errorData) : ''}`);
    }

    const data = await response.json();
    const generatedText = data.choices[0].message.content;

    // Process the AI response to handle potential formatting issues
    let displayResponse = generatedText;
    
    // First try to get content after "DRAFT EMAIL:" marker
    if (generatedText.includes("DRAFT EMAIL:")) {
      displayResponse = generatedText.split("DRAFT EMAIL:")[1].trim();
    }
    
    // Also handle the case where response contains "### Draft Email:"
    if (displayResponse.includes("### Draft Email:")) {
      displayResponse = displayResponse.split("### Draft Email:")[1].trim();
    }
    
    // Remove any analysis section if present
    if (displayResponse.includes("### Analysis:")) {
      displayResponse = displayResponse.split("### Analysis:")[1].trim();
      // If the response contains both analysis and draft email sections
      if (displayResponse.includes("### Draft Email:")) {
        displayResponse = displayResponse.split("### Draft Email:")[1].trim();
      }
    }
    
    // Display the processed response
    responseContainer.innerText = displayResponse;
    responseContainer.style.display = "block";

    // Show copy and reply buttons
    copyButton.style.display = "inline-block";
    replyButton.style.display = "inline-block";

    updateStatus("Response generated!");

    // Hide the loader after successful response
    hideLoader();

  } catch (error) {
    console.error("Error:", error);
    updateStatus(`Error: ${error.message}`);
    hideLoader(); // Hide loader on any error
  }
}

export default {
  initialize
}; 