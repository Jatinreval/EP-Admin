/**
 * Initial Reminder handler
 * This module handles email generation for initial reminders to clients
 */

import { extractClientInfo, updateStatus, createNewEmail, getApiKey } from '../utils/email-utils';

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
  console.log("Initializing Reminder re Initial Email handler");
  
  // Remove any existing event listeners first to prevent duplicates
  const generateButton = document.getElementById("reminder-initial-generate");
  const replyButton = document.getElementById("reminder-initial-reply-button");
  
  if (generateButton) {
    // Create a new clone to remove all event listeners
    const newGenerateButton = generateButton.cloneNode(true);
    generateButton.parentNode.replaceChild(newGenerateButton, generateButton);
    
    // Add event listener with debounce for generate button
    let generateTimeout = null;
    newGenerateButton.addEventListener("click", function(event) {
      event.preventDefault();
      
      // Debounce to prevent double-clicking
      if (generateTimeout) {
        clearTimeout(generateTimeout);
      }
      
      generateTimeout = setTimeout(() => {
        generateInitialReminder();
        generateTimeout = null;
      }, 300);
    });
  }
  
  if (replyButton) {
    // Create a new clone to remove all event listeners
    const newReplyButton = replyButton.cloneNode(true);
    replyButton.parentNode.replaceChild(newReplyButton, replyButton);
    
    // Add event listener with debounce for reply button
    let replyTimeout = null;
    newReplyButton.addEventListener("click", function(event) {
      event.preventDefault();
      event.stopPropagation();
      
      // Debounce to prevent double-clicking
      if (replyTimeout) {
        clearTimeout(replyTimeout);
      }
      
      replyTimeout = setTimeout(() => {
        replyWithResponse();
        replyTimeout = null;
      }, 300);
    });
  }
  
  console.log("Event listeners initialized for Reminder re Initial Email handler");
}

/**
 * Generates an initial reminder email
 */
async function generateInitialReminder() {
  const statusElement = document.querySelector('.reminder-status');
  const responseContainer = document.getElementById('reminder-initial-response-container');
  const replyButton = document.getElementById('reminder-initial-reply-button');
  const copyButton = document.getElementById('reminder-initial-copy-button');
  
  // Show the loader
  showLoader();
  
  updateStatus("Generating initial reminder email...", '.reminder-status');
  
  // Hide previous response and buttons
  if (responseContainer) responseContainer.style.display = 'none';
  if (replyButton) replyButton.style.display = 'none';
  if (copyButton) copyButton.style.display = 'none';
  
  try {
    // Extract client info from the current email
    const clientInfo = await extractClientInfo();
    
    // Generate email from template using the AI prompt
    const emailContent = await generateEmailFromPrompt(clientInfo);
    
    // Show the generated content in the response container
    if (responseContainer) {
      responseContainer.innerText = emailContent;
      responseContainer.style.display = 'block';
    }
    
    // Show reply and copy buttons
    if (replyButton) replyButton.style.display = 'inline-block';
    if (copyButton) copyButton.style.display = 'inline-block';
    
    // Format a nice subject
    const subject = `Reminder - ${clientInfo.subject || clientInfo.name || 'Client'}`;
    
    // Format the email as HTML
    const formattedHtml = formatEmailAsHtml(emailContent, subject);
    
    // Create a new email message - but don't open a reply window automatically
    createNewEmail({
      toRecipients: [clientInfo.email].filter(Boolean),
      subject: subject,
      body: formattedHtml,
      openWindow: false // Don't open window automatically
    });
    
    updateStatus("Initial reminder email created successfully!", '.reminder-status');
  } catch (error) {
    updateStatus(`Error: ${error.message}`, '.reminder-status');
    console.error("Error generating initial reminder email:", error);
  } finally {
    // Hide the loader when done
    hideLoader();
  }
}

/**
 * Reply to the current email with the generated response
 */
function replyWithResponse() {
  // Show loader during reply creation
  showLoader();
  
  try {
    const responseContainer = document.getElementById('reminder-initial-response-container');
    
    // Get the response text
    const responseText = responseContainer.innerText;
    
    if (!responseText) {
      updateStatus("No response to reply with. Please generate a response first.", '.reminder-status');
      hideLoader();
      return;
    }
    
    // Format the email as HTML directly from the text (don't use stored HTML)
    const formattedHtml = formatEmailAsHtml(responseText);
    
          // Use the Office API to display a reply all form - exactly as in initial-email.js
      Office.context.mailbox.item.displayReplyAllForm(formattedHtml);
      
      updateStatus("Reply All created with the generated response!", '.reminder-status');
    hideLoader();
  } catch (error) {
    console.error("Error in replyWithResponse:", error);
    updateStatus(`Error creating reply all: ${error.message}`, '.reminder-status');
    hideLoader();
  }
}

/**
 * Formats the response text as HTML with proper styling
 * @param {string} text - The text to format
 * @param {string} [defaultSubject] - Optional default subject if none found in text
 * @returns {string} - HTML formatted response
 */
function formatEmailAsHtml(text, defaultSubject = '') {
  // Extract subject if available
  let subject = defaultSubject;
  let html = text;
  
  // Extract subject line if it exists in the format **Subject: XXX**
  const subjectMatch = text.match(/\*\*Subject: (.+?)\*\*/);
  if (subjectMatch && subjectMatch[1]) {
    subject = subjectMatch[1];
    // Remove the subject line from the text
    html = text.replace(/\*\*Subject: .+?\*\*\s*/g, '');
  }
  
  // Remove any markdown headers (lines starting with #)
  html = html.replace(/^#{1,6}\s+(.*)$/gm, '$1');
  
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
  
  // Convert bullet points to HTML list items (hyphen style)
  const bulletPattern = /<br>- (.*?)(?=<br>|$)/g;
  if (html.match(bulletPattern)) {
    html = html.replace(bulletPattern, '<br>• $1');
  }
  
  // Also handle asterisk bullets
  const asteriskPattern = /<br>\* (.*?)(?=<br>|$)/g;
  if (html.match(asteriskPattern)) {
    html = html.replace(asteriskPattern, '<br>• $1');
  }
  
  // Create a simple, clean HTML structure with just one wrapper div
  return `<div style="font-family: Arial, sans-serif; font-size: 10pt;">${html}</div>`;
}

/**
 * Generates email content using a prompt sent to the OpenAI API
 * @param {Object} clientInfo - Information about the client
 * @returns {Promise<string>} - The generated email text
 */
async function generateEmailFromPrompt(clientInfo) {
  try {
    // Hard-coded endpoint for testing
    const endpoint = "https://epmfl.openai.azure.com/openai/deployments/gpt-4o/chat/completions?api-version=2024-02-15-preview";
    
    // Get API key using shared utility function
    const apiKey = getApiKey();
    
    // Prepare the request payload for chat completions
    const payload = {
      messages: [
        {
          role: "system",
          content: "You are an AI assistant specialized in drafting professional legal communications. Your task is to create a well-structured initial reminder email to a client that requires their attention. Format your response using markdown for bold text (**bold**) and proper structure. IMPORTANT: Output ONLY the email content itself - do NOT include any analysis, explanations, or headers/notes before the actual email."
        },
        {
          role: "user",
          content: `Draft a professional initial reminder email to a client. Include these details:

Client Name: ${clientInfo.name || 'Valued Client'}
Client Email: ${clientInfo.email || '[Client Email]'}

Email Content Guidelines:
1. Begin with a subject line: "**Subject: Reminder - ${clientInfo.subject || 'Your Matter'}**"
2. Next, start with "**Private and Confidential**"
3. Begin with a formal greeting to the client
4. Politely remind them about their pending matter that requires their attention
5. Mention that their response or action is needed to proceed
6. Provide a brief summary of what action is needed from them
7. Suggest a timeframe for their response
8. Offer assistance if they have any questions
9. Thank them for their attention to this matter
10. End with a formal closing and signature

Keep the tone professional yet approachable, emphasizing the importance of their response while maintaining a helpful approach.`
        }
      ],
      temperature: 0.7,
      max_tokens: 500,
      top_p: 0.95,
      frequency_penalty: 0,
      presence_penalty: 0
    };
    
    // Make the API call
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
    let generatedText = data.choices[0].message.content;
    
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
    
    return displayResponse;
  } catch (error) {
    console.error("Error generating email from prompt:", error);
    throw error;
  }
}

export default {
  initialize
}; 