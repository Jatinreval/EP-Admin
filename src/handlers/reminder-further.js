/**
 * Reminder re Further Information handler
 * This module handles email generation for reminders about further information
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
  console.log("Initializing Reminder re Further Information handler");
  
  // Remove any existing event listeners first to prevent duplicates
  const generateButton = document.getElementById("reminder-further-generate");
  const replyButton = document.getElementById("reminder-further-reply-button");
  
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
        generateFurtherReminder();
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
  
  console.log("Event listeners initialized for Reminder re Further Information handler");
}

/**
 * Generates a further information reminder email
 */
async function generateFurtherReminder() {
  const statusElement = document.querySelector('.reminder-status');
  const responseContainer = document.getElementById('reminder-further-response-container');
  const replyButton = document.getElementById('reminder-further-reply-button');
  const copyButton = document.getElementById('reminder-further-copy-button');
  
  // Show the loader
  showLoader();
  
  updateStatus("Generating further information reminder email...", '.reminder-status');
  
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
    const subject = `Request for Further Information - ${clientInfo.subject || clientInfo.name || 'Client'}`;
    
    // Format the email as HTML
    const formattedHtml = formatEmailAsHtml(emailContent, subject);
    
    // Store the formatted HTML for later use
    if (responseContainer) {
      responseContainer.setAttribute('data-formatted-html', formattedHtml);
    }
    
    // Create a new email message without opening a window
    createNewEmail({
      toRecipients: [clientInfo.email].filter(Boolean),
      subject: subject,
      body: formattedHtml,
      openWindow: false // Don't open window automatically
    });
    
    updateStatus("Further information reminder email created successfully!", '.reminder-status');
  } catch (error) {
    updateStatus(`Error: ${error.message}`, '.reminder-status');
    console.error("Error generating further information reminder email:", error);
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
    const responseContainer = document.getElementById('reminder-further-response-container');
    
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
  
  // Convert bullet points (lines starting with "- ") to HTML list items
  const bulletPattern = /<br>- (.*?)(?=<br>|$)/g;
  if (html.match(bulletPattern)) {
    html = html.replace(bulletPattern, '<br>• $1');
  }
  
  // Wrap the content in an HTML structure with Arial 10px font
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
          content: "You are an AI assistant specialized in drafting professional legal communications. Your task is to create a well-structured reminder email to a client regarding further information that is needed to proceed with their matter. Format your response using markdown for bold text (**bold**) and proper structure. IMPORTANT: Output ONLY the email content itself - do NOT include any analysis, explanations, or headers/notes before the actual email."
        },
        {
          role: "user",
          content: `Draft a professional reminder email to a client about providing further information. Include these details:

Client Name: ${clientInfo.name || 'Valued Client'}
Client Email: ${clientInfo.email || '[Client Email]'}

Email Content Guidelines:
1. Begin with a subject line: "**Subject: Reminder: Further Information Required - ${clientInfo.subject || 'Your Matter'}**"
2. Next, start with "**Private and Confidential**"
3. Begin with a formal greeting to the client
4. Reference your previous communication requesting specific information
5. Remind them that you are still waiting for them to provide that information
6. Emphasize that this information is essential to proceed with their matter
7. Provide a suggested deadline for their response
8. Offer to discuss any questions or concerns they may have about what's needed
9. Request they contact your office at their earliest convenience
10. End with a formal closing and signature

Keep the tone professional yet friendly, emphasizing the importance of the requested information while maintaining a helpful approach.`
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