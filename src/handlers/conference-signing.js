/**
 * Conference for Signing Documents handler
 * This module handles the generation of emails for a conference for signing documents
 */

import { extractClientInfo, updateStatus, createNewEmail, DEFAULT_MEETING_TIMES, getApiKey } from '../utils/email-utils';

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
  document.getElementById("signing-generate").onclick = generateSigningConferenceEmail;
  document.getElementById("signing-reply-button").onclick = replyWithResponse;
}

/**
 * Generates an email for a conference for signing documents
 */
async function generateSigningConferenceEmail() {
  // Get references to UI elements
  const statusElement = document.querySelector('.conference-status');
  const responseContainer = document.getElementById('conference-signing-response-container');
  const replyButton = document.getElementById('signing-reply-button');
  const copyButton = document.getElementById('conference-signing-copy-button');
  
  // Show the loader
  showLoader();
  
  updateStatus("Generating conference for signing documents email...", '.conference-status');
  
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
    
    // Extract subject from the response content
    let subject = `Conference for Signing Documents - ${clientInfo.subject || "Estate Planning"}`;
    const subjectMatch = emailContent.match(/\*\*Subject: (.+?)\*\*/);
    if (subjectMatch && subjectMatch[1]) {
      subject = subjectMatch[1];
    }
    
    // Format the email as HTML
    const formattedHtml = formatEmailAsHtml(emailContent);
    
    // Store the formatted HTML in the response container for later use
    if (responseContainer) {
      responseContainer.setAttribute('data-formatted-html', formattedHtml);
    }
    
    // Create a new email but don't automatically open a window
    createNewEmail({
      toRecipients: [clientInfo.email].filter(Boolean),
      subject: subject,
      body: formattedHtml,
      openWindow: false // Don't open window automatically
    });
    
    updateStatus("Conference for signing documents email created successfully!", '.conference-status');
  } catch (error) {
    updateStatus(`Error: ${error.message}`, '.conference-status');
    console.error("Error generating conference for signing documents email:", error);
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
  
  const statusElement = document.querySelector('.conference-status');
  const responseContainer = document.getElementById('conference-signing-response-container');
  
  try {
    // Check if we have a cached formatted HTML version
    const cachedFormattedHtml = responseContainer.getAttribute('data-formatted-html');
    let formattedHtml;
    
    if (!cachedFormattedHtml) {
      // Fall back to formatting from text content
      const responseText = responseContainer.innerText;
      
      if (!responseText) {
        updateStatus("No response to reply with. Please generate a response first.", '.conference-status');
        hideLoader();
        return;
      }
      
      // Format response as HTML
      formattedHtml = formatEmailAsHtml(responseText);
    } else {
      formattedHtml = cachedFormattedHtml;
    }
    
    // Use a setTimeout to ensure any previous operations have completed
    setTimeout(() => {
      try {
        // Use Office API to create a properly formatted reply - only call this once
        Office.context.mailbox.item.displayReplyAllForm(formattedHtml);
        
        // Update status after successful reply all creation
        updateStatus("Reply All created with the generated response!", '.conference-status');
      } catch (replyError) {
        console.error("Error creating reply all:", replyError);
        updateStatus(`Error creating reply all: ${replyError.message}`, '.conference-status');
      } finally {
        hideLoader();
      }
    }, 100);
  } catch (error) {
    console.error("Error in replyWithResponse:", error);
    updateStatus(`Error creating reply all: ${error.message}`, '.conference-status');
    hideLoader();
  }
}

/**
 * Formats the response text as HTML with proper styling
 * @param {string} text - The text to format
 * @returns {string} - HTML formatted response
 */
function formatEmailAsHtml(text) {
  // Extract subject if available
  let subject = '';
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
    // Get API key using the shared utility function
    const apiKey = getApiKey();
    
    // Format meeting times for the prompt
    const meetingTimes = [
      "Wednesday, 15 May 2025 at 10:30am or 11am; or",
      "Thursday, 16 May 2025 at 10:30am, 11am, 2pm or 3pm; or",
      "Friday, 17 May 2025 at 9am, 12:30pm, 1pm or 2pm"
    ].map(time => `- ${time}`).join('\n');
    
    // Prepare the request payload for chat completions
    const payload = {
      messages: [
        {
          role: "system",
          content: "You are an AI assistant specialized in analyzing legal correspondence and generating appropriate follow-up emails. IMPORTANT: Output ONLY the email content itself - do NOT include any analysis, explanations, or headers/notes before the actual email. DO NOT include text like 'ANALYSIS:', 'EMAIL:', or 'DRAFT EMAIL:' in your output. The output should begin with '**Private and Confidential**' and then proceed directly with the email content."
        },
        {
          role: "user",
          content: `AI Prompt for Conference Email Generation (Signing Documents)

The client information is:
Name: ${clientInfo.name}
Email: ${clientInfo.email}
Matter: ${clientInfo.matter || "Estate Planning"}

Based on this information, draft a professional email to schedule a conference regarding signing documents. Use the following template structure:

**Subject: Conference for Signing Documents - ${clientInfo.subject || "Estate Planning"}**

**Private and Confidential**

Dear ${clientInfo.name || "[Client Name]"},

**Signing [Document Type]**

We are pleased to advise that your documents are now ready for signing.

**Conference**

We would like to arrange a time to conference with you regarding the signing of the documents. We have the following appointment times available:

${meetingTimes}

Please let us know what date and times work best for your schedule. Alternatively, please let us know of a few dates and times more suitable for you.

We look forward to hearing from you shortly.

Please email us ensuring that you select \`Reply All\` to our email so our team can assist you.

Response Generation Guidelines:
- Begin with the Subject line as shown above
- Then "**Private and Confidential**"
- Address the client by name
- Maintain the exact formatting from the template including bold headings
- Present meeting times as bullet points as provided
- OUTPUT FORMAT: Generate ONLY the email content itself, starting with the Subject line followed by "**Private and Confidential**". Do not include any analysis or explanations before the actual email content.`
        }
      ],
      temperature: 0.7,
      max_tokens: 800,
      top_p: 0.95,
      frequency_penalty: 0,
      presence_penalty: 0
    };
    
    // Make the API call
    const response = await fetch("https://epmfl.openai.azure.com/openai/deployments/gpt-4o/chat/completions?api-version=2024-02-15-preview", {
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