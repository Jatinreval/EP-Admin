/**
 * Conference for Draft Documents handler
 * This module handles email generation for Conference for Draft Documents
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
  console.log("Initializing Conference for Draft Documents handler");
  
  // Remove any existing event listeners first to prevent duplicates
  const generateButton = document.getElementById("draft-generate");
  const replyButton = document.getElementById("draft-reply-button");
  
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
        generateDraftConferenceEmail();
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
  
  console.log("Event listeners initialized for Conference for Draft Documents handler");
}

/**
 * Generates a Conference for Draft Documents email
 */
async function generateDraftConferenceEmail() {
  const statusElement = document.querySelector('.draft-status');
  const responseContainer = document.getElementById('draft-response-container');
  const replyButton = document.getElementById('draft-reply-button');
  const copyButton = document.getElementById('draft-copy-button');
  
  // Show the loader
  showLoader();
  
  updateStatus("Generating conference for draft documents email...", '.draft-status');
  
  // Hide previous response and buttons
  if (responseContainer) responseContainer.style.display = 'none';
  if (replyButton) replyButton.style.display = 'none';
  if (copyButton) copyButton.style.display = 'none';
  
  try {
    // Extract client info from the current email
    const clientInfo = await extractClientInfo();
    
    // Generate email from AI prompt
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
    let subject = `Conference - ${clientInfo.subject || "Draft Documents"}`;
    const subjectMatch = emailContent.match(/\*\*Subject: (.+?)\*\*/);
    if (subjectMatch && subjectMatch[1]) {
      subject = subjectMatch[1];
    }
    
    // Store the subject in a data attribute (but not the formatted HTML)
    if (responseContainer) {
      responseContainer.setAttribute('data-subject', subject);
    }
    
    // Format the email as HTML (only for creating the new email, not for storing)
    const formattedHtml = formatEmailAsHtml(emailContent);
    
    // Create a new email but don't automatically open a window
    createNewEmail({
      toRecipients: [clientInfo.email].filter(Boolean),
      subject: subject,
      body: formattedHtml,
      openWindow: false // Don't open window automatically
    });
    
    updateStatus("Conference for draft documents email created successfully!", '.draft-status');
  } catch (error) {
    updateStatus(`Error: ${error.message}`, '.draft-status');
    console.error("Error generating conference for draft documents email:", error);
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
    const responseContainer = document.getElementById('draft-response-container');
    
    // Get the response text
    const responseText = responseContainer.innerText;
    
    if (!responseText) {
      updateStatus("No response to reply with. Please generate a response first.", '.draft-status');
      hideLoader();
      return;
    }
    
    // Format the email as HTML directly from the text (don't use stored HTML)
    const formattedHtml = formatEmailAsHtml(responseText);
    
    // Use the Office API to display a reply form - exactly as in initial-email.js
    Office.context.mailbox.item.displayReplyAllForm(formattedHtml);
    
          updateStatus("Reply All created with the generated response!", '.draft-status');
    hideLoader();
  } catch (error) {
    console.error("Error in replyWithResponse:", error);
    updateStatus(`Error creating reply: ${error.message}`, '.draft-status');
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
  
  // First try to get subject from data attribute
  const responseContainer = document.getElementById('draft-response-container');
  if (responseContainer) {
    const storedSubject = responseContainer.getAttribute('data-subject');
    if (storedSubject) {
      subject = storedSubject;
    }
  }
  
  // If no subject in data attribute, try to extract from text
  if (!subject) {
    const subjectMatch = text.match(/\*\*Subject: (.+?)\*\*/);
    if (subjectMatch && subjectMatch[1]) {
      subject = subjectMatch[1];
      // Remove the subject line from the text
      html = text.replace(/\*\*Subject: .+?\*\*\s*/g, '');
    }
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
  
  // Convert bullet points to HTML list items
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
    // Get API key using the shared utility function
    const apiKey = getApiKey();
    
    // Format meeting times for the prompt as bullet points
    const meetingTimes = [
      "Monday, 28 April 2025 at 10:30am or 11am",
      "Thursday, 1 May 2025 at 10:30am, 11am, 2pm or 3pm",
      "Monday, 5 May 2025 at 9am, 12:30pm, 1pm or 2pm"
    ].map(time => `* ${time}`).join('; or\n');
    
    // Prepare the original subject line
    const originalSubject = clientInfo.subject || "Draft Documents";
    
    // Prepare the request payload for chat completions
    const payload = {
      messages: [
        {
          role: "system",
          content: `You are an AI assistant specialized in analyzing legal correspondence and generating appropriate follow-up emails. Your task is to analyze an email to a client enclosing draft legal documents and generate a draft response following specific template and guidelines.
In this task, we have to draft an email to propose times for a meeting with the lawyer Bhavesh Mistry. 
In the emails, you will notice that we are using the words such as execution, testamentary, superannuation, last wishes, etc. Since the correspondence can be about drafting and signing of wills, please ensure that these are not misinterpreted in any other context.`
        },
        {
          role: "user",
          content: `Please draft an email for a conference about draft documents using the following information:

Initial Email Body:
${clientInfo.body || "N/A"}

Subject: ${originalSubject}

Available Meeting Times:
${meetingTimes}

Client Name: ${clientInfo.name || "[CLIENT NAME]"}
Client Email: ${clientInfo.email || "[CLIENT EMAIL]"}

Template Reference:
**Private and Confidential**

Hi [client name]

Further to our previous correspondence enclosing your draft [document type], please let us know if you have any queries or amendments so we can finalise your documents for signing without delay.

[Text from input email with heading for Further Information. The heading to be bold.]

**Conference**

We are available to discuss the drafts with you at the following dates and times:

- [date/time option 1]; or
- [date/time option 2]; or
- [date/time option 3].

Please let us know what date and times work best for your schedule. Alternatively, please let us know of a few dates and times more suitable for you.

We look forward to hearing from you shortly.

Please email us ensuring that you select \`Reply All\` to our email so our team can assist you.

Analysis Requirements:
1. Extract all client names mentioned in the email greeting.
2. Identify the type of document being discussed (e.g., testamentary documents, deed of amendment, trust documents).
3. Note any further information sought in the input email.

Response Generation Guidelines:
1. Email Structure:
   - Begin with "Private and Confidential."
   - Address the client by name.
   - Reference the specific document type in the follow-up line.
   - Include a "Conference" section with available meeting times.
   - End with standard closing and "Reply All" instruction.
   - The email subject line should have "Conference -" added at the beginning of the original subject line.
2. Formatting Requirements:
   - Maintain the exact formatting from the templates.
   - Present meeting times as bullet points.
   - Preserve all bold formatting for headings.
3. Strict Adherence to Template:
   - Only include sections explicitly outlined in the template.
   - Omit any additional sections not part of the template.
4. Further Information Section:
   - Use the exact content if a "Further Information" section exists in the input email.
   - If no such section is present, omit the "Further Information" heading entirely.

IMPORTANT: Your response should begin with the subject line: "**Subject: Conference - ${originalSubject}**" followed by the template content. Do NOT include any analysis, explanations, or notes before the actual email content.`
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
