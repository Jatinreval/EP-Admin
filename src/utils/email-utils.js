/**
 * Utilities for email extraction and client information functions
 * These are shared across all email generation modules
 */

/**
 * Extracts client information from the current email
 * @returns {Promise<Object>} - Promise resolving to client information
 */
function extractClientInfo() {
  return new Promise((resolve, reject) => {
    try {
      const item = Office.context.mailbox.item;
      
      if (!item) {
        reject(new Error("No email selected"));
        return;
      }
      
      // Create an object to store client data
      const clientInfo = {
        name: "",
        email: "",
        subject: item.subject || "",
        body: "",
        receivedTime: item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : null
      };
      
      // Try to extract name from greeting
      item.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const bodyText = result.value;
          clientInfo.body = bodyText;
          
          // Extract name from greeting (e.g., "Hi Barry" -> "Barry")
          const greetingMatch = bodyText.match(/(?:Hi|Hello|Dear)\s+([A-Za-z]+)/);
          if (greetingMatch && greetingMatch[1]) {
            clientInfo.name = greetingMatch[1];
          }
          
          // Extract email if available in the sender
          if (item.sender) {
            clientInfo.email = item.sender.emailAddress;
          }
          
          resolve(clientInfo);
        } else {
          reject(new Error(`Error getting email body: ${result.error.message}`));
        }
      });
    } catch (error) {
      reject(error);
    }
  });
}

/**
 * Gets the API key from various sources
 * @returns {string} - The API key or error if not found
 */
function getApiKey() {
  try {
    // First check window.__env from runtime-config.js
    if (window.__env && window.__env.AZURE_OPENAI_API_KEY) {
      console.log("Using API key from runtime config");
      return window.__env.AZURE_OPENAI_API_KEY;
    } 
    // Fallback to process.env if available
    else if (typeof process !== 'undefined' && process.env && process.env.AZURE_OPENAI_API_KEY) {
      console.log("Using API key from process.env");
      return process.env.AZURE_OPENAI_API_KEY;
    } 
    // Try REACT_APP_OPENAI_API_KEY 
    else if (typeof process !== 'undefined' && process.env && process.env.REACT_APP_OPENAI_API_KEY) {
      console.log("Using API key from REACT_APP_OPENAI_API_KEY");
      return process.env.REACT_APP_OPENAI_API_KEY;
    }
    // Try window.OPENAI_API_KEY
    else if (window.OPENAI_API_KEY) {
      console.log("Using API key from window.OPENAI_API_KEY");
      return window.OPENAI_API_KEY;
    }
    // Final fallback
    else {
      console.log("No API key found in any source");
      throw new Error("Azure OpenAI API key not configured. Please add your key to the .env file.");
    }
  } catch (e) {
    console.error("Error accessing API key:", e);
    throw new Error("Error accessing API key. Please check your configuration.");
  }
}

/**
 * Default available meeting times - these could be configured or loaded from a settings file
 */
const DEFAULT_MEETING_TIMES = [
  "Thursday, 27 March 2025 at 10:30am",
  "Thursday, 27 March 2025 at 1:30pm",
  "Thursday, 27 March 2025 at 2:30pm"
];

/**
 * Updates the status message in the specified container
 * @param {string} message - The message to display
 * @param {string} containerSelector - The selector for the status container
 */
function updateStatus(message, containerSelector = '#status') {
  const statusElement = document.querySelector(containerSelector);
  if (statusElement) {
    statusElement.innerText = message;
  }
}

/**
 * Creates and displays a new email message
 * @param {Object} emailData - Object containing email details
 * @param {string} emailData.subject - Email subject
 * @param {string} emailData.body - Email body text
 * @param {Array<string>} emailData.toRecipients - Array of recipient email addresses
 * @param {boolean} [emailData.openWindow=true] - Whether to open a new window (default: true)
 */
function createNewEmail(emailData) {
  // Only open a new email window if openWindow is true or not specified
  if (emailData.openWindow !== false) {
    Office.context.mailbox.displayNewMessageForm({
      toRecipients: emailData.toRecipients || [],
      subject: emailData.subject || "",
      htmlBody: emailData.body.replace(/\n/g, '<br>')
    });
  }
}

// Export functions for use in other modules
export {
  extractClientInfo,
  DEFAULT_MEETING_TIMES,
  updateStatus,
  createNewEmail,
  getApiKey
}; 