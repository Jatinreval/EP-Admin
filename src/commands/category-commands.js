/**
 * Category-specific commands for the Outlook add-in
 * This file contains functions and templates for different email categories
 */

// Templates for different email types
const EMAIL_TEMPLATES = {
  // Conference for Draft Documents template
  DRAFT_CONFERENCE: {
    subject: "Conference for Draft Documents - [Client Name]",
    body: `Private and Confidential

Dear [Client Name],

I am writing to confirm that we have prepared the draft documents for your review.

Conference
I would like to schedule a conference to discuss these documents with you. Please let me know if any of the following times would suit you:
- [Date/Time Option 1]
- [Date/Time Option 2]
- [Date/Time Option 3]

The conference can be held via Microsoft Teams or in person at our office.

Please reply to all recipients to confirm your preferred time.

Yours faithfully,
[Attorney Name]
[Title]
`
  },
  
  // Conference for Signing Documents template
  SIGNING_CONFERENCE: {
    subject: "Conference for Signing Documents - [Client Name]",
    body: `Private and Confidential

Dear [Client Name],

I am writing to confirm that your documents are now ready for signing.

Conference
I would like to schedule a conference for the signing. Please let me know if any of the following times would suit you:
- [Date/Time Option 1]
- [Date/Time Option 2]
- [Date/Time Option 3]

The conference should be held in person at our office.

Please reply to all recipients to confirm your preferred time.

Yours faithfully,
[Attorney Name]
[Title]
`
  },
  
  // Reminder templates
  REMINDER_INITIAL: {
    subject: "Reminder: Initial Email - [Client Name]",
    body: `Private and Confidential

Dear [Client Name],

This is a friendly reminder regarding our initial email sent to you on [Date]. We haven't received a response yet and wanted to ensure you received our previous communication.

Please let us know if you would like to proceed with the matter or if you have any questions.

Please reply to all recipients.

Yours faithfully,
[Attorney Name]
[Title]
`
  },
  
  REMINDER_FURTHER: {
    subject: "Reminder: Further Information Required - [Client Name]",
    body: `Private and Confidential

Dear [Client Name],

This is a reminder that we are still waiting to receive the following information from you:
- [Outstanding Item 1]
- [Outstanding Item 2]
- [Outstanding Item 3]

We cannot proceed with your matter until we receive this information. Please provide the above details at your earliest convenience.

Please reply to all recipients.

Yours faithfully,
[Attorney Name]
[Title]
`
  },
  
  REMINDER_DRAFT: {
    subject: "Reminder: Conference for Draft Documents - [Client Name]",
    body: `Private and Confidential

Dear [Client Name],

This is a friendly reminder regarding our scheduled conference to discuss your draft documents.

Conference Details:
Date: [Conference Date]
Time: [Conference Time]
Format: [Conference Format]

Please confirm your attendance or let us know if you need to reschedule.

Please reply to all recipients.

Yours faithfully,
[Attorney Name]
[Title]
`
  },
  
  REMINDER_SIGNING: {
    subject: "Reminder: Conference for Signing Documents - [Client Name]",
    body: `Private and Confidential

Dear [Client Name],

This is a friendly reminder regarding our scheduled conference for signing your documents.

Conference Details:
Date: [Conference Date]
Time: [Conference Time]
Location: Our office

Please remember to bring appropriate identification documentation.

Please reply to all recipients.

Yours faithfully,
[Attorney Name]
[Title]
`
  },
  
  REMINDER_ATTORNEY: {
    subject: "Reminder to Attorney/Guardian - [Client Name] Matter",
    body: `Private and Confidential

Dear Attorney/Guardian,

This is a reminder regarding [Client Name]'s matter. We require your attention on the following:
- [Action Item]

Your prompt attention to this matter would be greatly appreciated.

Please reply to all recipients.

Yours faithfully,
[Attorney Name]
[Title]
`
  }
};

/**
 * Generates an email draft based on a template and client information
 * @param {Object} client - Client information
 * @param {String} templateKey - Key for the template to use
 * @param {Array} availableTimes - Array of available meeting times
 * @returns {Object} - Object containing subject and body
 */
function generateEmailFromTemplate(client, templateKey, availableTimes = []) {
  if (!EMAIL_TEMPLATES[templateKey]) {
    throw new Error(`Template "${templateKey}" not found`);
  }
  
  const template = EMAIL_TEMPLATES[templateKey];
  let subject = template.subject;
  let body = template.body;
  
  // Replace placeholders with client information
  subject = subject.replace('[Client Name]', client.name || 'Client');
  body = body.replace(/\[Client Name\]/g, client.name || 'Client');
  
  // Replace attorney information
  body = body.replace('[Attorney Name]', Office.context.mailbox.userProfile.displayName || 'Attorney');
  body = body.replace('[Title]', client.attorneyTitle || 'Attorney');
  
  // Replace available times if provided
  if (availableTimes && availableTimes.length > 0) {
    const timesString = availableTimes.map(time => `- ${time}`).join('\n');
    body = body.replace(/- \[Date\/Time Option 1\]\n- \[Date\/Time Option 2\]\n- \[Date\/Time Option 3\]/g, timesString);
  }
  
  // Replace today's date if needed
  body = body.replace(/\[Date\]/g, new Date().toLocaleDateString('en-US', { 
    weekday: 'long', 
    year: 'numeric', 
    month: 'long', 
    day: 'numeric' 
  }));
  
  return {
    subject,
    body
  };
}

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
        subject: item.subject || ""
      };
      
      // Try to extract name from greeting
      item.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const bodyText = result.value;
          
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

// Export functions for use in taskpane.js
export {
  EMAIL_TEMPLATES,
  generateEmailFromTemplate,
  extractClientInfo
}; 