import { saveAs } from 'file-saver';

// Import all handlers
import initialEmailHandler from '../handlers/initial-email';
import conferenceDraftHandler from '../handlers/conference-draft';
import conferenceSigningHandler from '../handlers/conference-signing';
import reminderInitialHandler from '../handlers/reminder-initial';
import reminderFurtherHandler from '../handlers/reminder-further';
import reminderDraftHandler from '../handlers/reminder-draft';
import reminderSigningHandler from '../handlers/reminder-signing';
import reminderAttorneyHandler from '../handlers/reminder-attorney';

// Map task IDs to their respective handlers
const taskHandlers = {
  'initial-email': initialEmailHandler,
  'conference-draft': conferenceDraftHandler,
  'conference-signing': conferenceSigningHandler,
  'reminder-initial': reminderInitialHandler,
  'reminder-further': reminderFurtherHandler,
  'reminder-draft': reminderDraftHandler,
  'reminder-signing': reminderSigningHandler,
  'reminder-attorney': reminderAttorneyHandler
};

// Active handler reference
let activeHandler = null;

Office.onReady((info) => {
  // Check if we're in Outlook
  if (info.host === Office.HostType.Outlook) {
    // Show initial loader
    showLoader();
    
    // Setup navigation
    setupNavigation();
    
    // Add event handlers for common buttons
    document.getElementById("save-button").onclick = saveEmailAsJson;
    
    // Initialize the default handler (Initial Email)
    activateHandler('initial-email');
    
    // Hide the loader after initialization
    hideLoader();
    
    console.log("Add-in initialized in Outlook");
  }
});

/**
 * Shows the loading animation
 */
function showLoader() {
  const loaderContainer = document.getElementById('loader-container');
  if (loaderContainer) {
    loaderContainer.style.display = 'flex';
  }
}

/**
 * Hides the loading animation
 */
function hideLoader() {
  const loaderContainer = document.getElementById('loader-container');
  if (loaderContainer) {
    loaderContainer.style.display = 'none';
  }
}

/**
 * Activates the specified handler
 * @param {string} taskId - The ID of the task to activate
 */
function activateHandler(taskId) {
  // Get the handler
  const handler = taskHandlers[taskId];
  
  if (handler) {
    // Save reference to active handler
    activeHandler = handler;
    
    try {
      // Initialize the handler
      handler.initialize();
      
      console.log(`Activated handler for ${taskId}`);
    } catch (error) {
      console.error(`Error initializing handler for ${taskId}:`, error);
      
      // Hide loader in case of error
      hideLoader();
    }
  } else {
    console.error(`No handler found for task ${taskId}`);
    
    // Hide loader if no handler found
    hideLoader();
  }
}

/**
 * Sets up the navigation between different sections
 */
function setupNavigation() {
  // Get all dropdown links
  const dropdownLinks = document.querySelectorAll('.dropdown-content a');
  
  // Add click event listeners to each link
  dropdownLinks.forEach(link => {
    link.addEventListener('click', function(e) {
      e.preventDefault();
      
      // Get the task ID from data attribute
      const taskId = this.getAttribute('data-task');
      
      // Show the loader when clicking a menu item
      showLoader();
      
      // Hide all task sections
      document.querySelectorAll('.task-section').forEach(section => {
        section.classList.remove('active');
      });
      
      // Show the selected section
      const selectedSection = document.getElementById(taskId);
      if (selectedSection) {
        selectedSection.classList.add('active');
      }
      
      // Clear status messages
      document.querySelectorAll('#status, .reminder-status, .draft-status, .signing-status').forEach(el => {
        if (el) el.innerText = '';
      });
      
      // Hide response container when switching sections
      const responseContainer = document.getElementById('response-container');
      if (responseContainer) {
        responseContainer.style.display = 'none';
      }
      
      // Hide buttons when switching sections
      const copyButton = document.getElementById('copy-button');
      const replyButton = document.getElementById('reply-button');
      if (copyButton) copyButton.style.display = 'none';
      if (replyButton) replyButton.style.display = 'none';
      
      // Activate the handler for this task
      activateHandler(taskId);
      
      // Hide the loader as UI has updated and handler is initialized
      hideLoader();
    });
  });
  
  // Add dropdown toggle behavior
  document.addEventListener('DOMContentLoaded', function() {
    // Get all dropdown buttons
    const dropdownBtns = document.querySelectorAll('.dropdown-btn');
    
    // Add click event to each dropdown button
    dropdownBtns.forEach(btn => {
      btn.addEventListener('click', function(e) {
        e.preventDefault();
        e.stopPropagation();
        
        // Get the parent dropdown container
        const dropdown = this.parentElement;
        
        // Check if this dropdown is currently open
        const isOpen = dropdown.classList.contains('open');
        
        // Close all dropdowns first
        document.querySelectorAll('.dropdown').forEach(d => {
          d.classList.remove('open');
        });
        
        // If it wasn't open before, open it now
        if (!isOpen) {
          dropdown.classList.add('open');
        }
      });
    });
    
    // Close dropdown when clicking anywhere else on the page
    document.addEventListener('click', function() {
      document.querySelectorAll('.dropdown').forEach(d => {
        d.classList.remove('open');
      });
    });
  });
}

/**
 * Gets the current email item and extracts its content to save as JSON
 */
function saveEmailAsJson() {
  const statusElement = document.getElementById("status");
  statusElement.innerText = "Processing...";

  try {
    // Get the current item (email)
    const item = Office.context.mailbox.item;
    
    if (!item) {
      statusElement.innerText = "No email selected";
      return;
    }

    // Create an object to store email data
    const emailData = {
      subject: item.subject,
      sender: item.sender ? item.sender.emailAddress : "Unknown",
      receivedTime: item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : null,
      bodyContent: null
    };

    // Get the email body content
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        // Add the body text to our data object
        emailData.bodyContent = result.value;
        
        // Convert the email data to a JSON string
        const jsonData = JSON.stringify(emailData, null, 2);
        
        // Create a Blob from the JSON string
        const blob = new Blob([jsonData], { type: "application/json" });
        
        // Generate a filename using the subject (or a default name if no subject)
        const filename = `${emailData.subject || "email"}_${new Date().getTime()}.json`;
        
        // Save the file
        saveAs(blob, filename);
        
        statusElement.innerText = "Email saved as JSON successfully!";
      } else {
        statusElement.innerText = `Error getting email body: ${result.error.message}`;
      }
    });
  } catch (error) {
    statusElement.innerText = `Error: ${error.message}`;
  }
} 