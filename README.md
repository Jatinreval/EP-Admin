# EP Admin - Outlook Email Assistant

An Outlook Add-in that leverages Azure OpenAI to generate professional legal emails for various scenarios including initial correspondence, conference scheduling, and reminder emails.

## Features

- **AI-Powered Email Generation**: Uses Azure OpenAI to create contextually appropriate legal emails
- **Multiple Email Templates**:
  - Initial Email - For first client communications
  - Conference for Draft Documents - For scheduling document review meetings
  - Conference for Signing Documents - For scheduling document signing
  - Multiple Reminder Types - For following up on various stages of legal processes
- **Smart Context Extraction**: Automatically analyzes the content of received emails
- **One-Click Email Responses**: Generate and reply with professionally formatted emails
- **HTML Formatting**: All generated emails are properly formatted with HTML styling
- **Works Across Platforms**: Compatible with Outlook Desktop (Windows/Mac) and Outlook Web

## Prerequisites

- [Node.js](https://nodejs.org) (LTS version recommended)
- Microsoft Outlook (Desktop or Web)
- OpenSSL (for certificate generation during development)
- Azure OpenAI API access

## Project Structure

```
EP-Admin/
├── src/
│   ├── commands/         # Menu commands and ribbon integration
│   ├── handlers/         # Email generation logic for different scenarios
│   ├── taskpane/         # Main UI interface
│   └── utils/            # Shared utility functions
├── assets/               # Images and static resources
├── certs/                # Development certificates
├── manifest.xml          # Add-in manifest for production
└── dev-manifest.xml      # Add-in manifest for development
```

## Setup for Development

1. Clone this repository
2. Navigate to the project directory
3. Install the dependencies:
   ```
   npm install
   ```
4. Generate self-signed certificates for HTTPS development:
   ```
   # Create certs directory
   mkdir -p certs
   cd certs

   # Generate server key and certificate
   openssl req -newkey rsa:2048 -nodes -keyout server.key -x509 -days 365 -out server.crt -subj "/CN=localhost"

   # Create certificate serial number file
   echo "01" > ca.srl

   # Generate CA key
   openssl genrsa -out ca.key 2048

   # Generate CA certificate
   openssl req -new -x509 -key ca.key -out ca.crt -days 365 -subj "/CN=localhost"

   # Return to project root
   cd ..
   ```
5. Configure Azure OpenAI credentials:
   - Create a `.env` file in the project root
   - Add your Azure OpenAI credentials:
     ```
     AZURE_OPENAI_ENDPOINT=your_full_endpoint_url_here  # e.g. https://your-resource-name.openai.azure.com/openai/deployments/your-deployment/chat/completions?api-version=2024-02-15-preview
     AZURE_OPENAI_API_KEY=your_api_key_here
     ```
6. Start the development server:
   ```
   npm start
   ```

## Azure OpenAI Setup

1. Create an Azure OpenAI resource in the [Azure Portal](https://portal.azure.com)
2. Deploy a model in your Azure OpenAI resource (recommended: GPT-4o)
3. Get your full endpoint URL (including deployment name and API version) and API key from the Azure Portal
4. Add these credentials to your `.env` file as shown above

## Sideloading the Add-in in Outlook

### Outlook Desktop (Windows)
1. Open Outlook
2. Click on the gear icon (Settings) and select "Manage Add-ins"
3. Click "My add-ins" in the left navigation
4. Click the "+" icon and select "Add from file..."
5. Browse to the `dev-manifest.xml` file for local development or `manifest.xml` for production
6. Follow the prompts to install the add-in

### Outlook Desktop (Mac)
1. Open Outlook
2. Click "Tools" > "Add-ins"
3. Click "My add-ins" in the left navigation
4. Click the "+" icon and select "Add from file..."
5. Browse to the `dev-manifest.xml` file for local development or `manifest.xml` for production
6. Follow the prompts to install the add-in

### Outlook Web
1. Go to [Outlook Web](https://outlook.office.com)
2. Click on the gear icon (Settings) and select "Manage Add-ins"
3. Click "My add-ins" in the left navigation
4. Click the "+" icon and select "Add from file..."
5. Browse to the `dev-manifest.xml` file for local development or `manifest.xml` for production
6. Follow the prompts to install the add-in

## Usage

1. Open Outlook and select an email or create a new message
2. Click the "EP Admin" button in the ribbon to open the taskpane
3. In the taskpane:
   - Select the type of email you want to generate from the dropdown menus
   - Click "Generate" to create the email using AI
   - Review the generated content
   - Click "Reply with Response" to insert the content into a reply
   - Or click "Copy" to copy the content to your clipboard

### Available Email Templates

- **Sending Times**:
   - Initial Email: Generate a response to a client's first contact
   - Conference for Draft Documents: Schedule a meeting to discuss draft documents
   - Conference for Signing Documents: Schedule a meeting for document signing
- **Reminders**:
  - Initial Email Reminder: Follow up on an unanswered initial contact
  - Draft Document Reminder: Remind about reviewing draft documents
  - Signing Document Reminder: Remind about signing documents
  - Attorney/Guardian Reminder: Specific reminder for enduring attorney or guardian appointments
  - Further Information Reminder: Request additional information

## Deployment

To deploy the add-in for production:

1. Build the production version:
   ```
   npm run build
   ```
2. Host the built files on a secure HTTPS server
3. Update the URLs in `manifest.xml` to point to your hosted version
4. Share the `manifest.xml` file with users who need to install the add-in

## Troubleshooting

### Certificate Issues
If you encounter certificate-related errors during development:

1. Make sure you've generated all the required certificate files in the `certs` directory
2. Verify that the certificates are valid and not expired
3. Check that the certificates are properly referenced in your development environment

### Azure OpenAI Issues
If you encounter issues with the Azure OpenAI integration:

1. Verify that your Azure OpenAI credentials in the `.env` file are correct
2. Check that your Azure OpenAI deployment is active and available
3. Ensure that you have sufficient quota available for your Azure OpenAI resource
4. Check the browser console for any API-related error messages

### UI Issues
If buttons are not working correctly:

1. Check the browser console for JavaScript errors
2. Make sure the latest version of the add-in is installed
3. Try reloading the add-in or restarting Outlook

## License

ISC
