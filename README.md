# OAuthEmailer

OAuthEmailer is a Node.js-based email automation tool that reads data from an Excel sheet and sends personalized emails using Gmail's OAuth2 for secure authentication. This tool is ideal for bulk email sending, such as job applications or personalized announcements.

## Features

- Reads recipient data from an Excel file.
- Sends personalized emails with dynamic content.
- Uses Gmail's OAuth2 for secure and reliable email sending.
- Includes random delays to avoid spam detection.

## Prerequisites

Before running this project, ensure you have:

1. **Node.js** installed on your system. [Download here](https://nodejs.org/).
2. **Gmail OAuth2 credentials**:
   - Create a project in the [Google Cloud Console](https://console.cloud.google.com/).
   - Enable the Gmail API for your project.
   - Generate OAuth2 credentials (Client ID, Client Secret, Redirect URI, and Refresh Token).
3. An Excel file (`Senior Software Engineer.xlsx`) with the required data.

## Installation

### 1. Clone the Repository


git clone https://github.com/yourusername/OAuthEmailer.git
cd OAuthEmailer



## 2. Install Dependencies

npm install

## 3. Configure Environment Variables

Create a .env file in the root directory with the following content:

CLIENT_ID=your_client_id
CLIENT_SECRET=your_client_secret
REDIRECT_URI=your_redirect_uri
REFRESH_TOKEN=your_refresh_token
ACCESSTOKEN=your_access_token

Replace ***your_client_id*** , ***your_client_secret***, ***your_redirect_uri***, ***your_refresh_token***, and ***your_access_token*** with the values from your Google Cloud Console project.
 
## 4. Add Your Excel File
Ensure your Excel file (Senior Software Engineer.xlsx) is in the root directory.

Usage
To send emails, run:

```bash
node index.js

The script will:

Read data from Senior Software Engineer.xlsx.
Use Gmail OAuth2 to authenticate and send emails.
Log the status of each email sent.

# File Structure

## File Structure

- `OAuthEmailer/`
  - `index.js`  # Main script for sending emails
  - `package.json`  # Dependencies and scripts
  - `.env`  # Environment variables for OAuth configuration
  - `Senior Software Engineer.xlsx`  # Excel file with email data
  - `README.md`  # Documentation


# Dependencies

- `xlsx`: For reading Excel files.
- `nodemailer`: For sending emails.
- `googleapis`: For Gmail OAuth2 integration.
- `dotenv`: For managing environment variables.

# Customization

- Excel Columns: Update the sendEmail function to match your Excel column names.
- Delay Between Emails: Adjust the delay logic in sendEmailsSynchronously for optimal email sending.

# Notes

- Gmail has daily limits for sending emails. Be cautious when sending large batches.
- Double-check your OAuth2 setup to ensure valid credentials.
