# Generate Import Sensei Lesson from SharePoint

A simple command-line tool to generate a CSV file containing lessons from SharePoint videos, ready to be imported into the [Sensei LMS Plugin WordPress](https://wordpress.org/plugins/sensei-lms/).

## Description

This tool connects to the Microsoft Graph API to access video files shared via SharePoint, then creates a CSV file containing information about the lessons, including titles and video embed codes. This CSV file can be used to bulk import lessons into a learning management system.

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/generate-import-sensei-lesson-from-sharepoint.git
   cd generate-import-sensei-lesson-from-sharepoint
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Create an `.env` file from the template:
   ```bash
   cp .env.template .env
   ```

4. Update the `.env` file with your Microsoft Graph API authentication information:
   ```
   TENANT_ID=your_tenant_id
   CLIENT_ID=your_client_id
   CLIENT_SECRET=your_client_secret
   ```

## Usage

1. Run the application:
   ```bash
   npm start
   ```

2. When prompted, enter the SharePoint sharing URL containing the lesson videos:
   ```bash
   Input your sharepoint url: https://your-sharepoint-url.com/shared-folder
   ```

3. The application will generate a CSV file in the `output` directory with the filename format `data-YYYY-MM-DD_HH-MM.csv`.

## Output CSV Structure

This file used to import lession to Sensei LMS

The output CSV file will have the following columns:
- **Lesson**: The lesson title (derived from the video filename)
- **Description**: HTML embed code to display the video
- **Status**: Publication status (defaults to "publish")
- **Prerequisite**: Prerequisites (defaults to empty)

Reference: [Lesson Import Schema](
https://github.com/Automattic/sensei/wiki/Lesson-Import-Schema)

## Requirements

- Node.js v18
- Microsoft account with access to SharePoint
- Registered application in Azure AD with Microsoft Graph API access permissions

## License

ISC
