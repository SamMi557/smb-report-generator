
# SMB Lab Report Generator – Google Apps Script Automation

This project is a fully automated report generation system built using Google Apps Script. It pulls data from a Google Form (via a linked Sheet), maps and cleans it, and generates a structured report in Google Docs – perfect for creating company summaries, client onboarding, or academic portfolio cases.

## Features

- 1. **Live Report Generation** on Google Form submission
- 2. **Date-Range Filtering** with a custom Sheets menu
- 3. **Test Function** to simulate submissions for offline development
- 4. **Field Mapping Logic** to clean and convert messy form keys to report-friendly placeholders
- 5. **Dynamic Google Docs Output** with placeholder replacement
- 6. Organized output saved to a specified Google Drive folder

## Use Case

Built for the **CSMB at ASU**, but adaptable to:
- Client onboarding reports
- Intake summaries for nonprofits or incubators
- Academic grading rubrics
- Custom analytics summaries

## Tech Stack

- **Google Apps Script** (JavaScript variant)
- **Google Sheets** – for response intake
- **Google Forms** – for data collection
- **Google Docs** – templated report output
- **Google Drive API** – file creation and storage

## Key Components

### 1. `onFormSubmit(e)`
Runs automatically on each Google Form submission. It:
- Cleans up field names via `cleanResponseKeys()`
- Copies a Google Docs template
- Replaces all placeholders like `{{Company Name}}` with actual data
- Saves the new Doc to Drive

### 2. `generateReportsInDateRange(startDate, endDate)`
Prompts user in Sheets to enter a start and end date. Generates reports only for rows within the date range (based on `"Date"` column).

### 3. `testCreateReport()`
Simulates a full submission event with fake data. Perfect for local testing without filling the form each time.

## Folder Structure

```
SMB-Report-Automation/
├── Google Sheet (Form linked)
├── Google Form (Client Intake)
├── Google Docs Template (with {{placeholders}})
├── Report Output Folder
└── Code.gs (this script)
```

## Setup Instructions

1. **Create a Google Form** and link it to a Sheet
2. **Create a Google Docs template** with placeholders like `{{Company Name}}`
3. **Create a Drive folder** to store generated reports
4. **Copy this script into the Sheet's Apps Script editor**
5. Replace `TEMPLATE_ID` and `DESTINATION_FOLDER_ID` with your real values
6. Run `testCreateReport()` to verify setup
7. Set a trigger for `onFormSubmit(e)` under **Triggers → Add Trigger**


## Screenshots

Check screenshots of the Google Form, Sheet, Menu UI, Generated Reports



## Testing

- Run `testCreateReport()` to validate the output
- Use the `Generate Reports → By Date Range` menu in Sheets to create historical reports



## Template Example

Your Google Docs template should include placeholders like:

```
Company Name: {{Company Name}}
Founder Story: {{Founder Story}}
Mission: {{Mission}}
Revenue: {{Revenue}}
Signed by: {{Digital Signature}}, Date: {{Date}}
```

## Contact / Credit

Built by Saurabh Mishra as part of the ASU CSMB Lab.  
For feedback, contact Saurabh_Mishra@asu.edu


## License

Licensed by Saurabh Mishra
