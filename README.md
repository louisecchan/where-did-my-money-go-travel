# Travel Bookings Tracker for Gmail (Agoda

A Google Apps Script that pulls your Agoda booking details from Gmail and saves them into a Google Sheet. It helps you keep track of your travel spending automatically.

## How It Works

1. Searches Gmail for Agoda booking confirmations.
2. Extracts details like hotel name, dates, price, and location.
3. Adds the data to a connected Google Sheet.

## Getting Started

### Prerequisites

- A Google account
- Access to Gmail and Google Sheets
- Basic familiarity with Google Apps Script

### Setup

1. Go to [Google Apps Script](https://script.google.com/) and create a new project.
2. Copy-paste the code from [`code.gs`](https://github.com/louisecchan/where-did-my-money-go-travel/blob/main/code.gs).
3. Create a new Google Sheet to hold your data.
4. In the script, replace the `SPREADSHEET_ID` with your Sheet's ID.
5. Run the `main` function.
6. Grant script permissions when prompted.

## Customization

- Modify the Gmail search query to target other booking platforms.
- Adjust the script to extract different data points.
- Format the Google Sheet for better readability or summaries.

## Troubleshooting

- Make sure there are Agoda booking emails in your inbox.
- Double-check the Google Sheet ID.
- Review the execution logs if the script fails.
- Ensure you’ve granted Gmail and Sheets access.

