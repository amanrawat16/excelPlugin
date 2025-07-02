# Invalid Contact Checker - Excel Add-in

This is a Microsoft Excel plugin built with React and Office.js. It helps you quickly check a table of leads or contacts for missing or invalid data.

## What this plugin does
- Adds a sample table of leads/contacts to the worksheet for testing.
- Checks each row for missing or invalid Name, Email, Phone, or Company fields.
- Highlights valid rows in green and invalid rows in yellow in Excel.
- Shows a summary and details of invalid rows in the task pane.
- Lets you clear all background color formatting from the selected range.

## Technologies used
- React (frontend, task pane UI)
- Office.js (Excel integration)
- JavaScript

## How to use
1. Click **Add Sample Data** to insert a demo table (A1:D7).
2. Select the table (A1:D7) and click **Check Contacts**.
3. The plugin will highlight valid/invalid rows and show a summary in the task pane.
4. Use **Clear Formatting** to remove highlights.

---

This project is a proof-of-concept for building Excel plugins with React and Office. 