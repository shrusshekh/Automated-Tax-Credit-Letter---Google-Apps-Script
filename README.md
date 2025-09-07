# Automated-Tax-Credit-Letter - Google Apps Script
This project automates the creation of personalized tax credit letters directly from a Google Sheet using Google Apps Script. This project was developed as part of my summer job to streamline the process of generating personalized letters for my coworkers. To make the script understandable for non-technical users I added clear in-code comments indicating which parts of the script can and cannot be changed, making it easy for my coworkers to adapt without breaking the automation.

Below is a guide I prepared for my coworkers following the completion of this project. 

## **How Do You Use this Automation?**
1. Open the spreadsheet associated with the Google Apps Script.
2. An “Automate” menu appears at the top.
3. Select the rows you want to generate a letter for. 
4. Click “Generate Letters” from the “Automate” menu.
5. Wait. Each letter takes 1 second to generate. A message will appear once the automation has been completed. 
6. The selected letters will be generated into a PDF format and saved into the specified Google Drive folder. 
7. Every time a letter is generated, it creates a timestamp of when that letter was generated to the right side of the sheet. Double clicking on it will show the exact time of generation. 

## **What You Can Safely Edit**
Note: JavaScript is case-sensitive which means anything you rename or change in the document template or spreadsheet should have an exact replication in the Apps Script. 
Sheet Tab Name

If you rename the spreadsheet tab (e.g. from “Sheet1” to “Co-op Students”), you must update the line highlighted in red to the name of the tab. 
const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");

### Google Doc Template ID
If you're using a different template:
Copy the new Google Docs File ID (the long string in the URL after /d/)
Replace this line:
const templateId = "1_u1YSDQbunU7-cLDSwvmOynpMH2nYPskDSlp7IJKMhc";

### Google Drive Folder ID
To save the generated files to a different folder, copy that folder’s ID from its URL (after /folder/) and replace this line: 
const folderId = "1GIyeODY9oH6ZmPr74dakXMUWsT9Tpz5m";

### Changing the Names of the Files Generated
When the letters are generated, they are saved with the name “Tax Credit Letter - [Student Name]”.
If you want to change how these letters are named (for example, to include employer name or program name), look for this line in the script (line 34):
.makeCopy(`Tax Credit Letter - ${rowData["studentName"]}`, folder);

### Adding a New Column in the Spreadsheet
If your sheet and template include a new variable (i.e. jobTitle), you’ll need to:
Add a new body.replaceText AFTER the last body.replaceText. For jobTitle:
body.replaceText("{{jobTitle}}", rowData["jobTitle"]);

1. Make sure the new variable exists (i.e. jobTitle):
2. In your sheet header (no extra spaces or typos)
3. As a placeholder (in curly brackets) in your template Doc → {{jobTitle}}

## **What You Should Not Change**
Unless you know JavaScript, do not change:
1. Anything labeled “DO NOT CHANGE” in the code.
2. The order of replaceText() calls. 
3. Placeholder names unless you also update your Google Doc to match. 
4. DO NOT copy paste placeholders in the document template as this will carry over any hidden formatting, making it unreadable by Apps Script. If you are copy pasting anything in the template, ensure you remove formatting:
- Highlight the text you have copy pasted.
- Navigate to “Format” on Google Docs. 
- Click “Clear Formatting”). 

