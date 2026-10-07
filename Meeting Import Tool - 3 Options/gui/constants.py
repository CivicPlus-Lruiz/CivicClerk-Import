"""
Application Constants
"""

CHECKLIST_ITEMS = [
    {"section": "Step 1: Pre-Import Checklist", "items": [
        {"text": "Did the customer sign their Import Agreement", "url": "https://implementation-civicclerk.app.transform.civicplus.com/forms/58093?parentTile=Implementation%20Forms"},
        {"text": "Has the customer confirmed ALL files have been uploaded to the Azure Storage?"},
        {"text": "Create a Local Folder for Customer's Import (This will be the Parent Folder)"},
        {"text": 'Create a "meetings" folder in Parent Folder, dump all files into "/meetings" folder'},
        {"text": "Are all files downloaded?"},
        {"text": "Are all files in their respective Meeting Body Folder?"},
        {"text": "Are all files in their respective File Type Folder?"}
    ]},
    {"section": "Step 2: Configuration", "items": [
        {"text": "In Tab 2 (Meeting Bodies): Browse and select the Base Directory (parent folder)"},
        {"text": "Click 'Auto-Scan' to automatically populate Meeting Bodies from /meetings folder"},
        {"text": "Verify Meeting Types match what the customer has on their site", "needs_siteid": True,
         "url": "https://{SITEID}.v8.civicclerk.com/SiteSettings/Meetings/MeetingTypes"},
        {"text": "If not, manually adjust Meeting Type in Tab 2"},
        {"text": "Verify Event Categories match what is on the customers site", "needs_siteid": True,
         "url": "https://{SITEID}.v8.civicclerk.com/SiteSettings/EventsAndPublicPortal/Events"},
        {"text": "If not, manually adjust Event Category in Tab 2"},
        {"text": "Set Default Time for each Meeting Body using time pickers"},
        {"text": "Select Date Format (US DATE or EU DATE) if needed"},
        {"text": "In Tab 3 (Subtypes) - OPTIONAL: Add file aliases (e.g., 'WorkSession', 'Special', 'Joint')"},
        {"text": "In Tab 3 (Subtypes) - OPTIONAL: Set Override Times for subtypes if different from default"}
    ]},
    {"section": "Step 3: Dry Run (Preview)", "items": [
        {"text": "In Tab 4 (Run): Click '🔍 Dry Run (Preview)' button at top right",
         "hint": "Dry Run shows errors but doesn't move files or create output"},
        {"text": "Review preview results in the table"},
        {"text": "Check Warnings and Duplicates tabs below for any issues"},
        {"text": "If errors appear, switch to Tab 5 (File Errors)"},
        {"text": "In Tab 5: For Ambiguous Dates - Rename files with correct date format"},
        {"text": "In Tab 5: For Weekend Dates - Confirm date or rename file"},
        {"text": "In Tab 5: For Duplicates - Select files and click 'Remove (Duplicates)' button"},
        {"text": "In Tab 5: For Unrecognized Formats - Files auto-moved, check warning"},
        {"text": "Return to Tab 4 and click 'Dry Run' again until no errors remain"}
    ]},
    {"section": "Step 4: Run Import", "items": [
        {"text": "Once Dry Run is clean, click '▶ Run Import' button at top right of Tab 4",
         "hint": "This will create the output file and move files to their final locations"},
        {"text": "Review the import results in the preview table"},
        {"text": "In Tab 4: Click '📊 Export Excel' button to save as Meetings_Output.xlsx"},
        {"text": "If there were renamed files, move files from the 'Move Files back to Azure Storage' folder into the customer's Azure Storage"}
    ]},
    {"section": "Step 5: Upload to CivicClerk", "items": [
        {"text": 'Upload "Meetings_Output.xlsx" and click "Import"',
         "hint": "This is the file exported from Tab 4 (Run)"},
        {"text": "Check for Errors in CivicClerk"},
        {"text": "If no errors you should see a progress bar"}
    ]},
    {"section": "Step 6: Indexing & Cleanup", "items": [
        {"text": "Once all events imported, navigate to Site Manager for indexing",
         "url": "https://civicclerksitemanager.azurewebsites.net/"},
        {"text": "Check the following boxes for indexing:"},
        {"text": "Force Re-index document content"},
        {"text": "Agenda Indexer"},
        {"text": "Agenda Items Indexer"},
        {"text": "Agenda Files Indexer"},
        {"text": "Attachment Indexer"},
        {"text": "Event Indexer"},
        {"text": 'Click "Start Indexing"'},
        {"text": "Inform the customer that import is complete and files are indexing"},
        {"text": "Open Folder Site Deletion List and move Customer Name or Site ID to bottom",
         "url": "https://civicplus.sharepoint.com/:x:/r/sites/CPOregonWebMeetingsTeam/Shared%20Documents/Operations%20-%20Meetings/_Luis/FTP%20Logins.xlsx?d=wdf801a75c8154734b09c8c9e8d076ce2&csf=1&web=1&e=cahH2U"},
        {"text": "Congrats you are done with the import!", "type": "celebration"}
    ]}
]
