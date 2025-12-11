# Service-to-Sales Bridge Dashboard
## Setup Instructions

This guide will walk you through setting up the Service-to-Sales Bridge Dashboard for Union Park Buick GMC.

---

## Prerequisites

- A Google account with access to Google Sheets
- The email addresses for authorized users:
  - Brian Callahan: `bcallahan@unionparkgmc.com`
  - Dan Testa: `dtesta620@gmail.com`

---

## Step 1: Create a New Google Sheet

1. Go to [Google Sheets](https://sheets.google.com)
2. Click **"+ Blank"** to create a new spreadsheet
3. Name it: **"Union Park - Service-to-Sales Bridge Dashboard"**

---

## Step 2: Set Up Apps Script

1. In your new Google Sheet, go to **Extensions > Apps Script**
2. This opens the Apps Script editor in a new tab

### 2a: Replace the Code

1. Delete any existing code in `Code.gs`
2. Copy the **entire contents** of the `Code.gs` file from this repository
3. Paste it into the Apps Script editor

### 2b: Update the Manifest

1. In the Apps Script editor, click on **Project Settings** (gear icon) in the left sidebar
2. Check the box: **"Show 'appsscript.json' manifest file in editor"**
3. Go back to the Editor view
4. Click on `appsscript.json` in the left sidebar
5. Replace its contents with the `appsscript.json` file from this repository
6. Click **Save** (or press Ctrl+S)

---

## Step 3: Run Initial Setup

1. In the Apps Script editor, select **`initialSetup`** from the function dropdown (next to the Run button)
2. Click **Run**
3. You'll be prompted to authorize the script:
   - Click **"Review Permissions"**
   - Choose your Google account
   - Click **"Advanced"** (if you see a warning)
   - Click **"Go to Union Park - Service-to-Sales Bridge Dashboard (unsafe)"**
   - Click **"Allow"**

4. Wait for the setup to complete (may take 30-60 seconds)
5. You should see a confirmation dialog when done

---

## Step 4: Share the Spreadsheet

1. Go back to your Google Sheet
2. Click the **"Share"** button (top right)
3. Add the following editors:
   - `bcallahan@unionparkgmc.com`
   - `dtesta620@gmail.com`
4. Set their permission to **"Editor"**
5. Uncheck "Notify people" if you want to inform them separately
6. Click **"Send"** or **"Share"**

---

## Step 5: Verify Setup

After running `initialSetup`, verify that:

### Sheets Created
You should see 6 tabs at the bottom:
- Dashboard
- Dealer Trade Re-PDIs
- Customer Accessory Installs
- New Car Parts Installation
- Service Drive Appraisals
- Completed Archive

### Test the Menu
1. Refresh the Google Sheet page
2. You should see a custom menu: **"Bridge Dashboard"**
3. Click it to see available options:
   - Refresh Dashboard
   - Run Archive Now
   - Setup Sheet Structure
   - Apply Protections
   - Setup Triggers

### Test Email Notifications
1. Go to **Extensions > Apps Script**
2. Select **`testEmailNotification`** from the function dropdown
3. Click **Run**
4. Check your email for a test notification

---

## Step 6: Configure Triggers (Optional Verification)

Triggers should be automatically set up, but you can verify:

1. In Apps Script, click on **Triggers** (clock icon) in the left sidebar
2. You should see 3 triggers:
   - `onEditHandler` - runs on every edit
   - `archiveCompletedItems` - runs daily at 2 AM
   - `refreshDashboard` - runs every hour

---

## Troubleshooting

### "Authorization required" errors
- Make sure you've authorized all requested permissions
- Try running `initialSetup` again

### Emails not being sent
- Check that the email addresses in `CONFIG` are correct
- Verify the script has the `mail.google.com` scope
- Check the Executions log in Apps Script for errors

### Menu not appearing
- Refresh the Google Sheet
- Make sure the script is saved
- Try running `onOpen` manually from Apps Script

### Protections not working
- Run `applyProtections` from the Bridge Dashboard menu
- Make sure both authorized emails are added as editors to the sheet

### Conditional formatting not applied
- Run `applyConditionalFormatting` from Apps Script editor

---

## Updating the Script

If you need to update the script later:

1. Go to **Extensions > Apps Script**
2. Make your changes
3. Click **Save**
4. If you added new triggers, run `setupTriggers`
5. If you changed sheet structure, run `setupAllSheets`

---

## Support

If you encounter issues:

1. Check the **Executions** log in Apps Script (click on "Executions" in left sidebar)
2. Review error messages in the log
3. Verify all email addresses and permissions are correct

---

## Quick Reference

| Function | What It Does |
|----------|--------------|
| `initialSetup` | Full setup - run this first! |
| `setupAllSheets` | Creates/resets all sheet tabs |
| `applyProtections` | Locks sheets to authorized users |
| `setupTriggers` | Sets up automatic triggers |
| `refreshDashboard` | Updates dashboard counts |
| `archiveCompletedItems` | Archives old completed items |
| `applyConditionalFormatting` | Applies color rules |
| `testEmailNotification` | Sends test email to yourself |
| `checkAuthorization` | Shows your auth status |

---

## Customization

### Changing Authorized Users

Edit the `CONFIG` object at the top of `Code.gs`:

```javascript
const CONFIG = {
  BRIAN_EMAIL: 'bcallahan@unionparkgmc.com',
  DAN_EMAIL: 'dtesta620@gmail.com',
  // ...
};
```

### Changing Archive Period

```javascript
ARCHIVE_DAYS: 7,  // Change to desired number of days
```

### Changing Priority Thresholds

```javascript
URGENT_HOURS: 24,       // Delivery within 24 hours = URGENT
HIGH_PRIORITY_HOURS: 48 // Delivery within 48 hours = HIGH
```

### Changing Colors

Modify the `COLORS` object in `CONFIG` to match your brand.
