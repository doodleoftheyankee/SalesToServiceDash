# Service-to-Sales Bridge Dashboard

A Google Sheets + Apps Script solution for Union Park Buick GMC that bridges communication between Sales and Service departments.

## Problem

The Sales and Service departments are physically separated by over a block, causing:
- Missed communications on time-sensitive dealer trades needing re-PDI
- Lost customer accessory orders requiring technician installation
- Delays in parts installation for new vehicles before delivery
- Missed appraisal opportunities from the service drive

## Solution

A centralized Google Sheets dashboard with automated email notifications that ensures nothing falls through the cracks.

## Features

### Dashboard Tab
- Real-time summary counts (Pending, Scheduled, Completed Today)
- Priority section highlighting SOLD units delivering within 48 hours
- Color-coded status indicators (Red = Urgent, Yellow = Pending, Green = Completed)

### Tracking Tabs
1. **Dealer Trade Re-PDIs** - Track dealer trades needing re-PDI before delivery
2. **Customer Accessory Installs** - Track accessory orders requiring tech installation
3. **New Car Parts Installation** - Track new vehicles needing parts before delivery
4. **Service Drive Appraisals** - Track hot/warm/cold leads from service customers

### Automation
- **Email Notifications**: Instant alerts when items are added or completed
- **Auto-Timestamp**: Automatic date and user tracking
- **Priority Calculation**: Auto-flags URGENT items based on delivery dates
- **Conditional Formatting**: Visual cues for sold units, priorities, and heat levels
- **Auto-Archive**: Completed items archived after 7 days

### Access Control
- Locked to two authorized editors only:
  - Brian Callahan (bcallahan@unionparkgmc.com) - Sales Manager
  - Dan Testa (dtesta620@gmail.com) - Service Manager

## Files

| File | Description |
|------|-------------|
| `Code.gs` | Main Apps Script code with all automation |
| `appsscript.json` | Apps Script manifest with required permissions |
| `SETUP_INSTRUCTIONS.md` | Step-by-step deployment guide |
| `USER_GUIDE.md` | How-to guide for Brian and Dan |

## Quick Start

1. Create a new Google Sheet
2. Go to Extensions > Apps Script
3. Copy `Code.gs` into the script editor
4. Update `appsscript.json` manifest
5. Run `initialSetup` function
6. Share sheet with authorized users

See [SETUP_INSTRUCTIONS.md](SETUP_INSTRUCTIONS.md) for detailed steps.

## Email Notification Flow

```
Brian adds Dealer Trade Re-PDI ──────► Email to Dan
Brian adds Accessory Install (tech needed) ──► Email to Dan
Brian adds New Car Parts Installation ──► Email to Dan
Dan adds Service Drive Appraisal ──────► Email to Brian
Any item marked Completed ─────────────► Email to submitter
```

## Priority System

| Condition | Priority | Color |
|-----------|----------|-------|
| SOLD + delivery within 24 hours | URGENT | Red |
| SOLD + delivery within 48 hours | HIGH | Orange |
| All other items | Normal | Standard |

## Heat Levels (Service Appraisals)

| Level | Description | Use When |
|-------|-------------|----------|
| Hot | Strong buying signals | Customer asked about new vehicles, complained about payments |
| Warm | Some interest | High repair bill, mentioned wanting something newer |
| Cold | Worth following up | Long-time customer, just gathering info |

## Customization

Edit the `CONFIG` object in `Code.gs` to customize:
- Authorized user emails
- Color scheme
- Archive period (default: 7 days)
- Priority thresholds (default: 24/48 hours)

## License

Private - Union Park Buick GMC
