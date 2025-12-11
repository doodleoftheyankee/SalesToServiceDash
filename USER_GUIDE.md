# Service-to-Sales Bridge Dashboard
## User Guide for Brian & Dan

Welcome to the Service-to-Sales Bridge Dashboard! This tool helps Sales and Service communicate better on time-sensitive items.

---

## Getting Started

### Accessing the Dashboard

1. Open the shared Google Sheet link (saved in your bookmarks)
2. You'll see 6 tabs at the bottom:
   - **Dashboard** - Summary view with priority items
   - **Dealer Trade Re-PDIs** - For dealer trade vehicles needing re-PDI
   - **Customer Accessory Installs** - For customer accessory orders
   - **New Car Parts Installation** - For new cars needing parts before delivery
   - **Service Drive Appraisals** - For appraisal opportunities from service
   - **Completed Archive** - Historical records

### Mobile Access

- The dashboard works on mobile browsers (Chrome, Safari)
- Bookmark the link for quick access
- Headers are frozen so you can always see column names while scrolling

---

## Dashboard Overview

The **Dashboard** tab shows you everything at a glance:

### Priority Section (Red Header)
- Shows all **SOLD units** with delivery within **48 hours**
- **Red rows** = delivery within 24 hours (URGENT!)
- **Orange rows** = delivery within 24-48 hours

### Summary Counts
| Category | Pending | Scheduled | Completed Today |
|----------|---------|-----------|-----------------|
| Dealer Trade Re-PDIs | X | X | X |
| Accessory Installs | X | X | X |
| New Car Parts Needed | X | X | X |
| Service Drive Appraisals | Hot/Warm/Cold counts |

The dashboard updates automatically every hour, or you can manually refresh it from the **Bridge Dashboard** menu.

---

## For Brian (Sales Manager)

### Adding a Dealer Trade Re-PDI

1. Go to the **"Dealer Trade Re-PDIs"** tab
2. Click in the first empty row under **Stock Number** (Column C)
3. Fill in the information:
   - Stock Number
   - VIN (last 8 digits is fine)
   - Year/Make/Model
   - Is it SOLD? (Yes/No dropdown)
   - Customer Name (if sold)
   - Delivery Date & Time
4. The **Date Submitted** and **Submitted By** will auto-fill
5. **Dan will receive an email notification immediately**

**Note:** If marked SOLD with delivery within 48 hours, the **Priority Flag** will automatically show URGENT or HIGH.

### Adding a Customer Accessory Install

1. Go to the **"Customer Accessory Installs"** tab
2. Click in the first empty row under **Customer Name** (Column C)
3. Fill in:
   - Customer Name, Phone, Email
   - Vehicle info
   - Part Number(s) and Description
   - Part Ordered? (Yes/No)
   - Part Received? (Yes/No)
   - **Requires Tech Install?** (Yes/No) - Important!
4. If **Requires Tech Install = Yes**, Dan receives an email notification

### Adding New Car Parts Installation

1. Go to the **"New Car Parts Installation"** tab
2. Fill in vehicle and part information
3. Mark if SOLD and add delivery date
4. Dan will receive an email notification

---

## For Dan (Service Manager)

### Adding a Service Drive Appraisal

1. Go to the **"Service Drive Appraisals"** tab
2. Click in the first empty row under **Customer Name** (Column C)
3. Fill in:
   - Customer Name, Phone, Email
   - Vehicle they're driving (Year/Make/Model)
   - Mileage
   - Service Being Performed (why they're there)
   - **Heat Level** (Hot/Warm/Cold)
   - Reason for Heat Level (e.g., "complained about payments")
4. **Brian will receive an email notification immediately**

### Heat Level Guide

| Level | Description | Example |
|-------|-------------|---------|
| **Hot** | Customer expressed strong interest | "Asked about new trucks", "Complained about payments", "Said they're looking" |
| **Warm** | Some buying signals | "High mileage repair bill", "Mentioned wanting something newer" |
| **Cold** | Worth following up, not urgent | "Just gathering info", "Long-time customer" |

### Updating Status on Items

When you schedule or complete work:

1. Find the item in the appropriate tab
2. Change the **Status** dropdown:
   - For Re-PDIs: Pending → Scheduled → In Progress → Completed
   - For Accessory Installs: Part Ordered → Part Received → Appointment Scheduled → Completed
   - For Parts Installation: Pending → Parts Received → Scheduled → Completed
3. The original submitter will receive an email when marked **Completed**

---

## Status Workflows

### Dealer Trade Re-PDIs
```
Pending → Scheduled → In Progress → Completed
```

### Customer Accessory Installs
```
Part Ordered → Part Received → Appointment Scheduled → Completed
```

### New Car Parts Installation
```
Pending → Parts Received → Scheduled → Completed
```

### Service Drive Appraisals
```
New Lead → Contacted → Appointment Set → Sold/Lost
```

---

## Email Notifications

You'll receive automatic emails for:

| Event | Who Gets Notified |
|-------|-------------------|
| New Dealer Trade Re-PDI | Dan |
| New Accessory Install (needs tech) | Dan |
| New Car Parts Installation | Dan |
| New Service Drive Appraisal | Brian |
| Any item marked "Completed" | Original submitter |

### Email Format

Each email includes:
- What was added/changed
- Key details (stock #, customer, priority)
- Direct link to the dashboard

---

## Color Coding

### Row Colors
| Color | Meaning |
|-------|---------|
| Yellow | SOLD unit |
| Red | URGENT - delivery within 24 hours |
| Orange | HIGH priority - delivery within 48 hours |
| Gray | Completed |

### Heat Level Colors (Service Appraisals)
| Color | Meaning |
|-------|---------|
| Red | Hot lead |
| Yellow | Warm lead |
| Blue | Cold lead |

---

## Tips & Best Practices

### For Both Users

1. **Check the Dashboard daily** - Priority section shows what needs attention
2. **Update status promptly** - This keeps everyone in sync
3. **Add notes** - Include relevant details in the Notes column
4. **Complete items when done** - Don't leave things hanging

### For Brian

1. **Enter delivery dates accurately** - This drives the priority calculations
2. **Mark SOLD immediately** - This highlights time-sensitive items
3. **Include customer phone** - Makes it easier for Service to schedule

### For Dan

1. **Be specific about heat level reasons** - Helps Sales prioritize follow-up
2. **Capture customer info in service drive** - Phone # is crucial for Sales follow-up
3. **Update appointment times** - Keeps everyone informed

---

## Automatic Features

### Auto-Timestamp
When you add a new row, the date and your email are automatically recorded.

### Auto-Priority
For SOLD units, the Priority Flag automatically calculates:
- **URGENT** = Delivery within 24 hours
- **HIGH** = Delivery within 48 hours
- **Normal** = Everything else

### Auto-Archive
Items marked **Completed** (or Sold/Lost for appraisals) are automatically moved to the archive after 7 days. This keeps your working tabs clean.

### Auto-Dashboard Refresh
The Dashboard updates every hour automatically. You can also refresh it manually from the Bridge Dashboard menu.

---

## Troubleshooting

### "I can't edit the sheet"
- Make sure you're logged into the correct Google account
- Contact the administrator to verify your access

### "I didn't receive an email notification"
- Check your spam folder
- Verify the email address in the Submitted By column is correct
- Email notifications only trigger for certain actions (new items, completions)

### "The colors aren't showing"
- Try refreshing the page
- The conditional formatting should apply automatically

### "The dashboard counts seem wrong"
- Go to Bridge Dashboard menu → Refresh Dashboard
- Make sure items have the correct status selected

---

## Quick Reference Card

### Brian's Tasks
| Action | Tab | Dan Gets Email? |
|--------|-----|-----------------|
| Add dealer trade for re-PDI | Dealer Trade Re-PDIs | Yes |
| Add accessory needing install | Customer Accessory Installs | Yes (if Requires Tech = Yes) |
| Add new car needing parts | New Car Parts Installation | Yes |

### Dan's Tasks
| Action | Tab | Brian Gets Email? |
|--------|-----|-------------------|
| Add service drive appraisal | Service Drive Appraisals | Yes |
| Update status to Completed | Any tab | Yes (to original submitter) |
| Schedule appointments | Update Status column | No |

---

## Need Help?

If something isn't working:
1. Refresh the page first
2. Make sure you're logged into the right Google account
3. Check that all required fields are filled in
4. Contact the administrator if issues persist

---

**Happy communicating!**
*This dashboard was built to bridge the block between Sales and Service.*
