# 📧 Bulk Email Sender for Google Sheets

A Google Apps Script that adds bulk email functionality to Google Sheets. Send personalized emails to multiple recipients using customizable templates.

## ✨ Features

- **Bulk email sending** via Gmail
- **Customizable templates** with variable placeholders
- **Default values** for missing recipient data
- **Email preview** before sending
- **Test send** to verify your template
- **Status tracking** with automatic updates
- **Visual feedback** with color-coded rows

## 📋 Requirements

- Google Account
- Google Sheets
- Gmail (for sending)

## 🚀 Quick Start

### Installation

1. Open a new or existing **Google Sheet**
2. Go to **Extensions → Apps Script**
3. Delete any existing code in the editor
4. Copy and paste the entire contents of `bulk_email_sender.js`
5. Click **Save** (💾)
6. Refresh your Google Sheet
7. You'll see a new menu: **📧 Email Sender**

### First-Time Setup

1. Click **📧 Email Sender → 🔧 Initialize Sheets**
2. Confirm to create the required tabs
3. When prompted, **authorize** the script to access Gmail

### Usage

1. **Edit Template** - Go to the "Template" tab and customize:
   - Email subject
   - Email body (use `{{variable}}` placeholders)
   - Default greeting (used when name is empty)
   - Default message (used when message is empty)

2. **Add Recipients** - Go to the "Recipients" tab and add:
   - `email` - Recipient's email address
   - `greeting_first_name` - Name for greeting (optional)
   - `message` - Personalized message (optional)

3. **Preview** - Click **👁️ Preview** to see how emails will look

4. **Test** - Click **🧪 Test Send** to send a test email to yourself

5. **Send** - Click **✉️ Send Emails** to send to all recipients

## 📝 Template Variables

| Variable | Description |
|----------|-------------|
| `{{greeting_first_name}}` | Recipient's name (uses default if empty) |
| `{{message}}` | Personalized message (uses default if empty) |
| `{{email}}` | Recipient's email address |

### Example Template

```
Dear {{greeting_first_name}},

{{message}}

Best regards
```

## 📊 Sheet Structure

### Recipients Tab

| email | greeting_first_name | message | status | sent_date |
|-------|---------------------|---------|--------|-----------|
| john@example.com | John | Great work on your project! | | |
| jane@example.com | Jane | Loved your presentation | | |
| user@example.com | | | | |

### Template Tab

Contains editable fields for:
- Subject line
- Body template
- Default greeting
- Default message

## ⚠️ Limits

- **Gmail**: 500 emails/day
- **Google Workspace**: 2,000 emails/day

## 🔒 Permissions

The script requires the following permissions:
- **Gmail** - To send emails
- **Spreadsheets** - To read recipient data and update status

## 🐛 Troubleshooting

### "Permission denied" error

1. Go to [Google Account Permissions](https://myaccount.google.com/permissions)
2. Remove the Apps Script project
3. Run the script again and re-authorize

### Emails not sending

1. Check that recipient emails are valid
2. Verify you haven't exceeded daily sending limits
3. Check the "status" column for error messages

## 📄 License

MIT License - feel free to use and modify.

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
