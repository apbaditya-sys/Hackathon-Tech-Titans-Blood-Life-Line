# Email Configuration for Blood Life Line

# IMPORTANT: To enable email sending, you need to configure the following:

## Option 1: Using Gmail (Recommended for testing)
1. Go to your Google Account settings
2. Enable 2-Step Verification
3. Generate an App Password:
   - Go to https://myaccount.google.com/apppasswords
   - Select "Mail" and "Other (Custom name)"
   - Copy the generated 16-character password
4. Update server.js lines 17-20 with your credentials:
   - user: 'your-email@gmail.com'
   - pass: 'your-16-character-app-password'

## Option 2: Using Other Email Services
You can use other services like:
- Outlook/Hotmail
- Yahoo Mail
- SendGrid
- Mailgun

Update the transporter configuration in server.js accordingly.

## For Testing Without Real Email
If you just want to test without sending real emails, you can use Ethereal Email:
- Visit https://ethereal.email/
- Create a free account
- Use the provided SMTP credentials in server.js

## Current Status
Email sending is currently configured but will fail until you add real credentials.
The registration will still work and save to Excel, but emails won't be sent.
