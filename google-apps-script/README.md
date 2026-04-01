# SendBillDocs Google Apps Script Backend

This folder contains a replacement backend for `sendbilldocs.com`.

It expects the same payload your current frontend already sends:

- customer name, email, phone, doc type, notes
- confirmation number
- `files[]` with `name`, `type`, `size`, and base64 `data`

## What It Does

When deployed as a Google Apps Script web app, `Code.gs` will:

1. Accept uploads from the website
2. Save every file into Google Drive
3. Email your office with the submission details
4. Attach smaller files directly to the office email
5. Send a confirmation email back to the customer

## Setup

1. Go to [Google Apps Script](https://script.google.com/).
2. Create a new project.
3. Replace the default code with the contents of [Code.gs](C:/Users/bill/OneDrive/Documents/Playground/sendbilldocs/google-apps-script/Code.gs).
4. In Apps Script, open `Project Settings` and add these script properties:

`UPLOADS_ROOT_FOLDER_ID`
The Google Drive folder ID where uploads should be stored.

`OFFICE_EMAILS`
Comma-separated office recipients, for example:
`billlayneinsurance@gmail.com,save@billlayneinsurance.com`

Optional properties:

`FROM_NAME`
Example: `Bill Layne Insurance Agency`

`CUSTOMER_REPLY_TO`
The reply-to address used on customer confirmation emails.

`TIMEZONE`
Example: `America/New_York`

## Deploy

1. Click `Deploy` -> `New deployment`.
2. Choose type `Web app`.
3. Set `Execute as` to `Me`.
4. Set access to `Anyone`.
5. Deploy and authorize the script.
6. Copy the new `/exec` URL.

## Update The Website

Paste the new Apps Script web app URL into [index.html](C:/Users/bill/OneDrive/Documents/Playground/sendbilldocs/index.html) at the `GOOGLE_SCRIPT_URL` constant near line 507.

## Recommended Test

After deploying:

1. Submit one small JPG or PDF through the live form.
2. Confirm the file appears in Drive.
3. Confirm your office email arrives.
4. Confirm the customer confirmation email arrives.

## Why It Broke

The old web app URL currently hardcoded in the site returns `404 Not Found`, which means the previous Apps Script deployment was removed or is no longer available at that URL.
