# iCloud to Exchange Calendar Sync

A Vercel serverless function that synchronizes specific iCloud calendars into a dedicated Exchange calendar using a full-overwrite model. The sync runs automatically every 15 minutes via Vercel Cron Jobs.

## Overview

This application:
- Connects to iCloud via CalDAV
- Reads events from two personal iCloud calendars
- Deletes all events in the target Exchange calendar within the sync window
- Recreates all events as Busy blocks with full titles and descriptions
- Ensures Calendly can accurately evaluate availability

## Features

- **Full Overwrite Model**: Each sync deletes all events in the window and recreates them from iCloud
- **Privacy**: All synced events are marked as `private` and `busy` to prevent coworkers from seeing details
- **Automatic Sync**: Runs every 15 minutes via Vercel Cron Job
- **Configurable Window**: Syncs events from 1 day in the past to 90 days in the future (configurable)
- **Error Handling**: Comprehensive error logging and graceful failure handling

## Environment Variables

Configure these environment variables in your Vercel project settings:

### iCloud Configuration

- `ICLOUD_USERNAME` - Your iCloud email address (e.g., `jaypszeto@icloud.com`)
- `ICLOUD_APP_PASSWORD` - Apple app-specific password (see setup instructions below)
- `ICLOUD_CAL1` - First calendar name to sync (e.g., `"Shared Calendar"`)
- `ICLOUD_CAL2` - Second calendar name to sync (e.g., `"Jay - Personal"`)

### Microsoft Graph Configuration

- `MS_TENANT_ID` - Your Microsoft tenant ID
- `MS_CLIENT_ID` - Azure AD application client ID
- `MS_CLIENT_SECRET` - Azure AD application client secret
- `MS_USER_ID` - Exchange user email (e.g., `jay@altvina.com`)
- `MS_TARGET_CALENDAR_NAME` - Name of the target Exchange calendar (e.g., `"Personal Busy (iCloud Sync)"`)

### Sync Configuration

- `SYNC_LOOKBACK_DAYS` - Days to look back (default: `1`)
- `SYNC_LOOKAHEAD_DAYS` - Days to look ahead (default: `90`)
- `TIMEZONE` - Timezone for date conversions (e.g., `America/Los_Angeles`)

## Setup Instructions

### 1. Apple App-Specific Password

iCloud requires an app-specific password for CalDAV access:

1. Go to [appleid.apple.com](https://appleid.apple.com)
2. Sign in with your Apple ID
3. Navigate to **Sign-In and Security** → **App-Specific Passwords**
4. Click **Generate an app-specific password**
5. Enter a label (e.g., "Vercel Calendar Sync")
6. Copy the generated password (format: `xxxx-xxxx-xxxx-xxxx`)
7. Set this as `ICLOUD_APP_PASSWORD` in Vercel

**Note**: Do not use your regular Apple ID password. App-specific passwords are required for CalDAV access.

### 2. Microsoft Graph API Setup

#### Create Azure AD Application

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** → **App registrations**
3. Click **New registration**
4. Enter a name (e.g., "Calendar Sync")
5. Select **Accounts in this organizational directory only**
6. Click **Register**
7. Note the **Application (client) ID** and **Directory (tenant) ID**

#### Configure API Permissions

1. In your app registration, go to **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph**
4. Choose **Application permissions** (not Delegated)
5. Search for and add: **Calendars.ReadWrite**
6. Click **Grant admin consent** (requires admin privileges)

#### Create Client Secret

1. Go to **Certificates & secrets**
2. Click **New client secret**
3. Enter a description and expiration
4. Click **Add**
5. **Copy the secret value immediately** (it won't be shown again)
6. Set this as `MS_CLIENT_SECRET` in Vercel

#### Create Target Calendar

1. Log in to Outlook/Exchange as the user specified in `MS_USER_ID`
2. Create a new calendar named exactly as specified in `MS_TARGET_CALENDAR_NAME`
3. Ensure the calendar is accessible via Microsoft Graph API

### 3. Vercel Configuration

#### Deploy the Project

1. Install Vercel CLI: `npm i -g vercel`
2. Run `vercel` in the project directory
3. Follow the prompts to link your project

#### Set Environment Variables

1. Go to your Vercel project dashboard
2. Navigate to **Settings** → **Environment Variables**
3. Add all required environment variables listed above
4. Ensure they're set for **Production**, **Preview**, and **Development** environments as needed

#### Configure Cron Job

The cron job is configured in `vercel.json`:

```json
{
  "crons": [{
    "path": "/api/sync-icloud-to-exchange",
    "schedule": "*/15 * * * *"
  }]
}
```

This runs the sync every 15 minutes. The cron job is automatically enabled when you deploy to Vercel.

## Local Development

### Prerequisites

- Node.js 18+ (for native fetch support)
- npm or yarn

### Setup

1. Clone the repository
2. Install dependencies:
   ```bash
   npm install
   ```

3. Create a `.env.local` file with your environment variables:
   ```
   ICLOUD_USERNAME=your-email@icloud.com
   ICLOUD_APP_PASSWORD=your-app-password
   ICLOUD_CAL1=Shared Calendar
   ICLOUD_CAL2=Jay - Personal
   MS_TENANT_ID=your-tenant-id
   MS_CLIENT_ID=your-client-id
   MS_CLIENT_SECRET=your-client-secret
   MS_USER_ID=your-user@domain.com
   MS_TARGET_CALENDAR_NAME=Personal Busy (iCloud Sync)
   SYNC_LOOKBACK_DAYS=1
   SYNC_LOOKAHEAD_DAYS=90
   TIMEZONE=America/Los_Angeles
   ```

4. Run the development server:
   ```bash
   npm run dev
   ```

5. Test the function:
   ```bash
   curl http://localhost:3000/api/sync-icloud-to-exchange
   ```

### Type Checking

```bash
npm run type-check
```

## How It Works

### Sync Process

1. **Calculate Window**: Determines the time range based on `SYNC_LOOKBACK_DAYS` and `SYNC_LOOKAHEAD_DAYS`
2. **Fetch iCloud Events**: 
   - Discovers calendars matching `ICLOUD_CAL1` and `ICLOUD_CAL2`
   - Excludes "Holidays", "Birthdays", and "Reminders" calendars
   - Fetches all events in the sync window using CalDAV REPORT
3. **Authenticate with Microsoft Graph**: Gets OAuth token using client credentials flow
4. **Find Target Calendar**: Locates the Exchange calendar by name
5. **Delete Existing Events**: Removes all events in the target calendar within the sync window
6. **Create New Events**: Creates new events from iCloud data with:
   - Full title and description
   - `showAs: "busy"`
   - `sensitivity: "private"`
   - Original iCloud UID appended to description

### Event Properties

All synced events have:
- **Subject**: Full event title from iCloud
- **Body**: Full description + `Synced UID: <uid>` at the bottom
- **Show As**: Busy (for availability checking)
- **Sensitivity**: Private (prevents external visibility)
- **Location**: Preserved if present in iCloud event

## Logging and Debugging

### Vercel Logs

View logs in the Vercel dashboard:
1. Go to your project
2. Navigate to **Deployments**
3. Click on a deployment
4. Go to **Functions** → **sync-icloud-to-exchange**
5. View **Logs** tab

### Log Output

The function logs:
- Sync start time
- Sync window range
- Number of events fetched from iCloud
- Number of events deleted from Exchange
- Number of events created in Exchange
- Total sync duration
- Any errors encountered

### Manual Trigger

You can manually trigger the sync by making a GET or POST request to:
```
https://your-project.vercel.app/api/sync-icloud-to-exchange
```

## Troubleshooting

### "Calendar not found" Error

- Verify the calendar name in `MS_TARGET_CALENDAR_NAME` matches exactly (case-sensitive)
- Ensure the calendar exists in the Exchange account
- Check that the app has `Calendars.ReadWrite` permission

### "Failed to discover calendars" Error

- Verify `ICLOUD_USERNAME` and `ICLOUD_APP_PASSWORD` are correct
- Ensure you're using an app-specific password, not your regular password
- Check that two-factor authentication is enabled on your Apple ID

### "Failed to get access token" Error

- Verify `MS_TENANT_ID`, `MS_CLIENT_ID`, and `MS_CLIENT_SECRET` are correct
- Ensure the client secret hasn't expired
- Check that admin consent has been granted for `Calendars.ReadWrite` permission

### Events Not Syncing

- Check Vercel logs for errors
- Verify the calendar names in `ICLOUD_CAL1` and `ICLOUD_CAL2` match exactly
- Ensure events exist in the sync window (1 day back to 90 days forward)
- Check that the target calendar is not read-only

### Timezone Issues

- Verify `TIMEZONE` is set to a valid IANA timezone (e.g., `America/Los_Angeles`)
- All dates are converted to this timezone before syncing

## Security Notes

- Never commit `.env` files or environment variables to version control
- App-specific passwords are required for iCloud CalDAV access
- Client secrets should be rotated periodically
- The sync only modifies the specified target calendar, never other calendars
- All synced events are marked as private to prevent external visibility

## License

ISC


