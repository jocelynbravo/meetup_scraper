# meetup_scraper

Live Web App: https://script.google.com/macros/s/AKfycbxVwvqnC6BZilQR0SEwfKGAYRLlsmS_HhlVPZjYlN-UhAQE7ALmDsN449SedEPw9_azqg/exec

A Google Apps Script that automatically scrapes Techlahoma events from Meetup and syncs them into a Google Spreadsheet.

## What It Does

- Fetches upcoming and past events from Techlahoma Meetup groups
- Populates two sheets: **Upcoming Events** and **All Events**
- Tracks event details including name, group, date, venue, city, link, status, first seen, last seen, image URL, address, and description
- Automatically creates the sheets if they do not already exist

## Sheets Structure

| Column | Description |
|---|---|
| Event ID | Unique identifier for the event |
| Event Name | Title of the event |
| Group | Meetup group name |
| Date | Event date |
| Venue | Venue name |
| City | City where the event is held |
| Link | URL to the Meetup event page |
| Status | Current status of the event |
| First Seen | Date the event was first recorded |
| Last Seen | Date the event was last updated |
| Image URL | URL to the event image |
| Address | Full address of the venue |
| Description | Event description |

## Setup

1. Open [Google Apps Script](https://script.google.com) and create a new project
2. Copy the contents of `code` into the script editor, replacing any existing code
3. Open the Google Sheet you want to use (or create a new one)
4. In the Apps Script editor, click **Project Settings** and confirm the script is bound to your spreadsheet, or update the spreadsheet reference in the code
5. Click **Save**, then run `syncTechlahomaEvents` manually once to authorize permissions
6. Approve the required Google permissions when prompted

## Running the Script

**Manually:**
- In the Apps Script editor, select `syncTechlahomaEvents` from the function dropdown and click **Run**

**Automatically (recommended):**
- In the Apps Script editor, go to **Triggers** (clock icon on the left)
- Click **Add Trigger**
- Set the function to `syncTechlahomaEvents`
- Set the event source to **Time-driven** and choose your preferred interval (e.g., daily or hourly)

## Requirements

- A Google account
- A Google Spreadsheet
- Access to the Techlahoma Meetup event data (no API key required if using the public endpoint)

## Notes

- The script uses the timezone from your Google account session (`Session.getScriptTimeZone()`)
- Re-running the script will update existing events rather than creating duplicates, based on Event ID
- The deployed web app URL (`/macros/s/.../exec`) can be used to trigger the script externally if the web app is published
