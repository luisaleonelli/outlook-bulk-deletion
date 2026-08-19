# Outlook Calendar Bulk Deletion

Code snippet for in-browser bulk deletion of Outlook Web Calendar events.

This is an adaptation of [alex-gru's google-calendar-bulk-deletion](https://github.com/alex-gru/google-calendar-bulk-deletion) script, rewritten for Outlook Web Calendar's different UI and interaction flow (right-click context menu instead of a click-to-open popup, a confirmation dialog step, pagination, retry logic, and logging).

## Why you might need this script

* You imported events (e.g. from an `.ics` file) into the wrong target calendar.
* You want to clean up your calendar.
* You don't want to delete the events manually.

## What this script does

It searches for calendar events matching a text string you provide, and deletes them one by one across multiple calendar pages (months/weeks), simulating the click → right-click menu → confirmation dialog flow that Outlook Web requires.

## Limitations

* Selectors are based on Outlook Web's current UI (as of August 2026) and may break if Microsoft changes their interface.
* Recurring events may trigger an additional dialog beyond what's been tested — proceed cautiously if your search matches any.
* This deletes without a second per-event confirmation once it starts running. **Test on a throwaway event first.**

## What you have to do

1. Open Outlook Web Calendar and decide on a search string that matches only the events you want to delete.
2. Open developer tools (`F12`, or `Cmd+Option+J` / `Ctrl+Shift+J`) and go to the Console tab.
3. Paste the snippet below and hit `Enter`. (Some browsers require typing "allow pasting" first.)
4. Follow the prompts — you'll be asked to confirm before anything is deleted.

## The Snippet

Paste this into your browser's console and hit `Enter`.

```javascript
// See outlook-bulk-delete.js in this repo for the full script.
```

## Credit

Original concept and Google Calendar implementation: [alex-gru/google-calendar-bulk-deletion](https://github.com/alex-gru/google-calendar-bulk-deletion).

This adaptation was made with the help of Claude.

