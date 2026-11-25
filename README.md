# Employee Anniversaries

An SPFx web part that shows employees whose work anniversary falls within a configurable window (default: 7 days before to 7 days after), based on data in a SharePoint list.

![version](https://img.shields.io/badge/version-1.24.1-green.svg)

## Features

- Reads from any generic list on the current site
- Maps a **Person** column (employee) and a **Date** column (start date) from the properties pane
- Configurable window: days before and days after the anniversary (0-60 each)
- Shows photo, name, role (job title), department and years of service; each can be toggled on or off
- Shows relative timing ("Today", "In 3 days", "2 days ago"), sorted by date
- Handles windows that cross New Year, and 29 Feb start dates (treated as 28 Feb in non-leap years)
- People whose first anniversary has not yet arrived are excluded

## SharePoint list setup

Create a list with at least:

| Column     | Type                    | Notes                         |
| ---------- | ----------------------- | ----------------------------- |
| Employee   | Person or Group         | Single selection              |
| Start Date | Date (date only is fine)| The employee's start date     |

Role and department come from the user information list (the user's SharePoint profile fields `JobTitle` and `Department`), so they must be populated in the profile.

## Web part properties

**Data source**
- Web part title
- Anniversary list
- Employee field (Person columns only)
- Start date field (Date columns only)
- Days before / days after anniversary

**Display**
- Show photo, name, role, department, years of service

## Build and deploy

```
npm install
npm run build
```

Upload `sharepoint/solution/employee-anniversaries.sppkg` to your App Catalog, deploy it, then add the web part to a page.

## Preview and debug

SPFx has no local workbench, so test against SharePoint:

```
npx heft trust-dev-cert   # once
npm start
```

Then open `https://<tenant>.sharepoint.com/sites/<site>/_layouts/15/workbench.aspx`, or append
`?debugManifestsFile=https://localhost:4321/temp/build/manifests.js` to a modern page URL.

## Project structure

- `src/webparts/employeeAnniversaries/EmployeeAnniversariesWebPart.ts`: web part and properties pane
- `src/webparts/employeeAnniversaries/components/EmployeeAnniversaries.tsx`: UI and list query
- `src/webparts/employeeAnniversaries/components/anniversaryUtils.ts`: anniversary window logic
- `src/webparts/employeeAnniversaries/loc/`: UI strings

## Permissions

Uses the signed-in user's permissions via `SPHttpClient`; no API permissions need approving.

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
