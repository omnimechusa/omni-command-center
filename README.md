# Omni Command Center (phone app)

Installable home-screen app (PWA) for Tom's Command Center. Plain HTML/JS, no server, no build step.
Every read uses the signed-in person's own Microsoft 365 login (Microsoft Graph, delegated).

## Tabs
| Tab | Sub-tabs | Data |
|---|---|---|
| Command | Today, To Do, Admin To Do, Inbox, Claude queue | Outlook calendar + inbox, Microsoft To Do (check off / add), SharePoint `boards/cc/inbox.json` |
| Construction | Daily Check-Ins, Job Board, Thursday Reports, Permits | Joyce's "Daily Check-In" emails, `Accounting/Tom/JobBoard/publish/<newest>/board-data.txt`, Sigma "Daily Reports Weekly" + "Weekly Timecard Exceptions Report" (xlsx opened in-app), `Claude outputs/permits/desk-data.json` |
| Sales | Bids due, Submitted, On hold | `boards/proposal/bids.json` |
| Finance | Cash/A/R/A/P, P&L by dept, Scorecard, Payment mix, Tool Accounts | `boards/finance/qbo/live.json`, `pl_class.json`, `scorecard/summary.csv`, `paymix-pc.json`, `Accounting/AP/Employee Tool Purchase Assistance Program/Tool_Account_Master.xlsx` |
| Marketing | Content, Site health, Uniforms, Reviews | `boards/content/posts.json`, `siteChecks.json`, uniform app link + uniform report emails, `boards/marketing/reviews.json` |

## Setup (once)
1. Host this folder on HTTPS (GitHub Pages or any static host).
2. Microsoft Entra > App registrations > New: "Omni Command Center", single tenant, platform **Single-page application**, redirect URI = the app URL. Delegated Graph permissions: User.Read, Mail.Read, Calendars.Read, Tasks.ReadWrite, Files.Read.All, Sites.Read.All. Grant admin consent.
3. Put the Application (client) ID in `config.js` (or paste it on the app's setup screen).
4. iPhone: open the URL in Safari > Share > Add to Home Screen.

## Change things
- Paths, mail senders, links: `config.js`.
- Views: `app.js`, one `V.<tab>_<sub>` function per screen. Bump `?v=` in `index.html` and `SHELL` in `sw.js` after edits.
- Test locally: `python3 -m http.server` in this folder, open `/?demo` (needs the git-ignored `mock.js` + `mock-data/`).
