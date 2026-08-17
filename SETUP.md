# Family Heritage Tree — setup

**Build Once. Grow Forever.**

This guide takes you from nothing to a working family website. It assumes no
technical background. Read it straight through once before starting anything.

Total time: about an hour. Cost: nothing.

---

## What you are building

Four pieces, each doing one job.

| Piece | What it holds | Who touches it |
|---|---|---|
| **Google Sheets** | Every fact about every person | You, daily |
| **Google Drive** | The photograph files themselves | You, when photos arrive |
| **GitHub Pages** | The website family members visit | You, once |
| **Apps Script** | Admin tools + the upload form | You, once |

The important thing to understand: **the website reads the spreadsheet every
time someone opens it.** Adding a relative means adding a row to the
spreadsheet. You will never edit the website again after setup.

---

## What you need before you start

- A Google account (free)
- A GitHub account (free — sign up at github.com)
- The two files delivered with this guide:
  - `Family_Heritage_Tree_Database.xlsx`
  - `index.html`
  - `apps-script/Code.gs`

---

## Part 1 — The spreadsheet

**1.1** Go to [drive.google.com](https://drive.google.com). Click **New ▸ File
upload** and upload `Family_Heritage_Tree_Database.xlsx`.

**1.2** Find the uploaded file, right-click it, and choose
**Open with ▸ Google Sheets**. Then **File ▸ Save as Google Sheets**.

> This conversion matters. The `.xlsx` is only a delivery format — the live
> database must be a Google Sheet so the website can read it.

**1.3** Rename it to `Family Heritage Tree — Master Database`. You can delete
the original `.xlsx` from Drive afterwards.

**1.4** Read the **START HERE** tab. It is short and it is the most important
page in the whole system.

**1.5** Look at the **PEOPLE** tab. Nineteen rows are already there — the
skeleton from the plan, with placeholder names. Replace them with the real ones
as you learn them. Never change a Person ID.

---

## Part 2 — Let the website read the spreadsheet

**2.1** In the spreadsheet, click **Share** (top right).

**2.2** Under *General access*, change **Restricted** to
**Anyone with the link**, and leave the role as **Viewer**. Click **Done**.

**2.3** Copy the spreadsheet's ID from the address bar. The URL looks like:

```
https://docs.google.com/spreadsheets/d/1a2B3cD4eFgHiJkLmNoPqRsTuVwXyZ/edit#gid=0
                                       └──────── this part is the ID ────────┘
```

Keep it somewhere handy. You need it in Part 3.

> **Is this safe?** The spreadsheet becomes readable by anyone who has the ID.
> Since the website itself will be public, the data is public either way. What
> protects your family is the **Privacy** and **Living** columns — see
> *Privacy* below. Do not put anything in this spreadsheet that you would not
> want a stranger to read.

---

## Part 3 — The website

**3.1** Open `index.html` in a plain text editor (Notepad, TextEdit, or
GitHub's own editor). Near the top you will find:

```js
const CONFIG = {
  SHEET_ID: "",
  APPS_SCRIPT_URL: "",
```

Paste your spreadsheet ID between the first pair of quotes. Leave
`APPS_SCRIPT_URL` empty for now.

**3.2** Go to [github.com/new](https://github.com/new) and create a repository.

- Name it something like `family-tree`
- Set it to **Public** (GitHub Pages needs this on a free account)
- Tick **Add a README file**
- Click **Create repository**

**3.3** In the new repository click **Add file ▸ Upload files**, drag in your
edited `index.html`, and click **Commit changes**.

**3.4** Go to **Settings ▸ Pages**. Under *Build and deployment ▸ Source*
choose **Deploy from a branch**, pick branch **main** and folder **/ (root)**,
then **Save**.

**3.5** Wait two or three minutes, then visit:

```
https://YOUR-USERNAME.github.io/family-tree/
```

You should see the family tree with your own data.

> **Still showing the sample family?** The site could not read the sheet. Work
> through the checklist in *Troubleshooting* at the end of this guide.

---

## Part 4 — Admin tools and the upload form

This part is optional. The tree works without it. But it gives you the
administration menu and lets family members send photographs.

**4.1** In the spreadsheet, choose **Extensions ▸ Apps Script**.

**4.2** Delete whatever is in the editor. Paste in the entire contents of
`apps-script/Code.gs`. Click the save icon.

**4.3** Reload the spreadsheet tab. A new **Family Tree** menu appears next to
*Help*. The first time you use it, Google asks for permission — click through
**Advanced ▸ Go to (project name) ▸ Allow**. It is your own script asking to
edit your own spreadsheet and Drive.

**4.4** Run **Family Tree ▸ Create the Google Drive folders**. This builds the
`FAMILY TREE` folder structure and records its ID in SETTINGS.

**4.5** Run **Family Tree ▸ Turn on nightly backups**. A dated copy of the
whole database will be saved to Drive every night, keeping the last 30. Do not
skip this step.

**4.6** To enable the upload form, go back to the Apps Script editor and click
**Deploy ▸ New deployment**. Click the gear icon and choose **Web app**, then:

- *Execute as*: **Me**
- *Who has access*: **Anyone**

Click **Deploy** and copy the **Web app URL** (it ends in `/exec`).

**4.7** Paste that URL into `APPS_SCRIPT_URL` in `index.html`, and re-upload the
file to GitHub. A **Contribute** link now appears in the site menu.

> Whenever you change `Code.gs`, you must run **Deploy ▸ Manage deployments ▸
> Edit ▸ Version: New version ▸ Deploy** for the change to reach the website.

---

## Everyday use

### Adding a person

Add a row to **PEOPLE**. Give them the next free Person ID. Fill in
`FatherID` and `MotherID`. That is all. Siblings, cousins, grandparents,
descendants and the whole tree are worked out from those two columns.

Or use **Family Tree ▸ Add a person…**, which fills in the ID, generation and
branch for you.

### Adding detail to a life

Each of these sheets takes **one row per thing**, all keyed by Person ID:

- **PLACES** — every village, town and city across a life
- **OCCUPATIONS** — every job, trade or role, in date order
- **EDUCATION** — schools, universities, apprenticeships, qualifications
- **EVENTS** — dated moments, which become the life timeline
- **STORIES** — memories and oral history
- **PHOTOS** — captions and Drive file IDs

The **EXAMPLES** tab shows what a filled-in row looks like for each.

### Replacing a silhouette with a photograph

Everyone starts with a silhouette. It is a placeholder, not a gap.

1. Put the photograph in `FAMILY TREE ▸ PEOPLE ▸ P0xx` in Drive
2. Right-click it ▸ **Share** ▸ **Anyone with the link ▸ Viewer**
3. Copy its file ID from the link (the long code between `/d/` and `/view`)
4. Add a row to **PHOTOS** with that ID, set `ApprovalStatus` to `Approved`
   and `IsProfile` to `Yes`

Or select the row and use **Family Tree ▸ Make the selected photo the profile
photo**, which handles the flags for you.

The person keeps the same Person ID. Their old photographs stay on record.

### Approving what family members send

Submissions land in **PHOTOS** and **STORIES** marked `Pending` and are
invisible on the site until you act. Use **Family Tree ▸ Review the approval
queue** to see what is waiting, then select the rows and use **Approve the
selected rows** or **Reject the selected rows**.

### Keeping the record honest

Run **Family Tree ▸ Check the record for problems** every so often. It finds
duplicate IDs, parent links pointing at people who do not exist, circular
parentage, children born before their parents, and anyone floating free of the
tree.

---

## Privacy

Genealogy sites leak information about living people. Three things guard
against that:

**The `Living` column.** Set it to `Yes` for anyone alive. With
`hide_living_details` set to `Yes` in SETTINGS (the default), the public site
shows a living person's name and their place in the family, and withholds their
dates, places, work, schooling and biography.

**The `Privacy` column.** Set it to `Private` for anyone who should not appear
at all. They keep their place in the tree structure but the site shows only
"Private record".

**What you choose to type.** Nothing protects against a full address or a date
of birth typed into a public spreadsheet. Ask living relatives before you
publish anything about them. Some will not want to be there at all, and that
is their call to make.

---

## Why it is built this way

The plan proposed routing everything through Apps Script. This build reads the
spreadsheet directly instead, and uses Apps Script only for uploads and admin.
Three reasons:

**Speed.** An Apps Script web app cold-starts in two to five seconds. Reading
the published sheet is close to instant. On a site people visit occasionally,
almost every visit is a cold start.

**Quotas.** Apps Script has daily execution limits. Published-sheet reads have
none. A family site that goes quiet for months and then gets shared around at a
funeral or a wedding is exactly the traffic pattern that trips a quota.

**Survivability.** This matters most for a project meant to last generations.
If the Apps Script deployment breaks — a Google policy change, an expired
authorisation, an account transferred to the next custodian — the tree keeps
working. Only uploads stop. The site also carries a built-in copy of the
starting data, so it renders something even if every remote piece fails.

The same instinct is behind the website having no external dependencies. No
chart library, no font service, no CDN. One file, no links to anything that
could disappear. A CDN going dark in fifteen years should not take the family
tree with it.

---

## Troubleshooting

**The site shows the sample family instead of mine.**
Check, in order:
1. Is `SHEET_ID` filled in, with no spaces and no `/edit` on the end?
2. Is the spreadsheet shared as *Anyone with the link ▸ Viewer*?
3. Are the tab names still `PEOPLE`, `PLACES` and so on, in capitals?
4. Open the site, press F12, click Console, and read the warning there.

If it still fails, try **File ▸ Share ▸ Publish to web ▸ Entire document ▸
Publish** in the spreadsheet, then reload the site.

**Photographs show as broken images.**
The file is not link-shared. Right-click it in Drive ▸ **Share** ▸
**Anyone with the link ▸ Viewer**. You can set this on the whole `PEOPLE`
folder once and it will apply to everything inside.

**The Family Tree menu is missing.**
Reload the spreadsheet tab. If it is still absent, open **Extensions ▸ Apps
Script** and check the code was pasted and saved.

**A relative is missing from the tree.**
Their `FatherID` or `MotherID` is blank or points at an ID that does not exist.
Run **Family Tree ▸ Check the record for problems**.

**The tree looks cramped.**
Use the `+` and `−` buttons, or the percentage button to fit everything. On a
phone, **Outline view** is easier to read than the chart.

---

## Handing it on

This archive will outlive its first administrator. Write down, somewhere your
family can find it:

- the Google account that owns the spreadsheet and Drive folder
- the GitHub account that owns the website
- where the nightly backups are kept
- that this file explains how it all fits together

Then add a second person as an **Editor** on the spreadsheet and Drive folder,
and as a collaborator on the GitHub repository. One custodian is a single point
of failure; two is a succession plan.

The tree is never finished.
