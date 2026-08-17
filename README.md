# Family Heritage Tree

**Build Once. Grow Forever.**

A free, expandable digital family tree and heritage archive — ancestry,
relationships, photographs, places, professions, education, stories and life
histories, preserved for present and future generations.

---

## What is here

```
index.html                            The whole website. One file, no dependencies.
Family_Heritage_Tree_Database.xlsx    The master database, ready to upload to Google Sheets.
apps-script/Code.gs                   Admin tools + the photo/story upload backend.
apps-script/appsscript.json           Apps Script manifest (permissions and web app settings).
SETUP.md                              Step-by-step setup. Start here.
build/                                Source files used to generate the two above. Not needed to run the site.
```

**New here? Read [SETUP.md](SETUP.md).**

---

## How it works

```
        Family members
              │
              ▼
      GitHub Pages  ──────── index.html, one self-contained file
              │
              ├──── reads ────► Google Sheets      the master family database
              │
              ├──── reads ────► Google Drive       photograph files
              │
              └──── writes ───► Apps Script        uploads, stories, admin tools
                                     │
                                     └──► Google Sheets + Drive
```

The website reads the spreadsheet directly on every visit, so adding a person
to the spreadsheet is all it takes for them to appear on the site — with their
siblings, cousins, ancestors and descendants worked out automatically.

Apps Script handles only writes and administration. If it ever breaks, the tree
keeps working.

---

## The database

Nine sheets, all keyed by a permanent **Person ID** that never changes:

| Sheet | Holds |
|---|---|
| `PEOPLE` | One row per person — names, parents, dates, branch, privacy |
| `RELATIONSHIPS` | Marriages, partnerships, adoptions |
| `PLACES` | Every place across a life, with dates and type |
| `OCCUPATIONS` | Career history, one row per role |
| `EDUCATION` | Schools, universities, apprenticeships, qualifications |
| `EVENTS` | Dated moments, which become the life timeline |
| `PHOTOS` | Captions and Drive file IDs, with an approval flag |
| `STORIES` | Memories and oral history, with who told them |
| `SETTINGS` | Site title, root person, privacy switch, configuration |

A person's parentage is two columns: `FatherID` and `MotherID`. Everything else
about the tree's shape is derived from those.

---

## What the site does

- Interactive family tree — chart view with zoom-to-fit, collapsible branches,
  and an outline view for phones
- A profile page per person: family, life, places, career, education, timeline,
  stories and photographs
- Ancestry breadcrumb from the earliest known ancestor down to anyone
- "How am I related?" — works out siblings, cousins, removals, aunts, uncles
  and great-grandparents from the parentage data
- Search across names, nicknames, places, professions, schools and stories
- Silhouettes for everyone without a photograph, replaced in place when a real
  one is approved
- A privacy screen that withholds details of living relatives
- Light and dark, desktop and phone

---

## Principles this build holds to

**Unknown stays unknown.** Nothing in the system requires you to invent a
detail to fill a blank. The `Status` column records how certain each fact is,
so oral history is preserved as oral history rather than promoted to fact.

**Person IDs are permanent.** Names change, spellings vary, relationships get
corrected. The ID never moves.

**No dependencies.** No chart library, no CDN, no web fonts. One HTML file that
will still open in a browser decades from now.

**Photographs never enter this repository.** Media lives in Google Drive under
the administrator's control. The repository holds code.

---

## Rebuilding from source

Only needed if you want to change the schema or the seed data.

```bash
cd build
python3 make_workbook.py     # schema.py  →  the .xlsx and seed-data.json
python3 make_site.py         # template + seed-data.json  →  index.html
node verify.mjs              # 51 checks: parsing, relationships, privacy, rendering
```

`build/schema.py` is the single source of truth. The spreadsheet and the
website's built-in fallback data are both generated from it, so they cannot
drift apart.

---

Every branch matters. Every person has a place. Every generation can add to
what came before.

The tree is never finished.
