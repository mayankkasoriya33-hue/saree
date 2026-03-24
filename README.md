# CRM KPI Dashboard — Google Apps Script

A multi-page KPI dashboard built with Google Apps Script + HTML Service.

---

## File Structure

```
Code.gs            ← All backend functions (getData, getTeamData, getEnrollmentData, etc.)
index.html         ← Shell: sidebar, shared CSS, calculator, shared JS
IndividualKPI.html ← Page 1 — Individual KPI table + filters + chart + export
TeamKPI.html       ← Page 2 — Team KPI (B2B) + AI Gemini chatbox
Enrollment.html    ← Page 3 — Enrollment table + Pivot Modal
OnFieldKPI.html    ← Page 4 — On-Field KPI table
DeptList.html      ← Page 5 — Department Directory (team cards + member profiles)
```

---

## How to Deploy

1. Open [script.google.com](https://script.google.com) and create a new project.
2. Copy **Code.gs** content into the default `Code.gs` file.
3. Create **6 HTML files** (File → New → HTML file) named exactly:
   - `index`
   - `IndividualKPI`
   - `TeamKPI`
   - `Enrollment`
   - `OnFieldKPI`
   - `DeptList`
4. Paste each file's content from this repo into the matching Apps Script HTML file.
5. Run `setupProperties` once to save your Gemini API key:
   - Replace `APNI_GEMINI_API_KEY_YAHAN_DAALO` with your real key inside `setupProperties()`.
   - Select `setupProperties` from the dropdown → click **Run**.
6. Run `authorizeMe` once to grant permissions.
7. Deploy → **New deployment** → Web app → Execute as **Me** → Who has access: **Anyone**.
8. Copy the Web App URL and share it.

---

## How the Multi-File System Works

`index.html` uses the Apps Script **templating tag** `<?!= include('Page1') ?>` to inject the raw HTML of each page file at serve time. This means:

- **`index.html`** = Shell (sidebar, global CSS, global JS, calculator)
- **`Page1.html`** = Only Page 1 `<div id="page-1">` + its `<style>` + its `<script>`
- Same pattern for Page2–Page5

When Google Apps Script serves `index.html`, it calls `include('Page1')`, which runs `HtmlService.createHtmlOutputFromFile('Page1').getContent()` and injects the result inline.

---

## Editing a Specific Page

| Want to change... | Edit only... |
|---|---|
| Page 1 — Individual KPI table/filters | `IndividualKPI.html` |
| Page 2 — Team KPI table or AI chat | `TeamKPI.html` |
| Page 3 — Enrollment filters, pivot modal | `Enrollment.html` |
| Page 4 — On-Field KPI | `OnFieldKPI.html` |
| Page 5 — Department List | `DeptList.html` |
| Sidebar, global CSS, calculator, helpers | `index.html` |
| Any backend data function | `Code.gs` |

---

## Spreadsheet Column Map

### Sheet: `Digi- Sales Induvisual Daily KPI Report`
| Col | Data |
|---|---|
| A | Sr. No. |
| B | RM Name |
| C | Deposit |
| D | Old App |
| E | New App |
| F | With MOU |
| G | Without MOU |
| H | Status |
| I | Date |
| J | Team |

### Sheet: `Enrollment sheet`
Auto-detected by header keywords: `Payment Date`, `Student Status`, `B2B RM`, `Team Leader`, `Onboarded By`, `Managed By`.

### Sheet: `Team List`
| Col | Data |
|---|---|
| A | Sr. |
| B | Name |
| C | CRM ID |
| D | Designation |
| E | Team |
| F | Department |
| G | Email |
| H | Contact |
| I | Status |
