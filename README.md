# EOD Waste & Production

Staff form: https://rupat001.github.io/thb-eod-waste/
Manager page: https://rupat001.github.io/thb-eod-waste/manager.html

## Files and what each one is for

| File | What it does | Who changes it |
|---|---|---|
| `items.js` | **The item list.** The only place items are defined. | Manager |
| `config.js` | Sheet link (`SCRIPT_URL`), first entry date (`START_DATE`), API version. | Rarely |
| `common.js` | Shared code: talking to the Sheet, PIN screen, dates. | Rarely |
| `index.html` | Staff form (Produced / Waste). | Feature work |
| `manager.html` | Manager review and approval. | Feature work |
| `Code.gs` | Google Apps Script behind the Sheet (paste into Apps Script, not used by GitHub). | Feature work |
| `favicon.svg`, `apple-touch-icon.png` | Icons. | — |

Keeping these apart means a change to the manager page can't break the staff form,
and changing items never touches page code.

## Adding, renaming or removing an item

Edit only `items.js`:

- **Add:** add `["new_key", "Label"]` in the right group. The Sheet gets a new column
  for it automatically on the next save. Keys: lowercase letters, numbers, `_`.
- **Rename what staff see:** change the label. Never change a key once it's been used.
- **Remove:** delete the line. Its old numbers stay in the Sheet.

Then tell Claude the new key's ClearView name so "enter approved waste" knows where it goes.

## Rules that keep things from breaking

1. Numbers are matched to Sheet columns by **header name**, never by position.
   Don't rename or reorder the header row by hand.
2. `Code.gs` keeps old actions working when new ones are added. A breaking change
   bumps `API_VERSION` in both `config.js` and `Code.gs`.
3. Every save is also copied to the **History** tab, which is never overwritten.
   If the Submissions tab is ever damaged, the numbers can be rebuilt from History.
4. Status (column E) only takes: Submitted, Approved, Entered, Changed after entry, Recorded.

## Undo a bad update

- **Pages (GitHub):** open the repo's commit history, pick the last good commit,
  and restore those files (or ask Claude to). The site updates in about a minute.
- **Apps Script:** Deploy → Manage deployments → pencil → Version: pick the previous
  version → Deploy. Same `/exec` link, takes effect right away.
- **Sheet data:** File → Version history in Google Sheets, or the History tab.

## PINs

Set in Apps Script → Project Settings → Script properties: `STAFF_PIN`, `MANAGER_PIN`.
Never put PINs in this repo — it's public. Changing a PIN there logs everyone out;
they'll be asked for the new one.
