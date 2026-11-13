# EOD Waste & Production

Staff form: https://rupat001.github.io/thb-eod-waste/
Manager page: https://rupat001.github.io/thb-eod-waste/manager.html

## Files and what each one is for

| File | What it does | Who changes it |
|---|---|---|
| Sheet tab **Items** | **The item list** staff see. Edited on the Manager page → Items. | Manager |
| `items.js` | Built-in fallback list, used only until the Items tab exists or if the Sheet can't be reached. | Rarely |
| `config.js` | Sheet link (`SCRIPT_URL`), first entry date (`START_DATE`), API version. | Rarely |
| `common.js` | Shared code: talking to the Sheet, PIN screen, dates. | Rarely |
| `index.html` | Staff form (Produced / Waste). | Feature work |
| `manager.html` | Manager review and approval. | Feature work |
| `Code.gs` | Google Apps Script behind the Sheet (paste into Apps Script, not used by GitHub). | Feature work |
| `favicon.svg`, `apple-touch-icon.png` | Icons. | — |

Keeping these apart means a change to the manager page can't break the staff form,
and changing items never touches page code.

## Adding, renaming or hiding an item

Manager page → **Items** (or open `manager.html#items`). The first time, tap **Save items**
once to copy the built-in list into the Sheet's **Items** tab.

- **Add:** type the name under its category → Add, fill in its **ClearView row**
  (the exact row name on the ClearView Bakery sheet), Save items. The Sheet gets a new
  column for it right away.
- **Rename / note / reorder:** edit the boxes or use ↑ ↓, then Save items.
- **Remove:** tap **Hide**. Saved items can't be deleted, so their past numbers stay readable.
  Tap **Show** to bring one back.
- **New category:** "+ Add category" at the bottom.

Staff phones pick up the new list the next time they open the form (not while they
have unsaved numbers). Don't edit the Items tab by hand; use the Manager page.

## How entries work

- **Produced** (green page) and **Waste** (red page) work the same way: staff enter only
  what they just made / are throwing out. Each entry is added to the day's total and kept
  in the Sheet's **Batches** tab (column Type says Produced or Waste).
- Entered too many earlier? Tap − below 0 (e.g. −2). It can't go below the day's total.
- Each item shows its **Last** entry (amount and time).
- Every save asks "Add to production / waste for …?" with the numbers and the person's name.
- A new waste entry on an Approved day sets it back to Submitted; on an Entered day it
  becomes Changed after entry.
- The name on a shared iPad is forgotten after 2 hours of no use.

## Publishing an update to the staff form

Bump the build in **both** `version.json` and `BUILD` in `index.html` (same text, e.g.
`2026-11-12.3`), and the `?v=` on its three `<script>` tags. Open forms see the new
build within 10 minutes (or as soon as the iPad wakes) and reload themselves, never
while someone has unsaved numbers.

## Rules that keep things from breaking

1. Numbers are matched to Sheet columns by **header name**, never by position.
   Don't rename or reorder the header row by hand.
2. `Code.gs` keeps old actions working when new ones are added. A breaking change
   bumps `API_VERSION` in both `config.js` and `Code.gs`.
3. Every save is also copied to the **History** tab, which is never overwritten.
   If the Submissions tab is ever damaged, the numbers can be rebuilt from History.
4. Status (column E) only takes: Submitted, Approved, Entered, Changed after entry, Recorded.
5. Item keys (the Sheet column names) never change once saved, even when an item is renamed.

## Undo a bad update

- **Pages (GitHub):** open the repo's commit history, pick the last good commit,
  and restore those files (or ask Claude to). The site updates in about a minute.
- **Apps Script:** Deploy → Manage deployments → pencil → Version: pick the previous
  version → Deploy. Same `/exec` link, takes effect right away.
- **Sheet data:** File → Version history in Google Sheets, or the History tab.

## Settings (Manager page > Settings)

- How long a shared iPad remembers the person's name (or ask every save).
- What the Review page opens with (7 / 14 / 30 days / since start).
- Change the staff or manager PIN (stored in Script properties, never shown on a page).

## PINs

Set in Apps Script → Project Settings → Script properties: `STAFF_PIN`, `MANAGER_PIN`.
Never put PINs in this repo — it's public. Changing a PIN there logs everyone out;
they'll be asked for the new one.
