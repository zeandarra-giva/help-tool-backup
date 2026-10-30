# help-tool-backup

Turns the HELP+ Access database (`HELP+.accdb`) into **one self-contained HTML file** that looks and works like the HELP+ tool. It has category tabs, a sidebar and a **Find** search box. Share that single file with anyone: it opens offline in Chrome or Edge, with no folders, installs or internet needed.

The export reads the tables directly through the Access database engine. It never opens forms or runs VBA, so it works even when the Access front-end is broken, and it doesn't change the database.

## Usage (Windows)

1. Double-click `export.bat`. It finds `HELP+.accdb` automatically.
2. The snapshot is saved as `snapshots\HELP+ snapshot <date time>.html` and opens in your browser.
3. Send that one `.html` file to whoever needs it, by email, chat or a shared drive.

To use a specific database file instead, drag the `.accdb` onto `export.bat`.

## Using the snapshot

| | |
|---|---|
| **Top tabs / sidebar** | Switch between categories and tables, like in HELP+ |
| **Find** | Filters the current table as you type. Several words must all match (any column, any order). Matches are highlighted. |
| **Exact match** | Only rows where a whole cell equals the search text |
| **Search all tables** | Searches every table at once, grouped by table, with **Open table →** links |
| **Click a row** | Opens the full record with a **Copy** button for each field, plus **Copy all** |
| **Keyboard** | `/` jumps to the search box. `Esc` clears the search or closes the record. `Enter` opens the selected row. |
| **Light / Dark mode** | Button in the top-right corner. The choice is remembered. |

The snapshot needs JavaScript, which is built into the file. Save the file and open it in a browser. Email preview panes won't run it.

## Tabs and categories: `categories.psd1`

`categories.psd1` decides which table appears under which tab, by matching **table names** with `*` wildcards:

```powershell
@{ Label = 'Applications'; Match = '*Application*' }   # e.g. tblApplications
```

- Categories with `Placement = 'Sidebar'` show as buttons in the left sidebar. All others are top tabs.
- A table goes to the **first** pattern that matches, so list specific patterns before general ones.
- Tables that match nothing still appear under **Other Tables**, and saved queries under **Queries**, so nothing is ever left out.

The patterns shipped here are **guesses based on the HELP+ screen**. The real table names aren't known yet. Each export prints where every table went, for example:

```
[Search Tools] (top tab): Applications, SecOPS
[Other Tables] (top tab): tblMisc
```

Adjust the patterns after the first run if something lands in the wrong place, then run `export.bat` again.

## How the database is found

The script looks for `HELP+.accdb` in this order and uses the first match:

1. The location found on the last run, saved in `last-db-path.txt`
2. The script's folder, then each parent folder above it
3. Desktop, Documents, Downloads and OneDrive, up to 4 levels deep
4. Every local and mapped network drive. This can take a few minutes.

If a folder has more than one match, the most recently modified file wins. To look for a different file name, change `NAME` in `export.bat`.

Tip: put this folder on the same share as the database (for example `\\server\share\help-tool-backup\`). Step 2 then finds it instantly. Unmapped `\\server\share` paths can't be searched.

## Files

| File | Purpose |
|---|---|
| `export.bat` | Double-click to run the export and open the result |
| `export-to-html.ps1` | Finds and reads the database, then builds the snapshot |
| `categories.psd1` | Which tables go under which tabs |
| `template.html`, `styles.css`, `app.js` | The page layout, look and behaviour. All three are built into the snapshot file. |

You can also run the script directly:

```powershell
powershell -ExecutionPolicy Bypass -File export-to-html.ps1 -OutDir ".\snapshots" [-DbPath "C:\path\db.accdb"] [-DbName "HELP+.accdb"] [-SkipQueries] [-NoOpen]
```

## Notes

- **Each snapshot is a point in time.** Run `export.bat` again to refresh it. Each run creates a new dated file, so older snapshots are kept.
- **Anyone with the file can see every record in it.** Check what it contains before sharing it widely.
- **Queries that use forms or VBA functions can't run outside Access.** The sidebar notes how many were skipped.
- **The Access driver comes with Office.** If you get "Access driver not available", the batch file retries with 32-bit PowerShell automatically. If it still fails, install the [Microsoft Access Database Engine](https://www.microsoft.com/en-us/download/details.aspx?id=54920) at the same 32- or 64-bit version as Office.
- **Pictures and attachment fields show as `[binary]`.**
- **`snapshots/` and `last-db-path.txt` are git-ignored** because they contain database records and your network paths. Don't commit them to this public repo.
