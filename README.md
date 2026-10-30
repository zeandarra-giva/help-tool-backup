# help-tool-backup

Exports every table in an Access database (`.accdb` / `.mdb`) to static, read-only HTML pages styled with plain CSS. Use it to view the data offline when the Access front-end won't open, for example when its VBA modules are corrupted.

The export reads the tables directly through the Access database engine. It never opens forms or runs VBA, and it doesn't change the database.

## Files

| File | Purpose |
|---|---|
| `export.bat` | Double-click to run the export and open the result |
| `export-to-html.ps1` | Reads the database and writes one HTML page per table or query |
| `styles.css` | Styling for the generated pages (copied into the output) |

## Usage (Windows)

1. Double-click `export.bat`. It finds `HELP+.accdb` automatically.
2. The pages are written to an `html\` folder next to the script, and `index.html` opens in your browser.

To use a specific database file instead, drag the `.accdb` onto `export.bat`.

### How the database is found

The script looks for a file named `HELP+.accdb` in this order and uses the first match:

1. The location found on the last run, saved in `last-db-path.txt`
2. The script's folder, then each parent folder above it
3. Desktop, Documents, Downloads and OneDrive, up to 4 levels deep
4. Every local and mapped network drive. This can take a few minutes.

If a folder has more than one match, the most recently modified file wins. To look for a different file name, change `NAME` in `export.bat`.

Tip: put this folder on the same share as the database (for example `\\server\share\help-tool-backup\`). Step 2 then finds it instantly.

Run it again any time to refresh the snapshot.

You can also run the script directly:

```powershell
powershell -ExecutionPolicy Bypass -File export-to-html.ps1 -OutDir ".\html" [-DbPath "C:\path\db.accdb"] [-DbName "HELP+.accdb"] [-SkipQueries]
```

## Notes

- **The output is a snapshot.** It doesn't update live.
- **Queries that use forms or VBA functions can't run outside Access.** They're listed on the index page as unavailable.
- **The Access driver comes with Office.** If you get "Access driver not available", the batch file retries with 32-bit PowerShell automatically. If it still fails, install the [Microsoft Access Database Engine](https://www.microsoft.com/en-us/download/details.aspx?id=54920) at the same 32- or 64-bit version as Office.
- **Pictures and attachment fields show as `[binary]`.**
- **`html/` and `last-db-path.txt` are git-ignored** because they contain database records and your network paths. Don't commit them to this public repo.
