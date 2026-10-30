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

1. Open `export.bat` in Notepad and set the database path:
   ```bat
   set "DB=\\server\share\HELP+.accdb"
   ```
2. Double-click `export.bat`.
3. The pages are written to an `html\` folder next to the script, and `index.html` opens in your browser.

Run it again any time to refresh the snapshot.

You can also run the script directly:

```powershell
powershell -ExecutionPolicy Bypass -File export-to-html.ps1 -DbPath "C:\path\db.accdb" -OutDir ".\html" [-SkipQueries]
```

## Notes

- **The output is a snapshot.** It doesn't update live.
- **Queries that use forms or VBA functions can't run outside Access.** They're listed on the index page as unavailable.
- **The Access driver comes with Office.** If you get "Access driver not available", the batch file retries with 32-bit PowerShell automatically. If it still fails, install the [Microsoft Access Database Engine](https://www.microsoft.com/en-us/download/details.aspx?id=54920) at the same 32- or 64-bit version as Office.
- **Pictures and attachment fields show as `[binary]`.**
- **`html/` is git-ignored** because it contains database records. Don't commit exported data to this public repo.
