# Keyboard guide

Everything in this workbook can be driven from the keyboard. These are standard Excel for Windows shortcuts. I did not test each one on a real keyboard while building this project (the filter tests were run through automation), so try them once and correct this file if your Excel version differs.

Alt sequences are pressed one key after another, not held together. Press Alt, release, then the next keys.

## Use the dashboard

| Keys | What it does |
|---|---|
| `Ctrl+PageUp` / `Ctrl+PageDown` | Move to the previous or next sheet. Dashboard is the first sheet. |
| `Ctrl+G` then type `B5` then `Enter` | Jump to the Country dropdown cell. Other dropdown cells are D5 (Month), F5 (Customer type), H5 (Metric). |
| `Alt+Down` | Open the dropdown list on the selected cell. |
| `Down` / `Up` then `Enter` | Pick a value from the open list. |
| `Esc` | Close the list without changing anything. |
| `F9` | Recalculate. Normally not needed; calculation is automatic. |

## Find your way around

| Keys | What it does |
|---|---|
| `Ctrl+Arrow` | Jump to the edge of the data. |
| `Ctrl+Home` / `Ctrl+End` | First cell / last used cell. |
| `Ctrl+A` | Select the current region (press twice for the whole sheet). |
| `Ctrl+Shift+Arrow` | Select to the edge of the data. |
| `Ctrl+Space` / `Shift+Space` | Select the column / the row. |
| `F5` or `Ctrl+G`, then a name such as `cRevenue` | Jump to a named range. |
| `Ctrl+F3` | Open the Name Manager to see all named ranges. |

## Read formulas

| Keys | What it does |
|---|---|
| `F2` | Edit the selected cell and show its formula. |
| `Ctrl+`` (backtick) | Show formulas instead of values on the sheet. Press again to switch back. |
| `Ctrl+[` | Jump to the cells the selected formula points at. |
| `F4` | While editing a formula, cycle `A1`, `$A$1`, `A$1`, `$A1`. |
| `F3` | While editing a formula, paste a named range. |
| `Alt+=` | Insert `SUM` for the cells above or beside. |

## Work with the data

| Keys | What it does |
|---|---|
| `Ctrl+T` | Turn a range into a table. |
| `Ctrl+Shift+L` | Switch filter arrows on or off. |
| `Alt+Down` on a header | Open that column's filter menu. |
| `Alt+A`, `V`, `V` | Open Data Validation (this is how the dropdowns are made). |
| `Alt+A`, `M` | Remove Duplicates. |
| `Alt+A`, `R`, `A` | Refresh All (matters if you build the Power Query version). |
| `Alt+N`, `V` | Insert a PivotTable. |
| `Alt+H`, `O`, `I` | Auto-fit column width. |
| `Alt+W`, `F`, `F` | Freeze panes at the selected cell. |
| `Ctrl+Shift+F3` | Create names from the selected row or column labels. |

## Charts

| Keys | What it does |
|---|---|
| `Tab` on the sheet when nothing is selected, or `F6` | Move between ribbon, sheet and other panes. |
| `Alt+F1` | Insert a chart for the selected data on the same sheet. |
| `F11` | Insert a chart for the selected data on its own sheet. |

## Not verified

- Every shortcut above comes from standard Excel for Windows behaviour. None of them were run by hand during this project.
- Ribbon letters (`Alt+A`, `Alt+H`, `Alt+N`, `Alt+W`) can differ in other languages and in Excel for Mac or the web. When you press Alt, Excel shows the letters on screen; trust those over this file.
