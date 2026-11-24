# Excel trimmer

Adds a `Trimmed` tab to the Keyloop "Used Vehicle Stock" export that is open in Excel. The tab has only Veh Description, Reg Number, Mileage, Selling Price and Disposal Method, with tick boxes, in Retail, Demo, Courtesy, Auction (and Other) blocks, each sorted by reg. The export's own sheet is never changed and nothing is saved.

## Trim the export
1. Have the Keyloop export open in Excel.
2. Open PowerShell: press the Windows key, type `PowerShell`, press Enter.
3. Paste this line (right-click pastes) and press Enter:

   `irm https://raw.githubusercontent.com/ryangraham-1999/excel-trimmer/main/run.ps1 | iex`

4. Excel asks `Trim "<sheet>" in "<workbook>"?`: press **Yes**. If nothing shows, click Excel on the taskbar; the question may be behind PowerShell.
5. A box says how many vehicles went into each block: **OK**. The `Trimmed` tab is the last tab of the export. Print or save it yourself if you want to.

## Self-test (first time, and whenever the coordinator asks)
Same as above with this line instead:

`irm https://raw.githubusercontent.com/ryangraham-1999/excel-trimmer/main/selftest.ps1 | iex`

Excel must be open (any workbook). After a little while a box titled **Trimmer self-test** shows a line starting `TRIMMER PASS`. The text is already selected: press **Ctrl+C**, **OK**, and paste it to the coordinator. It holds no stock data.

## Updating
Nothing to do: every run downloads the latest trimmer.

## What the script does
It loads the trimmer into a hidden Excel workbook, runs it, then closes that workbook without saving. To load it, it switches on Excel's "Trust access to the VBA project object model" for your Windows user, and switches it back as it was when it finishes, even after an error.

## Messages
- `Trimmer: open the Keyloop export in Excel first...`: Excel isn't open. Open the export, run the line again.
- `Trimmer: your IT has switched off macro import...`: this route cannot work on this PC. Tell the coordinator.
- `Trimmer: Excel refused to load the trimmer...`: close every Excel window, open the export again, run the line again. If it happens again, tell the coordinator.
- `Trimmer: could not download ...` or `Trimmer: stopped: ...`: send the coordinator the message's words.
