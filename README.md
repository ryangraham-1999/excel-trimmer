# Excel trimmer

Adds a `Trimmed` tab to the Keyloop "Used Vehicle Stock" export open in Excel: only Veh Description, Reg Number, Mileage, Selling Price and Disposal Method, with tick boxes, in Retail, Demo, Courtesy, Auction (and Other) blocks, each sorted by reg. The export's own sheet is never changed and nothing is saved.

## Install (once)
1. Open PowerShell: press the Windows key, type `PowerShell`, press Enter.
2. Paste this line (right-click pastes) and press Enter:

   `irm https://raw.githubusercontent.com/ryangraham-1999/excel-trimmer/main/install.ps1 | iex`

3. It prints `Trimmer: installed...`. Close every Excel window.

It builds the trimmer as an add-in in your Excel add-ins folder and switches it on. To do that it switches on Excel's "Trust access to the VBA project object model" for your Windows user, and puts it back as it was when it finishes.

## Use
1. Open the Keyloop export in Excel.
2. Press **Ctrl+Shift+T**.
3. Press **Yes** to "Trim ... ?". A box says how many vehicles went into each block: **OK**. The `Trimmed` tab is the last tab. Print or save it yourself if you want to.

## Self-test (after installing or updating)
Press **Ctrl+Shift+Y** in any open workbook. A box titled **Trimmer self-test** shows a line starting `TRIMMER PASS`. It is selected: press **Ctrl+C**, **OK**, and paste it to the coordinator.

## Update
Paste the install line again, then close every Excel window. The same keys keep working.

## Remove
In Excel: **File** > **Options** > **Add-ins** > **Go...** next to Manage Excel Add-ins, untick **Excel Trimmer**, **OK**.

## If it doesn't work
- `Trimmer: your IT has switched off macro import...`: this route cannot work on this PC. Tell the coordinator.
- `Trimmer: Excel refused to load...` or `Trimmer: stopped: ...`: close every Excel window, run the install line again. If it happens again, send the coordinator the message's words.
- Ctrl+Shift+T does nothing: Excel may show a security bar for the add-in. Press **Enable Content**, or check **File** > **Options** > **Add-ins** shows **Excel Trimmer**.
