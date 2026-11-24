# The trimmer, started from PowerShell (car-coordinator designs/stage-B-trimmer.md section 12).
# Use: irm https://raw.githubusercontent.com/ryangraham-1999/excel-trimmer/main/run.ps1 | iex
# It finds the Excel that is already open, loads the three trimmer modules into a hidden workbook,
# runs the macro there, then closes that workbook without saving, puts Excel's
# "Trust access to the VBA project object model" setting back as it was and deletes its temp folder.
# It never starts Excel, never saves anything and contacts no address but the repo below.
& {
    $macro = 'TrimStock'
    $base = 'https://raw.githubusercontent.com/ryangraham-1999/excel-trimmer/main/'
    $modules = @('Trimmer.bas', 'TrimmerSelfTest.bas', 'TrimmerTestData.bas')
    $ErrorActionPreference = 'Stop'
    function Say([string]$text) { Write-Host ('Trimmer: ' + $text) }
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }

    try { $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') }
    catch { Say 'open the Keyloop export in Excel first, then run the line again.'; return }

    $ver = $xl.Version
    $userKey = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
    foreach ($policyKey in @("HKCU:\Software\Policies\Microsoft\Office\$ver\Excel\Security", "HKLM:\Software\Policies\Microsoft\Office\$ver\Excel\Security")) {
        $policy = Get-ItemProperty -Path $policyKey -Name AccessVBOM -ErrorAction SilentlyContinue
        if ($policy -and $policy.AccessVBOM -eq 0) { Say 'your IT has switched off macro import, so this route cannot work. Use the USB route.'; return }
    }

    $folder = Join-Path $env:TEMP ('excel-trimmer-' + [guid]::NewGuid().ToString('N'))
    $old = Get-ItemProperty -Path $userKey -Name AccessVBOM -ErrorAction SilentlyContinue
    $changed = $false
    $wb = $null
    try {
        New-Item -ItemType Directory -Path $folder | Out-Null
        foreach ($m in $modules) {
            $url = $base + $m
            try { Invoke-WebRequest -Uri $url -OutFile (Join-Path $folder $m) -UseBasicParsing }
            catch { Say ('could not download ' + $url + ' : ' + $_.Exception.Message); return }
        }

        if (-not (Test-Path $userKey)) { New-Item -Path $userKey -Force | Out-Null }
        $changed = $true
        Set-ItemProperty -Path $userKey -Name AccessVBOM -Value 1 -Type DWord

        $wb = $xl.Workbooks.Add()
        $wb.Windows.Item(1).Visible = $false
        try {
            $project = $wb.VBProject
            foreach ($m in $modules) { [void]$project.VBComponents.Import((Join-Path $folder $m)) }
        }
        catch { Say 'Excel refused to load the trimmer. Close Excel completely, open the export again, then run the line again.'; return }

        Say 'running. If nothing shows, click Excel on the taskbar.'
        [void]$xl.Run("'" + $wb.Name + "'!" + $macro)
        Say 'finished.'
    }
    catch { Say ('stopped: ' + $_.Exception.Message) }
    finally {
        if ($wb) { try { $wb.Close($false) } catch { } }
        if ($changed) {
            if ($old) { Set-ItemProperty -Path $userKey -Name AccessVBOM -Value $old.AccessVBOM -Type DWord }
            else { Remove-ItemProperty -Path $userKey -Name AccessVBOM -ErrorAction SilentlyContinue }
        }
        Remove-Item -Path $folder -Recurse -Force -ErrorAction SilentlyContinue
    }
}
