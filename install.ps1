# Installs the trimmer into Excel once (car-coordinator designs/stage-B-trimmer.md section 13).
# Use: irm https://raw.githubusercontent.com/ryangraham-1999/excel-trimmer/main/install.ps1 | iex
# It builds the add-in Trimmer.xlam from the four trimmer modules in a hidden Excel of its own (your open
# Excel windows are not touched), saves it in your Excel add-ins folder, switches it on for your Windows user,
# puts Excel's "Trust access to the VBA project object model" setting back as it was and deletes its temp folder.
# Afterwards Ctrl+Shift+T inside Excel trims the open export. Run the line again to update.
# It never saves or changes any of your workbooks and contacts no address but the repo below.
& {
    $base = 'https://raw.githubusercontent.com/ryangraham-1999/excel-trimmer/main/'
    $modules = @('Trimmer.bas', 'TrimmerSelfTest.bas', 'TrimmerTestData.bas', 'TrimmerKeys.bas')
    $addInName = 'Trimmer.xlam'
    $ErrorActionPreference = 'Stop'
    function Say([string]$text) { Write-Host ('Trimmer: ' + $text) }
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }

    $ver = '16.0'
    $curVer = Get-ItemProperty -Path 'Registry::HKEY_CLASSES_ROOT\Excel.Application\CurVer' -ErrorAction SilentlyContinue
    if ($curVer -and $curVer.'(default)' -match '(\d+)$') { $ver = $Matches[1] + '.0' }
    $userKey = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
    foreach ($policyKey in @("HKCU:\Software\Policies\Microsoft\Office\$ver\Excel\Security", "HKLM:\Software\Policies\Microsoft\Office\$ver\Excel\Security")) {
        $policy = Get-ItemProperty -Path $policyKey -Name AccessVBOM -ErrorAction SilentlyContinue
        if ($policy -and $policy.AccessVBOM -eq 0) { Say 'your IT has switched off macro import, so this route cannot work. Use the USB route.'; return }
    }

    $addInFolder = Join-Path $env:APPDATA 'Microsoft\AddIns'
    $target = Join-Path $addInFolder $addInName
    $work = Join-Path $env:TEMP ('excel-trimmer-' + [guid]::NewGuid().ToString('N'))
    $old = Get-ItemProperty -Path $userKey -Name AccessVBOM -ErrorAction SilentlyContinue
    $changed = $false
    $xl = $null
    try {
        New-Item -ItemType Directory -Path $work | Out-Null
        foreach ($m in $modules) {
            $url = $base + $m
            try { Invoke-WebRequest -Uri $url -OutFile (Join-Path $work $m) -UseBasicParsing }
            catch { Say ('could not download ' + $url + ' : ' + $_.Exception.Message); return }
        }

        if (-not (Test-Path $userKey)) { New-Item -Path $userKey -Force | Out-Null }
        $changed = $true
        Set-ItemProperty -Path $userKey -Name AccessVBOM -Value 1 -Type DWord

        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        foreach ($a in @($xl.AddIns)) { if ($a.Name -eq $addInName) { $a.Installed = $false } }
        foreach ($w in @($xl.Workbooks)) { if ($w.Name -eq $addInName) { $w.Close($false) } }

        $wb = $xl.Workbooks.Add()
        try { foreach ($m in $modules) { [void]$wb.VBProject.VBComponents.Import((Join-Path $work $m)) } }
        catch { Say 'Excel refused to load the trimmer code. Close Excel completely, then run the line again.'; return }
        $wb.Title = 'Excel Trimmer'
        if (-not (Test-Path $addInFolder)) { New-Item -ItemType Directory -Path $addInFolder | Out-Null }
        $wb.SaveAs($target, 55)
        $wb.Close($false)
        $added = $xl.AddIns.Add($target, $false)
        $added.Installed = $true
        Say 'installed. Close every Excel window, open the Keyloop export again, then press Ctrl+Shift+T to trim it (Ctrl+Shift+Y runs the self-test).'
    }
    catch { Say ('stopped: ' + $_.Exception.Message + ' If the trimmer is already installed, close every Excel window and run the line again.') }
    finally {
        if ($xl) { try { $xl.Quit() } catch { }; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch { } }
        if ($changed) {
            if ($old) { Set-ItemProperty -Path $userKey -Name AccessVBOM -Value $old.AccessVBOM -Type DWord }
            else { Remove-ItemProperty -Path $userKey -Name AccessVBOM -ErrorAction SilentlyContinue }
        }
        Remove-Item -Path $work -Recurse -Force -ErrorAction SilentlyContinue
    }
}
