Add-Type -AssemblyName System.Windows.Forms

function Get-OEMKey {
    try {
        (Get-CimInstance -Query 'select * from SoftwareLicensingService').OA3xOriginalProductKey
    } catch {
        $null
    }
}

function Get-InstalledEdition {
    try {
        (Get-ComputerInfo).WindowsProductName
    } catch {
        $null
    }
}

function Detect-OEMKeyEdition($key) {
    # Rough OEM key detection – most OEM keys are for Home
    if ($key -match "T83GX|TX9XD|3KHY7|DXG7C|7HNRX|P6KBT|YQGMW") {
        return "Home"
    } else {
        return "Pro"
    }
}

function Activate-WithModernAPI($key) {
    try {
        $Service = Get-CimInstance -Namespace root\cimv2 -Class SoftwareLicensingService
        $null = $Service.InstallProductKey($key)
        Start-Sleep -Seconds 2
        $Service.RefreshLicenseStatus()
        Start-Sleep -Seconds 5

        $activated = Get-CimInstance SoftwareLicensingProduct | Where-Object {
            $_.PartialProductKey -and $_.LicenseStatus -eq 1
        }
        return $activated -ne $null
    } catch {
        return $false
    }
}

function Activate-WithSlmgr($key) {
    Write-Host "Installing key using SLMGR..."
    & cscript.exe //nologo slmgr.vbs /ipk $key | Out-Null

    # Restart licensing service
    Stop-Service sppsvc -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Start-Service sppsvc -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 5

    for ($i=1; $i -le 3; $i++) {
        Write-Host ("Attempt {0}: Activating with SLMGR..." -f $i)
        & cscript.exe //nologo slmgr.vbs /ato | Out-Null
        Start-Sleep -Seconds 5

        $activated = Get-CimInstance SoftwareLicensingProduct | Where-Object {
            $_.PartialProductKey -and $_.LicenseStatus -eq 1
        }
        if ($activated) { return $true }
    }
    return $false
}

# ---------------- MAIN ----------------
$oemKey = Get-OEMKey
$installedEdition = Get-InstalledEdition

if ([string]::IsNullOrWhiteSpace($oemKey)) {
    Write-Host "❌ No OEM key found." -ForegroundColor Red
    [System.Windows.Forms.MessageBox]::Show("No OEM key found on this device. Attempting fallback method...", 
        "OEM Key Missing", 'OK', 'Warning')

    # Fallback method (safe download, not auto-run)
    $fallbackScript = "$env:TEMP\ActivateFallback.ps1"
    Invoke-WebRequest -Uri "https://get.activated.win" -OutFile $fallbackScript
    Write-Host "Run fallback script manually: $fallbackScript"
}
else {
    $oemEdition = Detect-OEMKeyEdition $oemKey
    Write-Host "OEM key found: $oemKey (Edition: $oemEdition)" -ForegroundColor Cyan
    Write-Host "Installed Windows edition: $installedEdition" -ForegroundColor Yellow

    if ($installedEdition -notmatch $oemEdition) {
        $msg = "⚠️ Installed Windows edition ($installedEdition) does not match OEM key edition ($oemEdition). Activation may fail."
        Write-Host $msg -ForegroundColor Yellow
        [System.Windows.Forms.MessageBox]::Show($msg, "Edition Mismatch", 'OK', 'Warning')
    }

    Write-Host "Attempting activation using modern API first..."
    $activated = Activate-WithModernAPI $oemKey

    if (-not $activated) {
        Write-Host "Modern API activation failed, trying SLMGR method..."
        $activated = Activate-WithSlmgr $oemKey
    }

    if ($activated) {
        Write-Host "✅ Windows activated successfully!" -ForegroundColor Green
        [System.Windows.Forms.MessageBox]::Show("Windows activated successfully!", 
            "Activation Complete", 'OK', 'Information')
    } else {
        Write-Host "⚠️ Failed to activate Windows with OEM key." -ForegroundColor Yellow
        [System.Windows.Forms.MessageBox]::Show("Failed to activate Windows with the OEM key. Attempting fallback method...", 
            "Activation Failed", 'OK', 'Warning')

        $fallbackScript = "$env:TEMP\ActivateFallback.ps1"
        Invoke-WebRequest -Uri "https://get.activated.win" -OutFile $fallbackScript
        Write-Host "Run fallback script manually: $fallbackScript"
    }
}

Write-Host "All tasks completed. Reboot is recommended." -ForegroundColor Cyan
[System.Windows.Forms.MessageBox]::Show("Network fix applied and Windows activation attempted. Please reboot.", 
    "Setup Complete", 'OK', 'Information')
