# Run this script as Administrator

function Get-OEMKey {
    try {
        $key = (Get-WmiObject -query 'select * from SoftwareLicensingService').OA3xOriginalProductKey
        return $key
    } catch {
        return $null
    }
}

$oemKey = Get-OEMKey

if ([string]::IsNullOrWhiteSpace($oemKey)) {
    Write-Host "No OEM key found." -ForegroundColor Red
} else {
    Write-Host "OEM key found: $oemKey" -ForegroundColor Cyan

    # Install the OEM key silently
    & cscript.exe //nologo slmgr.vbs /ipk $oemKey | Out-Null
    Start-Sleep -Seconds 3

    # Activate Windows silently
    & cscript.exe //nologo slmgr.vbs /ato | Out-Null
    Start-Sleep -Seconds 3

    # Check activation status
    $activated = Get-CimInstance SoftwareLicensingProduct | Where-Object {
        $_.PartialProductKey -and $_.LicenseStatus -eq 1
    }

    if ($activated) {
        Write-Host "✅ Windows activated successfully!" -ForegroundColor Green

        # Show success popup
        [System.Windows.Forms.MessageBox]::Show("Windows activated successfully!", "Activation Complete", 'OK', 'Information')
    } else {
        Write-Host "⚠️ Failed to activate Windows with the OEM key." -ForegroundColor Yellow
    }
}
