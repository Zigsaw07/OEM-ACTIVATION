# Requires elevated privileges
Add-Type -AssemblyName System.Windows.Forms

# Check if running as Administrator
$runningAsAdmin = [System.Security.Principal.WindowsIdentity]::GetCurrent().Groups -match 'S-1-5-32-544'
if (-not $runningAsAdmin) {
    # Restart the script as Administrator
    $arguments = "& '" + $myinvocation.MyCommand.Path + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    return
}

function Get-OEMKey {
    try {
        $key = (Get-CimInstance -Query 'select * from SoftwareLicensingService').OA3xOriginalProductKey
        return $key
    } catch {
        return $null
    }
}

$oemKey = Get-OEMKey

if ([string]::IsNullOrWhiteSpace($oemKey)) {
    Write-Host "❌ No OEM key found." -ForegroundColor Red
    [System.Windows.Forms.MessageBox]::Show("No OEM key found on this device. Attempting fallback method...", "OEM Key Missing", 'OK', 'Warning')

    # Fallback method
    irm https://get.activated.win | iex
    return
} else {
    Write-Host "OEM key found: $oemKey" -ForegroundColor Cyan

    & cscript.exe //nologo slmgr.vbs /ipk $oemKey | Out-Null
    Start-Sleep -Seconds 3

    & cscript.exe //nologo slmgr.vbs /ato | Out-Null
    Start-Sleep -Seconds 3

    $activated = Get-CimInstance SoftwareLicensingProduct | Where-Object {
        $_.PartialProductKey -and $_.LicenseStatus -eq 1
    }

    if ($activated) {
        Write-Host "✅ Windows activated successfully!" -ForegroundColor Green
        [System.Windows.Forms.MessageBox]::Show("Windows activated successfully!", "Activation Complete", 'OK', 'Information')
    } else {
        Write-Host "⚠️ Failed to activate Windows with the OEM key." -ForegroundColor Yellow
        [System.Windows.Forms.MessageBox]::Show("Failed to activate Windows with the OEM key. Attempting fallback method...", "Activation Failed", 'OK', 'Warning')

        # Fallback method
        irm https://get.activated.win | iex
    }
}
