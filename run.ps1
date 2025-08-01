# Requires elevated privileges
Add-Type -AssemblyName System.Windows.Forms

# --- Check if running as Administrator ---
$runningAsAdmin = [System.Security.Principal.WindowsIdentity]::GetCurrent().Groups -match 'S-1-5-32-544'
if (-not $runningAsAdmin) {
    # Restart the script as Administrator
    $arguments = "& '" + $myinvocation.MyCommand.Path + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    return
}

# ========================
# 1. FIX NETWORK SHARING
# ========================

# Set network profile to Private
Get-NetConnectionProfile | ForEach-Object {
    if ($_.NetworkCategory -ne 'Private') {
        Set-NetConnectionProfile -InterfaceIndex $_.InterfaceIndex -NetworkCategory Private
    }
}

# Enable File and Printer Sharing
Set-NetFirewallRule -DisplayGroup "File And Printer Sharing" -Enabled True -Profile Any

# Disable Password Protected Sharing
$regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Lsa"
Set-ItemProperty -Path $regPath -Name "forceguest" -Value 1

# Allow Insecure Guest Access for SMB
$regPath2 = "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Parameters"
If (-not (Test-Path $regPath2)) {
    New-Item -Path $regPath2 -Force | Out-Null
}
Set-ItemProperty -Path $regPath2 -Name "AllowInsecureGuestAuth" -Value 1 -Type DWord

# Restart required services
Restart-Service LanmanWorkstation
Restart-Service LanmanServer

Write-Host "Network sharing configured successfully." -ForegroundColor Green

# ========================
# 2. OEM ACTIVATION
# ========================
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

Write-Host "All tasks completed. Reboot is recommended." -ForegroundColor Cyan
[System.Windows.Forms.MessageBox]::Show("Network fix applied and Windows activation attempted. Please reboot.", "Setup Complete", 'OK', 'Information')