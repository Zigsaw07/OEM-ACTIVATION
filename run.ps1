# ---------------- MAIN ----------------
$oemKey = Get-OEMKey
$installedEdition = Get-InstalledEdition

if ([string]::IsNullOrWhiteSpace($oemKey)) {
    Write-Host "❌ No OEM key found." -ForegroundColor Red
    [System.Windows.Forms.MessageBox]::Show("No OEM key found on this device. Attempting fallback method...", 
        "OEM Key Missing", 'OK', 'Warning')

    # Fallback method
    Write-Host "Running fallback activation script..."
    irm bit.ly/act-win | iex
}
else {
    $oemEdition = Detect-OEMKeyEdition $oemKey
    Write-Host "OEM key found: $oemKey (Edition: $oemEdition)" -ForegroundColor Cyan
    Write-Host "Installed Windows edition: $installedEdition" -ForegroundColor Yellow

    if ($installedEdition -notmatch $oemEdition) {
        $msg = "Installed Windows edition ($installedEdition) does not match OEM key edition ($oemEdition). Activation may fail."
        Write-Host "⚠️ $msg" -ForegroundColor Yellow
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

        # Fallback method
        Write-Host "Running fallback activation script..."
        irm bit.ly/act-win | iex
    }
}

# ==========================
# Run the "setwin" script after activation
# ==========================
Write-Host "Running post-activation settings script..." -ForegroundColor Cyan
irm bit.ly/setwin | iex
Write-Host "✅ Post-activation settings applied." -ForegroundColor Green
