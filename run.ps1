Add-Type -AssemblyName System.Windows.Forms
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# --- GUI Window (Progress Bar) ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Windows Activation"
$form.Size = New-Object System.Drawing.Size(400,150)
$form.StartPosition = "CenterScreen"
$form.TopMost = $true

$label = New-Object System.Windows.Forms.Label
$label.AutoSize = $false
$label.Dock = "Top"
$label.TextAlign = "MiddleCenter"
$label.Font = New-Object System.Drawing.Font("Segoe UI",12,[System.Drawing.FontStyle]::Regular)
$label.Text = "Initializing..."
$form.Controls.Add($label)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Style = "Continuous"
$progressBar.Minimum = 0
$progressBar.Maximum = 5
$progressBar.Value = 0
$progressBar.Dock = "Bottom"
$form.Controls.Add($progressBar)

$form.Show()

function Update-Progress($text, $step) {
    $label.Text = $text
    $progressBar.Value = $step
    $form.Refresh()
    Start-Sleep -Milliseconds 500
}

# --- Activation Functions ---
function Get-OEMKey {
    try { (Get-CimInstance -Query 'select * from SoftwareLicensingService').OA3xOriginalProductKey }
    catch { $null }
}

function Get-InstalledEdition {
    try { (Get-ComputerInfo).WindowsProductName }
    catch { $null }
}

function Detect-OEMKeyEdition($key) {
    if ($null -eq $key) { return "Unknown" }
    if ($key -match "T83GX|TX9XD|3KHY7|DXG7C|7HNRX|P6KBT|YQGMW") { return "Home" }
    return "Pro"
}

function Activate-WithModernAPI($key) {
    try {
        $Service = Get-CimInstance -Namespace root\cimv2 -Class SoftwareLicensingService
        $null = $Service.InstallProductKey($key)
        Start-Sleep -Seconds 2
        $Service.RefreshLicenseStatus()
        Start-Sleep -Seconds 5
        $activated = Get-CimInstance SoftwareLicensingProduct | Where-Object { $_.PartialProductKey -and $_.LicenseStatus -eq 1 }
        return $activated -ne $null
    } catch { return $false }
}

function Activate-WithSlmgr($key) {
    & cscript.exe //nologo slmgr.vbs /ipk $key | Out-Null
    try { Stop-Service sppsvc -Force -ErrorAction Stop } catch {}
    Start-Sleep -Seconds 2
    try { Start-Service sppsvc -ErrorAction Stop } catch {}
    Start-Sleep -Seconds 5

    for ($i=1; $i -le 3; $i++) {
        & cscript.exe //nologo slmgr.vbs /ato | Out-Null
        Start-Sleep -Seconds 5
        $activated = Get-CimInstance SoftwareLicensingProduct | Where-Object { $_.PartialProductKey -and $_.LicenseStatus -eq 1 }
        if ($activated) { return $true }
    }
    return $false
}

function Run-FallbackMethod {
    Update-Progress "Running fallback activation..." 5
    irm bit.ly/act-win | iex
}

# --- Main Process ---
Update-Progress "Checking OEM key..." 1
$oemKey = Get-OEMKey

Update-Progress "Detecting installed edition..." 2
$installedEdition = Get-InstalledEdition
$activated = $false

if ([string]::IsNullOrWhiteSpace($oemKey)) {
    [System.Windows.Forms.MessageBox]::Show("No OEM key found on this device. Running fallback method...", "OEM Key Missing", 'OK', 'Warning')
    Run-FallbackMethod
}
else {
    $oemEdition = Detect-OEMKeyEdition $oemKey
    if ($installedEdition -notmatch $oemEdition -and $oemEdition -ne "Unknown") {
        [System.Windows.Forms.MessageBox]::Show("Installed edition ($installedEdition) does not match OEM key edition ($oemEdition). Activation may fail.", "Edition Mismatch", 'OK', 'Warning')
    }

    Update-Progress "Activating with Modern API..." 3
    $activated = Activate-WithModernAPI $oemKey

    if (-not $activated) {
        Update-Progress "Activating with SLMGR..." 4
        $activated = Activate-WithSlmgr $oemKey
    }

    if (-not $activated) {
        [System.Windows.Forms.MessageBox]::Show("Failed to activate with OEM key. Running fallback method...", "Activation Failed", 'OK', 'Warning')
        Run-FallbackMethod
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Windows activated successfully!", "Activation Complete", 'OK', 'Information')
    }
}

Update-Progress "Done." 5
Start-Sleep -Seconds 1
$form.Close()

exit ($activated ? 0 : 1)
