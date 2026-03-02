#Requires -RunAsAdministrator

# Ensure this script can run regardless of the current execution policy.
# This only affects the current process — it does not change system-wide settings.
Set-ExecutionPolicy Bypass -Scope Process -Force

<#
.SYNOPSIS
    Installs Adobe Creative Cloud and Microsoft 365 silently.

.DESCRIPTION
    1. Installs Adobe Creative Cloud manager via Winget.
    2. Downloads the Office Deployment Tool (ODT) via Winget.
    3. Generates a Microsoft 365 configuration XML.
    4. Uses ODT to download and install Microsoft 365.

.NOTES
    Must be run as Administrator.

    ADOBE CREATIVE CLOUD:
    Only the Creative Cloud manager/launcher is installed. Sign in with
    your Adobe ID on first launch to install individual apps (Photoshop,
    Illustrator, etc.).

    MICROSOFT 365:
    This script installs Microsoft 365 Home Premium (the consumer
    subscription). If you have a Microsoft 365 Personal subscription,
    this is the same installer — just sign in and your license will
    activate automatically.

    If you have a Microsoft 365 Business subscription instead, change
    the ProductID in the XML config section below from
    "O365HomePremRetail" to "O365BusinessRetail".

    The installer downloads Office directly from Microsoft's servers
    (~3-4 GB), so this step will take a while depending on your
    connection speed.
#>

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Write-Header {
    param([string]$Text)
    $line = "=" * 60
    Write-Host ""
    Write-Host $line -ForegroundColor Cyan
    Write-Host "  $Text" -ForegroundColor Cyan
    Write-Host $line -ForegroundColor Cyan
    Write-Host ""
}

function Write-Step {
    param([string]$Text)
    Write-Host "--> $Text" -ForegroundColor Yellow
}

function Write-OK {
    param([string]$Text)
    Write-Host "[OK] $Text" -ForegroundColor Green
}

function Write-Fail {
    param([string]$Text)
    Write-Host "[FAIL] $Text" -ForegroundColor Red
}

# ---------------------------------------------------------------------------
# Step 1 — Adobe Creative Cloud
# ---------------------------------------------------------------------------

Write-Header "Step 1: Adobe Creative Cloud"

Write-Step "Installing Adobe Creative Cloud manager via Winget..."

$confirm = Read-Host "Ready to install Adobe Creative Cloud. Proceed? (Y/N)"
if ($confirm -notmatch "^[Yy]$") {
    Write-Host "Skipped Adobe Creative Cloud." -ForegroundColor DarkGray
} else {
    winget install --id Adobe.CreativeCloud --source winget --silent --accept-package-agreements --accept-source-agreements

    if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq -1978335189) {
        Write-OK "Adobe Creative Cloud installed. Sign in with your Adobe ID on first launch."
    } else {
        Write-Fail "Adobe Creative Cloud install returned exit code $LASTEXITCODE."
        Write-Host "You can install it manually from: https://creativecloud.adobe.com/apps/download/creative-cloud" -ForegroundColor DarkGray
    }
}

# ---------------------------------------------------------------------------
# Step 2 — Figma Desktop
# ---------------------------------------------------------------------------

Write-Header "Step 2: Figma Desktop"

Write-Step "Installing Figma Desktop via Winget..."

$confirm = Read-Host "Ready to install Figma Desktop. Proceed? (Y/N)"
if ($confirm -notmatch "^[Yy]$") {
    Write-Host "Skipped Figma Desktop." -ForegroundColor DarkGray
} else {
    winget install --id Figma.Figma --source winget --silent --accept-package-agreements --accept-source-agreements

    if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq -1978335189) {
        Write-OK Figma Desktop installed. Sign in with your Figma account on first launch."
    } else {
        Write-Fail "Figma Desktop install returned exit code $LASTEXITCODE."
        Write-Host "You can install it manually from: https://www.figma.com/downloads/" -ForegroundColor DarkGray
    }
}

# ---------------------------------------------------------------------------
# Step 2 — Office Deployment Tool
# ---------------------------------------------------------------------------

Write-Header "Step 2: Office Deployment Tool (ODT)"

$confirm = Read-Host "Ready to install Microsoft 365 (requires ~3-4 GB download). Proceed? (Y/N)"
if ($confirm -notmatch "^[Yy]$") {
    Write-Host "Skipped Microsoft 365 installation." -ForegroundColor DarkGray
} else {

    Write-Step "Installing Office Deployment Tool via Winget..."
    winget install --id Microsoft.OfficeDeploymentTool --source winget --silent --accept-package-agreements --accept-source-agreements

    # ODT extracts setup.exe to an OfficeDeploymentTool folder under Program Files.
    # Check both 64-bit and 32-bit Program Files locations.
    $possiblePaths = @(
        "C:\Program Files (x86)\OfficeDeploymentTool\setup.exe"
        "C:\Program Files\OfficeDeploymentTool\setup.exe"
    )

    $setupPath = $possiblePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

    # If still not found, do a broad recursive search across both Program Files dirs
    if (-not $setupPath) {
        Write-Step "setup.exe not found in expected locations. Searching broadly..."
        $setupPath = Get-ChildItem -Path @("C:\Program Files", "C:\Program Files (x86)") `
                        -Filter "setup.exe" -Recurse -ErrorAction SilentlyContinue |
                     Where-Object { $_.FullName -like "*OfficeDeploymentTool*" } |
                     Select-Object -First 1 -ExpandProperty FullName
    }

    # Diagnostic: show what's actually in the ODT folders to help troubleshoot
    if (-not $setupPath) {
        Write-Host ""
        Write-Host "Diagnostic — contents of Program Files\OfficeDeploymentTool (if exists):" -ForegroundColor DarkGray
        Get-ChildItem "C:\Program Files\OfficeDeploymentTool" -ErrorAction SilentlyContinue | ForEach-Object { Write-Host "  $($_.FullName)" -ForegroundColor DarkGray }
        Write-Host "Diagnostic — contents of Program Files (x86)\OfficeDeploymentTool (if exists):" -ForegroundColor DarkGray
        Get-ChildItem "C:\Program Files (x86)\OfficeDeploymentTool" -ErrorAction SilentlyContinue | ForEach-Object { Write-Host "  $($_.FullName)" -ForegroundColor DarkGray }
        Write-Host ""
    }

    if (-not $setupPath -or -not (Test-Path $setupPath)) {
        Write-Fail "Could not locate ODT setup.exe. See diagnostic output above."
        Write-Host "Download ODT manually from: https://www.microsoft.com/en-us/download/details.aspx?id=49117" -ForegroundColor DarkGray
    } else {
        $odtDir = Split-Path $setupPath -Parent
        Write-OK "ODT found at: $setupPath"

        # -----------------------------------------------------------------------
        # Step 3 — Generate Microsoft 365 Configuration XML
        # -----------------------------------------------------------------------

        Write-Header "Step 3: Generating Microsoft 365 Configuration"

        # To switch to Microsoft 365 Business, change ProductID to "O365BusinessRetail".
        # Full list of product IDs: https://learn.microsoft.com/en-us/microsoft-365/admin/misc/product-ids
        $configXml = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="Current">
    <Product ID="O365HomePremRetail">
      <Language ID="en-us" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="Bing" />
    </Product>
  </Add>
  <Updates Enabled="TRUE" Channel="Current" />
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="AUTOACTIVATE" Value="0" />
</Configuration>
"@

        $configPath = Join-Path $odtDir "M365-config.xml"
        $configXml | Out-File -FilePath $configPath -Encoding UTF8 -Force
        Write-OK "Configuration written to: $configPath"

        # -----------------------------------------------------------------------
        # Step 4 — Download and Install Microsoft 365
        # -----------------------------------------------------------------------

        Write-Header "Step 4: Installing Microsoft 365"

        Write-Host "Starting Office installation. This typically takes 10-20 minutes..." -ForegroundColor White
        Write-Host ""

        # Run the ODT installer as a background job so we can animate while it runs
        $installJob = Start-Job -ScriptBlock {
            param($setup, $config)
            & $setup /configure $config
            return $LASTEXITCODE
        } -ArgumentList $setupPath, $configPath

        # Spinner + elapsed time displayed while the job runs
        $spinner  = @('|', '/', '-', '\')
        $frame    = 0
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

        while ($installJob.State -eq 'Running') {
            $elapsed = $stopwatch.Elapsed
            $display = "{0:D2}:{1:D2}:{2:D2}" -f $elapsed.Hours, $elapsed.Minutes, $elapsed.Seconds
            Write-Host "`r  $($spinner[$frame % 4])  Installing Microsoft 365...  Elapsed: $display   " -NoNewline -ForegroundColor Yellow
            $frame++
            Start-Sleep -Milliseconds 250
        }

        $stopwatch.Stop()
        $elapsed  = $stopwatch.Elapsed
        $total    = "{0:D2}:{1:D2}:{2:D2}" -f $elapsed.Hours, $elapsed.Minutes, $elapsed.Seconds

        # Collect the exit code from the job
        $jobResult = Receive-Job -Job $installJob
        Remove-Job -Job $installJob

        Write-Host "`r" -NoNewline
        Write-Host "  Completed in $total" -ForegroundColor DarkGray
        Write-Host ""

        if ($jobResult -eq 0) {
            Write-OK "Microsoft 365 installed successfully."
            Write-Host "Open any Office app and sign in with your Microsoft account to activate." -ForegroundColor DarkGray
        } else {
            Write-Fail "Microsoft 365 installation returned exit code $jobResult."
            Write-Host "Check logs at: C:\Windows\Temp\Microsoft Office" -ForegroundColor DarkGray
        }
    }
}

Write-Header "Installation Complete"

Write-Host "Summary:" -ForegroundColor White
Write-Host "  - Adobe Creative Cloud: Sign in at first launch to install your apps."
Write-Host "  - Microsoft 365: Open Word, Excel, etc. and sign in to activate your license."
Write-Host ""
Write-Host "Useful links:" -ForegroundColor DarkGray
Write-Host "  Adobe:     https://creativecloud.adobe.com" -ForegroundColor DarkGray
Write-Host "  Microsoft: https://account.microsoft.com/services" -ForegroundColor DarkGray
Write-Host ""
