#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Windows 11 Application Installer — Base / Streaming / Presentation Profiles
.DESCRIPTION
    Installs a curated set of applications using Winget, Chocolatey, and direct
    web downloads. Prompts for setup type (Base / Streaming / Presentation) and
    handles per-app confirmation, scope (machine vs. admin-user), checksum error
    fallbacks, and manual-browser prompts where automated download is not possible.
.NOTES
    Must be run as Administrator.
    Internet connection required.
    Execution policy is set to Unrestricted at script start.
#>

# ─────────────────────────────────────────────────────────────────────────────
# 0. Execution Policy
# ─────────────────────────────────────────────────────────────────────────────

Set-ExecutionPolicy Unrestricted -Scope LocalMachine -Force
Set-ExecutionPolicy Bypass       -Scope Process      -Force


# ─────────────────────────────────────────────────────────────────────────────
# HELPER FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

function Write-Header {
    param([string]$Text)
    $line = "=" * 72
    Write-Host ""
    Write-Host $line                    -ForegroundColor Cyan
    Write-Host "  $Text"               -ForegroundColor Cyan
    Write-Host $line                    -ForegroundColor Cyan
    Write-Host ""
}

function Write-Step  { param([string]$T) Write-Host "--> $T"     -ForegroundColor Yellow   }
function Write-OK    { param([string]$T) Write-Host "[OK]   $T"  -ForegroundColor Green    }
function Write-Fail  { param([string]$T) Write-Host "[FAIL] $T"  -ForegroundColor Red      }
function Write-Info  { param([string]$T) Write-Host "[INFO] $T"  -ForegroundColor Cyan     }
function Write-Skip  { param([string]$T) Write-Host "[SKIP] $T"  -ForegroundColor DarkGray }
function Write-Warn  { param([string]$T) Write-Host "[WARN] $T"  -ForegroundColor Magenta  }

# Prompt for a yes/no confirmation before installing an app.
function Confirm-Install {
    param([string]$AppName)
    $ans = Read-Host "  Install $AppName? (Y/N)"
    return ($ans -match '^[Yy]$')
}

# Runs winget and captures output + exit code.
# Returns a hashtable: Success, HashError, ExitCode, Output
function Invoke-WingetInstall {
    param(
        [string]$Id,
        [string]$Scope = ""          # "machine", "user", or "" (winget default)
    )

    $extraArgs = @()
    if ($Scope) { $extraArgs += @("--scope", $Scope) }

    $rawLines  = @()
    $process   = $null

    # Use Start-Process to capture output cleanly
    $tmpOut = [System.IO.Path]::GetTempFileName()
    $tmpErr = [System.IO.Path]::GetTempFileName()

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName               = "winget"
    $psi.Arguments              = "install --id `"$Id`" --exact --silent " +
                                  "--accept-package-agreements " +
                                  "--accept-source-agreements " +
                                  ($extraArgs -join " ")
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute        = $false
    $psi.CreateNoWindow         = $true

    $proc = [System.Diagnostics.Process]::Start($psi)
    $stdout = $proc.StandardOutput.ReadToEnd()
    $stderr = $proc.StandardError.ReadToEnd()
    $proc.WaitForExit()
    $exitCode = $proc.ExitCode

    $combined = "$stdout`n$stderr"

    # Already installed: winget returns 0x8A150021 (-1978335199 signed or 2316632097 unsigned)
    $alreadyInstalled = ($exitCode -eq -1978335199) -or
                        ($exitCode -eq -1978335189) -or
                        ($combined -match "already installed")

    # Hash / checksum mismatch detection
    $hashError = ($combined -imatch "hash") -or
                 ($combined -imatch "checksum") -or
                 ($combined -imatch "0x8A15010") -or   # various hash codes
                 ($exitCode -eq -1978334981)            # APPINSTALLER_CLI_ERROR_INSTALLER_HASH_MISMATCH

    $success = ($exitCode -eq 0) -or $alreadyInstalled

    return @{
        Success       = $success
        AlreadyThere  = $alreadyInstalled
        HashError     = $hashError
        ExitCode      = $exitCode
        Output        = $combined
    }
}

# Full install wrapper: winget with automatic fallback handling.
#   -Name         : Display name shown in prompts
#   -WingetId     : Winget package identifier
#   -Scope        : "machine" | "user" | "" (default)
#   -ChocoId      : Chocolatey package name for fallback (optional)
#   -IsChrome     : Switch — enables the special Chrome/Chromium choice dialog
function Install-App {
    param(
        [string]$Name,
        [string]$WingetId,
        [string]$Scope     = "",
        [string]$ChocoId   = "",
        [switch]$IsChrome
    )

    Write-Step "Installing $Name  ($WingetId)..."
    $r = Invoke-WingetInstall -Id $WingetId -Scope $Scope

    if ($r.AlreadyThere) {
        Write-OK "$Name is already installed — skipping."
        return
    }

    if ($r.Success) {
        Write-OK "$Name installed successfully."
        return
    }

    # ── FAILURE PATH ──────────────────────────────────────────────────────────

    Write-Fail "$Name installation failed (exit code: $($r.ExitCode))."

    if ($r.HashError) {
        Write-Warn "Installer hash mismatch detected. The winget manifest may not yet be synced with the publisher's latest release."
    }

    # Chrome / Chromium special handling
    if ($IsChrome) {
        Write-Host ""
        Write-Host "  Choose an alternative for Google Chrome:" -ForegroundColor White
        Write-Host "    1) Google Chrome  via Chocolatey  (identical browser, different source)"
        Write-Host "    2) Chromium       via Winget      (open-source, no Google services)"
        Write-Host "    3) Skip"
        Write-Host ""
        $choice = Read-Host "  Enter choice (1/2/3)"
        switch ($choice.Trim()) {
            "1" {
                Write-Step "Installing Google Chrome via Chocolatey..."
                choco install googlechrome -y --force
                if ($LASTEXITCODE -eq 0) { Write-OK "Google Chrome installed via Chocolatey." }
                else { Write-Fail "Chocolatey install also failed (exit: $LASTEXITCODE)." }
            }
            "2" {
                Write-Step "Installing Chromium via Winget..."
                $rc = Invoke-WingetInstall -Id "Hibbiki.Chromium" -Scope $Scope
                if ($rc.Success) { Write-OK "Chromium installed." }
                else { Write-Fail "Chromium install failed (exit: $($rc.ExitCode))." }
            }
            default { Write-Skip "Chrome/Chromium skipped." }
        }
        return
    }

    # Standard Chocolatey fallback
    if ($ChocoId) {
        $try = Read-Host "  Try $Name via Chocolatey instead? (Y/N)"
        if ($try -match '^[Yy]$') {
            Write-Step "Installing $Name via Chocolatey ($ChocoId)..."
            choco install $ChocoId -y --force
            if ($LASTEXITCODE -eq 0) { Write-OK "$Name installed via Chocolatey." }
            else { Write-Fail "Chocolatey also failed for $Name (exit: $LASTEXITCODE)." }
        } else {
            Write-Skip "$Name skipped."
        }
    } else {
        Write-Info "No Chocolatey fallback configured for $Name. Please install manually."
    }
}

# Install a Chocolatey package, with basic success/fail reporting.
function Install-Choco {
    param([string]$Name, [string]$Id)
    Write-Step "Installing $Name via Chocolatey ($Id)..."
    choco install $Id -y --force 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) { Write-OK "$Name installed." }
    else { Write-Fail "$Name Chocolatey install failed (exit: $LASTEXITCODE)." }
}


# ─────────────────────────────────────────────────────────────────────────────
# 1. SETUP TYPE SELECTION
# ─────────────────────────────────────────────────────────────────────────────

Write-Header "Windows 11 Application Installer"

Write-Host "  Select a setup profile:"                                           -ForegroundColor White
Write-Host ""
Write-Host "    1) Base          — Core applications for all purposes"
Write-Host "    2) Streaming     — Base + HandBrake, Kdenlive, vMix,"
Write-Host "                       DeckLink Desktop Video, NDI Tools"
Write-Host "    3) Presentation  — Base + Office 365 (optional),"
Write-Host "                       ProPresenter, NDI Tools"
Write-Host ""
Write-Host "  All profiles are prompted for: Microsoft Office 365, X32 Edit"  -ForegroundColor DarkGray
Write-Host ""

do {
    $setupChoice = (Read-Host "  Enter choice (1/2/3)").Trim()
} while ($setupChoice -notin @("1","2","3"))

$setupType = switch ($setupChoice) {
    "1" { "Base"         }
    "2" { "Streaming"    }
    "3" { "Presentation" }
}

Write-Info "Setup profile selected: $setupType"


# ─────────────────────────────────────────────────────────────────────────────
# 2. WINGET BOOTSTRAP
# ─────────────────────────────────────────────────────────────────────────────

Write-Header "Step 1 of 3 — Winget"

$wingetCmd = Get-Command winget -ErrorAction SilentlyContinue

if (-not $wingetCmd) {
    Write-Step "Winget not found. Fetching latest release from GitHub..."
    try {
        $apiUrl  = "https://api.github.com/repos/microsoft/winget-cli/releases/latest"
        $release = Invoke-RestMethod -Uri $apiUrl -UseBasicParsing
        $asset   = $release.assets |
                   Where-Object { $_.name -like "Microsoft.DesktopAppInstaller_*.msixbundle" } |
                   Select-Object -First 1

        if (-not $asset) {
            # Fallback: look for any .msixbundle
            $asset = $release.assets |
                     Where-Object { $_.name -like "*.msixbundle" } |
                     Select-Object -First 1
        }

        if ($asset) {
            $installer = "$env:TEMP\AppInstaller.msixbundle"
            Write-Step "Downloading $($asset.name)..."
            Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $installer -UseBasicParsing
            Add-AppxPackage -Path $installer
            Remove-Item $installer -Force -ErrorAction SilentlyContinue

            # Refresh PATH
            $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH","Machine") + ";" +
                        [System.Environment]::GetEnvironmentVariable("PATH","User")

            if (Get-Command winget -ErrorAction SilentlyContinue) {
                Write-OK "Winget installed: $(winget --version)"
            } else {
                Write-Fail "Winget still not available after install attempt."
                Write-Info "Please install App Installer from the Microsoft Store, then re-run this script."
                exit 1
            }
        } else {
            Write-Fail "No suitable release asset found on GitHub."
            exit 1
        }
    } catch {
        Write-Fail "Failed to install Winget: $_"
        exit 1
    }
} else {
    Write-OK "Winget present: $(winget --version)"
}

# Remove msstore source to prevent conflicts
Write-Step "Removing 'msstore' source (avoids Store-only installs and popups)..."
winget source remove --name msstore 2>&1 | Out-Null
Write-OK "msstore source removed (or was not present)."

Write-Step "Updating winget sources..."
winget source update 2>&1 | Out-Null
Write-OK "Sources updated."


# ─────────────────────────────────────────────────────────────────────────────
# 3. CHOCOLATEY BOOTSTRAP
# ─────────────────────────────────────────────────────────────────────────────

Write-Header "Step 2 of 3 — Chocolatey"

if (Get-Command choco -ErrorAction SilentlyContinue) {
    Write-OK "Chocolatey present: $(choco --version)"
    Write-Step "Upgrading Chocolatey..."
    choco upgrade chocolatey -y 2>&1 | Out-Null
} else {
    Write-Step "Installing Chocolatey via official install script..."
    try {
        [System.Net.ServicePointManager]::SecurityProtocol =
            [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        Invoke-Expression ((New-Object System.Net.WebClient).DownloadString(
            'https://community.chocolatey.org/install.ps1'))

        # Refresh PATH
        $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH","Machine") + ";" +
                    [System.Environment]::GetEnvironmentVariable("PATH","User")

        if (Get-Command choco -ErrorAction SilentlyContinue) {
            Write-OK "Chocolatey installed: $(choco --version)"
        } else {
            Write-Warn "Chocolatey installation may have failed. Choco fallbacks may not work."
        }
    } catch {
        Write-Warn "Chocolatey install threw an exception: $_"
        Write-Warn "Chocolatey fallbacks may not work in this session."
    }
}


# ─────────────────────────────────────────────────────────────────────────────
# 4. BASE APPLICATIONS
#    Installed for all three setup profiles.
#    Machine scope: Firefox, Chrome, Zoom, VLC, 7-Zip, PowerShell,
#                   OpenJDK, FFmpeg, Adobe Acrobat Reader
#    Admin-user scope: Notepad++, Audacity
# ─────────────────────────────────────────────────────────────────────────────

Write-Header "Step 3 of 3 — Applications  ($setupType Profile)"
Write-Host "── BASE APPLICATIONS ────────────────────────────────────────────" -ForegroundColor DarkGray
Write-Host ""

# Firefox  (prompt required, machine)
if (Confirm-Install "Mozilla Firefox") {
    Install-App -Name "Firefox" -WingetId "Mozilla.Firefox" -Scope "machine" -ChocoId "firefox"
} else { Write-Skip "Firefox skipped." }

# Google Chrome  (machine, Chrome/Chromium fallback dialog)
Install-App -Name "Google Chrome" -WingetId "Google.Chrome" -Scope "machine" -IsChrome

# Zoom  (prompt required, machine)
if (Confirm-Install "Zoom") {
    Install-App -Name "Zoom" -WingetId "Zoom.Zoom" -Scope "machine" -ChocoId "zoom"
} else { Write-Skip "Zoom skipped." }

# VLC  (machine)
Install-App -Name "VLC Media Player" -WingetId "VideoLAN.VLC" -Scope "machine" -ChocoId "vlc"

# 7-Zip  (machine)
Install-App -Name "7-Zip" -WingetId "7zip.7zip" -Scope "machine" -ChocoId "7zip"

# PowerShell Core  (machine)
Install-App -Name "PowerShell 7" -WingetId "Microsoft.PowerShell" -Scope "machine" -ChocoId "powershell-core"

# Notepad++  (admin-user scope)
Install-App -Name "Notepad++" -WingetId "Notepad++.Notepad++" -Scope "user" -ChocoId "notepadplusplus"

# OpenJDK 21  (machine)
Install-App -Name "Microsoft OpenJDK 21" -WingetId "Microsoft.OpenJDK.21" -Scope "machine" -ChocoId "microsoft-openjdk21"

# FFmpeg  (machine)
Install-App -Name "FFmpeg" -WingetId "Gyan.FFmpeg" -Scope "machine" -ChocoId "ffmpeg"

# Adobe Acrobat Reader 64-bit  (machine)
Install-App -Name "Adobe Acrobat Reader (64-bit)" -WingetId "Adobe.Acrobat.Reader.64-bit" -Scope "machine" -ChocoId "adobereader"

# Audacity  (admin-user scope)
Install-App -Name "Audacity" -WingetId "Audacity.Audacity" -Scope "user" -ChocoId "audacity"


# ─────────────────────────────────────────────────────────────────────────────
# 5. STREAMING APPLICATIONS
# ─────────────────────────────────────────────────────────────────────────────

if ($setupType -eq "Streaming") {

    Write-Host ""
    Write-Host "── STREAMING APPLICATIONS ───────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""

    # HandBrake  (prompt, admin-user)
    if (Confirm-Install "HandBrake") {
        Install-App -Name "HandBrake" -WingetId "HandBrake.HandBrake" -Scope "user" -ChocoId "handbrake"
    } else { Write-Skip "HandBrake skipped." }

    # Kdenlive  (prompt, admin-user)
    if (Confirm-Install "Kdenlive") {
        Install-App -Name "Kdenlive" -WingetId "KDE.Kdenlive" -Scope "user" -ChocoId "kdenlive"
    } else { Write-Skip "Kdenlive skipped." }

    # vMix  (machine, no individual prompt — part of Streaming profile)
    Write-Step "Installing vMix (fetching latest download URL)..."
    $vmixUrl = $null

    try {
        $vmixPage = Invoke-WebRequest -Uri "https://www.vmix.com/software/download.aspx" `
                        -UseBasicParsing -TimeoutSec 20

        # Search anchor hrefs first
        $vmixLink = $vmixPage.Links |
                    Where-Object { $_.href -imatch "vmix\.com/download/vmix[\d.]+\.exe" } |
                    Select-Object -First 1
        if ($vmixLink) { $vmixUrl = $vmixLink.href }

        # Fallback: regex on raw HTML
        if (-not $vmixUrl) {
            if ($vmixPage.RawContent -imatch '(https?://www\.vmix\.com/download/vmix[\d.]+\.exe)') {
                $vmixUrl = $Matches[1]
            } elseif ($vmixPage.RawContent -imatch '"(/download/vmix[\d.]+\.exe)"') {
                $vmixUrl = "https://www.vmix.com" + $Matches[1]
            }
        }
    } catch {
        Write-Warn "Could not fetch vMix download page: $_"
    }

    if ($vmixUrl) {
        Write-Info "Found vMix installer URL: $vmixUrl"
        $vmixDest = "$env:TEMP\vmix_setup.exe"
        try {
            Write-Step "Downloading vMix..."
            Invoke-WebRequest -Uri $vmixUrl -OutFile $vmixDest -UseBasicParsing
            Write-Step "Running vMix installer (silent)..."
            # vMix uses Inno Setup — /VERYSILENT suppresses all UI
            $proc = Start-Process -FilePath $vmixDest `
                        -ArgumentList "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART" `
                        -Wait -PassThru -NoNewWindow
            if ($proc.ExitCode -eq 0) {
                Write-OK "vMix installed."
            } else {
                Write-Fail "vMix installer exited with code $($proc.ExitCode)."
            }
        } catch {
            Write-Fail "vMix install threw an exception: $_"
        } finally {
            Remove-Item $vmixDest -Force -ErrorAction SilentlyContinue
        }
    } else {
        Write-Fail "Could not automatically determine the vMix download URL."
        $openVmix = Read-Host "  Open the vMix download page in your browser? (Y/N)"
        if ($openVmix -match '^[Yy]$') {
            Start-Process "https://www.vmix.com/software/download.aspx"
            Write-Host "  Download and run the vMix installer manually." -ForegroundColor White
            Read-Host "  Press Enter once vMix is installed to continue..."
        } else {
            Write-Skip "vMix skipped."
        }
    }

    # Blackmagic DeckLink Desktop Video  (prompt, browser-only — token-gated URL)
    Write-Host ""
    if (Confirm-Install "Blackmagic DeckLink Desktop Video") {
        Write-Info "DeckLink Desktop Video uses a token-protected download URL and cannot be"
        Write-Info "downloaded automatically. The Blackmagic support page will open in your browser."
        Write-Host ""
        Write-Step "Opening Blackmagic Design support page..."
        Start-Process "https://www.blackmagicdesign.com/support/family/capture-and-playback"
        Write-Host ""
        Write-Host "  Instructions:" -ForegroundColor White
        Write-Host "  1. Find 'Desktop Video' in the list and click the download icon."
        Write-Host "  2. Fill in the registration form if prompted."
        Write-Host "  3. Download the ZIP file for Windows."
        Write-Host "  4. Extract the ZIP and run the .MSI installer inside."
        Write-Host ""
        Read-Host "  Press Enter once DeckLink Desktop Video is installed to continue..."
        Write-OK "Continuing (DeckLink install state assumed complete — manual step)."
    } else {
        Write-Skip "DeckLink Desktop Video skipped."
    }

    # NDI Tools  (prompt, machine)
    Write-Host ""
    if (Confirm-Install "NDI Tools") {
        Install-App -Name "NDI Tools" -WingetId "NDI.NDITools" -Scope "machine" -ChocoId "ndi-tools"
    } else { Write-Skip "NDI Tools skipped." }
}


# ─────────────────────────────────────────────────────────────────────────────
# 6. PRESENTATION APPLICATIONS
# ─────────────────────────────────────────────────────────────────────────────

if ($setupType -eq "Presentation") {

    Write-Host ""
    Write-Host "── PRESENTATION APPLICATIONS ────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""

    # ProPresenter  (machine, no individual prompt — part of Presentation profile)
    Write-Step "Installing ProPresenter (fetching latest Windows installer URL)..."
    $ppUrl = $null

    try {
        $ppPage = Invoke-WebRequest -Uri "https://renewedvision.com/propresenter/download" `
                      -UseBasicParsing -TimeoutSec 20

        # Check anchor hrefs
        $ppLink = $ppPage.Links |
                  Where-Object { $_.href -imatch "renewedvision\.com/downloads/propresenter/win/.*\.exe" } |
                  Select-Object -First 1
        if ($ppLink) { $ppUrl = $ppLink.href }

        # Fallback: regex on raw content
        if (-not $ppUrl) {
            if ($ppPage.RawContent -imatch '(https://renewedvision\.com/downloads/propresenter/win/[^"''<>\s]+\.exe)') {
                $ppUrl = $Matches[1]
            }
        }
    } catch {
        Write-Warn "Could not fetch ProPresenter download page: $_"
    }

    if ($ppUrl) {
        Write-Info "Found ProPresenter installer URL: $ppUrl"
        $ppDest = "$env:TEMP\ProPresenter_Setup.exe"
        try {
            Write-Step "Downloading ProPresenter (this may take a moment)..."
            Invoke-WebRequest -Uri $ppUrl -OutFile $ppDest -UseBasicParsing
            Write-Step "Running ProPresenter installer (silent)..."
            # ProPresenter uses Inno Setup
            $proc = Start-Process -FilePath $ppDest `
                        -ArgumentList "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART" `
                        -Wait -PassThru -NoNewWindow
            if ($proc.ExitCode -eq 0) {
                Write-OK "ProPresenter installed."
            } else {
                Write-Fail "ProPresenter installer exited with code $($proc.ExitCode)."
                Write-Info "If installation appears complete but the exit code is non-zero,"
                Write-Info "ProPresenter may still have installed correctly. Launch it to verify."
            }
        } catch {
            Write-Fail "ProPresenter install threw an exception: $_"
        } finally {
            Remove-Item $ppDest -Force -ErrorAction SilentlyContinue
        }
    } else {
        Write-Fail "Could not automatically find the ProPresenter download URL."
        $ppBrowser = Read-Host "  Open the ProPresenter download page in your browser? (Y/N)"
        if ($ppBrowser -match '^[Yy]$') {
            Start-Process "https://renewedvision.com/propresenter/download"
            Write-Host "  Download the Windows installer and run it." -ForegroundColor White
            Read-Host "  Press Enter once ProPresenter is installed to continue..."
        } else {
            Write-Skip "ProPresenter skipped."
        }
    }

    # NDI Tools  (prompt, machine)
    Write-Host ""
    if (Confirm-Install "NDI Tools") {
        Install-App -Name "NDI Tools" -WingetId "NDI.NDITools" -Scope "machine" -ChocoId "ndi-tools"
    } else { Write-Skip "NDI Tools skipped." }
}


# ─────────────────────────────────────────────────────────────────────────────
# 7. UNIVERSAL OPTIONAL INSTALLS  (all three profiles)
#    Microsoft Office 365 — prompted separately
#    X32 Edit             — prompted separately
# ─────────────────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "── OPTIONAL: MICROSOFT OFFICE 365 ──────────────────────────────────" -ForegroundColor DarkGray
Write-Host ""

if (Confirm-Install "Microsoft Office 365  (~3-4 GB download, sign in to activate)") {

    Write-Step "Installing Office Deployment Tool (ODT) via Winget..."
    winget install --id Microsoft.OfficeDeploymentTool --exact --silent `
        --accept-package-agreements --accept-source-agreements 2>&1 | Out-Null

    # Locate setup.exe — ODT can land in either Program Files location
    $odtSetup = $null
    $odtSearchRoots = @("$env:ProgramFiles", "${env:ProgramFiles(x86)}")
    foreach ($root in $odtSearchRoots) {
        $candidate = Join-Path $root "OfficeDeploymentTool\setup.exe"
        if (Test-Path $candidate) { $odtSetup = $candidate; break }
    }

    # Broader search if still not found
    if (-not $odtSetup) {
        $found = Get-ChildItem $odtSearchRoots -Filter "setup.exe" -Recurse -ErrorAction SilentlyContinue |
                 Where-Object { $_.DirectoryName -imatch "OfficeDeploymentTool" } |
                 Select-Object -First 1
        if ($found) { $odtSetup = $found.FullName }
    }

    if (-not $odtSetup) {
        Write-Fail "Could not locate ODT setup.exe. Falling back to winget direct install..."
        $r = Invoke-WingetInstall -Id "Microsoft.Office" -Scope "machine"
        if ($r.Success) {
            Write-OK "Microsoft 365 installed via Winget. Sign in with your Microsoft account."
        } else {
            Write-Fail "Office 365 install failed. Please install manually:"
            Write-Info "https://www.microsoft.com/en-us/microsoft-365"
        }
    } else {
        Write-OK "ODT found: $odtSetup"

        # Write ODT configuration XML
        $configPath = "$env:TEMP\office365-install.xml"
        $officeXml  = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="Current">
    <Product ID="O365HomePremRetail">
      <Language ID="en-us" />
      <ExcludeApp ID="Groove"  />
      <ExcludeApp ID="Lync"   />
      <ExcludeApp ID="OneDrive"/>
      <ExcludeApp ID="Teams"  />
    </Product>
  </Add>
  <Property Name="AUTOACTIVATE"     Value="0"     />
  <Property Name="FORCEAPPSHUTDOWN" Value="FALSE" />
  <Display  Level="None" AcceptEULA="TRUE"        />
  <Logging  Level="Standard" Path="%temp%"        />
</Configuration>
"@
        Set-Content -Path $configPath -Value $officeXml -Encoding UTF8

        Write-Info "If you have a Microsoft 365 Business subscription, change 'O365HomePremRetail'"
        Write-Info "to 'O365BusinessRetail' in the XML config, then rerun the script."
        Write-Host ""
        Write-Step "Downloading and installing Microsoft 365 (this will take several minutes)..."

        & $odtSetup /configure $configPath

        if ($LASTEXITCODE -eq 0) {
            Write-OK "Microsoft 365 installed. Open any Office app and sign in to activate."
        } else {
            Write-Fail "Office 365 install returned exit code $LASTEXITCODE."
            Write-Info "Check logs at: $env:TEMP\Microsoft Office*.log"
        }
        Remove-Item $configPath -Force -ErrorAction SilentlyContinue
    }
} else {
    Write-Skip "Microsoft Office 365 skipped."
}


Write-Host ""
Write-Host "── OPTIONAL: X32 EDIT (Behringer) ───────────────────────────────────" -ForegroundColor DarkGray
Write-Host ""

if (Confirm-Install "X32 Edit  (Behringer X32 remote control software)") {
    Install-Choco -Name "X32 Edit" -Id "x32-edit"
} else {
    Write-Skip "X32 Edit skipped."
}


# ─────────────────────────────────────────────────────────────────────────────
# DONE
# ─────────────────────────────────────────────────────────────────────────────

Write-Header "Installation Complete"

Write-Host "  Profile  : $setupType"  -ForegroundColor White
Write-Host ""
Write-Host "  Post-install checklist:" -ForegroundColor White
Write-Host "  ─────────────────────────────────────────────────────────────────"
Write-Host "  • Restart Windows to ensure all PATH changes and drivers take effect."
Write-Host "  • Adobe Acrobat Reader — may request a restart to finalize install."
Write-Host "  • PowerShell 7 — open a new terminal to use pwsh instead of powershell."
Write-Host "  • OpenJDK — verify with: java -version  (in a new terminal)"
Write-Host "  • FFmpeg   — verify with: ffmpeg -version  (in a new terminal)"
if ($setupType -eq "Streaming") {
    Write-Host "  • vMix — 60-day trial starts on first launch. Enter your license key if owned."
    Write-Host "  • DeckLink — confirm driver version in the Desktop Video Setup app."
    Write-Host "  • NDI Tools — open NDI Access Manager to configure network permissions."
}
if ($setupType -eq "Presentation") {
    Write-Host "  • ProPresenter — sign in with your Renewed Vision account on first launch."
    Write-Host "  • Office 365 — sign in with your Microsoft account in any Office app."
    Write-Host "  • NDI Tools — open NDI Access Manager to configure network permissions."
}
Write-Host "  ─────────────────────────────────────────────────────────────────"
Write-Host ""
Write-Host "  Run this command periodically to keep apps up to date:" -ForegroundColor DarkGray
Write-Host "  winget upgrade --all --silent --accept-package-agreements" -ForegroundColor DarkGray
Write-Host ""
