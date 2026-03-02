#Requires -RunAsAdministrator

# Ensure this script can run regardless of the current execution policy.
# This only affects the current process — it does not change system-wide settings.
Set-ExecutionPolicy Bypass -Scope Process -Force

<#
.SYNOPSIS
    Installs a curated set of applications using Winget and Chocolatey.

.DESCRIPTION
    1. Verifies/updates Winget (App Installer)
    2. Installs Chocolatey via official method
    3. Installs all listed applications

.NOTES
    Must be run as Administrator.
    Audiveris requires the OpenJDK runtime, which is included in the
    Winget installation list below (Microsoft.OpenJDK.21).

    FFmpeg NOTE: FFmpeg is installed via Winget (Gyan.FFmpeg). If you run
    into errors, the Chocolatey version is a reliable fallback:
        choco install ffmpeg -y
    Chocolatey is still installed by this script, so the fallback is
    always available without any extra setup.
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

function Install-WingetPackage {
    param(
        [string]$Name,
        [string]$Id
    )
    Write-Step "Installing $Name ($Id)..."
    winget install --id $Id --source winget --silent --accept-package-agreements --accept-source-agreements
    if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq -1978335189) {
        # Exit code -1978335189 (0x8A150021) means already installed
        Write-OK "$Name installed or already present."
    } else {
        Write-Fail "$Name installation returned exit code $LASTEXITCODE. Continuing..."
    }
}

function Install-ChocoPackage {
    param([string]$Name)
    Write-Step "Installing $Name via Chocolatey..."
    choco install $Name -y
    if ($LASTEXITCODE -eq 0) {
        Write-OK "$Name installed."
    } else {
        Write-Fail "$Name installation returned exit code $LASTEXITCODE. Continuing..."
    }
}

# ---------------------------------------------------------------------------
# Step 1 — Ensure Winget is installed and up to date
# ---------------------------------------------------------------------------

Write-Header "Step 1: Verifying Winget (App Installer)"

$wingetCmd = Get-Command winget -ErrorAction SilentlyContinue

if (-not $wingetCmd) {
    Write-Step "Winget not found. Attempting to install via Microsoft Store (App Installer)..."

    # Winget ships as part of the App Installer package. The most reliable
    # approach on Windows 10/11 is to grab the latest release from GitHub.
    $apiUrl  = "https://api.github.com/repos/microsoft/winget-cli/releases/latest"
    $release = Invoke-RestMethod -Uri $apiUrl -UseBasicParsing
    $asset   = $release.assets | Where-Object { $_.name -like "*.msixbundle" } | Select-Object -First 1

    if ($asset) {
        $installerPath = "$env:TEMP\AppInstaller.msixbundle"
        Write-Step "Downloading $($asset.name)..."
        Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $installerPath -UseBasicParsing
        Add-AppxPackage -Path $installerPath
        Remove-Item $installerPath -Force

        # Refresh PATH
        $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" +
                    [System.Environment]::GetEnvironmentVariable("PATH", "User")

        $wingetCmd = Get-Command winget -ErrorAction SilentlyContinue
        if ($wingetCmd) {
            Write-OK "Winget installed successfully."
        } else {
            Write-Fail "Winget still not detected after install. Aborting."
            exit 1
        }
    } else {
        Write-Fail "Could not locate Winget release asset. Please install App Installer from the Microsoft Store manually."
        exit 1
    }
} else {
    Write-OK "Winget is present: $(winget --version)"
    Write-Step "Upgrading Winget sources..."
    winget source update
}

# ---------------------------------------------------------------------------
# Step 2 — Install Chocolatey
# ---------------------------------------------------------------------------

Write-Header "Step 2: Installing Chocolatey"

if (Get-Command choco -ErrorAction SilentlyContinue) {
    Write-OK "Chocolatey is already installed: $(choco --version)"
    Write-Step "Upgrading Chocolatey..."
    choco upgrade chocolatey -y
} else {
    Write-Step "Installing Chocolatey via official install script..."

    # Official Chocolatey install method
    Set-ExecutionPolicy Bypass -Scope Process -Force
    [System.Net.ServicePointManager]::SecurityProtocol = `
        [System.Net.ServicePointManager]::SecurityProtocol -bor 3072

    Invoke-Expression ((New-Object System.Net.WebClient).DownloadString(
        'https://community.chocolatey.org/install.ps1'
    ))

    # Refresh environment so 'choco' is available in this session
    $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" +
                [System.Environment]::GetEnvironmentVariable("PATH", "User")

    if (Get-Command choco -ErrorAction SilentlyContinue) {
        Write-OK "Chocolatey installed: $(choco --version)"
    } else {
        Write-Fail "Chocolatey installation failed. Aborting."
        exit 1
    }
}

# ---------------------------------------------------------------------------
# Step 3 — Install applications via Winget
# ---------------------------------------------------------------------------

Write-Header "Step 3: Installing Applications via Winget"

# Each entry: Display Name => Winget Package ID
$wingetApps = [ordered]@{
    "Audacity"        = "Audacity.Audacity"
    "Audiveris"       = "audiveris.org.Audiveris" # Requires OpenJDK — installed below
    "VSCodium"        = "VSCodium.VSCodium"
    "Steam"           = "Valve.Steam"
    "GIMP"            = "GIMP.GIMP.3"
    "Inkscape"        = "Inkscape.Inkscape"
    "Upscayl"         = "Upscayl.Upscayl"
    "Discord"         = "Discord.Discord"
    "Firefox"         = "Mozilla.Firefox"
    "Google Chrome"   = "Google.Chrome"
    "Thunderbird"     = "Mozilla.Thunderbird"
    "Transmission"    = "Transmission.Transmission"
    "Zoom"            = "Zoom.Zoom"
    "Slack"           = "SlackTechnologies.Slack"
    "HandBrake"       = "HandBrake.HandBrake"
    "Kdenlive"        = "KDE.Kdenlive"
    "VLC"             = "VideoLAN.VLC"
    "Bitwarden"       = "Bitwarden.Bitwarden"
    "7-Zip"           = "7zip.7zip"
    "FontBase"        = "Levitsky.FontBase"
    "Git"             = "Git.Git"
    "Node.js LTS"     = "OpenJS.NodeJS.LTS"
    "PowerShell Core" = "Microsoft.PowerShell"
    "Notepad++"       = "Notepad++.Notepad++"
    "OpenJDK 21"      = "Microsoft.OpenJDK.21"
    "FFmpeg"          = "Gyan.FFmpeg"           # Fallback: choco install ffmpeg -y
}

foreach ($app in $wingetApps.GetEnumerator()) {
    Install-WingetPackage -Name $app.Key -Id $app.Value
}

# ---------------------------------------------------------------------------
# Step 4 — Install applications via Chocolatey
# ---------------------------------------------------------------------------

Write-Header "Step 4: Installing Applications via Chocolatey"

# FileZilla — not available via Winget, installed via Chocolatey instead.
# If FFmpeg via Winget gives you trouble, add "ffmpeg" to this list as a fallback.
$chocoApps = @(
    "filezilla"
)

foreach ($pkg in $chocoApps) {
    Install-ChocoPackage -Name $pkg
}

# ---------------------------------------------------------------------------
# Step 5 — Chris Titus Tech Windows Utility (optional)
# ---------------------------------------------------------------------------

Write-Header "Step 5: Chris Titus Tech Windows Utility"

Write-Host "The Chris Titus Tech Windows Utility is a popular open-source tool for" -ForegroundColor White
Write-Host "tweaking Windows settings, removing bloatware, and optimizing your system." -ForegroundColor White
Write-Host ""
Write-Host "Source: https://christitus.com/win" -ForegroundColor DarkGray
Write-Host ""

$response = Read-Host "Would you like to launch it now? (Y/N)"

if ($response -match "^[Yy]$") {
    Write-Step "Launching Chris Titus Tech Windows Utility..."
    irm "https://christitus.com/win" | iex
} else {
    Write-Host "Skipped. You can run it any time with:" -ForegroundColor DarkGray
    Write-Host '  irm "https://christitus.com/win" | iex' -ForegroundColor DarkGray
}

# ---------------------------------------------------------------------------
# Done
# ---------------------------------------------------------------------------

Write-Header "Installation Complete"

Write-Host "All installations have been attempted." -ForegroundColor Cyan
Write-Host ""
Write-Host "Recommended next steps:" -ForegroundColor White
Write-Host "  - Restart your terminal (or your PC) to pick up updated PATH entries."
Write-Host "  - Check any [FAIL] entries above and install those manually if needed."
Write-Host "  - If FFmpeg (Winget) gives errors, the Chocolatey fallback is ready:"
Write-Host "      choco install ffmpeg -y" -ForegroundColor DarkGray
Write-Host "  - Run 'winget upgrade --all' periodically to keep apps up to date."
Write-Host ""
