#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Launches Chris Titus Tech's Windows Utility (WinUtil).
.DESCRIPTION
    Sets execution policy and invokes the WinUtil script directly from
    GitHub via the official one-liner. WinUtil opens as a GUI window.
.LINK
    https://github.com/ChrisTitusTech/winutil
.NOTES
    Must be run as Administrator.
    Requires an active internet connection.
#>

Set-ExecutionPolicy Bypass -Scope Process -Force

[System.Net.ServicePointManager]::SecurityProtocol =
    [System.Net.ServicePointManager]::SecurityProtocol -bor 3072

Write-Host ""
Write-Host "  Launching Chris Titus Tech WinUtil..." -ForegroundColor Cyan
Write-Host "  https://github.com/ChrisTitusTech/winutil" -ForegroundColor DarkGray
Write-Host ""

Invoke-RestMethod "https://christitus.com/win" | Invoke-Expression
