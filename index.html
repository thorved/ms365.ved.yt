<#
.SYNOPSIS
    Interactive Microsoft Office 365 installation script with auto-elevation
.DESCRIPTION
    This script provides an interactive menu to configure and install Microsoft Office 365
    by automatically downloading the Office Deployment Tool and creating a configuration XML file.
    It will automatically request elevation to administrator privileges if needed.
.NOTES
    Author: Script User
    Date: March 17, 2025
#>

# Initialize script variables with defaults at the beginning
$script:Architecture = "64"
$script:Language = "en-us"
$script:Channel = "Current"
$script:ExcludedApps = @("Groove", "OneDrive", "Access", "Lync", "OneNote", "Publisher")
$script:TempDir = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "Office365_Install_$(Get-Random)")
$script:ConfigFilePath = [System.IO.Path]::Combine($script:TempDir, "config.xml")
$script:InstallerPath = [System.IO.Path]::Combine($script:TempDir, "setup.exe")
$script:ODTUrl = "https://ms365.ved.yt/ODTsetup.exe"  # Office Deployment Tool

# Check if running as administrator and self-elevate if needed
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Admin)) {
    Write-Host "This script requires administrator privileges. Attempting to elevate..." -ForegroundColor Yellow
    Start-Sleep -Seconds 1

    # Get the path of the currently running script
    $scriptPath = $MyInvocation.MyCommand.Definition
    
    # Prepare to relaunch the script with elevated privileges
    $proc = Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -Verb RunAs -PassThru

    # Exit the non-elevated script
    exit
}

Write-Host "Running with administrator privileges!" -ForegroundColor Green

# Create temp directory
if (-not (Test-Path -Path $script:TempDir)) {
    New-Item -ItemType Directory -Path $script:TempDir -Force | Out-Null
    Write-Host "Created temporary directory: $script:TempDir"
}

# Function to download Office Deployment Tool
function Get-OfficeTool {
    try {
        Write-Host "Downloading Office Deployment Tool from: $script:ODTUrl" -ForegroundColor Cyan
        $odtSetupPath = [System.IO.Path]::Combine($script:TempDir, "setup.exe")
        
        Invoke-WebRequest -Uri $script:ODTUrl -OutFile $odtSetupPath
        
        if (Test-Path $odtSetupPath) {
            Write-Host "Office Deployment Tool downloaded successfully." -ForegroundColor Green
            return $true
        } else {
            Write-Host "Failed to download Office Deployment Tool" -ForegroundColor Red
            return $false
        }
    } catch {
        Write-Host "Error downloading Office Deployment Tool: $_" -ForegroundColor Red
        return $false
    }
}

function Show-Menu {
    param (
        [string]$Title = 'Office 365 Installation Configuration'
    )
    Clear-Host
    Write-Host "================ $Title ================"
    Write-Host
    Write-Host "Current Configuration:"
    Write-Host "  1. Architecture: $script:Architecture-bit"
    Write-Host "  2. Language: $script:Language"
    Write-Host "  3. Update Channel: $script:Channel"
    Write-Host "  4. Excluded Apps: $($script:ExcludedApps -join ', ')"
    Write-Host
    Write-Host "Options:"
    Write-Host "  5. Proceed with Installation"
    Write-Host "  6. Exit"
    Write-Host
}

function Select-Architecture {
    Clear-Host
    Write-Host "=== Select Office Architecture ==="
    Write-Host "  1. 32-bit"
    Write-Host "  2. 64-bit"
    
    $selection = Read-Host "Enter your choice (1-2)"
    
    switch ($selection) {
        "1" { $script:Architecture = "32" }
        "2" { $script:Architecture = "64" }
        default { Write-Host "Invalid selection, keeping current setting." -ForegroundColor Yellow }
    }
}

function Select-Language {
    Clear-Host
    Write-Host "=== Select Office Language ==="
    Write-Host "  1. English (US) - en-us"
    Write-Host "  2. English (UK) - en-gb"
    Write-Host "  3. French - fr-fr"
    Write-Host "  4. German - de-de"
    Write-Host "  5. Spanish - es-es"
    Write-Host "  6. Custom"
    
    $selection = Read-Host "Enter your choice (1-6)"
    
    switch ($selection) {
        "1" { $script:Language = "en-us" }
        "2" { $script:Language = "en-gb" }
        "3" { $script:Language = "fr-fr" }
        "4" { $script:Language = "de-de" }
        "5" { $script:Language = "es-es" }
        "6" { 
            $customLang = Read-Host "Enter language code (e.g., ja-jp)"
            if ($customLang -match "^[a-z]{2}-[a-z]{2}$") {
                $script:Language = $customLang 
            }
            else {
                Write-Host "Invalid language code format, keeping current setting." -ForegroundColor Yellow
            }
        }
        default { Write-Host "Invalid selection, keeping current setting." -ForegroundColor Yellow }
    }
}

function Select-Channel {
    Clear-Host
    Write-Host "=== Select Office Update Channel ==="
    Write-Host "  1. Current (Monthly updates with latest features)"
    Write-Host "  2. MonthlyEnterprise (Monthly updates with security fixes only)"
    Write-Host "  3. SemiAnnual (Updates every six months)"
    Write-Host "  4. SemiAnnualPreview (Preview of semi-annual features)"
    
    $selection = Read-Host "Enter your choice (1-4)"
    
    switch ($selection) {
        "1" { $script:Channel = "Current" }
        "2" { $script:Channel = "MonthlyEnterprise" }
        "3" { $script:Channel = "SemiAnnual" }
        "4" { $script:Channel = "SemiAnnualPreview" }
        default { Write-Host "Invalid selection, keeping current setting." -ForegroundColor Yellow }
    }
}

function Manage-ExcludedApps {
    $appOptions = @{
        "1" = "Access"
        "2" = "Excel"
        "3" = "Groove"
        "4" = "Lync"
        "5" = "OneDrive"
        "6" = "OneNote"
        "7" = "Outlook"
        "8" = "PowerPoint"
        "9" = "Publisher"
        "10" = "Teams"
        "11" = "Word"
    }
    
    $done = $false
    
    while (-not $done) {
        Clear-Host
        Write-Host "=== Manage Excluded Apps ==="
        Write-Host "Current exclusions: $($script:ExcludedApps -join ', ')"
        Write-Host
        
        foreach ($key in $appOptions.Keys | Sort-Object) {
            $app = $appOptions[$key]
            $status = if ($script:ExcludedApps -contains $app) { "[Excluded]" } else { "[Included]" }
            Write-Host "  $key. $app $status"
        }
        
        Write-Host
        Write-Host "  0. Done"
        
        $selection = Read-Host "Enter number to toggle app inclusion/exclusion (0 to finish)"
        
        if ($selection -eq "0") {
            $done = $true
        }
        elseif ($appOptions.ContainsKey($selection)) {
            $app = $appOptions[$selection]
            if ($script:ExcludedApps -contains $app) {
                $script:ExcludedApps = $script:ExcludedApps | Where-Object { $_ -ne $app }
                Write-Host "$app will be included" -ForegroundColor Green
            }
            else {
                $script:ExcludedApps += $app
                Write-Host "$app will be excluded" -ForegroundColor Yellow
            }
            Start-Sleep -Seconds 1
        }
        else {
            Write-Host "Invalid selection" -ForegroundColor Red
            Start-Sleep -Seconds 1
        }
    }
}

function Create-ConfigXML {
    $excludeXml = ""
    foreach ($app in $script:ExcludedApps) {
        $excludeXml += "      <ExcludeApp ID=`"$app`" />`r`n"
    }
    
    $configXml = @"
<Configuration ID="$(New-Guid)">
  <Add OfficeClientEdition="$script:Architecture" Channel="$script:Channel">
    <Product ID="O365ProPlusRetail">
      <Language ID="$script:Language" />
$excludeXml
    </Product>
  </Add>
  <Updates Enabled="TRUE" />
  <RemoveMSI />
  <Display Level="Full" AcceptEULA="TRUE" />
</Configuration>
"@
    
    $configXml | Out-File -FilePath $script:ConfigFilePath -Encoding utf8
    Write-Host "Configuration file created at: $script:ConfigFilePath" -ForegroundColor Green
}

# Download Office Deployment Tool
Write-Host "`nChecking for Office Deployment Tool..." -ForegroundColor Cyan
$toolResult = Get-OfficeTool
if (-not $toolResult) {
    Write-Host "Failed to download Office Deployment Tool. Please check your internet connection and try again." -ForegroundColor Red
    Write-Host "Alternatively, you can manually download it from: https://www.microsoft.com/en-us/download/details.aspx?id=49117" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

# Main menu loop
$exit = $false

while (-not $exit) {
    Show-Menu
    $selection = Read-Host "Please make a selection (1-6)"
    
    switch ($selection) {
        "1" { Select-Architecture }
        "2" { Select-Language }
        "3" { Select-Channel }
        "4" { Manage-ExcludedApps }
        "5" { 
            Create-ConfigXML
            
            Write-Host "`nPreparing to install Office 365 with these settings:" -ForegroundColor Cyan
            Write-Host "  Architecture: $script:Architecture-bit"
            Write-Host "  Language: $script:Language"
            Write-Host "  Update Channel: $script:Channel"
            Write-Host "  Excluded Apps: $($script:ExcludedApps -join ', ')"
            
            $confirm = Read-Host "`nProceed with installation? (Y/N)"
            if ($confirm -eq "Y" -or $confirm -eq "y") {
                Write-Host "`nStarting Office 365 installation. This may take several minutes..." -ForegroundColor Cyan
                try {
                    Start-Process -FilePath $script:InstallerPath -ArgumentList "/configure `"$script:ConfigFilePath`"" -Wait
                    Write-Host "Microsoft Office 365 installation completed." -ForegroundColor Green
                    
                    # Verify installation
                    $officeInstalled = $false
                    $regPaths = @(
                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
                    )
                    
                    foreach ($path in $regPaths) {
                        foreach ($key in (Get-ChildItem -Path $path -ErrorAction SilentlyContinue)) {
                            if ($key.GetValue('DisplayName') -like '*Microsoft 365*') {
                                $officeVersion = $key.GetValue('DisplayName')
                                $officeInstalled = $true
                                break
                            }
                        }
                        if ($officeInstalled) { break }
                    }
                    
                    if ($officeInstalled) {
                        Write-Host "Verification successful: $officeVersion installed." -ForegroundColor Green
                    }
                    else {
                        Write-Host "Verification notice: Microsoft 365 installation not detected in registry. This may be normal if you selected specific components." -ForegroundColor Yellow
                    }
                }
                catch {
                    Write-Host "Error during installation: $_" -ForegroundColor Red
                }
                
                # Clean up temporary files
                Write-Host "`nCleaning up temporary files..." -ForegroundColor Cyan
                try {
                    Remove-Item -Path $script:TempDir -Recurse -Force
                    Write-Host "Temporary files removed successfully." -ForegroundColor Green
                }
                catch {
                    Write-Host "Warning: Unable to remove all temporary files: $_" -ForegroundColor Yellow
                }
                
                $exit = $true
            }
        }
        "6" { 
            $exit = $true 
            Write-Host "Installation cancelled." -ForegroundColor Yellow
            
            # Clean up temporary files if cancelling
            if (Test-Path -Path $script:TempDir) {
                Write-Host "Cleaning up temporary files..." -ForegroundColor Cyan
                Remove-Item -Path $script:TempDir -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
        default { Write-Host "Invalid selection, please try again." -ForegroundColor Red }
    }
}