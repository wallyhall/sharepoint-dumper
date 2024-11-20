# README:
#
# This script attempts to export/mirror/archive a Sharepoint site to a local drive/directory
# (which may be a mounted remote network drive).
#
# It is only known to capture "library" contents, i.e., documents/files/site pages (ASPX), etc.
#
# Known things it does _NOT_ capture, include "lists" and 3rd party "connected application data"
# (like LucidCharts).
#
# Library contents is downloaded in a single-threaded manner.  This may take a very long time
# for particularly large Sharepoint sites.
#
# This script relies on the "PnP.PowerShell" module, which in turn requires an Entra ID
# application.  You can use this script to try and create that for you, or otherwise follow
# their published documentation:
#
#   https://pnp.github.io/powershell/articles/registerapplication.html
#
# Otherwise, glhf ... -DryRun is always recommended in the first instance.
#
#
# TL;DR / example usage:
# 
#   Install-Module -Name `"PnP.PowerShell`" -AllowClobber
#   sharepoint.ps1 -CreateAppId -TenantId <YOUR_AZURE_TENANT_ID>
#   sharepoint.ps1 \
#      -AppId <YOUR_NEW_AZURE_APP_ID> \
#      -CompanyUrl <YOUR_COMPANY>.sharepoint.com \
#      -ExportPath c:\temp\sharepoint_exports \
#      -Site <YOUR_SITE> \
#      -DryRun
#

# Parse command-line inputs
[CmdletBinding(DefaultParameterSetName = 'export')]
param(
    [Parameter(ParameterSetName = 'export', Mandatory = $true)][String]$ExportPath,
    [Parameter(ParameterSetName = 'export', Mandatory = $true)][String]$CompanyUrl,
    [Parameter(ParameterSetName = 'export', Mandatory = $true)][String]$Site,
    [Parameter(ParameterSetName = 'export')][String]$AppId = $null,
    [Parameter(ParameterSetName = 'export')][switch]$DryRun = $False,

    [Parameter(ParameterSetName = 'createAppId', Mandatory = $true)][switch]$CreateAppId,
    [Parameter(ParameterSetName = 'createAppId', Mandatory = $true)][String]$TenantId
)

#########################

# Sanity check PS version.
if (!($PSVersionTable.PSVersion.Major -ge 7)) {
    Write-Error -Message "This script requires PS >= 7."
    Exit 1
}

#########################

# Check the necessary PowerShell module is found
try {
    Import-Module PnP.PowerShell || Throw
} catch {
    Write-Host "This script requires the PnP.PowerShell module to be installed, but it has not been found."
    Write-Host "To install it, run the following command:"
    Write-Host -ForegroundColor Yellow "`n  Install-Module -Name ""PnP.PowerShell"" -AllowClobber`n"
    Exit 2
}

#########################

# "Create Entra App ID" mode
if ($CreateAppId) {
    Write-Host "Creating an Entra ID Application ... this may require interaction with your browser."

    try {
        Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "Sharepoint Archive PS Script" -Tenant $TenantId -Interactive
    } catch {
        Write-Host "Creation of Entra ID Application failed, reported error was:"
        Write-Host -ForegroundColor Yellow "`n  $_`n"
        Exit 4
    }

    Write-Host "Entra ID Application created.  You can now supply its App ID to this script."
    Exit
}

#########################

# Explain to the user why a AppId is required
if ([String]::IsNullOrWhiteSpace($AppId)) {
    Write-Host "An Entra ID Application (and associated App ID) is required for this script to operate."
    Write-Host "Details on why can be found in the PnP.PowerShell documentation:"
    Write-Host -ForegroundColor Yellow "`n  https://pnp.github.io/powershell/articles/registerapplication.html`n"
    Write-Host "You can follow the manual instructions linked above, or re-run this script with the '-CreateAppId' argument."
    Exit 3
}

#########################

# Connect to SharePoint Online.
$siteUrl = "$CompanyUrl/sites/$Site"
Write-Host "Connecting to '$siteUrl', this may require authenticating in your browser..."
Connect-PnPOnline -Url $siteUrl -Interactive -ClientId $AppId

#########################

function Create-SiteMap {
    Write-Host "Generating site map of resources from the Sharepoint site navigation menu."

    if (!$DryRun) {
        Get-PnPNavigationNode -Tree > $(Join-Path -Path $ExportPath -ChildPath "/sites/$($Site)_navigation.txt")
    } else {
        Write-Host "The navigation tree looks like this:`n"
        Get-PnPNavigationNode -Tree
        Write-Host ""
    }
}

# Recurse depth-first down every folder structure, downloading every file found.
function Download-Folder {
    param (
        $Folder
    )

    $mirroredFolderPath = Join-Path -Path $ExportPath -ChildPath $Folder

    # Get and mirror all sub-folders 
    Write-Host "Finding all sub-folders in '$Folder'..."
    Get-PnPFolderInFolder -Identity $Folder | ForEach-Object {
        $mirroredSubFolderPath = Join-Path -Path $mirroredFolderPath -ChildPath $_.Name
        "  ...mirroring folder '$($_.Name)' to '$mirroredSubFolderPath'"

        if (!$DryRun) {
            New-Item -ItemType Directory -Path $mirroredSubFolderPath -Force > $null
        }

        # Recurse (!!)
        Download-Folder -Folder "$Folder/$($_.Name)"
    }

    # Get and mirror all files
    Write-Host "Finding all files in '$Folder'..."
    Get-PnPFileInFolder -Identity $Folder | ForEach-Object {
        $mirroredFilePath = Join-Path -Path $mirroredFolderPath -ChildPath $_.Name
        "  ...downloading file '$($_.Name)' to '$mirroredFilePath'"

        if (!$DryRun) {
            Get-PnPFile -Url "$Folder/$($_.Name)" -Path $mirroredFolderPath -FileName $_.Name -AsFile
        }
    }
}

# Notify user if in dry-run mode
if ($DryRun) {
    Write-Host -ForegroundColor Cyan "Executing in dry-run mode.  Nothing will actually be downloaded."
}

# Test for permissions on the site
try {
    Get-PnPFileInFolder -Identity "/sites/$Site"
} catch {
    Write-Host "Investigating the Sharepoint site failed due to the following error:"
    Write-Host -ForegroundColor Yellow "`n  $_`n"
    Write-Host "The error above should indicate what went wrong, and imply how to fix it."
    Write-Host "If it is permissions related, ensure you have (and retry with) _owner_ rights on the site."
    Exit 5
}

# Create a directory mapping for future end-user reference
Create-SiteMap
# Download all documents, pages, etc
Download-Folder -Folder "/sites/$Site"

# Remind the user if in dry-run mode
if ($DryRun) {
    Write-Host -ForegroundColor Cyan "Executed in dry-run mode.  Nothing was actually be downloaded."
}
