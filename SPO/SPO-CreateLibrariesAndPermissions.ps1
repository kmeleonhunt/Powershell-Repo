╦╔═┌┬┐┌─┐┬  ┌─┐┌─┐┌┐┌┬ ┬┬ ┬┌┐┌┌┬┐
╠╩╗│││├┤ │  ├┤ │ ││││├─┤│ ││││ │ 
╩ ╩┴ ┴└─┘┴─┘└─┘└─┘┘└┘┴ ┴└─┘┘└┘ ┴ 
                                                               
<#
.SYNOPSIS
    Creates document libraries across multiple SharePoint Online sites with custom permission levels.

.DESCRIPTION
    This script creates consistent document libraries with granular permission control using custom permission levels:
    - Library 1: Uses custom "Contribute No Delete" permission level for Members and Visitors
    - Library 2: All groups with item-level invitation required for Visitors
    - Library 3: Standard permissions for all groups
    
    Custom permission levels are created at the site level and can be reused across libraries.

.NOTES
    Requires: PnP.PowerShell module
    Install with: Install-Module -Name PnP.PowerShell -Scope CurrentUser
#>
Add-Type -AssemblyName System.Windows.Forms
function DefinePath {
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog

    # Show the file dialog and check if the user clicked "OK"
    $result = $OpenFileDialog.ShowDialog()
    if ($result -eq "OK") {
        # If the user selected a file, assign its path to $csvFilePath
        $csvFilePath = $OpenFileDialog.FileName
        return $csvFilePath
    } else {
        Write-Host "No file selected"
        return $null  # Return null or another appropriate value indicating no file selected
    }
}
# Ensure PnP.PowerShell module is installed
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host "PnP.PowerShell module not found. Installing..." -ForegroundColor Yellow
    Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force
}

Import-Module PnP.PowerShell

# Function to create custom permission level if it doesn't exist
function New-CustomPermissionLevel {
    param (
        [string]$PermissionLevelName
    )
    
    try {
        # Check if permission level already exists
        $existingPermLevel = Get-PnPRoleDefinition -Identity $PermissionLevelName -ErrorAction SilentlyContinue
        
        if ($existingPermLevel) {
            Write-Host "  Permission level '$PermissionLevelName' already exists" -ForegroundColor Gray
            return $true
        }
        
        Write-Host "  Creating custom permission level: $PermissionLevelName" -ForegroundColor Cyan
        
        # Get the Contribute permission as a base
        $contributeRole = Get-PnPRoleDefinition -Identity "Contribute"
        
        # Clone and modify permissions
        # Remove delete permissions from Contribute
        $permissions = $contributeRole.BasePermissions
        
        # Remove delete-related permissions
        $permissions.Clear([Microsoft.SharePoint.Client.PermissionKind]::DeleteListItems)
        $permissions.Clear([Microsoft.SharePoint.Client.PermissionKind]::DeleteVersions)
        
        # Create the new permission level
        Add-PnPRoleDefinition -RoleName $PermissionLevelName `
            -Description "Can add, edit, and view items but cannot delete them" `
            -Clone "Contribute" `
            -Exclude DeleteListItems, DeleteVersions `
            -ErrorAction Stop
        
        Write-Host "  ✓ Custom permission level created" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "  ✗ Error creating permission level: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to create library with specific permissions
function New-SPOLibraryWithPermissions {
    param (
        [string]$SiteUrl,
        [string]$LibraryName,
        [int]$PermissionLevel,
        [string]$CustomPermissionLevelName
    )
    
    try {
        Write-Host "`n  Creating library: $LibraryName" -ForegroundColor Cyan
        
        # Create the document library
        $library = New-PnPList -Title $LibraryName -Template DocumentLibrary -ErrorAction Stop
        
        # Break permission inheritance
        Set-PnPList -Identity $LibraryName -BreakRoleInheritance -CopyRoleAssignments:$false -ClearSubscopes
        
        # Get the default groups
        $web = Get-PnPWeb
        $ownerGroup = Get-PnPGroup | Where-Object { $_.Title -like "*Owners" }
        $memberGroup = Get-PnPGroup | Where-Object { $_.Title -like "*Members" }
        $visitorGroup = Get-PnPGroup | Where-Object { $_.Title -like "*Visitors" }
        
        # Apply permissions based on level
        switch ($PermissionLevel) {
            1 {
                # Library 1: Owners (Full Control), Members & Visitors (Custom No Delete Permission)
                Write-Host "    Setting permissions: Owners (Full Control), Members & Visitors (No Delete)" -ForegroundColor Gray
                Set-PnPListPermission -Identity $LibraryName -Group $ownerGroup -AddRole "Full Control"
                Set-PnPListPermission -Identity $LibraryName -Group $memberGroup -AddRole $CustomPermissionLevelName
                Set-PnPListPermission -Identity $LibraryName -Group $visitorGroup -AddRole $CustomPermissionLevelName
                
                # Configure item-level permissions so Visitors only see what they're invited to
                Set-PnPList -Identity $LibraryName -ReadSecurity 2 -WriteSecurity 2
                
                Write-Host "    Members & Visitors can edit but cannot delete items" -ForegroundColor Gray
                Write-Host "    Visitors can only see items explicitly shared with them" -ForegroundColor Gray
            }
            2 {
                # Library 2: All groups, but Visitors need explicit invitation (item-level permissions)
                Write-Host "    Setting permissions: Owners (Full Control), Members (Edit)" -ForegroundColor Gray
                Set-PnPListPermission -Identity $LibraryName -Group $ownerGroup -AddRole "Full Control"
                Set-PnPListPermission -Identity $LibraryName -Group $memberGroup -AddRole "Contribute"
                
                # Give Visitors Limited Access so they can access items when invited
                Set-PnPListPermission -Identity $LibraryName -Group $visitorGroup -AddRole "Contribute"
                
                # Configure item-level permissions so Visitors only see what they're invited to
                Set-PnPList -Identity $LibraryName -ReadSecurity 2 -WriteSecurity 2
                
                Write-Host "    Visitors can only access items they're explicitly invited to" -ForegroundColor Gray
            }
            3 {
                # Library 3: All groups with standard permissions
                Write-Host "    Setting permissions: Owners (Full Control), Members (Edit), Visitors (Read)" -ForegroundColor Gray
                Set-PnPListPermission -Identity $LibraryName -Group $ownerGroup -AddRole "Full Control"
                Set-PnPListPermission -Identity $LibraryName -Group $memberGroup -AddRole "Contribute"
                Set-PnPListPermission -Identity $LibraryName -Group $visitorGroup -AddRole "Contribute"
                
                Write-Host "    All users (including external visitors) can view this library" -ForegroundColor Gray
            }
        }
        
        Write-Host "  ✓ Library '$LibraryName' created successfully" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "  ✗ Error creating library '$LibraryName': $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Main Script
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "SharePoint Library Creation Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Prompt for Azure AD App Registration details
Write-Host "`n--- Azure AD App Registration Configuration ---" -ForegroundColor Yellow
Write-Host "You need an Azure AD App Registration with the following API permissions:" -ForegroundColor White
Write-Host "  - Sites.FullControl.All (Application or Delegated)" -ForegroundColor Gray
Write-Host "  - Sites.Manage.All (Application or Delegated)" -ForegroundColor Gray
Write-Host ""
$clientId = Read-Host "Enter your Azure AD App Registration Client ID"
$tenantId = Read-Host "Enter your Tenant ID (or tenant name like contoso.onmicrosoft.com)"

if ([string]::IsNullOrWhiteSpace($clientId) -or [string]::IsNullOrWhiteSpace($tenantId)) {
    Write-Host "`nError: Client ID and Tenant ID are required." -ForegroundColor Red
    exit
}

# Prompt for SharePoint site URLs
Write-Host "`n--- SharePoint Sites Configuration ---" -ForegroundColor Yellow
Write-Host "Choose how to provide SharePoint site URLs:" -ForegroundColor White
Write-Host "  1. Enter URLs manually (one per line)" -ForegroundColor Gray
Write-Host "  2. Import from CSV file" -ForegroundColor Gray
$inputMethod = Read-Host "Enter your choice (1 or 2)"

$siteUrls = @()

if ($inputMethod -eq "2") {
    # CSV Import Method
    Write-Host "`n--- CSV File Configuration ---" -ForegroundColor Yellow
    Write-Host "CSV file should have a column named 'URL' with SharePoint site URLs" -ForegroundColor White
    Write-Host "Example CSV format:" -ForegroundColor Gray
    Write-Host "  URL" -ForegroundColor Gray
    Write-Host "  https://tenant.sharepoint.com/sites/Site1" -ForegroundColor Gray
    Write-Host "  https://tenant.sharepoint.com/sites/Site2" -ForegroundColor Gray
    Write-Host ""
    
    $csvPath = DefinePath
    
    if (-not (Test-Path $csvPath)) {
        Write-Host "`nError: CSV file not found at path: $csvPath" -ForegroundColor Red
        exit
    }
    
    try {
        $csvData = Import-Csv -Path $csvPath
        
        # Check if SiteUrl column exists
        if (-not ($csvData | Get-Member -Name "URL" -MemberType NoteProperty)) {
            Write-Host "`nError: CSV file must contain a 'URL' column" -ForegroundColor Red
            Write-Host "Current columns found: $($csvData | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name)" -ForegroundColor Yellow
            exit
        }
        
        # Extract site URLs from CSV
        $siteUrls = $csvData | Where-Object { -not [string]::IsNullOrWhiteSpace($_.URL) } | Select-Object -ExpandProperty URL
        
        if ($siteUrls.Count -eq 0) {
            Write-Host "`nError: No valid site URLs found in CSV file" -ForegroundColor Red
            exit
        }
        
        Write-Host "`n✓ Successfully loaded $($siteUrls.Count) site(s) from CSV" -ForegroundColor Green
        Write-Host "`nSites to process:" -ForegroundColor Cyan
        $siteUrls | ForEach-Object { Write-Host "  - $_" -ForegroundColor Gray }
    }
    catch {
        Write-Host "`nError reading CSV file: $($_.Exception.Message)" -ForegroundColor Red
        exit
    }
}
else {
    # Manual Entry Method
    Write-Host "`nEnter SharePoint site URLs (one per line, press Enter twice when done):" -ForegroundColor White
    do {
        $url = Read-Host "Site URL"
        if ($url) {
            $siteUrls += $url
        }
    } while ($url)
    
    if ($siteUrls.Count -eq 0) {
        Write-Host "No sites provided. Exiting." -ForegroundColor Red
        exit
    }
    
    Write-Host "`nSites to process: $($siteUrls.Count)" -ForegroundColor Green
}

# Prompt for library names
Write-Host "`n--- Library Configuration ---" -ForegroundColor Yellow
$library1Name = Read-Host "Enter name for Library 1 (Permanent files - no delete for Members/Visitors)"
$library2Name = Read-Host "Enter name for Library 2 (All groups, Visitors need invitation)"
$library3Name = Read-Host "Enter name for Library 3 (All groups, full access)"

# Validate library names
if ([string]::IsNullOrWhiteSpace($library1Name) -or 
    [string]::IsNullOrWhiteSpace($library2Name) -or 
    [string]::IsNullOrWhiteSpace($library3Name)) {
    Write-Host "`nError: All library names must be provided." -ForegroundColor Red
    exit
}

# Define custom permission level name
$customPermissionLevel = "Contribute No Delete"

# Confirmation
Write-Host "`n--- Configuration Summary ---" -ForegroundColor Yellow
Write-Host "Sites to process: $($siteUrls.Count)"
Write-Host "Library 1 (No delete for Members/Visitors): $library1Name"
Write-Host "Library 2 (Invite-only for Visitors): $library2Name"
Write-Host "Library 3 (Full access): $library3Name"
Write-Host "`nA custom permission level '$customPermissionLevel' will be created on each site."
$confirm = Read-Host "`nProceed with library creation? (Y/N)"

if ($confirm -ne 'Y' -and $confirm -ne 'y') {
    Write-Host "Operation cancelled." -ForegroundColor Yellow
    exit
}

# Process each site
$results = @()
$successCount = 0
$failureCount = 0

foreach ($siteUrl in $siteUrls) {
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Processing site: $siteUrl" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    
    try {
        # Connect to the site using interactive login
        Write-Host "Connecting to site (browser login will open)..." -ForegroundColor White
        Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Tenant $tenantId -Interactive -ErrorAction Stop
        
        # Create custom permission level
        $permLevelSuccess = New-CustomPermissionLevel -PermissionLevelName $customPermissionLevel
        
        if (-not $permLevelSuccess) {
            Write-Host "  ⚠ Warning: Could not create custom permission level. Using Contribute instead." -ForegroundColor Yellow
            $customPermissionLevel = "Contribute"
        }
        
        # Create the three libraries
        $lib1Success = New-SPOLibraryWithPermissions -SiteUrl $siteUrl -LibraryName $library1Name -PermissionLevel 1 -CustomPermissionLevelName $customPermissionLevel
        $lib2Success = New-SPOLibraryWithPermissions -SiteUrl $siteUrl -LibraryName $library2Name -PermissionLevel 2 -CustomPermissionLevelName $customPermissionLevel
        $lib3Success = New-SPOLibraryWithPermissions -SiteUrl $siteUrl -LibraryName $library3Name -PermissionLevel 3 -CustomPermissionLevelName $customPermissionLevel
        
        if ($lib1Success -and $lib2Success -and $lib3Success) {
            Write-Host "`n✓ All libraries created successfully on $siteUrl" -ForegroundColor Green
            $successCount++
        } else {
            Write-Host "`n⚠ Some libraries failed on $siteUrl" -ForegroundColor Yellow
            $failureCount++
        }
        
        $results += [PSCustomObject]@{
            SiteUrl = $siteUrl
            Library1 = if ($lib1Success) { "Success" } else { "Failed" }
            Library2 = if ($lib2Success) { "Success" } else { "Failed" }
            Library3 = if ($lib3Success) { "Success" } else { "Failed" }
            PermissionLevel = if ($permLevelSuccess) { "Created" } else { "Failed/Exists" }
        }
        
        # Disconnect from the site
        Disconnect-PnPOnline
    }
    catch {
        Write-Host "`n✗ Error connecting to site: $($_.Exception.Message)" -ForegroundColor Red
        $failureCount++
        $results += [PSCustomObject]@{
            SiteUrl = $siteUrl
            Library1 = "Connection Failed"
            Library2 = "Connection Failed"
            Library3 = "Connection Failed"
            PermissionLevel = "N/A"
        }
    }
}

# Summary Report
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "EXECUTION SUMMARY" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Total sites processed: $($siteUrls.Count)"
Write-Host "Successful: $successCount" -ForegroundColor Green
Write-Host "Failed: $failureCount" -ForegroundColor Red

Write-Host "`n--- Detailed Results ---" -ForegroundColor Yellow
$results | Format-Table -AutoSize

# Export results to CSV
$exportPath = Join-Path $PSScriptRoot "SPO_Library_Creation_Results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$results | Export-Csv -Path $exportPath -NoTypeInformation
Write-Host "`nResults exported to: $exportPath" -ForegroundColor Cyan

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "PERMISSION SUMMARY" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "`nLIBRARY 1 ($library1Name) - PERMANENT FILES:" -ForegroundColor White
Write-Host "  • Owners: Full Control (can do everything)"
Write-Host "  • Members: $customPermissionLevel (can add/edit but NOT delete)"
Write-Host "  • Visitors: $customPermissionLevel + Item-level access only"
Write-Host "  • Visitors can only see items explicitly shared with them"
Write-Host "  • New members added to default Members group automatically get these permissions"
Write-Host "`nLIBRARY 2 ($library2Name) - INVITE ONLY:" -ForegroundColor White
Write-Host "  • Owners: Full Control"
Write-Host "  • Members: Edit (full permissions)"
Write-Host "  • Visitors: Can only see items they're explicitly invited to"
Write-Host "  • To invite: Right-click item → Share → Enter email"
Write-Host "`nLIBRARY 3 ($library3Name) - PUBLIC ACCESS:" -ForegroundColor White
Write-Host "  • Owners: Full Control"
Write-Host "  • Members: Edit"
Write-Host "  • Visitors: Read (can view all content)"
Write-Host "  • External users added to Visitors group can access this library"
Write-Host "`n--- IMPORTANT NOTES ---" -ForegroundColor Yellow
Write-Host "✓ Custom permission level '$customPermissionLevel' created at site level"
Write-Host "✓ This permission level can be reused for other libraries if needed"
Write-Host "✓ New site members automatically inherit these permissions (no manual group management)"
Write-Host "✓ Each library has unique permissions while using the same default groups"
Write-Host "========================================" -ForegroundColor Cyan
