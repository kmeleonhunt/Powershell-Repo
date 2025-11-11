<#

╦╔═┌┬┐┌─┐┬  ┌─┐┌─┐┌┐┌┬ ┬┬ ┬┌┐┌┌┬┐
╠╩╗│││├┤ │  ├┤ │ ││││├─┤│ ││││ │ 
╩ ╩┴ ┴└─┘┴─┘└─┘└─┘┘└┘┴ ┴└─┘┘└┘ ┴ 

    .SYNOPSIS
      License management script for Microsoft 365 with support for multiple license types
    
    .DESCRIPTION
      Manage user licenses including Business Basic, Business Premium, and Threat Intelligence
      Supports single user operations and bulk CSV imports
    
    .REQUIREMENTS
      - Microsoft.Graph PowerShell module
      - Appropriate Graph API permissions
      - Admin privileges for license assignment
      - To edit the script with your own licenses requirement please refer to https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference for SkuPartName
#>

# Required scopes for license management
$Scopes = @('User.ReadWrite.All', 'Directory.ReadWrite.All', 'Organization.Read.All')

# Connect to Microsoft Graph
Connect-MgGraph -Scopes $Scopes

$Context = Get-MgContext
Write-Host "Connected with scopes: $($Context.Scopes -join ', ')" -ForegroundColor Green

# Load available licenses
$365BusinessBasic = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq 'O365_BUSINESS_ESSENTIALS' 
$365BusinessPremium = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq 'SPB'
$ThreatIntelligence = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq 'THREAT_INTELLIGENCE'

Write-Host "`n=== License Management Tool ===" -ForegroundColor Cyan
Write-Host "Type 'Get-One' to manage a single user" -ForegroundColor Yellow
Write-Host "Type 'Get-Multiple' to bulk process from CSV" -ForegroundColor Yellow
Write-Host "================================`n" -ForegroundColor Cyan

function Get-One {
    # Display unlicensed users option
    $Q1 = Read-Host "`nDo you wish to see all unlicensed users? (Y/N)"
    
    if ($Q1 -eq "Y") {
        Write-Host "`nRetrieving unlicensed users..." -ForegroundColor Yellow
        $Unlicensed = Get-MgUser -Filter 'assignedLicenses/$count eq 0' -ConsistencyLevel eventual -CountVariable unlicensedUserCount -All
        Write-Host "Found $unlicensedUserCount unlicensed users`n" -ForegroundColor Green
        $Unlicensed | Select-Object DisplayName, UserPrincipalName, JobTitle | Format-Table -AutoSize
        Pause
    }

    # Get target user
    $userUPN = Read-Host "`nEnter user email (UPN)"
    
    # Verify user exists
    try {
        $User = Get-MgUser -UserId $userUPN -ErrorAction Stop
        Write-Host "User found: $($User.DisplayName)" -ForegroundColor Green
        
        # Show current licenses
        $CurrentLicenses = Get-MgUserLicenseDetail -UserId $userUPN
        if ($CurrentLicenses) {
            Write-Host "`nCurrent licenses:" -ForegroundColor Cyan
            $CurrentLicenses | Select-Object SkuPartNumber, SkuId | Format-Table -AutoSize
        } else {
            Write-Host "`nUser has no licenses assigned" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "Error: User not found - $($_.Exception.Message)" -ForegroundColor Red
        return
    }
    
    # Show available licenses
    $Q2 = Read-Host "`nShow available licenses on tenant? (Y/N)"
    
    if ($Q2 -eq "Y") {
        Write-Host "`nAvailable licenses:" -ForegroundColor Yellow
        $AllLicenses = Get-MgSubscribedSku | Select-Object SkuPartNumber, SkuId, 
            @{N='Available';E={$_.PrepaidUnits.Enabled - $_.ConsumedUnits}}, 
            @{N='Total';E={$_.PrepaidUnits.Enabled}}
        $AllLicenses | Format-Table -AutoSize
        Pause
    }
    
    # Add or remove license
    $Q3 = Read-Host "`nDo you want to ADD or REMOVE a license? (add/remove)"
    $LType = Read-Host "Select license type:`n  [1] Basic (O365_BUSINESS_ESSENTIALS)`n  [2] Business Premium (SPB)`n  [3] Threat Intelligence (THREAT_INTELLIGENCE)`nEnter choice (1/2/3)"
    
    # Determine which license to use
    switch ($LType) {
        "1" { $SelectedLicense = $365BusinessBasic; $LicenseName = "Business Basic" }
        "2" { $SelectedLicense = $365BusinessPremium; $LicenseName = "Business Premium" }
        "3" { $SelectedLicense = $ThreatIntelligence; $LicenseName = "Threat Intelligence" }
        default { 
            Write-Host "Invalid selection" -ForegroundColor Red
            return
        }
    }
    
    if (-not $SelectedLicense) {
        Write-Host "Error: $LicenseName license not found in tenant" -ForegroundColor Red
        return
    }
    
    try {
        if ($Q3 -eq "add") {
            Set-MgUserLicense -UserId $userUPN -AddLicenses @{SkuId = $SelectedLicense.SkuId} -RemoveLicenses @()
            Write-Host "`nSuccess: Added $LicenseName license to $userUPN" -ForegroundColor Green
        }
        elseif ($Q3 -eq "remove") {
            Set-MgUserLicense -UserId $userUPN -RemoveLicenses @($SelectedLicense.SkuId) -AddLicenses @{}
            Write-Host "`nSuccess: Removed $LicenseName license from $userUPN" -ForegroundColor Green
        }
        else {
            Write-Host "Invalid action specified" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Get-Multiple {
    Write-Host "`n=== Bulk License Management ===" -ForegroundColor Cyan
    Write-Host "CSV format should have a 'UPN' column with user email addresses" -ForegroundColor Yellow
    Write-Host "Optional 'Action' column: 'add' or 'remove'" -ForegroundColor Yellow
    Write-Host "Optional 'LicenseType' column: 'Basic', 'Business', or 'ThreatIntel'" -ForegroundColor Yellow
    
    Pause
    
    # File picker
    Add-Type -AssemblyName System.Windows.Forms
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $OpenFileDialog.Title = "Select CSV file"
    
    if ($OpenFileDialog.ShowDialog() -ne "OK") {
        Write-Host "No file selected - operation cancelled" -ForegroundColor Yellow
        return
    }
    
    try {
        $CSV = Import-Csv $OpenFileDialog.FileName
        Write-Host "`nLoaded $($CSV.Count) users from CSV" -ForegroundColor Green
    }
    catch {
        Write-Host "Error reading CSV: $($_.Exception.Message)" -ForegroundColor Red
        return
    }
    
    # If CSV doesn't have Action/LicenseType columns, prompt for defaults
    $DefaultAction = Read-Host "`nDefault action for all users? (add/remove)"
    $DefaultLicense = Read-Host "Default license type? [1] Basic [2] Business [3] ThreatIntel (1/2/3)"
    
    $Results = @()
    
    foreach ($Row in $CSV) {
        $UserUPN = $Row.UPN
        $Action = if ($Row.Action) { $Row.Action } else { $DefaultAction }
        $LicenseChoice = if ($Row.LicenseType) { $Row.LicenseType } else { $DefaultLicense }
        
        Write-Host "`nProcessing: $UserUPN" -ForegroundColor Cyan
        
        # Select license
        switch ($LicenseChoice) {
            {$_ -in "1", "Basic"} { 
                $License = $365BusinessBasic
                $LicenseName = "Business Basic"
            }
            {$_ -in "2", "Business", "Premium"} { 
                $License = $365BusinessPremium
                $LicenseName = "Business Premium"
            }
            {$_ -in "3", "ThreatIntel", "Threat"} { 
                $License = $ThreatIntelligence
                $LicenseName = "Threat Intelligence"
            }
            default { 
                Write-Host "  Invalid license type for $UserUPN - skipped" -ForegroundColor Red
                $Results += [PSCustomObject]@{
                    UPN = $UserUPN
                    Action = $Action
                    License = "Unknown"
                    Status = "Failed - Invalid License"
                }
                continue
            }
        }
        
        if (-not $License) {
            Write-Host "  $LicenseName not available in tenant - skipped" -ForegroundColor Red
            $Results += [PSCustomObject]@{
                UPN = $UserUPN
                Action = $Action
                License = $LicenseName
                Status = "Failed - License Not Available"
            }
            continue
        }
        
        try {
            if ($Action -eq "add") {
                Set-MgUserLicense -UserId $UserUPN -AddLicenses @{SkuId = $License.SkuId} -RemoveLicenses @() -ErrorAction Stop
                Write-Host "  Added $LicenseName" -ForegroundColor Green
                $Status = "Success - Added"
            }
            elseif ($Action -eq "remove") {
                Set-MgUserLicense -UserId $UserUPN -RemoveLicenses @($License.SkuId) -AddLicenses @{} -ErrorAction Stop
                Write-Host "  Removed $LicenseName" -ForegroundColor Green
                $Status = "Success - Removed"
            }
            else {
                Write-Host "  Invalid action: $Action" -ForegroundColor Red
                $Status = "Failed - Invalid Action"
            }
            
            $Results += [PSCustomObject]@{
                UPN = $UserUPN
                Action = $Action
                License = $LicenseName
                Status = $Status
            }
        }
        catch {
            Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
            $Results += [PSCustomObject]@{
                UPN = $UserUPN
                Action = $Action
                License = $LicenseName
                Status = "Failed - $($_.Exception.Message)"
            }
        }
    }
    
    # Display summary
    Write-Host "`n=== Processing Summary ===" -ForegroundColor Cyan
    $Results | Format-Table -AutoSize
    
    # Export results
    $ExportPath = Join-Path (Split-Path $OpenFileDialog.FileName) "LicenseResults_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $Results | Export-Csv -Path $ExportPath -NoTypeInformation
    Write-Host "`nResults exported to: $ExportPath" -ForegroundColor Green
}

# Script execution entry point
# Uncomment one of the following to run:
# Get-One
# Get-Multiple
