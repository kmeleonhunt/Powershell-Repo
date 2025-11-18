<#
.SYNOPSIS
    Creates shared mailboxes with aliases and delegates from a CSV file.

.DESCRIPTION
    This script reads a CSV file containing shared mailbox information and:
    - Creates the shared mailbox
    - Adds email aliases
    - Assigns delegates with Full Access and Send As permissions

.EXAMPLE
    Just run the script from VS Code with F5 - it will open a file dialog to select your CSV

.NOTES
    CSV Format Required:
    DisplayName,PrimaryEmail,Aliases,Delegates
    "Sales Team","sales@contoso.com","salesinfo@contoso.com;sales.team@contoso.com","user1@contoso.com;user2@contoso.com"
    
    - Aliases: Multiple aliases separated by semicolon (;)
    - Delegates: Multiple delegates separated by semicolon (;)
    - Leave Aliases or Delegates empty if not needed
#>

# Load Windows Forms for file dialogs
Add-Type -AssemblyName System.Windows.Forms

# Function to select input file
function DefinePath {
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $OpenFileDialog.Title = "Select CSV file with shared mailbox information"
    
    $result = $OpenFileDialog.ShowDialog()
    if ($result -eq "OK") {
        return $OpenFileDialog.FileName
    } else {
        Write-Host "No file selected" -ForegroundColor Red
        return $null
    }
}

# Function to save output file
function SaveFile {
    param(
        [string]$DefaultFileName = "SharedMailbox_Results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    )
    
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.Filter = "CSV files (*.csv)|*.csv"
    $SaveFileDialog.FileName = $DefaultFileName
    $SaveFileDialog.Title = "Save results to CSV"
    
    if ($SaveFileDialog.ShowDialog() -eq "OK") {
        return $SaveFileDialog.FileName
    } else {
        return $null
    }
}

# Function to write colored output
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Type = "Info"
    )
    
    switch ($Type) {
        "Success" { Write-Host $Message -ForegroundColor Green }
        "Error"   { Write-Host $Message -ForegroundColor Red }
        "Warning" { Write-Host $Message -ForegroundColor Yellow }
        default   { Write-Host $Message -ForegroundColor Cyan }
    }
}

# Get CSV file path using file dialog
$CSVPath = DefinePath

if ([string]::IsNullOrWhiteSpace($CSVPath)) {
    Write-Host "Script cancelled - no file selected" -ForegroundColor Red
    exit 1
}

# Check if Exchange Online module is available
try {
    Write-ColorOutput "Checking for Exchange Online Management module..." "Info"
    
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-ColorOutput "Exchange Online Management module not found. Installing..." "Warning"
        Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
    }
    
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Write-ColorOutput "Exchange Online Management module loaded successfully." "Success"
}
catch {
    Write-ColorOutput "Failed to load Exchange Online Management module: $($_.Exception.Message)" "Error"
    exit 1
}

# Connect to Exchange Online
try {
    Write-ColorOutput "`nConnecting to Exchange Online..." "Info"
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-ColorOutput "Connected to Exchange Online successfully." "Success"
}
catch {
    Write-ColorOutput "Failed to connect to Exchange Online: $($_.Exception.Message)" "Error"
    exit 1
}

# Import CSV file
try {
    Write-ColorOutput "`nImporting CSV file from: $CSVPath" "Info"
    $mailboxesRaw = Import-Csv -Path $CSVPath -ErrorAction Stop
    
    # Filter out empty rows (rows where all required fields are empty)
    $mailboxes = $mailboxesRaw | Where-Object { 
        -not [string]::IsNullOrWhiteSpace($_.DisplayName) -and 
        -not [string]::IsNullOrWhiteSpace($_.PrimaryEmail) 
    }
    
    if ($mailboxes.Count -eq 0) {
        Write-ColorOutput "No valid mailboxes found in CSV file." "Error"
        Disconnect-ExchangeOnline -Confirm:$false
        exit 1
    }
    
    Write-ColorOutput "CSV imported successfully. Found $($mailboxes.Count) valid mailbox(es) to process." "Success"
    
    # Display mailboxes to be processed
    Write-ColorOutput "`nMailboxes to be created:" "Info"
    foreach ($mb in $mailboxes) {
        Write-ColorOutput "  - $($mb.DisplayName) ($($mb.PrimaryEmail))" "Info"
    }
}
catch {
    Write-ColorOutput "Failed to import CSV file: $($_.Exception.Message)" "Error"
    Disconnect-ExchangeOnline -Confirm:$false
    exit 1
}

# Validate CSV headers
$requiredHeaders = @('DisplayName', 'PrimaryEmail', 'Aliases', 'Delegates')
$csvHeaders = $mailboxes[0].PSObject.Properties.Name

foreach ($header in $requiredHeaders) {
    if ($header -notin $csvHeaders) {
        Write-ColorOutput "CSV file is missing required header: $header" "Error"
        Disconnect-ExchangeOnline -Confirm:$false
        exit 1
    }
}

# Process results tracking
$results = @()

# Process each mailbox
foreach ($mailbox in $mailboxes) {
    Write-ColorOutput "`n$('='*80)" "Info"
    Write-ColorOutput "Processing: $($mailbox.DisplayName)" "Info"
    Write-ColorOutput "$('='*80)" "Info"
    
    $result = [PSCustomObject]@{
        DisplayName = $mailbox.DisplayName
        PrimaryEmail = $mailbox.PrimaryEmail
        MailboxCreated = $false
        AliasesAdded = 0
        DelegatesAdded = 0
        Errors = @()
    }
    
    # Create shared mailbox
    try {
        Write-ColorOutput "`nCreating shared mailbox: $($mailbox.PrimaryEmail)" "Info"
        
        # Check if mailbox already exists
        $existingMailbox = Get-Mailbox -Identity $mailbox.PrimaryEmail -ErrorAction SilentlyContinue
        
        if ($existingMailbox) {
            Write-ColorOutput "Mailbox already exists: $($mailbox.PrimaryEmail)" "Warning"
            $result.Errors += "Mailbox already exists"
        }
        else {
            New-Mailbox -Shared -Name $mailbox.DisplayName -PrimarySmtpAddress $mailbox.PrimaryEmail -ErrorAction Stop | Out-Null
            Write-ColorOutput "Shared mailbox created successfully!" "Success"
            $result.MailboxCreated = $true
            
            # Wait a moment for mailbox to be fully provisioned
            Start-Sleep -Seconds 5
        }
    }
    catch {
        $errorMsg = "Failed to create mailbox: $($_.Exception.Message)"
        Write-ColorOutput $errorMsg "Error"
        $result.Errors += $errorMsg
        $results += $result
        continue
    }
    
    # Add aliases
    if (-not [string]::IsNullOrWhiteSpace($mailbox.Aliases)) {
        Write-ColorOutput "`nAdding email aliases..." "Info"
        $aliases = $mailbox.Aliases -split ';' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        
        foreach ($alias in $aliases) {
            $alias = $alias.Trim()
            try {
                Set-Mailbox -Identity $mailbox.PrimaryEmail -EmailAddresses @{Add = $alias} -ErrorAction Stop
                Write-ColorOutput "  - Added alias: $alias" "Success"
                $result.AliasesAdded++
            }
            catch {
                $errorMsg = "Failed to add alias '$alias': $($_.Exception.Message)"
                Write-ColorOutput "  - $errorMsg" "Error"
                $result.Errors += $errorMsg
            }
        }
    }
    else {
        Write-ColorOutput "`nNo aliases to add." "Info"
    }
    
    # Add delegates
    if (-not [string]::IsNullOrWhiteSpace($mailbox.Delegates)) {
        Write-ColorOutput "`nAdding delegates..." "Info"
        $delegates = $mailbox.Delegates -split ';' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        
        foreach ($delegate in $delegates) {
            $delegate = $delegate.Trim()
            try {
                # Verify delegate exists
                $delegateUser = Get-Mailbox -Identity $delegate -ErrorAction Stop
                
                # Add Full Access permission
                Add-MailboxPermission -Identity $mailbox.PrimaryEmail -User $delegate -AccessRights FullAccess -InheritanceType All -AutoMapping $true -ErrorAction Stop | Out-Null
                Write-ColorOutput "  - Added Full Access for: $delegate" "Success"
                
                # Add Send As permission
                Add-RecipientPermission -Identity $mailbox.PrimaryEmail -Trustee $delegate -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                Write-ColorOutput "  - Added Send As for: $delegate" "Success"
                
                $result.DelegatesAdded++
            }
            catch {
                $errorMsg = "Failed to add delegate '$delegate': $($_.Exception.Message)"
                Write-ColorOutput "  - $errorMsg" "Error"
                $result.Errors += $errorMsg
            }
        }
    }
    else {
        Write-ColorOutput "`nNo delegates to add." "Info"
    }
    
    $results += $result
}

# Display summary
Write-ColorOutput "`n`n$('='*80)" "Info"
Write-ColorOutput "SUMMARY" "Info"
Write-ColorOutput "$('='*80)" "Info"

foreach ($result in $results) {
    Write-ColorOutput "`nMailbox: $($result.DisplayName) ($($result.PrimaryEmail))" "Info"
    Write-ColorOutput "  Created: $($result.MailboxCreated)" $(if ($result.MailboxCreated) { "Success" } else { "Warning" })
    Write-ColorOutput "  Aliases Added: $($result.AliasesAdded)" "Info"
    Write-ColorOutput "  Delegates Added: $($result.DelegatesAdded)" "Info"
    
    if ($result.Errors.Count -gt 0) {
        Write-ColorOutput "  Errors:" "Error"
        foreach ($errors in $result.Errors) {
            Write-ColorOutput "    - $errors" "Error"
        }
    }
}

Write-ColorOutput "`n$('='*80)" "Info"
Write-ColorOutput "Total Mailboxes Processed: $($results.Count)" "Info"
Write-ColorOutput "Successfully Created: $(($results | Where-Object {$_.MailboxCreated}).Count)" "Success"
Write-ColorOutput "Failed: $(($results | Where-Object {-not $_.MailboxCreated}).Count)" $(if (($results | Where-Object {-not $_.MailboxCreated}).Count -gt 0) { "Error" } else { "Success" })

# Disconnect from Exchange Online
Write-ColorOutput "`nDisconnecting from Exchange Online..." "Info"
Disconnect-ExchangeOnline -Confirm:$false
Write-ColorOutput "Disconnected successfully." "Success"

# Export results to CSV using SaveFile dialog
$resultsPath = SaveFile -DefaultFileName "SharedMailbox_Results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

if (-not [string]::IsNullOrWhiteSpace($resultsPath)) {
    $results | Select-Object DisplayName, PrimaryEmail, MailboxCreated, AliasesAdded, DelegatesAdded, @{Name='Errors';Expression={$_.Errors -join '; '}} | Export-Csv -Path $resultsPath -NoTypeInformation
    Write-ColorOutput "Results exported to: $resultsPath" "Success"
} else {
    Write-ColorOutput "Results not saved - no path selected" "Warning"
}

Write-Host "`nPress any key to exit..." -ForegroundColor Cyan
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
