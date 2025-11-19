<#
╦╔═┌┬┐┌─┐┬  ┌─┐┌─┐┌┐┌┬ ┬┬ ┬┌┐┌┌┬┐
╠╩╗│││├┤ │  ├┤ │ ││││├─┤│ ││││ │ 
╩ ╩┴ ┴└─┘┴─┘└─┘└─┘┘└┘┴ ┴└─┘┘└┘ ┴ 
 
    .SYNOPSIS
       Export Mail FLow rules to XML file

    .DESCRIPTION
    This script exports the mail flow rules from Exchange Online to an XML file.

    .REQUIREMENTS
        

#>

<#
.FUNCTIONS
#>
function Export-TransportRuleXML{

Connect-ExchangeOnline

$Tenant = (Get-AcceptedDomain | Where-Object { $_.Default -eq $true } | Sort-Object -Property Name | Select-Object Name)

Write-Host "Informal : You are currently logged in $($Tenant) - Make sure you want to run the script on $($Tenant)" -ForegroundColor Yellow

function Save-File {
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.Filter = "CSV files (*.xml)|*.xml"
    $SaveFileDialog.FileName = "$(Get-Date -Format 'yyyyMMdd')_UserReport_$($client).xml"

    if ($SaveFileDialog.ShowDialog() -eq "OK") {
        $SaveFileDialog.FileName
    } else {
        "No path selected"
    }
}

$path = Save-File

$file = Export-TransportRuleCollection

[System.IO.File]::WriteAllBytes($path, $file.FileData)

Write-Host "O365 admin Transport Rules were exported to $($path)"

}

########
#SCRIPT#
########

Export-TransportRuleXML

