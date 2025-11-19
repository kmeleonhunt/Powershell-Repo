<#
╦╔═┌┬┐┌─┐┬  ┌─┐┌─┐┌┐┌┬ ┬┬ ┬┌┐┌┌┬┐
╠╩╗│││├┤ │  ├┤ │ ││││├─┤│ ││││ │ 
╩ ╩┴ ┴└─┘┴─┘└─┘└─┘┘└┘┴ ┴└─┘┘└┘ ┴ 

    .SYNOPSIS
       Import Mail FLow rules to Tenant

    .DESCRIPTION
    This script import the mailflow rules from an XML file to Exchange Online. After importing you will have to enable the rules manually and check the settings.

    .REQUIREMENTS
    You need the StandardRuleCollectionToImport.xml file in the C:\MailFlowRuleCollections folder.    

#>

<#
.FUNCTIONS
#>
Connect-ExchangeOnline

$directoryPath = "C:\MailFlowRuleCollections"

# Check if the directory exists
if (-not (Test-Path -Path $directoryPath)) {
    # If the directory does not exist, create it
    New-Item -Path $directoryPath -ItemType Directory
    Write-Output "Directory created at $directoryPath"
} else {
    # If the directory exists, output a message
    Write-Output "Directory already exists at $directoryPath"
}


[xml]$xml = Get-Content "C:\MailFlowRuleCollections\StandardRuleCollectionToImport.xml"

$rulesToImport = $xml.SelectNodes("//rules/rule")

if ($rulesToImport.Count -eq 0)

{
    Write-Host "There are no mail flow rules to be imported."

    return
}

Write-Host "Importing $($rulesToImport.Count) mail flow rules."

$index = 0

foreach ($rule in $rulesToImport)

{
    $index++

    Write-Host "Importing rule '$($rule.Name)' $index/$($rulesToImport.Count)."

    Invoke-Expression $($rule.version.commandBlock.InnerText) | Out-Null

}
