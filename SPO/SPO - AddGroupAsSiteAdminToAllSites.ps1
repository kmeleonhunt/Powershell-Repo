<#

╦╔═┌┬┐┌─┐┬  ┌─┐┌─┐┌┐┌┬ ┬┬ ┬┌┐┌┌┬┐
╠╩╗│││├┤ │  ├┤ │ ││││├─┤│ ││││ │ 
╩ ╩┴ ┴└─┘┴─┘└─┘└─┘┘└┘┴ ┴└─┘┘└┘ ┴ 
 
    .SYNOPSIS
       Add a O365 Security Group to all Sites as Site Admin Role

    .DESCRIPTION
     

    .REQUIREMENTS
        Permission string Object > Manually get it in a Site where the Group is already added as site admin.
        Site Permission > Advanced > Check Permission > input the email of the group : get the string that look like :

        c:0-.f|rolemanager|s-1-1-11-111111111-1111111111-1111111111-11111111

#>


# Connect to SharePoint Online
$url = Read-Host "Input Admin URL"

Connect-SPOService -Url $url

$Group = Read-Host "Input the string like : 'c:0o.c|federateddirectoryclaimprovider|882364d7e-c206-4e70-adbb-bd94eba851ad'"

# Get all SharePoint sites
$sites = Get-SPOSite -Limit All

# Iterate through each site
foreach ($site in $sites) {
    Write-Host "Processing site $($site.Url)"

    # Get site collection administrators for the current site
    Set-SPOUser -Site $site.Url -LoginName $Group -IsSiteCollectionAdmin $true

}
