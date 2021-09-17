<#
Synopsis - List out the permission levels and users for  
every non-empty sharepoint site export to excel spreadsheet
#>
Param (
[Parameter(Mandatory=$true)]
[string]$AdminUrl
)
$ErrorActionPreference = 'Stop'
Connect-AzureAD
Connect-SPOService -Url $AdminUrl
$sites = Get-SPOSite
$sites = $sites | where storageusagecurrent -gt '1'
$perms = [System.Collections.Generic.List[PsObject]]::new()
$i = 1
foreach ($site in $sites){
Write-Progress "Working on $($site.title)" -CurrentOperation "Working on site $i out of $($sites.count)"
$i++
$SiteGroups = Get-SPOSiteGroup -Site $site.url
ForEach($Group in $SiteGroups) {
     $perms.add([pscustomobject]@{
            'SiteName' = $site.title
            'Group Name' =$Group.Title
            'Permissions' =$Group.Roles -join ","
            'Users' = foreach ($user in $Group.users) {
            try {
            if([guid]$user){(Get-AzureADGroupMember -ObjectId $user).UserPrincipalName -join ','}
            }
            Catch {
            if ($user -match '(?im)^[{(]?[0-9A-F]{8}[-]?(?:[0-9A-F]{4}[-]?){3}[0-9A-F]{12}[_][]o[)}]?$'){
            Connect-PnPOnline -Url $site.url -PnPManagementShell
            (Get-PnPGroupMember -Group $Group.title).loginname -join ','
            }
            else {$User}
            }
            }})
        }
    
    }
$perms | select sitename,"Group Name",Permissions,@{n='Users';e={$_.users -join ','}} | export-excel 


