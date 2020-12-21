#Requires -Module MSOnline
#Requires -Module Microsoft.Online.SharePoint.PowerShell
<#Synopsis.
Grant User OneDrive Admin Access 
Connects to Sharepoint and grant access for a users personal OnedDrive to another specified user. 
.Parameter UPN
Specify the UserPrincipalName for the user whose onedrive you would like to grant admin access on
.Parameter Admin
.Example 
New-OneDriveAdmin -UPN user1@domain.com -Admin user2@domain.com
Specify the UserPrincipalName for the user you would like to grant admin access
Written By David Just
12/18/2020
#>
function New-OneDriveAdmin{
 [cmdletbinding()]
    Param (
        [Parameter(Mandatory,ValueFromPipeline=$true)][String]$UPN,
        [Parameter(Mandatory)][string[]]$Admin
        )
Connect-MsolService
$skuid = (Get-MsolAccountSku).AccountSkuID 
$spname = $skuid.Split(":")[0]
$spadminurl = "https://" + $spname + "-admin.sharepoint.com"
Try{
        Write-Host "Trying to connect to sharepoint at $spadminurl"
        Start-Sleep 2
            Connect-SpoService -url $spadminurl
        }
        Catch {
            Write-Host $_.Exception -ForegroundColor Red
        }
    $username = [mailaddress]$upn
    $odsite = "https://" + $spname + "-my.sharepoint.com/personal/" + $username.User + "_" +($username.host.split('.')[0]) +"_"+($username.host.split('.')[1])
    
    #Grant OnedriveAccess
    if ($Admin){
    try{
    $Admin | ForEach-Object {Set-SPOUser -site $ODsite -LoginName $_ -IsSiteCollectionAdmin $true}
    }
    Catch
    {
        Write-Host $_.exception -ForegroundColor Red
    }
}
    Write-Host "Users OneDrive URL : $ODSite"
    Pause
    Get-PSSession | Remove-PSSession
    }