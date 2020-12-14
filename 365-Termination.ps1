#Requires -Module ExchangeOnlineManagement
#Requires -Module AzureAD
#Requires -Module MSOnline
#Requires -Module Microsoft.Online.SharePoint.PowerShell
<#Synopsis.
User Termination for Microsoft 365
Sign out from all office apps, 
reset password, Remove groups, Remove from GAL
set mail forwarding, set mailbox permissions, set out of office message, grant onedrive permissions. 
.Parameter UPN
Specify the UserPrincipalName for the user being terminated
.Parameter Manager
Specify the terminated users manager by UPN
.Parameter Forwarding
Specify this switch to enable forwarding to the manager
.Parameter FullAccess
Will enable fullaccess to the users mailbox. Default is the manager. Can accept multiple strings. 
Written By David Just
#>
function Terminate-365User {
[cmdletbinding()]
Param (
    [Parameter(Mandatory,ValueFromPipeline=$true)][String]$UPN,
    [Parameter(Mandatory)][String]$Manager,
    [switch]$Forwarding,    
    [string[]]$fullaccess=$Manager
    )
$ErrorActionPreference="Stop"
$company = read-host "Enter Company Name"
$DesktopPath = [Environment]::GetFolderPath("Desktop")
Start-Transcript $DesktopPath\$upn-termination.log 
$user = $upn
Connect-MsolService
Connect-ExchangeOnline
Connect-AzureAD
#Get Tenant Name for SPO Url
$skuid = (Get-MsolAccountSku).AccountSkuID 
$spname = $skuid.Split(":")[0]
$spadminurl = "https://" + $spname + "-admin.sharepoint.com"
    Try{
    Write-Host "Trying to connect to sharepoint at $spadminurl"
    Start-Sleep 2
        Connect-SpoService -url $spadminurl
    }
    Catch {
        $_
    }

#If Deleted, Restore Deleted Account and Set ImmutibleID to Null
if ((Get-MsolUser -ReturnDeletedUsers -UserPrincipalName $upn).count -gt 0)
    {
        restore-msoluser -UserPrincipalName $upn
        Set-MsolUser -UserPrincipalName $upn -ImmutableId "$Null"
#If Account is restored without license, add license back        
        if ((Get-MsolUser -UserPrincipalName $upn).licenses.count -eq 0){
            Write-Host "Account was restored without license"
            Write-Host "Available Licenses:"
            Get-MsolAccountSku | Select-Object AccountSkuID,@{n="Available Count";e={$_.activeunits-$_.consumedunits}}
            $licenses = (Get-MsolAccountSku).AccountSkuID
            #If tenant only has 1 license sku, add that license   
            if ($licenses.count -eq 1)
                    {
                        Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses $licenses
                    }
            #If multiple licenses skus, prompt for which sku to add
                    else {   
                        Write-Host "License Count"
                        Get-MsolAccountSku | select AccountSkuID,@{n="available licenses";e={$_.ActiveUnits - $_.ConsumedUnits}}
            do { 
               
                   
           
                   $index = 1
                   foreach ($obj in $licenses) {
               
                       Write-Host "[$index] $obj"
                       $index++
           
                   }
               
                   $Selection = Read-Host "Please Select a license to add to restored account by number"
           
               } until ($licenses[$selection-1])
            Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses $licenses[$selection-1]
            }
        
        }
            do{
                write-host Waiting 30 Seconds for $upn mailbox to be restored
                Start-Sleep 30
                    }
                until(Get-Mailbox $upn -ea “silentlycontinue”)
    }
#Reset Password and sign out from 365 services
$managerobj = get-user $manager
$password = -join ((48..57) + (65..90) + (97..122) + (33,35,36,38,42) | Get-Random -Count 16 | foreach {[char]$_})
Get-AzureADUser -ObjectID $user | Revoke-AzureADUserAllRefreshToken 
Set-MsolUserPassword -UserPrincipalName $user -NewPassword $password -ForceChangePassword $false | Out-Null
Write-Output "New Password: $password"

#Block Sign In
Set-AzureADUser -ObjectId $user -AccountEnabled $False

#ListingGroups
Write-Output "$user Group Membership:"
<#
foreach ($Group in (Get-MsolGroup -all)) {if (Get-MsolGroupMember -all -GroupObjectId $Group.ObjectId | where {$_.EmailAddress -eq "$user"}) {$Group.Displayname}}
 $msolgroups = foreach ($Group in (Get-MsolGroup -all)) {if (Get-MsolGroupMember -all -GroupObjectId $Group.ObjectId | where {$_.EmailAddress -eq "$user"}) {$Group}}
 $userid = (Get-MsolUser -UserPrincipalName $user).objectid
    foreach ($group in $msolgroups){
    Write-Output "Removing $user from $($group.displayname)"
    Remove-MsoLGroupMember -GroupObjectId $Group.ObjectID.GUID -GroupMemberType User -GroupmemberObjectId $UserId.Guid -EA SilentlyContinue
    }
#>

#Get and List Group Membership
$AzGroups = foreach ($Group in (Get-AzureADGroup)) {if (Get-AzureADGroupMember -ObjectId $Group.ObjectId | where {$_.UserPrincipalName -eq "$user"}) {$Group}}
$DistroGroups = foreach ($Group in (Get-DistributionGroup)) {if (Get-DistributionGroupMember -identity $Group.name | where {$_.PrimarySmtpAddress -eq "$user"}) {$Group}}
     $GroupList = [PSCustomObject]@{
                                 User = $upn
                                 M365Groups = $AzGroups.Displayname
                                 DistroGroups = $DistroGroups.Displayname
                                                                            }
$GroupList

foreach ($group in $AzGroups){
    Write-Output "Removing $user from $($group.displayname)"
    Remove-AzureADGroupMember -ObjectId $Group.ObjectID -MemberID $Azuser.ObjectID 
    }

    foreach ($group in $DistroGroups){
        Write-Output "Removing $user from $($group.displayname)"
        Remove-DistributionGroupMember -identity $group.name -memberID $user -confirm:$false
        }

#Hide from GAL
Set-Mailbox -Identity $user -HiddenFromAddressListsEnabled $true

#Set out of office message message
$date = Get-Date -UFormat "%A %B %d %Y"
Set-MailboxAutoReplyConfiguration -identity $user -AutoReplyState enabled -InternalMessage "As of $date I am no longer with $company. Please send all communications to $Manager. Thank you" -ExternalMessage "As of $date I am no longer with $company. Please send all communications to $Manager. Thank you"

#Apply Forwarding
if ($forwarding){
Set-Mailbox $user -ForwardingAddress $manager -DeliverToMailboxandForward $true
}

#Add full access to manager
if ($fullaccess -ne $null){
    foreach ($object in $fullaccess){
    Add-MailboxPermission -identity $user -User $object -accessrights FullAccess}
    }

$upn -match "([a-z]+)+([a-z]+)"
$username = $matches[0]
$odsite = "https://" + $spname + "-my.sharepoint.com/personal/" + $username + "_" +$spname +"_com"

#Grant Manager OnedriveAccess
Set-SPOUser -site $ODsite -LoginName $manager -IsSiteCollectionAdmin $true

Write-Host "Users OneDrive URL : $ODSite"
Pause
Stop-Transcript
}