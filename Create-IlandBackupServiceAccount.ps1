 <#
    .SYNOPSIS
      Create the iland/Veeam service account for use with iland/Veeam O365 Backups for Exchange Sharepoint and Onedrive
      Create the Role Group for Veeam backups with the required roles and assign to the service account.
      Written By David Just
    
 #>
$exchangemodule = Get-InstalledModule -Name "ExchangeOnlineManagement"
$msolmodule = Get-InstalledModule -Name "MSOnline"

if ($exchangemodule -eq $null)
    {Write-Warning "Exhange Online Management Module is not installed. Installing Now"
        Install-Module -Name ExchangeOnlineManagement
        Connect-ExchangeOnline}
else {Connect-ExchangeOnline}

if ($msolmodule -eq $null)
    {Write-Warning "Microsoft Online Management Module is not installed. Installing Module"
        Install-Module -Name MSOnline
        Connect-MsolService}
else {Connect-MsolService}

$domain = Get-MsolDomain | where {$_.IsDefault -eq $true}
$Password = -join ((48..57) + (65..90) + (97..122) + (33,35,36,38,42) | Get-Random -Count 16 | foreach {[char]$_})
$secPwd = ConvertTo-SecureString -String $password -AsPlainText -Force
$upn='svc_ilandbackup@'+$domain.name
$msolaccount = Get-MSoluser -UserPrincipalName $upn -ErrorAction SilentlyContinue

if($msolaccount.count -eq 0){
                Write-Output "Creating account $upn"
                New-MSoluser -userprincipalname $upn -Displayname "SVC IlandBackup" -Firstname "SVC" -Lastname "IlandBackup" -Password $secPwd -PasswordNeverExpires $true -ForceChangePassword $false
                }
                else{write-output "Account Already Created, skipping creation"}

Try {
Get-RoleGroup "iland_backups" -erroraction STOP
    }Catch{New-Rolegroup -Name "iland_backups" -roles "ApplicationImpersonation","Role Management","Organization Configuration","View-Only Configuration","View-Only Recipients","Mailbox Search","Mail Recipients"
          }
New-AuthenticationPolicy -Name "Iland Backups" -AllowBasicAuthPowershell -AllowBasicAuthWebServices -EA SilentlyContinue

write-host "Waiting 60 seconds for account to sync to exchange directory"
Sleep 60
try {Get-User $upn -EA Stop | Out-Null
}catch{Write-Host "Account is not synced yet waiting 60 seconds"
    Sleep 60}
Add-RolegroupMember "iland_backups" -Member "SVC IlandBackup" 

Add-MsolRoleMember -RoleName "Sharepoint Service Administrator" -RoleMemberEmailAddress $upn

Set-User $upn -Authenticationpolicy "Iland Backups"

Write-Output "SVC Account is $UPN;
Password is $password"

