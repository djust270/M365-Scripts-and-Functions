<#
.SYNOPSIS Use this with the new Exchange V2 cmdlets. This script is designed to list out the majority of permissions in exchange
a user has, as well as permissions on the users mailbox. These permissions are listed on screen and output to file
for documentation purposes. Permissions can also be removed with this script. 
.EXAMPLE .\Get-UserExchPermissions.ps1 -user user@domain.com
Written by David Just
#>
#Requires -Module ExchangeOnlineManagement

[CmdletBinding()]
param (
	[Parameter(Mandatory=$true)]
	[string]$user
)
$ErrorActionPreference = 'Stop'
write-host "This Script will list out user permissions and save each output to a text file on the desktop"
set-location -path ([System.Environment]::GetFolderPath('Desktop'))

$username = Get-User ([mailaddress]$user).address
$mailboxes = Get-EXOCasMailbox -ResultSize unlimited
$i = 1

Write-Output "listing mailbox permissions this user has" | Tee-Object ($user + '-perms.txt') -Append
try
{
	foreach ($mbx in $mailboxes)
	{
		Write-progress -activity "Processing" -Status "Checking mailbox permissions on $($mbx.identity)" -PercentComplete (($i / $mailboxes.count) * 100)
		$mbxperm = Get-ExoMailboxpermission -Identity $mbx.identity | where { $_.user -like $user }
		$mbxperm | Tee-Object ($user + '-perms.txt') -Append
		$mbxperm | select-object identity, user, @{ n = "accessrights"; e = { ($_).accessrights } } | Export-Csv ($user + '-mbxperms.csv') -Append
		$i++
	}
}
Catch
{
	Write-Host $_.Exception -ForegroundColor Red
}

Write-Output "listing calendar permissions this user has" | Tee-Object ($user + '-perms.txt') -Append

$i = 1
try{
ForEach ($mbx in $mailboxes)
	{
		Write-progress -activity "Processing" -Status "Checking calendar permissions on $($mbx.identity)" -PercentComplete (($i / $mailboxes.count) * 100)
		$calendarperm = Get-MailboxFolderPermission (($mbx.PrimarySmtpAddress.ToString()) + ":\Calendar") | where { $_.User -like $username.name }  
		$calendarperm | select identity, user, accessrights | FL | Tee-Object ($user + '-perms.txt') -Append
		$calendarperm | Export-Csv ($user + '-calperms.csv') -Append
		#Get-MailboxFolderPermission (($mbx.PrimarySmtpAddress.ToString()) + ":\Calendar") -User $user -EA SilentlyContinue | select Identity, User, AccessRights | Tee-Object ($user + '-perms.txt') -Append
		$i++
	}
}
catch
{
	Write-Host $_.Exception -ForegroundColor Red
}

Write-Output "listing mailboxes user has send-as permissions on" | Tee-Object ($user + '-perms.txt') -Append

$i = 1
try
{
	foreach ($mbx in $mailboxes)
	{
		Write-progress -activity "Processing" -Status "Checking send as permissions on $($mbx.identity)" -PercentComplete (($i / $mailboxes.count) * 100)
		$sendas = Get-RecipientPermission -Identity $mbx.identity | where { ($_.trustee -like $user) -and ($_.trustee -notlike "*AUTHORITY*") }
		$sendas | select identity, @{ n = "Send As User"; e = { $_.trustee } } | Tee-Object ($user + '-perms.txt') -Append
		$sendas | Export-Csv ($user + '-sendasperms.csv') -Append
		$i++
		
	}
}
Catch
{
	Write-Host $_.Exception -ForegroundColor Red
}
Write-Output "listing contact folders user has permission on" | Tee-Object ($user + '-perms.txt') -Append
$i = 1
try
{
	foreach ($mbx in $mailboxes)
	{
		Write-progress -activity "Processing" -Status "Checking contact folder permissions on $($mbx.identity)" -PercentComplete (($i / $mailboxes.count) * 100)
		$contactperm = Get-MailboxFolderPermission (($mbx.PrimarySmtpAddress.ToString()) + ":\contacts") | where { $_.User -like $username.name }  
		$contactperm | select Identity, User, AccessRights | Tee-Object ($user + '-perms.txt') -Append
		$contactperm | export-csv ($user + '-contactperms.csv') -Append
		$i++
	}
}
catch
{
	Write-Host $_.Exception -ForegroundColor Red
}
Write-Output "listing users that have permission to this mailbox" | Tee-Object ($user + '-perms.txt') -Append

Get-EXOMailboxFolderPermission $user | select User,AccessRights | Tee-Object ($user + '-perms.txt') -Append
Write-Output "listing users that have permission to the terminated users Calendar" | Tee-Object ($user + '-perms.txt') -Append
Get-EXOMailboxFolderPermission ${user}:\calendar | Tee-Object ($user + '-perms.txt') -Append

Write-Output "listing users that have send as permission on this mailbox" | Tee-Object ($user + '-perms.txt') -Append
Get-EXORecipientPermission $user | where trustee -NotLike *AUTHORITY* |  Tee-Object ($user + '-perms.txt') -Append
notepad ($user + '-perms.txt')

Pause
$prompt = Read-Host "Would you like to remove the terminated users permissions? y/n"
switch ($prompt)
{
	Y {
		if (test-path ($user + '-mbxperms.csv')){
			$mbxperms = Import-Csv ($user + '-mbxperms.csv') -ea Continue
			$mbxpermscount = $mbxperms | measure
			if ($mbxpermscount.count -gt 0)
			{
				$mbxperms | foreach {
					Write-Host "Removing mailbox permission from $($_.identity)"
					remove-mailboxpermission -identity $_.identity -user $_.user -accessrights $_.accessrights -confirm:$false
				}
			}
		}
		
		if (Test-Path ($user + '-calperms.csv')){
			$calperms = Import-Csv ($user + '-calperms.csv') -ea Continue
			$calpermscount = $calperms | measure
			if ($calpermscount.count -gt 0)
			{
				$calperms | foreach {
					Write-Host "Removing Calendar permission from $($_.identity)"
					Remove-Mailboxfolderpermission -identity $_.identity -user $_.user -confirm:$false
				}
			}
		}
		if (Test-Path ($user + 'contactperms.csv')){
			$contactperms = Import-Csv ($user + '-contactperms.csv') -ea Continue
			$contactpermscount = $contactperms | measure
			if ($contactpermscount.count -gt 0)
			{
				$contactperms | foreach {
					Write-Host "Removing contact folder permission from $($_.identity)"
					Remove-Mailboxfolderpermission -identity $_.identity -user $_.user -confirm:$false
				}
			}
		}
		if (Test-Path ($user + '-sendasperms.csv')){
			$sendasperms = Import-Csv ($user + '-sendasperms.csv') -ea Continue
			$sendascount = $sendasperms | measure
			if ($sendascount.count -gt 0)
			{
				$sendasperms | foreach {
					Write-Host "Removing sendas permission from $($_.identity)"
					remove-recipientpermission -identity $_.identity -trustee $_.trustee -accessrights SendAs -confirm:$false
				}
			}
		}
		
	}
	N	{
		$csvs = ($user + '-sendasperms.csv'), ($user + '-calperms.csv'),($user + '-mbxperms.csv'), ($user + '-contactperms.csv')
		$csvs | foreach {Remove-Item $_ -Confirm:$false -ea Continue}
		Write-Host "Thank you for using the exchange permission tool. Have a nice day!"
		sleep 2 
		exit }
}



