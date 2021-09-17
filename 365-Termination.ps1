<#
.Synopsis
	
Connect to the various Microsoft 365 services upon a user's termination to break / remove access and provide any needed access to another user for exchange, SharePoint ect. 

.Parameter UPN

Specify the UserPrincipalName for the user being terminated

.Parameter Forwarding

Specify this switch to enable forwarding to the manager

.Parameter FullAccess

Will enable fullaccess to the users mailbox. Default is the manager. Can accept multiple strings. 

.Parameter EnableOOFMessage

Will enable Out of Office message to specified users

.PARAMETER CustomOOFMessage

Will prompt for a custom out of office message. Do not use with EnableOOFMessage

.PARAMETER EnableLitHold

Enables Litigation Hold on the mailbox

.Parameter OneDriveAccess

Will grant specified account full access to terminated users onedrive. Specify by UPN

.Parameter CustomOOFMesage

Creates a custom auto reply message.

.PARAMETER ConvertToShared

Will convert the mailbox to a shared mailbox

.NOTES
09/06/2021 - Updated AADGroup region. Replaced previous method of retreiving group membership by checking the membership of every group in AAD for the user. Replaced with 
a simple call to MSGraph. Added functions Get-AuthToken (take from IntuneGraph samples on Github) and created new function Get-GraphUserPrincipalGroupMembership
#>

function m365-termination
{
#Requires -Module ExchangeOnlineManagement
#Requires -Module AzureADPreview
#Requires -Module MSOnline
#Requires -Module Microsoft.Online.SharePoint.PowerShell

	
	[cmdletbinding()]
	Param (
		[Parameter(Mandatory, ValueFromPipeline = $true)]
		[String]$UPN,
		[string]$EnableOOFMessage,
		[string]$Forwarding,
		[string[]]$Fullaccess,
		[string[]]$OneDriveAccess,
		[switch]$CustomOOFMessage,
		[switch]$EnableLitHold,
		[switch]$ConvertToShared
	)
	if ($upn -notlike "*@*")
	{
		write-Host "Please use the full userprincipal name (i.e. email address)." -ForegroundColor Red
		Break
	}
	$ErrorActionPreference = "Continue"
	if ($EnableOOFMessage) { $company = read-host "Enter Company Name" }
	$DesktopPath = [Environment]::GetFolderPath("Desktop")
	$path = $DesktopPath + '\' + $upn + '-termination' + '-' + ((get-date).ToFileTime()) + '.log'
	Start-Transcript $path
	$user = $upn
	Write-Host "Connecting to MSGraph" -ForegroundColor Yellow
	$authToken = Get-AuthToken
	$AadAccessToken = $authtoken.Authorization -replace "Bearer ", ''
	Write-Host "Connecting to AzureAD" -foregroundcolor Green
	Connect-AzureAD 
	Write-Host "Connecting to MSOL" -ForegroundColor Yellow
	Connect-MsolService
	Write-Host "Connecting to Exchange Online" -ForegroundColor Yellow
	Connect-ExchangeOnline -ShowBanner:$false
		#Get Tenant Name for SPO Url
	$skuid = (Get-MsolAccountSku).AccountSkuID
	$spname = $skuid.Split(":")[0]
	$spadminurl = "https://" + $spname + "-admin.sharepoint.com"
	$UserDeleted = Get-MsolUser -ReturnDeletedUsers | where UserPrincipalName -eq $UPN
	$EnableOOFMessageDN = Get-User $EnableOOFMessage -Erroraction "SilentlyContinue" | select -ExpandProperty displayname
	
	
	#If Deleted, Restore Deleted Account and Set ImmutibleID to Null
	if ($UserDeleted.count -gt 0)
	{
		restore-msoluser -UserPrincipalName $upn 
		try
		{
			Set-MsolUser -UserPrincipalName $upn -ImmutableId "$Null"
		}
		Catch
		{
			Write-Host $_ -ForegroundColor Red
		}
		#If Account is restored without license, add license back        
		if ((Get-MsolUser -UserPrincipalName $upn).licenses.count -eq 0)
		{
			Write-Host "Account was restored without license"
			Write-Host "Available Licenses:"
			Get-MsolAccountSku | Select-Object AccountSkuID, @{ n = "Available Count"; e = { $_.activeunits - $_.consumedunits } }
			$licenses = (Get-MsolAccountSku).AccountSkuID
			#If tenant only has 1 license sku, add that license   
			if ($licenses.count -eq 1)
			{
				Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses $licenses
			}
			#If multiple licenses skus, prompt for which sku to add
			else
			{
				Write-Host "License Count"
				Get-MsolAccountSku | select AccountSkuID, @{ n = "available licenses"; e = { $_.ActiveUnits - $_.ConsumedUnits } }
				do
				{
					
					
					
					$index = 1
					foreach ($obj in $licenses)
					{
						
						Write-Host "[$index] $obj"
						$index++
						
					}
					
					$Selection = Read-Host "Please Select a license to add to restored account by number"
					
				}
				until ($licenses[$selection - 1])
				Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses $licenses[$selection-1]
			}
			
		}
		do
		{
			write-host "Waiting 60 Seconds for $upn mailbox to be restored. This may take awhile. Leave this window open" -ForegroundColor Yellow
			Start-Sleep 60
		}
		until (Get-Mailbox $upn -ea "silentlycontinue")
	}
	#Reset Password and sign out from 365 services
	try
	{
		$AZUser = Get-AzureADUser -ObjectId $user -ErrorAction 'Stop'
	}
	Catch
	{
		Write-Host "UNABLE TO LOCATE $($user). Check the UPN is correct. Check if the user was deleted and not restored" -ForegroundColor Red
		Break
	}
	$password = -join ((48 .. 57) + (65 .. 90) + (97 .. 122) + (33, 35, 36, 38, 42) | Get-Random -Count 16 | foreach { [char]$_ })
	Write-Host "Signing $upn out of all 365 services"
	Get-AzureADUser -ObjectId $user | Revoke-AzureADUserAllRefreshToken
	sleep 5
	Set-AzureADUserPassword -ObjectId $($AZUser.ObjectID) -password ($password | ConvertTo-SecureString -AsPlainText -Force) -ForceChangePasswordNextLogin $false
	Write-Output "New Password: $password"
	
	#Block Sign In
	Set-AzureADUser -ObjectId $user -AccountEnabled $False
	
	#Region ListingGroups
	Write-Output "$user Group Membership:"
		
	#Get and List Group Membership
	
	#region Distro groups
	$i = 1
	$DistributionGroups = Get-DistributionGroup -ResultSize unlimited | where name -NotLike "Iland Backup Users"
	$DistroGroups = foreach ($Group in $DistributionGroups)
	{
		Write-progress -activity "Processing" -Status "Checking group membership on $($group.displayname)" -PercentComplete (($i / $DistributionGroups.count) * 100)
		$i++
		if (Get-DistributionGroupMember -identity $Group.name | where { $_.PrimarySmtpAddress -eq "$user" }) { $Group }
	}
	
	Write-Output "`nDistro Groups `n `r-------------`r"
	$DistroGroups | select name, displayname, primarysmtpaddress
	
	#Remove Group Memberships
	
	foreach ($group in $DistroGroups)
	{
		Write-Output "Removing $user from $($group.displayname)"
		try
		{
			Remove-DistributionGroupMember -identity $group.name -member $user
		}
		catch
		{
			Write-Host $_ -ForegroundColor Red
		}
		
	}
	#endregion 
	
	#region AAD Groups
	# Get group memberhsip from MS Graph
	$AADGroups = Get-GraphUserPrincipalGroupMembership -UserID $AZUser.ObjectID
	# Compare AAD groups to Distribution Groups, since AAD will list exchange managed distribution groups. Exclude any groups that are distribution groups. 
	$Diff = Compare-Object -ReferenceObject $DistributionGroups.ExternalDirectoryObjectId -DifferenceObject $AADGroups.id | where SideIndicator -eq "=>" | select -ExpandProperty inputobject
	$i = 1
	# Get AAD groups that are not managed in AD
	# Checked each groups membership for user
	$AzGroups = foreach ($Group in $Diff)
	{
		Get-AzureADMSGroup -Id $Group
	}
	
	Write-Output "`nM365 Groups `n `r------------`r"
	$AzGroups | select id, displayname, @{ n = 'DirSynced'; e = { $_.OnPremisesSyncEnabled } }, @{ n = 'IsDynamic'; e = { if ($_.MembershipRule -ne $null) { $true } } } | FT -AutoSize
	
		foreach ($group in ($AzGroups | where { ($_.OnPremisesSyncEnabled -ne $true) -and ($_.MembershipRule -eq $null) }))
	{
		try
		{
			Write-Output "Removing $user from $($group.displayname)"
			Remove-AzureADGroupMember -ObjectId $Group.ID -MemberID $Azuser.ObjectID
		}
		Catch
		{
			Write-Host "Unable to remove user from group $($group.displayname)" -ForegroundColor Red
			Write-Host $Error[0]
		}
		
	}
	#endRegion 
	
	#region Mailbox Actions
	
	#Check if there is a mailbox. Skip mail related portion if no mailbox is found.     
	if (!(Get-mailbox $upn -ea 'SilentlyContinue')) { Write-Host "User does not have a mailbox, or mailbox has yet to be restored. Skipping mail related section." -ForegroundColor yellow }
	else
	{
		
		
#Hide from GAL
		try
		{
			Set-Mailbox -Identity $user -HiddenFromAddressListsEnabled $true
			Write-Host -ForegroundColor Magenta "User has been hidden from the GAL"
		}
		Catch
		{
			throw $_
			Write-Host -ForegroundColor Red "Unable to hide user from GAL"
		}
		
		
		#Try to set Litigation hold
		if ($EnableLitHold)
		{
			try
			{
				Set-Mailbox -Identity $user -LitigationHoldEnabled $true
			}
			Catch
			{
				
			}
		}
		
		#Set out of office message message
		if ($EnableOOFMessage)
		{
			try
			{
				$date = Get-Date -UFormat "%A %B %d %Y"
				Set-MailboxAutoReplyConfiguration -identity $user -AutoReplyState enabled -InternalMessage "As of $date, I am no longer with $company. Please send all communications to $EnableOOFMessageDN, $EnableOOFMessage. Thank you." -ExternalMessage "As of $date, I am no longer with $company. Please send all communications to $EnableOOFMessageDN, $EnableOOFMessage. Thank you."
			}
			Catch
			{
				Write-Host $_ -ForegroundColor Red
			}
		}
		
		if ($CustomOOFMessage)
		{
			try
			{
				Set-MailboxAutoReplyConfiguration -identity $user -AutoReplyState enabled -InternalMessage (read-host "Enter internal out of office Message") -ExternalMessage (Read-Host "Enter external out of office message")
			}
			catch
			{
				Write-Host $_ -ForegroundColor Red
			}
		}
		
		
#Apply Forwarding
		if ($forwarding)
		{
			Set-Mailbox $user -ForwardingAddress $forwarding -DeliverToMailboxandForward $true
		}
		
		#Add full access to specified accounts
		if ($fullaccess)
		{
			foreach ($object in $fullaccess)
			{
				try
				{
					Add-MailboxPermission -identity $user -User $object -accessrights FullAccess
				}
				Catch
				{
					Write-Host $_ -ForegroundColor Red
				}
			}
		}
#Convert to Shared Mailbox
		if ($ConvertToShared)
		{
			try
			{
				Set-Mailbox $user -type Shared
			}
			catch
			{
				Write-Host $_ -ForegroundColor Red
			}
		}
	}
	#endregion
	#$upn -match "([a-z]+)+([a-z]+)" use mailaddress type accelerator rather than regex. Regex here for reference
	$username = [mailaddress]$upn
	$odsite = "https://" + $spname + "-my.sharepoint.com/personal/" + $username.User + "_" + ($username.host.split('.')[0]) + "_" + ($username.host.split('.')[1])
	
#Region Grant OnedriveAccess
	if ($OneDriveAccess)
	{
		Try
		{
			Write-Host "Trying to connect to sharepoint at $spadminurl"
			Start-Sleep 2
			Connect-SpoService -url $spadminurl
		}
		Catch
		{
			Write-Host $_.Exception -ForegroundColor Red
		}
		
		try
		{
			$OnedriveAccess | ForEach-Object { Set-SPOUser -site $ODsite -LoginName $_ -IsSiteCollectionAdmin $true }
		}
		Catch
		{
			Write-Host $_.exception -ForegroundColor Red
		}
	}
	Write-Host "Users OneDrive URL : $ODSite"
	#endregion
	Disconnect-AzureAD -Confirm:$false
	Disconnect-ExchangeOnline -Confirm:$false
	Pause
	Stop-Transcript
}

function Get-AuthToken
{
	
<#
.SYNOPSIS
This function is used to authenticate with the Graph API REST interface
.DESCRIPTION
The function authenticate with the Graph API Interface with the tenant name
.EXAMPLE
Get-AuthToken
Authenticates you with the Graph API interface
.NOTES
NAME: Get-AuthToken
#>
	
	[cmdletbinding()]
	param
	(
		[String]$User
	)
	$ErrorActionPreference = 'Stop'
	$User = Read-Host "Enter the Admin UserPrincipalName to connect to MSGraph"
	$global:Admin = $User
	try
	{
		$emailaddress = [mailaddress]$user
	}
	catch
	{
		Write-Host $User "is not a valid UPN" -ForegroundColor Red
		Break
	}
	$userUpn = New-Object "System.Net.Mail.MailAddress" -ArgumentList $User
	
	$tenant = $userUpn.Host
	
	Write-Host "Checking for AzureAD module..."
	
	$AadModule = Get-Module -Name "AzureAD" -ListAvailable
	
	if ($AadModule -eq $null)
	{
		
		Write-Host "AzureAD PowerShell module not found, looking for AzureADPreview"
		$AadModule = Get-Module -Name "AzureADPreview" -ListAvailable
		
	}
	
	if ($AadModule -eq $null)
	{
		write-host
		write-host "AzureAD Powershell module not installed..." -f Red
		write-host "Install 'Install-Module AzureADPreview' from an elevated PowerShell prompt" -f Yellow
		write-host "Script can't continue..." -f Red
		write-host
		exit
	}
	
	# Getting path to ActiveDirectory Assemblies
	# If the module count is greater than 1 find the latest version
	
	if ($AadModule.count -gt 1)
	{
		
		$Latest_Version = ($AadModule | select version | Sort-Object)[-1]
		
		$aadModule = $AadModule | ? { $_.version -eq $Latest_Version.version }
		
		# Checking if there are multiple versions of the same module found
		
		if ($AadModule.count -gt 1)
		{
			
			$aadModule = $AadModule | select -Unique
			
		}
		
		$adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
		$adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
		
	}
	
	else
	{
		
		$adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
		$adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
		
	}
	
	[System.Reflection.Assembly]::LoadFrom($adal) | Out-Null
	
	[System.Reflection.Assembly]::LoadFrom($adalforms) | Out-Null
	
	$clientId = "d1ddf0e4-d672-4dae-b554-9d5bdfd93547"
	
	$redirectUri = "urn:ietf:wg:oauth:2.0:oob"
	
	$resourceAppIdURI = "https://graph.microsoft.com"
	
	$authority = "https://login.microsoftonline.com/$Tenant"
	
	try
	{
		
		$authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
		
		# https://msdn.microsoft.com/en-us/library/azure/microsoft.identitymodel.clients.activedirectory.promptbehavior.aspx
		# Change the prompt behaviour to force credentials each time: Auto, Always, Never, RefreshSession
		
		$platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters" -ArgumentList "Auto"
		
		$userId = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifier" -ArgumentList ($User, "OptionalDisplayableId")
		
		$authResult = $authContext.AcquireTokenAsync($resourceAppIdURI, $clientId, $redirectUri, $platformParameters, $userId).Result
		
		# If the accesstoken is valid then create the authentication header
		
		if ($authResult.AccessToken)
		{
			
			# Creating header for Authorization token
			
			$authHeader = @{
				'Content-Type'  = 'application/json'
				'Authorization' = "Bearer " + $authResult.AccessToken
				'ExpiresOn'	    = $authResult.ExpiresOn
			}
			
			return $authHeader
									
		}
		
		else
		{
			
			Write-Host
			Write-Host "Authorization Access Token is null, please re-run authentication..." -ForegroundColor Red
			Write-Host
			break
			
		}
		
	}
	
	catch
	{
		
		write-host $_.Exception.Message -f Red
		write-host $_.Exception.ItemName -f Red
		write-host
		break
		
	}
	
}

function Get-GraphUserPrincipalGroupMembership {
	<#
.SYNOPSIS
This function is used to return all groups a user is a member of
.EXAMPLE
Get-GraphUserPrincipalGroupMembership -UserID (Get-AzureADUser -searchstring user@contoso.com).objectid
.NOTES
Created by David Just 09/06/2021
#>
	param (
		[string]$UserID
		
	)
	$graphApiVersion = "v1.0"
	$Resource = "users/$UserID/memberOf/$/microsoft.graph.group?$select=id,displayName,securityEnabled,groupTypes,onPremisesSyncEnabled"
	$uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)"
	
	$groups = (Invoke-RestMethod -Uri $uri -Headers $authToken -Method Get).Value 
	return $groups
}
