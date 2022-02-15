<#
.Synopsis
	
Fetch group membership of a reference user. Compare to current group membership of a target user. Add any missing groups to the target user. 

.Parameter ReferenceUser

The reference user to base group membership comparison on. Enter the full userprincipal name (name@domain.com)

.PARAMETER TargetUser

The user to modify / compare group membership to. Enter the full userprincipal name (name@domain.com)

.PARAMETER ComparisonOnly

Switch parameter. Using this switch will only display a comparison of group memberships between the reference user and target user and will not modify group memberships. 

.NOTES
Usage: Download script. Open Windows PowerShell, path to download location ("cd downloads" for example). Then run script

.EXAMPLE
PS >.\MirrorGroupMemberships.ps1 -ReferenceUser userA@domain.com -TargetUser userB@domain.com

.EXAMPLE
PS >.\MirrorGroupMemberships.ps1 -ReferenceUser userA@domain.com -TargetUser userB@domain.com -ComparisonOnly
#>

#Requires -Module ExchangeOnlineManagement

param (
	[Parameter(Mandatory = $true)]
	[string]$TargetUser,
	[Parameter(Mandatory = $true)]
	[string]$ReferenceUser,
	[switch]$ComparisonOnly
)

$tempfile = New-TemporaryFile
Start-Transcript -Path $tempfile
#region HelperFunctions
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

function Get-GraphUserPrincipalGroupMembership
{
	<#
.SYNOPSIS
This function is used to return all groups a user is a member of
.EXAMPLE
Get-GraphUserPrincipalGroupMembership -UserID (Get-AzureADUser -searchstring user@contoso.com).objectid
.NOTES
Created by David Just 09/06/2021
#>
	[cmdletbinding()]
	param (
		[string]$UserID
		
	)
	$graphApiVersion = "v1.0"
	$Resource = "users/$UserID/memberOf/$/microsoft.graph.group?$select=id,displayName,securityEnabled,groupTypes,onPremisesSyncEnabled"
	$uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)"
	
	$groups = (Invoke-RestMethod -Uri $uri -Headers $authToken -Method Get).Value
	return $groups
}

function Add-Graph365GroupMember
{
	[cmdletbinding()]
	param
	(
		[String]$Identity,
		[String]$UserID
	)
	$graphApiVersion = "v1.0"
	$Resource = "groups/$($Identity)/members/`$ref"
	$uri = "https://graph.microsoft.com/v1.0/$($Resource)"
	
	$groupAdd = Invoke-RestMethod -Uri $uri -Headers $authToken -Body (@{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($UserID)" } | ConvertTo-Json) -ContentType "Application/JSON" -Method Post
	return $groupAdd
}

function Get-GraphUser
{
	[cmdletbinding()]
	param
	(
		[String]$User
	)
	$uri = "https://graph.microsoft.com/v1.0/users/$($User)"
	
	$GraphUser = Invoke-RestMethod -Uri $uri -Headers $authToken -Method GET
	return $GraphUser
	
}
#endregion

$authToken = Get-AuthToken

try
{
	$TargetUserID = Get-GraphUser -User $TargetUser -ea 'Stop'
}
Catch
{
	Throw "Unable to find $($TargetUser)!"
	break
}

$ReferenceUserGroups = Get-GraphUserPrincipalGroupMembership -UserID $ReferenceUser
$refOnPremGroups = $ReferenceUserGroups | where onPremisesSyncEnabled
$TargetUserGroups = Get-GraphUserPrincipalGroupMembership -UserID $TargetUser

# Warn if reference user is a member of synced on prem groups
if ($refOnPremGroups)
{
	Write-Warning "The reference user is a member of the following AD synced groups`n which will need to be added in AD:`n($refOnPremGroups | select -ExpandProperty displayname)"
	Write-Host "Continuing in 5 seconds"
	sleep 5
}

# Parse out reference and target group memberships separating distribution groups from Unified groups
$refDistros = $ReferenceUserGroups | where { -not $_.grouptypes -and $_.mailenabled } | select id, displayname, mail
$ref365 = $ReferenceUserGroups | where {-not $_.membershiprule} | where { $_.grouptypes -like "*Unified*" -or ($_.securityenabled -and -not $_.mailenabled -and -not $_.onPremisesSyncEnabled)} | select displayname,id

$targetDistros = $TargetUserGroups | where { -not $_.grouptypes -and $_.mailenabled } | select id, displayname, mail
$target365 = $TargetUserGroups | where { -not $_.membershiprule }  |  where { $_.grouptypes -like "*Unified*" -or ($_.securityenabled -and -not $_.mailenabled -and -not $_.onPremisesSyncEnabled) } | select displayname,id

# Compare current state of memberships and distill list down to whats needed
$neededDistroGroups = if ($targetDistros) { Compare-Object -ReferenceObject $refDistros -DifferenceObject $targetDistros -Property displayname, id | where SideIndicator -eq "<=" | select displayname, id }; else { $refDistros }
$needed365groups = if ($target365) { Compare-Object -ReferenceObject $ref365 -DifferenceObject $target365 -Property displayname, id | where SideIndicator -eq "<=" | select displayname, id }; else { $ref365 }


Write-Output "Needed Groups: 365 Groups"
$needed365groups
Write-Output "`nNeeded Groups: Distro Groups"
$neededDistrogroups

if ($ComparisonOnly)
{
	Stop-Transcript
	notepad $tempfile
	break
}


#region AddGroupMemberships

$365Success = @()
$365Error = @()

foreach ($Group in $needed365groups)
{
	Write-Host "Adding $($TargetUser) to $($Group.displayName)" -ForegroundColor Green -BackgroundColor Black
	try
	{
		Add-Graph365GroupMember -identity $Group.id -userID $TargetUserID.id -EA 'Stop'
		$365Success += $group
		Start-Sleep -Milliseconds 400
	}
	Catch
	{
		$365Error += $Group.displayname
		$Error[0]
	}
}

$i = 5
while ($i -ne 0)
{
	Write-Host "Connecting to ExchangeOnline in $($i)" -ForegroundColor Green -BackgroundColor Black
	$i = $i - 1
	sleep 1
}

Connect-ExchangeOnline -showbanner:$false
$distributionGroups = Get-DistributionGroup -resultsize unlimited
$distroError = @()
$distroSuccess = @()

foreach ($distroGroup in $neededDistrogroups)
{
	Write-Host "Adding $($TargetUser) to $($distroGroup.displayname)" -ForegroundColor Green -BackgroundColor Black
	try
	{
		Add-DistributionGroupMember -identity ($distributionGroups | where externaldirectoryobjectid -EQ $distroGroup.id | select -ExpandProperty guid | select -ExpandProperty guid) -member $TargetUser -EA 'Stop'
		$distroSuccess += $distroGroup.displayname
		Start-Sleep -Milliseconds 400
	}
	Catch
	{
		$distroError += $distroGroup.displayname
	}
}

#endregion

#region VerboseOutput

if ($365Success)
{
	Write-Host "Successfully added " $TargetUser "to aad groups:" -ForegroundColor Green -BackgroundColor Black
	$365Success
}

if ($distroSuccess)
{
	Write-Host "`n Successfully added " $TargetUser "to distribution groups:" -ForegroundColor Green -BackgroundColor Black
}


$distroSuccess

if ($365Error)
{
	$365Error | foreach {Write-Output "Unable to add " $TargetUser "to: " $_ }
}

if ($distroError)
{
	$distroError | foreach { Write-Output "Unable to add " $TargetUser "to: " $_ }
}

$finalGroupMembership = Get-GraphUserPrincipalGroupMembership -UserID $TargetUser

Write-Host "$($TargetUser) Group Membership:" -ForegroundColor Green -BackgroundColor Black
$finalGroupMembership | select displayname, grouptypes

$missingDistros = $finalGroupMembership | where { -not $_.grouptypes -and $_.mailenabled } | select id, displayname, mail
$missing365 = $finalGroupMembership | where { -not $_.membershiprule } | where { $_.grouptypes -like "*Unified*" -or ($_.securityenabled -and -not $_.mailenabled -and -not $_.onPremisesSyncEnabled) } | select displayname, id

$MissingDistroGroups = if ($missingDistros) { Compare-Object -ReferenceObject $refDistros -DifferenceObject $missingDistros -Property displayname,id | where SideIndicator -eq "<=" | select -ExpandProperty displayname}
$Missing365groups = if ($target365) { Compare-Object -ReferenceObject $ref365 -DifferenceObject $missing365 -Property displayname,id | where SideIndicator -eq "<=" | select -ExpandProperty displayname }

if ($MissingDistroGroups)
{
	Write-Warning "The following groups were not added:"
	$MissingDistroGroups
}

if ($Missing365groups)
{
	Write-Warning "The follwoing AADGroups were not added:"
	$Missing365groups
}

#endregion

Stop-Transcript

Write-Warning "Confirm group memberships are correct! The coder of this script is human and humans are inperfect beings. Do not assume this script is foolproof!"

Write-Host "Opening log file in 3 seconds"
Start-Sleep 3
notepad $tempfile




