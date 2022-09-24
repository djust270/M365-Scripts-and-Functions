$StartTime = Get-Date
<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2021 v5.8.195
	 Created on:   	9/21/2022 12:16 PM
	 Created by:   	david.just
	 Organization: 	
	 Filename: Export-O365MasterAuditReport.ps1
	===========================================================================
	.DESCRIPTION
		Uses the Microsoft Graph PowerShell SDK to generate a tenant audit report
#>
#region Check and load Required Modules
$isAdmin = [bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544")
$prereqmodules = @(
	'ImportExcel'
	'Microsoft.Online.Sharepoint.PowerShell'
	'Microsoft.Graph'
	'ExchangeOnlineManagement'
)
Write-Host "Checking Available Modules"
$available = Get-Module -ListAvailable | Select-Object -ExpandProperty Name
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
foreach ($module in $prereqmodules)
{
	if ($module -notin $available)
	{
		"{0} module not detected, attempting to install..." -f $module
		if (-not $isadmin)
		{ install-module $module -scope CurrentUser }
		else { Install-Module $module }
	}	
}
if ($PSEdition -eq "Core")
{
	Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell
}
else
{
	Import-Module Microsoft.Online.SharePoint.PowerShell
}
#endregion

#region Functions
#File Save box function 
Function Get-SaveFolderLocation
{
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |	Out-Null
	$SaveFolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
	#$SaveFolderDialog.Description = "Select a folder to save reports to" | out-null
	$SaveFolderDialog.ShowDialog() | out-null
	return $SaveFolderDialog.SelectedPath
}

function Get-MGTeamsUsageReport
{
	param (
		[Parameter(Mandatory)]
		[ValidateSet('D30', 'D90', 'D180')]
		$Period
	)
	try
	{
		$tempfile = New-TemporaryFile
		Invoke-MgGraphRequest -Uri "/beta/reports/getTeamsTeamActivityDetail(period='$Period'`)`?`$format=text/csv" -ErrorAction Stop -OutputFilePath $tempfile
		$TeamsReport = Import-csv $tempfile		
	}
	Catch
	{
		$_ ; break
	}
	Remove-Item $tempfile
	return $TeamsReport
}

function Get-MGUserPrincipalGroupMembership
{
<#
.SYNOPSIS
This function is used to return all groups a user is a member of
.EXAMPLE
Get-MGUserPrincipalGroupMembership -UserID (Get-AzureADUser -searchstring user@contoso.com).objectid
.NOTES
Created by David Just 09/06/2021
#>
	param (
		[string]$UserID
	)
	$graphApiVersion = "v1.0"
	$Resource = "users/$UserID/memberOf/$/microsoft.graph.group?$select=id,displayName,securityEnabled,groupTypes,onPremisesSyncEnabled"
	$uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)"
	$groups = (Invoke-MgGraphRequest -Uri $uri -Method Get).Value
	return $groups
}

function Export-SharedMailboxReport
{
	param (
		[string]$Workbook
	)
	#Check if connected to exchange
	try
	{
		Get-EXOMailbox -ResultSize 1 -ErrorAction 'Stop' | Out-Null
	}
	catch
	{
		"Connecting to ExchangeOnline"
		Connect-ExchangeOnline
	}
	
	$i = 0	
	$Shared = Get-Mailbox -Filter { isShared -eq 'true' }
	$SharedPermissions = foreach ($box in $shared)
	{
		Write-Progress -Activity "Processing Shared Mailbox Permissions" -Status "Working on $($box.displayname)" -PercentComplete (($i / $shared.Count) * 100)
		Get-MailboxPermission $box.identity | Where-Object { $_.User -notlike "*NT Authority*" } | Select-Object identity, user, @{ name = 'accessrights'; e = { ($_.accessrights -join ' , ') } }
	}
	$shared | Select-Object Identity, Displayname, PrimarySMTPAddress, @{
		n = 'EmailAddresses'; e = { ($_.EmailAddresses -join ' , ') }
	},HiddenFromAddressListsEnabled | Export-Excel -Path $Workbook -WorksheetName "SharedMailboxReport" -TableName "SharedMailboxes" -AutoSize
	$SharedPermissions | Export-Excel -Path $Workbook -WorksheetName "SharedMailboxPermissions" -TableName "SharedPerms" -AutoSize
}

function Export-MailboxReport #Export details on all non-shared mailboxes
{
	param (
		[string]$Workbook
	)
	$mailboxes = Get-Mailbox -Filter { isShared -eq 'false' }
	$i = 0
	$MailboxDetails = foreach ($box in $mailboxes)
	{
		Write-Progress -Activity "Processing Mailbox Report" -Status "Working on $($box.displayname)" -PercentComplete (($i / $mailboxes.Count) * 100)
		$TotalSize = (Get-MailboxStatistics -Identity $box.identity).TotalItemSize
		$box | Select-Object Identity, Displayname, PrimarySMTPAddress, @{
			n = 'EmailAddresses'; e = { ($_.EmailAddresses -join ' , ') }
		}, HiddenFromAddressListsEnabled, @{ n = 'TotalSize'; e = { $TotalSize } }
	}
	$MailboxDetails | Export-Excel -Path $Workbook -WorksheetName "MailboxReport" -TableName "Mailboxes" -AutoSize	
}
#endregion

$CompanyName = Read-Host "Enter Company Name for report title"
Write-Host "Please select a folder to save reports to:" -ForegroundColor White -BackgroundColor Black
$folderpath = Get-SaveFolderLocation
$ReportWorkbook = "$folderpath\$CompanyName-O365AuditReport.xlsx"

# Connect to Graph, SPO, EXO
$perms = @(
	'User.Read.All'
	'Directory.Read.All'
	'Application.Read.All'
	'Channel.ReadBasic.All'
	'Team.ReadBasic.All'
	'TeamMember.Read.All'
	'ReportSettings.ReadWrite.All'
	'Reports.Read.All'
	'Sites.Read.All'
)

Connect-MgGraph -ForceRefresh -Scopes $perms
Select-MgProfile beta
$SharepointAdminURL = (Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/v1.0/sites?$select=siteCollection,webUrl&$filter=siteCollection/root%20ne%20null').value.weburl -replace '.sharepoint', '-admin.sharepoint'
Connect-ExchangeOnline -ShowBanner:$false
Connect-SPOService -Url $SharepointAdminURL

#Region User Report
Write-Progress -Activity "Working on License Reports"
$FriendlyLicenses = @{
	'O365_BUSINESS_ESSENTIALS'		     = 'Office 365 Business Essentials'
	'O365_BUSINESS_PREMIUM'			     = 'Office 365 Business Premium'
	'DESKLESSPACK'					     = 'Office 365 (Plan K1)'
	'DESKLESSWOFFPACK'				     = 'Office 365 (Plan K2)'
	'LITEPACK'						     = 'Office 365 (Plan P1)'
	'EXCHANGESTANDARD'				     = 'Office 365 Exchange Online Only'
	'STANDARDPACK'					     = 'Enterprise Plan E1'
	'STANDARDWOFFPACK'				     = 'Office 365 (Plan E2)'
	'ENTERPRISEPACK'					 = 'Enterprise Plan E3'
	'ENTERPRISEPACKLRG'				     = 'Enterprise Plan E3'
	'ENTERPRISEWITHSCAL'				 = 'Enterprise Plan E4'
	'STANDARDPACK_STUDENT'			     = 'Office 365 (Plan A1) for Students'
	'STANDARDWOFFPACKPACK_STUDENT'	     = 'Office 365 (Plan A2) for Students'
	'ENTERPRISEPACK_STUDENT'			 = 'Office 365 (Plan A3) for Students'
	'ENTERPRISEWITHSCAL_STUDENT'		 = 'Office 365 (Plan A4) for Students'
	'STANDARDPACK_FACULTY'			     = 'Office 365 (Plan A1) for Faculty'
	'STANDARDWOFFPACKPACK_FACULTY'	     = 'Office 365 (Plan A2) for Faculty'
	'ENTERPRISEPACK_FACULTY'			 = 'Office 365 (Plan A3) for Faculty'
	'ENTERPRISEWITHSCAL_FACULTY'		 = 'Office 365 (Plan A4) for Faculty'
	'ENTERPRISEPACK_B_PILOT'			 = 'Office 365 (Enterprise Preview)'
	'STANDARD_B_PILOT'				     = 'Office 365 (Small Business Preview)'
	'VISIOCLIENT'					     = 'Visio Pro Online'
	'POWER_BI_ADDON'					 = 'Office 365 Power BI Addon'
	'POWER_BI_INDIVIDUAL_USE'		     = 'Power BI Individual User'
	'POWER_BI_STANDALONE'			     = 'Power BI Stand Alone'
	'POWER_BI_STANDARD'				     = 'Power-BI Standard'
	'PROJECTESSENTIALS'				     = 'Project Lite'
	'PROJECTCLIENT'					     = 'Project Professional'
	'PROJECTONLINE_PLAN_1'			     = 'Project Online'
	'PROJECTONLINE_PLAN_2'			     = 'Project Online and PRO'
	'ProjectPremium'					 = 'Project Online Premium'
	'ECAL_SERVICES'					     = 'ECAL'
	'EMS'							     = 'Enterprise Mobility Suite'
	'RIGHTSMANAGEMENT_ADHOC'			 = 'Windows Azure Rights Management'
	'MCOMEETADV'						 = 'PSTN conferencing'
	'SHAREPOINTSTORAGE'				     = 'SharePoint storage'
	'PLANNERSTANDALONE'				     = 'Planner Standalone'
	'CRMIUR'							 = 'CMRIUR'
	'BI_AZURE_P1'					     = 'Power BI Reporting and Analytics'
	'INTUNE_A'						     = 'Windows Intune Plan A'
	'PROJECTWORKMANAGEMENT'			     = 'Office 365 Planner Preview'
	'ATP_ENTERPRISE'					 = 'Exchange Online Advanced Threat Protection'
	'EQUIVIO_ANALYTICS'				     = 'Office 365 Advanced eDiscovery'
	'AAD_BASIC'						     = 'Azure Active Directory Basic'
	'RMS_S_ENTERPRISE'				     = 'Azure Active Directory Rights Management'
	'AAD_PREMIUM'					     = 'Azure Active Directory Premium'
	'MFA_PREMIUM'					     = 'Azure Multi-Factor Authentication'
	'STANDARDPACK_GOV'				     = 'Microsoft Office 365 (Plan G1) for Government'
	'STANDARDWOFFPACK_GOV'			     = 'Microsoft Office 365 (Plan G2) for Government'
	'ENTERPRISEPACK_GOV'				 = 'Microsoft Office 365 (Plan G3) for Government'
	'ENTERPRISEWITHSCAL_GOV'			 = 'Microsoft Office 365 (Plan G4) for Government'
	'DESKLESSPACK_GOV'				     = 'Microsoft Office 365 (Plan K1) for Government'
	'ESKLESSWOFFPACK_GOV'			     = 'Microsoft Office 365 (Plan K2) for Government'
	'EXCHANGESTANDARD_GOV'			     = 'Microsoft Office 365 Exchange Online (Plan 1) only for Government'
	'EXCHANGEENTERPRISE_GOV'			 = 'Microsoft Office 365 Exchange Online (Plan 2) only for Government'
	'SHAREPOINTDESKLESS_GOV'			 = 'SharePoint Online Kiosk'
	'EXCHANGE_S_DESKLESS_GOV'		     = 'Exchange Kiosk'
	'RMS_S_ENTERPRISE_GOV'			     = 'Windows Azure Active Directory Rights Management'
	'OFFICESUBSCRIPTION_GOV'			 = 'Office ProPlus'
	'MCOSTANDARD_GOV'				     = 'Lync Plan 2G'
	'SHAREPOINTWAC_GOV'				     = 'Office Online for Government'
	'SHAREPOINTENTERPRISE_GOV'		     = 'SharePoint Plan 2G'
	'EXCHANGE_S_ENTERPRISE_GOV'		     = 'Exchange Plan 2G'
	'EXCHANGE_S_ARCHIVE_ADDON_GOV'	     = 'Exchange Online Archiving'
	'EXCHANGE_S_DESKLESS'			     = 'Exchange Online Kiosk'
	'SHAREPOINTDESKLESS'				 = 'SharePoint Online Kiosk'
	'SHAREPOINTWAC'					     = 'Office Online'
	'YAMMER_ENTERPRISE'				     = 'Yammer for the Starship Enterprise'
	'EXCHANGE_L_STANDARD'			     = 'Exchange Online (Plan 1)'
	'MCOLITE'						     = 'Lync Online (Plan 1)'
	'SHAREPOINTLITE'					 = 'SharePoint Online (Plan 1)'
	'OFFICE_PRO_PLUS_SUBSCRIPTION_SMBIZ' = 'Office ProPlus'
	'EXCHANGE_S_STANDARD_MIDMARKET'	     = 'Exchange Online (Plan 1)'
	'MCOSTANDARD_MIDMARKET'			     = 'Lync Online (Plan 1)'
	'SHAREPOINTENTERPRISE_MIDMARKET'	 = 'SharePoint Online (Plan 1)'
	'OFFICESUBSCRIPTION'				 = 'Office ProPlus'
	'YAMMER_MIDSIZE'					 = 'Yammer'
	'DYN365_ENTERPRISE_PLAN1'		     = 'Dynamics 365 Customer Engagement Plan Enterprise Edition'
	'ENTERPRISEPREMIUM_NOPSTNCONF'	     = 'Enterprise E5 (without Audio Conferencing)'
	'ENTERPRISEPREMIUM'				     = 'Enterprise E5 (with Audio Conferencing)'
	'MCOSTANDARD'					     = 'Skype for Business Online Standalone Plan 2'
	'PROJECT_MADEIRA_PREVIEW_IW_SKU'	 = 'Dynamics 365 for Financials for IWs'
	'STANDARDWOFFPACK_IW_STUDENT'	     = 'Office 365 Education for Students'
	'STANDARDWOFFPACK_IW_FACULTY'	     = 'Office 365 Education for Faculty'
	'EOP_ENTERPRISE_FACULTY'			 = 'Exchange Online Protection for Faculty'
	'EXCHANGESTANDARD_STUDENT'		     = 'Exchange Online (Plan 1) for Students'
	'OFFICESUBSCRIPTION_STUDENT'		 = 'Office ProPlus Student Benefit'
	'STANDARDWOFFPACK_FACULTY'		     = 'Office 365 Education E1 for Faculty'
	'STANDARDWOFFPACK_STUDENT'		     = 'Microsoft Office 365 (Plan A2) for Students'
	'DYN365_FINANCIALS_BUSINESS_SKU'	 = 'Dynamics 365 for Financials Business Edition'
	'DYN365_FINANCIALS_TEAM_MEMBERS_SKU' = 'Dynamics 365 for Team Members Business Edition'
	'FLOW_FREE'						     = 'Microsoft Flow Free'
	'POWER_BI_PRO'					     = 'Power BI Pro'
	'O365_BUSINESS'					     = 'Office 365 Business'
	'DYN365_ENTERPRISE_SALES'		     = 'Dynamics Office 365 Enterprise Sales'
	'RIGHTSMANAGEMENT'				     = 'Rights Management'
	'PROJECTPROFESSIONAL'			     = 'Project Professional'
	'VISIOONLINE_PLAN1'				     = 'Visio Online Plan 1'
	'EXCHANGEENTERPRISE'				 = 'Exchange Online Plan 2'
	'DYN365_ENTERPRISE_P1_IW'		     = 'Dynamics 365 P1 Trial for Information Workers'
	'DYN365_ENTERPRISE_TEAM_MEMBERS'	 = 'Dynamics 365 For Team Members Enterprise Edition'
	'CRMSTANDARD'					     = 'Microsoft Dynamics CRM Online Professional'
	'EXCHANGEARCHIVE_ADDON'			     = 'Exchange Online Archiving For Exchange Online'
	'EXCHANGEDESKLESS'				     = 'Exchange Online Kiosk'
	'SPZA_IW'						     = 'App Connect'
	'WINDOWS_STORE'					     = 'Windows Store for Business'
	'MCOEV'							     = 'Microsoft Phone System'
	'VIDEO_INTEROP'					     = 'Polycom Skype Meeting Video Interop for Skype for Business'
	'SPE_E5'							 = 'Microsoft 365 E5'
	'SPE_E3'							 = 'Microsoft 365 E3'
	'ATA'							     = 'Advanced Threat Analytics'
	'MCOPSTN2'						     = 'Domestic and International Calling Plan'
	'FLOW_P1'						     = 'Microsoft Flow Plan 1'
	'FLOW_P2'						     = 'Microsoft Flow Plan 2'
	'DeveloperPack'					     = 'OFFICE 365 ENTERPRISE E3 DEVELOPER'
	'EMSPremium'						 = 'ENTERPRISE MOBILITY + SECURITY E5'
	'RightsManagemnt'				     = 'AZURE INFORMATION PROTECTION PLAN 1'
	'DYN365_ENTERPRISE_CUSTOMER_SERVICE' = 'DYNAMICS 365 FOR CUSTOMER SERVICE ENTERPRISE EDITION'
	'POWERFLOW_P1'					     = 'Microsoft PowerApps Plan 1'
	'POWERFLOW_P2'					     = 'Microsoft PowerApps Plan 2'
	'AAD_PREMIUM_P1'					 = 'Azure Active Directory Premium P1'
	'AAD_PREMIUM_P2'					 = 'Azure Active Directory Premium P2'
	'TEAMS_EXPLORATORY'				     = 'Teams Exploratory'
	'MDATP_XPLAT'					     = 'Microsoft Defender for Endpoint P2'
	
}

# Get all tenant skus
[Array]$Skus = Get-MgSubscribedSku
$TenantLicenseDetails = $Skus | foreach {
	[pscustomobject]@{
		SkuPartNumber = $_.SkuPartNumber
		TotalLicense  = $_.prepaidunits.enabled
		UsedLicense   = $_.ConsumedUnits
		FriendlyName = $FriendlyLicenses[$_.SkuPartNumber]
	}
}

[Array]$Users = Get-MGUser -All
$i = 0
foreach ($user in $Users)
{
	Write-Progress -Activity "Processing User License details" -Status "Working on $($user.displayname)" -PercentComplete (($i / $Users.Count) * 100)
	$user.LicenseDetails = Get-MgUserLicenseDetail -UserId $user.id
	if ($user.UserType -eq "Member") { $user.Add('GroupMemberships', ((Get-MGUserPrincipalGroupMembership -UserID $user.userprincipalname).displayname -join ' ; ')) }
	$i++
}
$UserLicenseDetails = $Users | Select-Object UserPrincipalName, AccountEnabled, Mail, PasswordPolicies, UserType, CreatedDateTime, SignInSessionsValidFromDateTime, OnPremisesSyncEnabled, MobilePhone,
@{
	name = 'Licenses'
	expression = {
		($_ | foreach {
				$_.licensedetails | foreach {
					if ($FriendlyLicenses[$_.SkuPartNumber]) { $FriendlyLicenses[$_.SkuPartNumber] }
					elseif ($.SkuPartID) { $_.SkuPartID }
					else {"No license"}
				}
			}) -join ' ; '
	}
},@{ name = "MemberOf"; Expression = { $_.additionalproperties.GroupMemberships } }

$UnLicensed = $Users | where { -not $_.LicenseDetails } | select UserPrincipalName, AccountEnabled, Mail, PasswordPolicies, UserType
$LicenseReport = "$folderpath\$CompanyName-UserLicenseAudit.csv"
$UserLicenseDetails | Export-Excel -path $ReportWorkBook -WorksheetName "UserDetails" -tablename "UserDetails" -Autosize
$TenantLicenseDetails | Export-Excel -path $ReportWorkBook -WorksheetName "TenantLicenseDetails" -TableName "LicenseDetails" -AutoSize

#endregion

#Region SSO Apps
Write-Progress -Activity "Working on SSO Apps"
$Apps = Get-MgServicePrincipal -All
$SSOApps = $Apps | Where-Object { $_.KeyCredentials.Displayname -eq "CN=Microsoft Azure Federated SSO Certificate" } | Select-Object displayname, LoginURL, Homepage, AppID, @{
	name	   = 'NotificaitonEmailaddresses'
	expression = { $_.NotificationEmailAddresses -join ' ; ' }
}, @{
	name	   = 'SSO Certificate Expiration Date';
	expression = { $_.keycredentials[0].EndDateTime.ToShortDateString() }
}
$SSOApps | Export-Excel -Path $ReportWorkbook -WorksheetName "SSOEnterpriseApps" -TableName "SSOApps" -AutoSize
#endregion

#region Admin Roles
Write-Progress -Activity "Working on Admins Report"
$AdminRoles = Get-MgDirectoryRole
$AdminRolesAndMembers = foreach ($Role in $AdminRoles)
{
	$Members = (Get-MgDirectoryRoleMember -DirectoryRoleId $Role.ID).AdditionalProperties.userPrincipalName
	[PSCustomobject]@{
		AdminRole = $Role.DisplayName
		RoleDescription = $Role.Description
		Members   = $Members -join ' ; '
	}
}
$AdminRolesAndMembers | Export-Excel -path $ReportWorkBook -WorksheetName "AdminRoles" -tablename "AdminRolesAndMembers" -Autosize
#endregion
#region Groups
Write-Progress -Activity "Working on Group Reports"

$groups = Get-MgGroup -All

$SecurityGroups = $groups | Where-Object { -not $_.GroupTypes -and $_.SecurityEnabled } | Select-Object DisplayName, MailEnabled, Mail, id, @{
	name	   = 'Source'
	expression = {
		if ($_.OnPremisesSyncEnabled) { "Windows Server AD" }
		else { "Cloud" }
	}
},
																										@{
	name	   = "Type"
	expression = { "Security" }
}, MembershipRule
$UnifiedGroups = $groups | Where-Object { $_.GroupTypes -eq 'Unified' } | Select-Object DisplayName, MailEnabled, Mail, id, @{
	name	   = 'Source'
	expression = {
		if ($_.OnPremisesSyncEnabled) { "Windows Server AD" }
		else { "Cloud" }
	}
},
																						@{
	name	   = "Type"
	expression = { "Microsoft365" }
}, MembershipRule
$DynamicGroups = $groups | Where-Object { $_.GroupTypes -eq 'DynamicMembership' } | Select-Object DisplayName, MailEnabled, Mail, id, @{
	name	   = 'Source'
	expression = {
		if ($_.OnPremisesSyncEnabled) { "Windows Server AD" }
		else { "Cloud" }
	}
},
																								  @{
	name	   = "Type"
	expression = { "Dynamic" }
},MembershipRule
$DistroGroups = $groups | Where-Object { -not $_.GroupTypes -and -not $_.SecurityEnabled } | Select-Object DisplayName, MailEnabled, Mail, id, @{
	name	   = 'Source'
	expression = {
		if ($_.OnPremisesSyncEnabled) { "Windows Server AD" }
		else { "Cloud" }
	}
},
																										   @{
	name	   = "Type"
	expression = { "Distribution" }
},MembershipRule

$AllGroups = @(
	$SecurityGroups
	$UnifiedGroups
	$DynamicGroups
	$DistroGroups
)

$AllGroups | Export-Excel -Path $ReportWorkbook -WorksheetName "GroupDetails" -TableName "GroupDetails" -AutoSize
$GroupMemberList = [System.Collections.Generic.List[PsObject]]::new()
$i = 0
foreach ($group in $AllGroups)
{
	Write-Progress -Activity "Processing GroupMemberships" -Status "Working on $($group.displayname)" -PercentComplete (($i / $AllGroups.Count) * 100)	
	$Members = Get-MgGroupMember -GroupId $group.id
	$Members | foreach {
		$GroupMemberList.Add(
			[pscustomobject]@{
				'GroupName' = $group.displayname
				'GroupType' = $group.type
				'Member'    = $_.AdditionalProperties.userPrincipalName
				'MemberID' = $_.id 
			}
		)
	}
	$i++
}
$GroupMemberList | Export-Excel -Path $ReportWorkbook -WorksheetName 'GroupMembers' -TableName 'GroupMembers' -AutoSize
#endregion

#region Teams
Write-Progress -Activity "Working on Teams Reports"
$Teams = Get-MgTeam -All
$TeamsDetails = [System.Collections.Generic.List[PsObject]]::new()
$TeamsChannels = [System.Collections.Generic.List[PsObject]]::new()
$i = 0
foreach ($Team in $Teams)
{
	Write-Progress -Activity "Processing Teams details" -Status "Working on $($Team.displayname)" -PercentComplete (($i / $Teams.Count) * 100)
	$TeamMembers = Get-MgTeamMember -TeamId $Team.id
	Write-Progress -Activity "Working on Team Channels"
	$Channels = Get-MgTeamChannel -TeamId $Team.id
	$i++
	$TeamsDetails.Add([pscustomobject]@{
			'TeamName' = $Team.displayname
			'Description' = $Team.Description
			'Visibility' = $Team.Visibility
			'Members' = $TeamMembers.AdditionalProperties.email -join ' ; '
		})
	$Channels | foreach {
			$TeamsChannels.Add(
				[pscustomobject]@{
					'Team' = $Team.displayname
					'Channel' = $_.displayname
				}
			)
		}
}
$TeamsDetails | Export-Excel -Path $ReportWorkbook -WorksheetName "Teams" -TableName 'Teams' -AutoSize
$TeamsChannels | Export-Excel -Path $ReportWorkbook -WorksheetName "TeamsChannels" -TableName 'TeamsChannels' -AutoSize

# Enable Display Concealed Names in Reports
Invoke-MGGraphRequest -Method PATCH -uri "/beta/admin/reportSettings" -body @{ "displayConcealedNames" = $false }
Get-MGTeamsUsageReport -Period D90 | Export-Excel -Path $ReportWorkbook -WorksheetName "TeamsUsageReport" -TableName "TeamsUsageReport" -AutoSize
#endregion

#region Sharepoint/OneDrive Sites
Write-Progress -Activity "Working on Sharepoint / OneDrive reports"
$sites = Get-SPOSite -IncludePersonalSite $true -Limit ALL
$OneDriveSites = $sites | Where-Object { $_.URL -like "*-my.sharepoint.com/personal/*" }
$sites = $sites | Where-Object { $_.URL -notlike "*-my*" }
$sites | Select-Object Title, Url, StorageQuota, StorageUsageCurrent, Owner, SharingCapability | Export-Excel -Path $ReportWorkbook -WorksheetName "SharePointSites" -TableName "SPSites" -AutoSize
$OneDriveSites | Select-Object Title, Url, StorageQuota, StorageUsageCurrent, Owner, SharingCapability | Export-Excel -Path $ReportWorkbook -WorksheetName "OneDrive-Summary" -TableName "OneDrive" -AutoSize

<#
$AdminUPN = read-host "Enter your User Principal Name (admin account)"
$perms = [System.Collections.Generic.List[PsObject]]::new()
$i = 0
foreach ($site in $sites)
{
	Set-SPOUser -site $ODsite -LoginName $AdminUPN -IsSiteCollectionAdmin $true	
	Write-Progress "Working on $($site.title) Permissions. This will take awhile" -CurrentOperation "Working on site $i out of $($sites.count)"
	$i++
	$SiteGroups = Get-SPOSiteGroup -Site $site.url
	ForEach ($Group in $SiteGroups)
	{
		$perms.add([pscustomobject]@{
				'SiteName' = $site.title
				'URL'	   = $site.url
				'Owner'    = $site.owner
				'Group Name' = $Group.Title
				'Permissions' = $Group.Roles -join ","
				'Users'    = foreach ($user in $Group.users)
				{
					try
					{
						if ([guid]$user) { (Get-MgGroupMember -GroupId $user).AdditionalProperties.userPrincipalName -join ',' }
					}
					Catch
					{
						if ($user -match '(?im)^[{(]?[0-9A-F]{8}[-]?(?:[0-9A-F]{4}[-]?){3}[0-9A-F]{12}[_][]o[)}]?$')
						{
							Connect-PnPOnline -Url $site.url -PnPManagementShell
							(Get-PnPGroupMember -Group $Group.title).loginname -join ','
						}
						else { $User }
					}
				}
			})
	}
	
}
$perms | Select-Object sitename, "Group Name", Permissions, @{ n = 'Users'; e = { $_.users -join ',' } } | Export-Excel -Path $ReportWorkbook -WorksheetName "SharePoint Site Permissions" -TableName "SPPermissions" -AutoSize
#>
#endregion 

#region Shared Mailboxes
"Working on Shared Mailbox Reports"
Export-SharedMailboxReport -Workbook $ReportWorkbook
"Working on Mailboxes"
Export-MailboxReport -Workbook $ReportWorkbook
#endregion

$EndTime = Get-Date
"Completed audit report in {0} Minutes, {1} Seconds" -f ($EndTime - $StartTime).Minutes, ($EndTime - $StartTime).Seconds
Disconnect-MgGraph
Disconnect-ExchangeOnline -Confirm:$false
Disconnect-SPOService
"Report saved to {0}" -f $ReportWorkbook