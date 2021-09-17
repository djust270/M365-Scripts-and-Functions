<# 
Requires AzureAD PS Module
Written by David Just 

.SYNOPSIS
Will create policy to block legacy authentication
Creates Legacy Auth Enabled group if not already created. 

.PARAMETER State
    Sets state of policy. Valid values are: Enabled , Disabled, OR enabledForReportingButNotEnforced
#>

param
(
	[string]$state = 'enabledForReportingButNotEnforced'
)

$LegacyAuthParams = @{
	DisplayName	    = "Legacy Auth Enabled"
	Description	    = "Users and Service Accounts using Legacy Authentication Methods"
	MailEnabled	    = $False
	MailNickName    = 'LegacyAuthEnabled'
	SecurityEnabled = $True
	OutVariable 	= 'LegacyAuthGroup'
}

if (!(Get-AzureADMSGroup -All $true | where displayname -like "Legacy Auth Enabled" -OutVariable LegacyAuthGroup))
{
	New-AzureADMSGroup @LegacyAuthParams
}

#Region Add Conditions

$conditions = New-Object -TypeName Microsoft.Open.MSGraph.Model.ConditionalAccessConditionSet
$conditions.Applications = New-Object -TypeName Microsoft.Open.MSGraph.Model.ConditionalAccessApplicationCondition
$conditions.locations = New-Object -TypeName Microsoft.Open.MSGraph.Model.ConditionalAccessLocationCondition
$conditions.Platforms = New-Object -TypeName Microsoft.Open.MSGraph.Model.ConditionalAccessPlatformCondition
$conditions.Users = New-Object -TypeName Microsoft.Open.MSGraph.Model.ConditionalAccessUserCondition

$conditions.Applications.IncludeApplications = "All"

$conditions.Users.IncludeUsers = "All"
$conditions.Users.ExcludeGroups = $LegacyAuthGroup.Id

$conditions.platforms.IncludePlatforms = "All"

$conditions.locations.IncludeLocations = "All"

$conditions.clientapptypes = 'ExchangeActiveSync','Other'

#Region Grant Controls
$GrantControls = New-Object -TypeName Microsoft.Open.MSGraph.Model.ConditionalAccessGrantControls
$GrantControls._Operator = "OR"
$GrantControls.BuiltInControls = "Block"

New-AzureADMSConditionalAccessPolicy -DisplayName "Block Legacy Authentication" -State $state -Conditions $conditions -GrantControls $GrantControls