<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2020 v5.7.182
	 Created on:   	4/4/2021 
	 Created by:   	David Just
	 Filename:    Get-MsolMFAStatus.ps1 	
	===========================================================================
	.SYNOPSIS
		Gets User MFA Status Report 
	.DESCRIPTION
		Gets MFA Status Enabled/Disabled. If user is part of a conditional access policy targeted group status is set to "Conditional"
	.EXAMPLE 
	PS Get-MsolMFAStatus -ConditionalAccessGroups 'Group1','Group2'

#>

function Get-MSOLMFAStatus
{
	[CmdletBinding()]
	param (
		[String[]]$ConditionalAccessGroups
	)
	#Connect-MsolService
	$Users = Get-MsolUser -MaxResults 10000 | Where-Object { ($_.islicensed -eq $true) -and (($_.licenses).accountskuid -like "*SPE*") -or (($_.licenses).accountskuid -like "*O365*") } | select userprincipalname, displayname, islicensed, @{
		n = 'MFAStatus'; e = {
			if ((($_.strongauthenticationrequirements).state) -eq $null) { "Disabled" }
			else { ($_.StrongAuthenticationRequirements).state }
		}
	}
	$conditionalGroups = if ($ConditionalAccessGroups.Count -gt 0)
	{
		foreach ($group in $ConditionalAccessGroups)
		{
			Get-MsolGroup | where { ($_.displayname -match $group) }
		}
	}
	$conditionalusers =
	foreach ($user in $users)
	{
		foreach ($group in $conditionalGroups)
		{
			if ($group -ne $null)
			{
				if (Get-MsolGroupMember -GroupObjectId $Group.ObjectId | where { $_.EmailAddress -eq ($($user).userprincipalname) }) { $user.userprincipalname }
			}
		}
	}
	
	$mfaReport = [System.Collections.Generic.List[PSObject]]::new()
		foreach ($user in $users)
		{
			$mfaReport.Add([pscustomobject]@{
					Email = $user.userprincipalname
					Displayname = $user.displayname
					MFAStatus = if ($conditionalusers -like $user.userprincipalname) { "Conditional" }else{ $user.mfastatus }
					Licensed = $user.islicensed
				}
			)
		
	}
	Write-Output $mfaReport
}

