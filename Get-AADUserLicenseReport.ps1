<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2020 v5.7.182
	 Created on:   	4/4/2021 1:22 PM
	 Created by:   	David Just
	 Filename:     	Get-AADUserLicenseReport.ps1
	===========================================================================
	
#>

<#
	.SYNOPSIS
		Exports Currently Licensed Users details
	
	.DESCRIPTION
		Gets all licensed users including license details, name, email and location
	
	.EXAMPLE
				PS C:\> Get-AADUserLicneseReport -ExportFolderPath C:\Users\Dave\Desktop
	
	.OUTPUTS
		String
	
	.NOTES
		Additional information about the function.
#>
function Get-AADUserLicneseReport
{
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[String]$ExportFolderPath
	)
	
	Connect-AzureAD
	$date = ((Get-Date).ToShortDateString()).Replace('/', '-')
	$TenantID = Get-AzureADTenantDetail | select -ExpandProperty ObjectID
	$users = Get-AzureADUser -all $true | where assignedlicenses -ne $null
	
	$list = [System.Collections.Generic.List[PsObject]]::new()
	foreach ($user in $users)
	{
		$licenses = ($user).assignedlicenses.skuid | foreach { (Get-AzureADSubscribedSku -ObjectID ($($TenantID) + '_' + $_)).skupartnumber }
		$lic = $licenses -join ' '
		$licComma = $lic -replace ' ', ','
		$list.add([PSCustomObject]@{
				Name = $user.displayname
				Email = $user.UserPrincipalName
				licenses = $licComma
				StreetAddress = $user.streetaddress
			})
	}
	
	$list | export-csv ($ExportFolderPath + '\' + 'AADLicenseReport' + $date + '.csv') -NoTypeInformation
	Disconnect-AzureAD
}

