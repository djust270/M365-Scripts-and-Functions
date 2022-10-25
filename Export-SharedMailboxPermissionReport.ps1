
Function Get-SaveFolderLocation
{
	[Cmdletbinding()]
	param(
		[String]$Description
	)
	[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |	Out-Null
	$SaveFolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
	$SaveFolderDialog.Description = $Description | out-null
	$SaveFolderDialog.ShowDialog() | out-null
	return $SaveFolderDialog.SelectedPath
}
Write-Host "Select a folder to save this report to:" -ForegroundColor Green -BackgroundColor Black
$ReportFolder = Get-SaveFolderLocation
$prereqmodules = @(
	'ImportExcel'	
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
		{
			install-module $module -scope CurrentUser			
		}
		else
		{
			Install-Module $module			
		}
	}
}
Write-Host "Connecting to Exhange Online..."
Connect-ExchangeOnline -ShowBanner:$false
$DefaultDomain = (Get-AcceptedDomain | Where-Object {$_.default}).Name
$SharedMailboxes = Get-Mailbox -Filter { isShared -eq 'true' } -Resultsize Unlimited
$i = 1
$PermissionReportArray = foreach ($box in $SharedMailboxes){
    Write-Progress -Activity "Working on Shared Mailbox Permission Report" -Status "Working on $($box.Identity)" -PercentComplete (($i / $SharedMailboxes.Count) * 100)
    $Permissions = Get-MailboxPermission -Identity $box.Guid.Guid | where-object {$_.user -notmatch "NT AUTHORITY\SELF"}
    $Permissions | ForEach-Object {
        [PSCustomObject]@{
            Identity = $_.identity
            PrimarySMTPAddress = $box.PrimarySMTPAddress
            User = $_.User
            AccessRights = $($_.AccessRights -join ',')
        }        
    }
    $i++
}
$PermissionReportArray | Export-Excel -Path "$ReportFolder\$DefaultDomain-SharedMailboxPermissionReport.xlsx" -WorksheetName "SharedMailboxReport" -TableName "SharedPermissions" -AutoSize
Write-Host "Report saved to "  $("$ReportFolder\$DefaultDomain-SharedMailboxPermissionReport.xlsx")
& "$ReportFolder\$DefaultDomain-SharedMailboxPermissionReport.xlsx"
Disconnect-ExchangeOnline
