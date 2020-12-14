<#
.Synopsis
Simple script leveraging the Graph SDK that will delete all contacts for a user. 
Written By David Just
#>
# Requires the Graph Powershell SDK.
Connect-Graph -Scopes "User.Read","User.ReadWrite.All","Mail.ReadWrite",`
            "Directory.Read.All","Chat.ReadWrite", "People.Read", `
            "Group.Read.All", "Tasks.ReadWrite", `
            "Sites.Manage.All","Contacts.ReadWrite","Contacts.Read","Contacts.Read.Shared","Contacts.ReadWrite.Shared"

for ($loop = 1; $loop lt #number of loops ; $loop++){
	Write-Host "Loop Iteration $loop"
	$i=0
$contacts = Get-MgUserContact -UserId user@domain.com -top 5000
foreach ($contact in $contacts){
write-progress -activity "Processing" -Status "Removing Contact $($contact.displayname)" -PercentComplete (($i / $contacts.count) * 100)
Remove-MgUserContact -ContactId $contact.id -UserId user@domain.com
		$i++
	}
}
	