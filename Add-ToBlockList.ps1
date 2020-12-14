#Requires -Module ExchangeOnlineManagement
<#
.Synopsis
Adds sender or domain to the spam filter blocked senders list
.Parameter SpamFilterName
Specify the SpamFilter name to modify. Default is Default
.Parameter BlockedSender
Add the email address(es) you wish to block
.Parameter BlockedDomain
Add the domain(s) you wish to block
Written By David Just
#>
function Add-ToBlockList {
[cmdletbinding()]
Param (
    [String]$SpamFilterName ="default",
    [string[]]$BlockedSender,
    [string[]]$BlockedDomain
)

$exchangemodule = Get-Module ExchangeOnlineManagement
$exchangemoduleinstalled = Get-InstalledModule ExchangeOnlineManagement

Connect-ExchangeOnline

$filter = Get-HostedContentFilterPolicy -Identity $blocklist
$senders = $filter.BlockedSenders.sender | select -expandproperty address 
$domains = $filter.BlockedSenderDomains.Domain
$nonfixedsenders = [System.Collections.Arraylist]@($senders)
$nonfixeddomains = [System.Collections.Arraylist]@($domains)

    foreach ($Sender in $BlockedSender){
        Write-Host "Adding $Sender to the $SpamFilterName Blocked List"
        Write-Host ""
        $nonfixedsenders.Add($Sender)
        Set-HostedContentFilterPolicy -Identity $SpamFilterName -BlockedSenders $nonfixedsenders 
        }
    

    foreach ($Domain in $BlockedDomain){
        Write-Host "Adding $Domain to the $SpamFilterName Blocked List"
        Write-Host ""
        $nonfixeddomains.Add($Domain)
        Set-HostedContentFilterPolicy -Identity $SpamFilterName -BlockedSenderDomains $nonfixeddomains
            }
     
}
