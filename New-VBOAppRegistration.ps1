#Requires -Module AzureAD
<#
    .SYNOPSIS
      This Script will create the AzureAD Application "VBO" for use with iland/Veeam O365 backups for Exchange, Sharepoint And Onedrive
      This will assign the required permissions and create a client secret password. Permissions must be consented by admin after
      creation. 
      Written By David Just
    
 #>
Connect-AzureAD

New-AzureADApplication -DisplayName "VBO"

$serviceprincipalEXO = Get-AzureADServicePrincipal -All $true | where displayname -eq "Office 365 Exchange Online"
$serviceprincipalSPO = Get-AzureADServicePrincipal -All $true | where displayname -eq "Office 365 SharePoint Online"
$serviceprincipalGRAPH = Get-AzureADServicePrincipal -All $true | where displayname -eq "Microsoft Graph"

$EXO = New-Object -TypeName "Microsoft.Open.AzureAD.Model.RequiredResourceAccess"
$EXO.ResourceAppId = $serviceprincipalEXO.AppId

$SPO = New-Object -TypeName "Microsoft.Open.AzureAD.Model.RequiredResourceAccess"
$SPO.ResourceAppId = $serviceprincipalSPO.AppId

$GRAPH = New-Object -TypeName "Microsoft.Open.AzureAD.Model.RequiredResourceAccess"
$GRAPH.ResourceAppId = $serviceprincipalGRAPH.AppId

#Exchange Permissions
$delPermission1 = New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList "dc890d15-9560-4a4c-9b7f-a736ec74ec40","Role" ##Exchange #full_access_as_app App
$delPermission2 = New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList "3b5f3d61-589b-4a3c-a359-5dd4b5ee5bd5","Scope" ##Exchange EWS.AccessAsUser.All Delegated 
#Sharepoint Permissions
$delPermission3 = New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList "82866913-39a9-4be7-8091-f4fa781088ae","Scope" ##Sharepoint User.ReadWrite.All Delegated
$delPermission4 = New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList "741f803b-c850-494e-b5df-cde7c675a1ca","Role" ##Sharepoint User.ReadWrite.All App
$delPermission5 = New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList "678536fe-1083-478a-9c59-b99265e6b0d3","Role" ##Sites.FullControl.All 
$delPermission6 = New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList "56680e0d-d2a3-4ae1-80d8-3c4f2100e3d0","Scope" ##AllSitesFullControl 
#Microsoft Graph Permissions
$delPermission7 = New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList "7427e0e9-2fba-42fe-b0c0-848c9e6a8182","Scope" #Graph offline_access Delegated
$delPermission8 = New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList "7ab1d382-f21e-4acd-a863-ba3e13f7da61","Role" #Graph #Directory.Read.All Delegated Application
$delPermission9 = New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList "62a82d76-70ea-41e2-9197-370581804d09","Role" ##Group.ReadWrite.All App
$delPermission10 = New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList "5b567255-7703-4780-807c-7be8301ae99b","Role" ##Group.Read.All Delegated
$delPermission11 = New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList "06da0dbc-49e2-44d2-8312-53f166ab848a","Scope" #Directory.Read.All Delegated
$delPermission12 = New-Object -TypeName "Microsoft.Open.AzureAD.Model.ResourceAccess" -ArgumentList "9492366f-7969-46a4-8d15-ed1a20078fff","Role" #Sites.ReadWrite.All Application

$EXO.ResourceAccess = $delPermission1,$delPermission2
$SPO.ResourceAccess = $delPermission3,$delPermission4,$delPermission5,$delPermission6
$GRAPH.ResourceAccess = $delPermission7,$delPermission8,$delPermission9,$delPermission10,$delPermission11,$delPermission12
$ADApplication = Get-AzureADApplication -All $true | ? { $_.Displayname -match "VBO"}
 
Set-AzureADApplication -ObjectId $ADApplication.ObjectId -RequiredResourceAccess $GRAPH,$EXO,$SPO

Write-Output "Generating Application Client Secret. Output will be below"
Sleep 2
$VBO = Get-AzureADApplication | where displayname -match "VBO"
$clientsecret = New-AzureADApplicationPasswordCredential -ObjectId $VBO.ObjectID -EndDate ((Get-Date).AddYears(100))
Write-Output "Client Secret:" 
$clientsecret.value
Write-Host "Please visit https://portal.azure.com/#blade/Microsoft_AAD_B2CAdmin/TenantManagementMenuBlade/registeredApps and give Admin Consent on VBO"
Pause
