param(
    [Parameter(Mandatory=$True)][string]$newServiceName
)

Write-Host "Creating new Tier 1 service $newServiceName.`n"	

.\Manage-AdObjects.ps1 -Restore -OU -Group -Permissions -GPO All -GpoLinks LinkEnabled -SettingsXml settings.xml

## Add 1 minute break before next step to allow replication to complete.
Start-Sleep -Seconds 60
## Add message with counter to announce the wait time. First announce that we have to wait 60 sec and then, every 15 sec write updated waiting time
Write-Host "Waiting for replication to complete. 60 seconds remaining."
for ($i=45; $i -ge 0; $i--) {
    if ($i % 15 -eq 0) {Write-Host "Waiting for replication to complete. $i seconds remaining."}
    Start-Sleep -Seconds 1
}

Write-Host "Replication complete. Proceeding with new service creation.`n"

# $newServiceName = "myBrandNewService"
$newGPOName = "AllowT1"+$newServiceName+"Access"
$newSubGPOName = "AllowT1"+$newServiceName+"SubServiceAccess"
$svcOUs = Get-ADOrganizationalUnit -Filter 'Name -like "T1-newservice*"' | Sort-Object -Descending

$AdminGRPName = "Tier 1 $newServiceName Admins"
$AdminGrpSAM = $AdminGRPName -replace " ",""
$AdminGrpDesc = "Members of this group administer Tier 1 $newservicename servers."

$SubAdminGRPName = "Tier 1 $newServiceName Admins - SubService"
$SubAdminGrpSAM = $SubAdminGRPName -replace " ",""
$SubAdminGrpDesc = "Members of this group administer Tier 1 $newservicename - SubService servers."

$SvcGrpName = "Tier 1 $newServiceName Service Accounts"
$SvcGrpSAM = $SvcGrpName -replace " ",""
$SvcGrpDesc = "Member of this group are Tier 1 $newServiceName service accounts."

$SrvGrpName = "Tier 1 $newServiceName Servers"
$SrvGrpSAM = $SrvGrpName -replace " ",""
$SrvGrpDesc = "Members of this group are Tier 1 $newServiceName servers."

rename-gpo CR-AllowT1NewServiceAccess -TargetName $newGPOName
rename-gpo CR-AllowT1NewServiceSubServiceAccess -TargetName $newSubGPOName

foreach ($svcOU in $svcOUs) {
    $newOUName = $svcOU.name -replace "NewService",$newServiceName
    Rename-ADObject -Identity $svcOU -NewName $newOUName
}

$newAdminGrp = Get-ADGroup -Identity Tier1NewServiceAdmins
$newSubAdminGrp = Get-ADGroup -Identity Tier1NewServiceAdmins-SubService
$newSvcGrp = Get-ADGroup -Identity Tier1NewServiceServiceAccounts
$newSrvGrp = Get-ADGroup -Identity Tier1NewServiceServers

rename-ADObject -Identity $newAdminGrp.distinguishedname -NewName $AdminGRPName
set-ADGroup -Identity Tier1NewServiceAdmins -Description $AdminGrpDesc -SamAccountName $AdminGrpSAM

rename-ADObject -Identity $newSubAdminGrp.distinguishedname -NewName $SubAdminGRPName
set-ADGroup -Identity Tier1NewServiceAdmins-SubService -Description $SubAdminGrpDesc -SamAccountName $SubAdminGrpSAM

rename-ADObject -Identity $newSvcGrp.distinguishedname -NewName $SvcGrpName
set-ADGroup -Identity Tier1NewServiceServiceAccounts -Description $SvcGrpDesc -SamAccountName $SvcGrpSAM

rename-ADObject -Identity $newSrvGrp.distinguishedname -NewName $SrvGrpName
set-ADGroup -Identity Tier1NewServiceServers -Description $SrvGrpDesc -SamAccountName $SrvGrpSAM

Add-ADGroupMember -Identity Tier1Accounts -Members $AdminGrpSAM,$SubAdminGRPSAM,$SvcGrpSAM



$admins = get-adgroupmember $AdminGrpSAM
$subadmins = Get-ADGroupMember $SubAdminGrpSAM

foreach ($admin in $admins) {Set-ADUser -Identity $admin.SamAccountName -AuthenticationPolicy $authA.name}
foreach ($subadmin in $subadmins) {Set-ADUser -Identity $subadmin.SamAccountName -AuthenticationPolicy $authA.name}

$svcaccs = get-adgroupmember $SvcGrpSAM

foreach ($svcacc in $svcaccs) {Set-ADUser -Identity $svcacc.SamAccountName -AuthenticationPolicy $authS.name}