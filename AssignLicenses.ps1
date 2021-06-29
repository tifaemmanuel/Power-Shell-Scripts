#needs installation of AzureAD module.  install-module AzureAD
#To get the planname you can run get-azureadsubscribedsku | FL and look for SkuPartNumber
$AzureAdCred = Get-Credential
Connect-AzureAD -Credential $AzureAdCred
$planName="ENTERPRISEPACKPLUS_STUUSEBNFT"
$License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
$License.SkuId = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $planName -EQ).SkuID
$LicensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
$LicensesToAssign.AddLicenses = $License
$inputFile = Import-CSV  C:\scripts\365\newstudents.csv
foreach($line in $inputFile){
$upn = $line.uniqueid+"@student.lasallehighschool.com"
$user = Get-AzureADUser -ObjectId $upn
Set-AzureADUser -ObjectId $user.ObjectId -UsageLocation "US"
Set-AzureADUserLicense -ObjectId $upn -AssignedLicenses $LicensesToAssign
}