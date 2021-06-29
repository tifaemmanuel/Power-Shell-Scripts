$AuthenticationPolicy = Get-OrganizationConfig | Select-Object DefaultAuthenticationPolicy

If (-not $AuthenticationPolicy.Identity) {
$AuthenticationPolicy = New-AuthenticationPolicy "Block Basic Auth";
Set-OrganizationConfig -DefaultAuthenticationPolicy $AuthenticationPolicy.Identity
}

Set-AuthenticationPolicy -Identity $AuthenticationPolicy.Identity -AllowBasicAuthActiveSync:$true -AllowBasicAuthAutodiscover:$true -AllowBasicAuthImap:$false -AllowBasicAuthMapi:$true -AllowBasicAuthOfflineAddressBook:$true -AllowBasicAuthOutlookService:$true -AllowBasicAuthPop:$false -AllowBasicAuthPowershell:$false -AllowBasicAuthReportingWebServices:$true -AllowBasicAuthRpc:$true -AllowBasicAuthSmtp:$false -AllowBasicAuthWebServices:$true

Get-User -ResultSize Unlimited | ForEach-Object { Set-User -Identity $_.Identity -AuthenticationPolicy $AuthenticationPolicy.Identity -STSRefreshTokensValidFrom $([System.DateTime]::UtcNow) }