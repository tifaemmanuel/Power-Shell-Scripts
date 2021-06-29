$dir = "c:\admin\alignment"
$report = Get-Mailbox -ResultSize Unlimited -Filter {RecipientTypeDetails -eq "UserMailbox"} | Select-Object Name, AuditEnabled
$report | Export-csv "$dir\Mailbox Auditing.csv" -NoTypeInformation
