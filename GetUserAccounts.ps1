import-module ActiveDirectory -ErrorAction SilentlyContinue
Get-ADUser -Filter * -Properties * | Select-Object -Property Name,SamAccountName,Enabled,UserPrincipalName,LastLogonDate | Export-CSV "C:\admin\AllADUsers.csv" -NoTypeInformation -Encoding UTF8

#import-module ActiveDirectory -ErrorAction SilentlyContinue
#Get-ADUser -Filter * -Properties lastlogondate | Select-Object Name,LastLogonDate,Enabled,SamAccountName | Sort-Object LastLogonDate