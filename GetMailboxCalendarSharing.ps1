$dir = "c:\admin\alignment"
Get-Mailbox -RecipientTypeDetails UserMailbox | ForEach-Object {Get-MailboxFolderPermission $($_.Identity + ":\Calendar")} |
     Select-Object Identity,User | Export-Csv -NoTypeInformation "$dir\Calendar Sharing.csv"
