$dir = "c:\admin\alignment"

get-mailbox -ResultSize Unlimited | Select-Object DisplayName,UserPrincipalName,ForwardingAddress,ForwardingSmtpAddress,DeliverToMailboxAndForward,@{Name='GrantSendOnBehalfTo';expression={[string]($_.GrantSendOnBehalfTo |
      ForEach-Object {$_.tostring().split("/")[-1]})}} | Export-Csv -NoTypeInformation "$dir\Mailbox Forwarding.csv"
