$domains = Get-AcceptedDomain

$mailboxes = Get-Mailbox

foreach ($mailbox in $mailboxes) {

    $forwardingRules = $null

    Write-Host "Checking rules for $($mailbox.displayname) - $($mailbox.primarysmtpaddress)"
    $rules = get-inboxrule -Mailbox $mailbox.primarysmtpaddress
    
    $forwardingRules = $rules | Where-Object {$_.forwardto -or $_.forwardasattachmentto}

    foreach ($rule in $forwardingRules) {
        $recipients = @()
        $recipients = $rule.ForwardTo | Where-Object {$_ -match "SMTP"}
        $recipients += $rule.ForwardAsAttachmentTo | Where-Object {$_ -match "SMTP"}
    
        $externalRecipients = @()

        foreach ($recipient in $recipients) {
            $email = ($recipient -split "SMTP:")[1].Trim("]")
            $domain = ($email -split "@")[1]

            if ($domains.DomainName -notcontains $domain) {
                $externalRecipients += $email
                $externalRecipients
            }    
        }

        if ($externalRecipients) {
            $extRecString = $externalRecipients -join ", "
            Write-Host "$($rule.Name) forwards to $extRecString" -ForegroundColor Yellow

            $ruleHash = $null
            $ruleHash = [ordered]@{
                PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                DisplayName        = $mailbox.DisplayName
                RuleName           = $rule.Name
                RuleDescription    = $rule.Description
                ExternalRecipients = $extRecString
            }
            $ruleObject = New-Object PSObject -Property $ruleHash

            $ruleObject | Export-Csv C:\admin\alignment\ForwardingRules.csv -NoTypeInformation -Append
        }

    }
}



