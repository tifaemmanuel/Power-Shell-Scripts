$report2 = @()
$dir = "c:\admin\"

#Distribution groups
$Groups = Get-DistributionGroup
$Groups | ForEach-Object {
$group = $_.Name
$groupemail = $_.PrimarySMTPAddress
Get-DistributionGroupMember $group | ForEach-Object {
      New-Object -TypeName PSObject -Property @{
       Group = $group
       GroupEmail = $groupemail
       Member = $_.Name
       EmailAddress = $_.PrimarySMTPAddress
       RecipientType= $_.RecipientType
}}} | Export-CSV "$dir\DistributionGroupMembers.csv" -NoTypeInformation -Encoding UTF8


#MS 365 Groups
$groups=Get-UnifiedGroup

foreach($group in $groups)
{

    $membersOfGroup = Get-UnifiedGroupLinks -Identity $group.Identity -LinkType Members

    foreach($member in $membersOfGroup)
    {
        $recip = Get-Recipient -Identity $member.Name | Select-Object PrimarySmtpAddress

        $comGroupObj3 = New-Object System.Object
                        $comGroupObj3 | Add-Member -MemberType NoteProperty -Name Group -Value $group.DisplayName
                        $comGroupObj3 | Add-Member -MemberType NoteProperty -Name GroupEmail -Value $group.primarySMTPAddress
                        $comGroupObj3 | Add-Member -MemberType NoteProperty -Name Member -Value $member.Name
                        $comGroupObj3 | Add-Member -MemberType NoteProperty -Name EmailAddress -Value $recip.primarySMTPAddress
                        $report2 += $comGroupObj3
    }

$report2 | Export-csv  "$dir\UnifiedGroupMembers.csv" -NoTypeInformation

}