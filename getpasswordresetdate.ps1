Connect-MsolService
get-msoluser -All | select-object DisplayName,userprincipalname,islicensed,LastPasswordChangeTimeStamp | Export-CSV c:\admin\LastPasswordChangeDate.csv -NoTypeInformation