
function Cleanup() {
    #Clean up by closing registry
    $HKEY_Users.Close()
    $remoteCURegKey.Close()
 }

 #Gets computername
 $computer = $env:computername

 #Set variables
 $outarrayop = @()
 
    #Used to get Outlook version
    $HKEY_LM = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine",$Computer)
    $OutKey = $HKEY_LM.OpenSubkey("SOFTWARE\Classes\Outlook.Application\CurVer")

    #Stop script if Outlook is not found on the machine.
    if($OutKey -eq $null) 
    {
        Write-Host "Outlook not found."
        $HKEY_LM.Close()
        Exit
    }
    else 
    {
        $version = $OutKey.GetValue("") #(Default)
        $HKEY_LM.Close()
    }

    #Used to get user SID's
    $HKEY_Users = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("Users",$Computer)
  
    #used to open Outlook Profile registry keys
    $remoteCURegKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("Users",$Computer)

    # Get list of SIDs
    $SIDs = $HKEY_Users.GetSubKeyNames() | Where-Object { ($_ -like "S-1-5-21*") -and ($_ -notlike "*_Classes") }

    # Associate SID with Username
    $TotalSIDs = ForEach ($SID in $SIDS) {
        Try {
            $SID = [system.security.principal.securityidentIfier]$SID
            $user = $SID.Translate([System.Security.Principal.NTAccount])
            New-Object PSObject -Property @{
                Name = $User.value
                SID = $SID.value
            }                 
        } 
        Catch {
            Write-Warning ("Unable to translate {0}.`n{1}" -f $UserName,$_.Exception.Message)
        }
    }

    $UserList = $TotalSIDs

    # Loop through user accounts
    ForEach($User in $UserList) 
    {

        # Get SID
        $UserSID = $User.SID

        #Open licensing location in registry based on Outlook version number.
        if ($version -eq "Outlook.Application.16")
        {
            $regKey = $remoteCURegKey.OpenSubkey("$UserSID\SOFTWARE\Microsoft\Office\16.0\Common\Licensing") 
            $ver ="16.0"

            #Office license type check
            $VerKey = $remoteCURegKey.OpenSubkey("$UserSID\SOFTWARE\Microsoft\Office\16.0\Common\Licensing\LastKnownC2RProductReleaseId") 
            $VerVal = $VerKey.GetValue("Word")

            if($VerVal -eq "HomeBusinessRetail")
            {
                Write-Host "Office Home and Business is not available through MS 365."
                #$VerKey = $remoteCURegKey.OpenSubkey("$UserSID\SOFTWARE\Microsoft\Office\ClickToRun\Configuration") 
                #$VerVal = $VerKey.GetValue("HomeBusinessRetail.EmailAddress")
                Cleanup
                Exit
            }
            #Write-Host "Outlook version 2016/19"
        }
        elseif ($version -eq "Outlook.Application.15")
        {
            #$regKey = $remoteCURegKey.OpenSubkey("$UserSID\SOFTWARE\Microsoft\Office\15.0\Common") 
            $ver ="15.0"
            Write-Host "Outlook version 2013 is not available through MS 365."
            Cleanup
            Exit
        }
        elseif ($version -eq "Outlook.Application.14")
        {
            #$regKey = $remoteCURegKey.OpenSubkey("$UserSID\SOFTWARE\Microsoft\Office\14.0\Common") 
            $ver ="14.0"
            Write-Host "Outlook version 2010 is not available through MS 365."
            Cleanup
            Exit
        }
        elseif ($version -eq "Outlook.Application.12")
        {
            #$regKey = $remoteCURegKey.OpenSubkey("$UserSID\SOFTWARE\Microsoft\Office\12.0\Common") 
            $ver ="12.0"
            Write-Host "Outlook version 2007 is not available through MS 365."
            Cleanup
            Exit
        }
        
        
        #Stop script if no Outlook Profiles are configured
        #if ($regKey.getsubkeynames().Count -lt 1) { Write-Host "Outlook version not supported or no profiles found."; Cleanup; Exit }

        try
        {
            #Open registry key that contains the licensed accounts.  It is possible to have more than one.
            $LicenseBase = $HKEY_Users.OpenSubKey("$UserSID\SOFTWARE\Microsoft\Office\$ver\Common\Licensing")
            $LicUserID = $LicenseBase.GetValue("NextUserLicensingLicensedUserIds")
            $UserIDs = $LicUserID.Split(",")

            foreach ($IDs in $UserIDs)
                {
                   
                    #Gets the email address for the matching ID.
                    $LicensedEmail = $HKEY_Users.OpenSubKey("$UserSID\SOFTWARE\Microsoft\Office\$ver\Common\Licensing\LicensingNext\LicenseIdToEmailMapping")
                    $LicensedEmailAddr = $LicensedEmail.GetValue($IDs)
                    #$LicensedEmailAddr
                    
                    #Create custom object
                    $comGroupObjOP = New-Object System.Object
                    $comGroupObjOP | Add-Member -MemberType NoteProperty -Name "Computer" -Value $Computer
                    $comGroupObjOP | Add-Member -MemberType NoteProperty -Name "Office Version" -Value $Ver
                    $comGroupObjOP | Add-Member -MemberType NoteProperty -Name "Licensed Email Address" -Value $LicensedEmailAddr
                    $outarrayop += $comGroupObjOP

                }    
            
        }

        catch
        {
            
            $comGroupObjOP = New-Object System.Object
                    $comGroupObjOP | Add-Member -MemberType NoteProperty -Name "Computer" -Value $Computer
                    $comGroupObjOP | Add-Member -MemberType NoteProperty -Name "Office Version" -Value $Ver
                    $comGroupObjOP | Add-Member -MemberType NoteProperty -Name "Licensed Email Address" -Value "Unknown Error"
                    $outarrayop += $comGroupObjOP

            $outarrayop
        }

        
    }

#Output information to screen        
$outarrayop

#Execute cleanup function   
Cleanup