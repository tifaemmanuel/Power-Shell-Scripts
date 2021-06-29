
     function Cleanup() {
        #Clean up by closing registry
        $HKEY_Users.Close()
        $remoteCURegKey.Close()
     }
     function OutlookProfileLoop()
     {
        $outarrayopt = @()
        $ascii0 = [char]0

         #Stop script if no Outlook Profiles are configured
         if ($regKey.getsubkeynames().Count -lt 1)
         { 
             Write-Host "Outlook version not supported or no profiles found.";
             Cleanup
         }
         else
         {
             
             #Loop through all Outlook profiles.
             foreach ($subName in ($regKey.getsubkeynames()))
             {

                 #Open registry for each Outlook profile
                 $OutlookProfiles = $HKEY_Users.OpenSubKey("$UserSID\SOFTWARE\Microsoft\Office\$ver\Outlook\Profiles\$subName")

                 # Loop Outlook Profile settings
                 ForEach($Profile in ($OutlookProfiles.GetSubKeyNames()))
                 {
             
                     #Open registry key that contains Outlook profile settings
                     $ProfileKey = $HKEY_Users.OpenSubKey("$UserSID\SOFTWARE\Microsoft\Office\$ver\Outlook\Profiles\$subName\$Profile")
     
                     #Gets all info for cache settings
                     If(($ProfileKey.GetValueNames() -contains "00036601") -eq $True)
                     {
                         $Result = $ProfileKey.GetValue("00036601") #Enabled or not
                         # Convert value to HEX
                         $Result = [System.BitConverter]::ToString($Result)

                         try 
                         {                         
                                $Result2 = $ProfileKey.GetValue("00036649") #Amount cached for greater than 3 months
                                # Convert value to HEX
                                $Result2 = [System.BitConverter]::ToString($Result2)
                         }
                        catch {
                            
                        }
                         
                         

                         #0003665a does not exist in Outlook 2013
                         if ($ver -eq "16.0")
                         {
                             try {
                                $Result3 = $ProfileKey.GetValue("0003665a") #Amount cached for less than 3 months
                                $Result3 = [System.BitConverter]::ToString($Result3)
                             }
                             catch {
                                $Result3 = "00-00-00"     
                             }
                         }
                         elseif ($ver -eq "15.0")
                         {
                            $Result3 = "00-00-00" 
                         }

                         $Result4 = $ProfileKey.GetValue("001e6750") #Profile name
                         #Email account name stored in Binary so we need to convert it.
                         $Result5 = ($ProfileKey.GetValue("001f3001") | ForEach-Object{ [char]$_ }) -join "" -replace $ascii0   
                         
                         # Determine if cache mode is enabled
                         If($Result -like "8*") {
                             $CacheMode = "Enabled"
                         }
                         Else {
                             $CacheMode = "Disabled"
                             $CacheSize = "Disabled"
                         }

                         #Get duration for cache mode
                         If($Result2 -like "01*") { $CacheSize = "1 month"}
                         If($Result2 -like "03*")
                         {
                             $CacheSize = "3 months"

                             #Uncomment to change sync duration.
                             #$OpenKey = $HKEY_Users.CreateSubKey("$UserSID\SOFTWARE\Microsoft\Office\$ver\Outlook\Profiles\$subName\$Profile")
                             #$OpenKey.SetValue("00036649", ([byte[]](0x06,0x00,0x00,0x00)),[Microsoft.Win32.RegistryValueKind]::Binary)
                             #$OpenKey.Close();

                         }
                         If($Result2 -like "06*") { $CacheSize = "6 months"}          
                         If($Result2 -like "0C*") { $CacheSize = "1 year"}
                         If($Result2 -like "18*") { $CacheSize = "2 years"}
                         If($Result2 -like "24*") { $CacheSize = "3 years"}
                         If($Result2 -like "3C*") { $CacheSize = "5 years"}
                         If($Result2 -like "00*" -and $Result3 -like "03*") {$CacheSize = "3 days"}
                         If($Result2 -like "00*" -and $Result3 -like "07*") {$CacheSize = "1 week"}
                         If($Result2 -like "00*" -and $Result3 -like "0E*") {$CacheSize = "2 weeks"}
                         If($Result2 -like "00*" -and $Result3 -like "00*") {$CacheSize = "All"}
                         #If($Result2 -like "FF*" -and $Result3 -like "00*") {$CacheSize = "Disabled"}
                         #If($Result2 -like "00*" -and $Result3 -like "FF*") {$CacheSize = "Disabled"}
                         if ($CacheSize -like "Disabled")
                         {
                             #Uncomment to enable cache and set to 6 months
                             #$OpenKey = $HKEY_Users.CreateSubKey("$UserSID\SOFTWARE\Microsoft\Office\$ver\Outlook\Profiles\$subName\$Profile")
                             #$OpenKey.SetValue("00036649", ([byte[]](0x06,0x00,0x00,0x00)),[Microsoft.Win32.RegistryValueKind]::Binary) #6 Months
                             #$OpenKey.SetValue("0003665a", ([byte[]](0x00,0x00,0x00,0x00)),[Microsoft.Win32.RegistryValueKind]::Binary) #We dont want anything set here
                             #$OpenKey.SetValue("00036601", ([byte[]](0x80,0x00,0x00,0x00)),[Microsoft.Win32.RegistryValueKind]::Binary) #Enable
                             #$OpenKey.Close();
                             
                         }

                         # Create custom object
                         $comGroupObjOPt = New-Object System.Object
                             $comGroupObjOPt | Add-Member -MemberType NoteProperty -Name "Username" -Value $User.Name
                             $comGroupObjOPt | Add-Member -MemberType NoteProperty -Name "Profile Name" -Value $Result4
                             $comGroupObjOPt | Add-Member -MemberType NoteProperty -Name "Email Account" -Value $Result5
                             $comGroupObjOPt | Add-Member -MemberType NoteProperty -Name "Cache Mode" -Value $CacheMode
                             $comGroupObjOPt | Add-Member -MemberType NoteProperty -Name "Cache Size" -Value $CacheSize
                             $outarrayopt += $comGroupObjOPt

                         
                     }
                    
                 }

             }
             $outarrayopt # | export-csv $dir\OutlookProfiles.csv -NoTypeInformation
         }
         
     }
     #Gets computername
     $computer = $env:computername

     #Set variables
     #$outarrayop = @()
     #$ascii0 = [char]0
     
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
            Write-Host "Outlook found."
            $version = $OutKey.GetValue("") #(Default)
            $HKEY_LM.Close()
        }

        #Used to get user SID's
        $HKEY_Users = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("Users",$Computer)
        
        #used to open Outlook Profile registry keys
        $remoteCURegKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("Users",$Computer)
        
        # Get list of SIDs
        $SIDs = $HKEY_Users.GetSubKeyNames() | Where-Object { ($_ -like "S-1-5-21*") -and ($_ -notlike "*_Classes") -or ($_ -like "S-1-12-1*") -and ($_ -notlike "*_Classes") }

        # Associate SID with Username
        $TotalSIDs = ForEach ($SID in $SIDS)
        {
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
    
        # Loop through user account
        ForEach($User in $UserList)
        {

            # Get SID
            $UserSID = $User.SID

            #Open Profile location in registry based on Outlook version number.
            if ($version -eq "Outlook.Application.16")
            {
                $regKey = $remoteCURegKey.OpenSubkey("$UserSID\SOFTWARE\Microsoft\Office\16.0\Outlook\Profiles") 
                $ver ="16.0"
                Write-Host "Outlook version 2016/19"
                OutlookProfileLoop
            }
            elseif ($version -eq "Outlook.Application.15")
            {
                $regKey = $remoteCURegKey.OpenSubkey("$UserSID\SOFTWARE\Microsoft\Office\15.0\Outlook\Profiles") 
                $ver ="15.0"
                Write-Host "Outlook version 2013"
                OutlookProfileLoop
            }
            elseif ($version -eq "Outlook.Application.14")
            {
                $regKey = $remoteCURegKey.OpenSubkey("$UserSID\SOFTWARE\Microsoft\Office\15.0\Outlook\Profiles") 
                $ver ="14.0"
                Write-Host "Outlook version 2010 does not use an offline cache."
                Cleanup
                Exit
            }
            elseif ($version -eq "Outlook.Application.12")
            {
                $regKey = $remoteCURegKey.OpenSubkey("$UserSID\SOFTWARE\Microsoft\Office\15.0\Outlook\Profiles") 
                $ver ="12.0"
                Write-Host "Outlook version 2007 does not use an offline cache"
                Cleanup
                Exit
            }
            
            
            #Stop script if no Outlook Profiles are configured
            try 
            {
                if ($regKey.getsubkeynames().Count -lt 1)
                {
                    Write-Host "Outlook version not supported or no profiles found."
                    Cleanup
                    Exit
                }
            }
            Catch
            {
                Write-Host "Outlook version not supported or no profiles found."
                Cleanup
                Exit
            }
            
            
        }

    #Output information to screen        
    #$outarrayop
 
 #Execute cleanup function   
 Cleanup