#Working Directory
$dir = "c:\admin\alignment"

#Create folder if it doesn't already exist
	if(!(Test-Path -Path $dir ))
	{
        New-Item -ItemType directory -Path $dir >$null 2>&1
        Write-Output "c:\admin\Alignment Folder created."
    }
    
Write-Output "Working directory is $dir"

#Functions
function ExportADUsers() {
    import-module ActiveDirectory -ErrorAction SilentlyContinue -ErrorVariable ADPSNotFound

    if($ADPSNotFound)
    {
        Write-Output "Active Directory Powershell module is not intalled, cannot continue."
        Exit 
    }   
    else
    {
        Get-ADUser -Filter * -Properties * | Select-Object -Property Name,SamAccountName,Enabled,UserPrincipalName,LastLogonDate | Where-Object {$_.enabled -like "true"} | Export-CSV "$dir\EnabledADUsers.csv" -NoTypeInformation -Encoding UTF8
    }
}
function ExportComputerAccounts() {
    import-module ActiveDirectory -ErrorAction SilentlyContinue -ErrorVariable ADPSNotFound

    if($ADPSNotFound)
    {
        Write-Output "Active Directory Powershell module is not intalled, cannot continue."
        Exit 
    }   
    else
    {
        Get-ADComputer -Filter * -Properties * | Select-Object -Property Name,DNSHostName,Enabled,LastLogonDate | Where-Object {$_.enabled -like "true"} | Export-CSV "$dir\AllADComputers.csv" -NoTypeInformation -Encoding UTF8
    }
}
function ExportReplicationType() {
    import-module ActiveDirectory -ErrorAction SilentlyContinue -ErrorVariable ADPSNotFound

    if($ADPSNotFound)
    {
        Write-Output "Active Directory Powershell module is not intalled, cannot continue."
        Exit 
    }   
    else
    {
        $currentDomain =(Get-ADDomainController).hostname

        $defaultNamingContext = (([ADSI]"LDAP://$currentDomain/rootDSE").defaultNamingContext)
        $searcher = New-Object DirectoryServices.DirectorySearcher
        $searcher.Filter = "(&(objectClass=computer)(dNSHostName=$currentDomain))"
        $searcher.SearchRoot = "LDAP://" + $currentDomain + "/OU=Domain Controllers," + $defaultNamingContext
        $dcObjectPath = $searcher.FindAll() | ForEach-Object {$_.Path}

        # DFSR
        $searchDFSR = New-Object DirectoryServices.DirectorySearcher
        $searchDFSR.Filter = "(&(objectClass=msDFSR-Subscription)(name=SYSVOL Subscription))"
        $searchDFSR.SearchRoot = $dcObjectPath
        $dfsrSubObject = $searchDFSR.FindAll()

        if ($dfsrSubObject -ne $null) {

            #[pscustomobject]@{
            #    "SYSVOL Replication Mechanism"= "DFSR"
            #    "Path:"= $dfsrSubObject| ForEach-Object {$_.Properties."msdfsr-rootpath"}
            $rep = "SYSVOL Replication Mechanism is DFSR"
            #}

        }

        # FRS
        $searchFRS = New-Object DirectoryServices.DirectorySearcher
        $searchFRS.Filter = "(&(objectClass=nTFRSSubscriber)(name=Domain System Volume (SYSVOL share)))"
        $searchFRS.SearchRoot = $dcObjectPath
        $frsSubObject = $searchFRS.FindAll()

        if($frsSubObject -ne $null){

            #[pscustomobject]@{
            #    "SYSVOL Replication Mechanism" = "FRS"
            #    "Path" = $frsSubObject| ForEach-Object {$_.Properties.frsrootpath}
            $rep = "SYSVOL Replication Mechanism is FRS"
            #}

        }
        $rep
        $rep | Out-File "$dir\ReplicationType.txt"
    }
}
function ExportADVersion() {
    #Windows 2000 Server 		13
    #Windows Server 2003 		30
    #Windows Server 2003 R2 	31
    #Windows Server 2008 		44
    #Windows Server 2008 R2 	47
    #Windows Server 2012 		56
    #Windows Server 2012 R2 	69
    #Windows Server 2016 		87
    #Windows Server 2019 		88
    import-module ActiveDirectory -ErrorAction SilentlyContinue -ErrorVariable ADPSNotFound

    if($ADPSNotFound)
    {
        Write-Output "Active Directory Powershell module is not intalled, cannot continue."
        Exit 
    }   
    else
    {
        $report2 = @()

        #Get Schema version
        $ver = Get-ADObject (Get-ADRootDSE).schemaNamingContext -Property objectVersion | Select-Object -ExpandProperty objectVersion
        switch ($ver) {
            "13" { $val = "Windows 2000 Server" }
            "30" { $val = "Windows Server 2003" }
            "31" { $val = "Windows Server 2003 R2" }
            "44" { $val = "Windows Server 2008" }
            "47" { $val = "Windows Server 2008 R2" }
            "56" { $val = "Windows Server 2012" }
            "69" { $val = "Windows Server 2012 R2" }
            "87" { $val = "Windows Server 2016" }
            "88" { $val = "Windows Server 2019" }
            Default { $val = "Unknown Schema Version: " + $ver}
        }

            #Get Domain Functionality Level
            $dom = Get-ADDomain | Select-Object -ExpandProperty DomainMode

            #Get Forest Functionality Level
            $func = Get-ADForest | Select-Object -ExpandProperty ForestMode

            #Build Object for file export
            $comGroupObj3 = New-Object System.Object
                        $comGroupObj3 | Add-Member -MemberType NoteProperty -Name Schema -Value $val
                        $comGroupObj3 | Add-Member -MemberType NoteProperty -Name DomainLevel -Value $dom
                        $comGroupObj3 | Add-Member -MemberType NoteProperty -Name ForestLevel-Value $func
                        $report2 += $comGroupObj3

            #Output to screen
            $val
            $dom
            $func

            #Output to file
            $report2 | Export-csv  "$dir\ADVersion.csv" -NoTypeInformation
    }
}
function ExportLocalAdmins() {

    import-module ActiveDirectory -ErrorAction SilentlyContinue -ErrorVariable ADPSNotFound

    if($ADPSNotFound)
    {
        Write-Output "Active Directory Powershell module is not intalled, cannot continue."
        Exit 
    }   
    else
    {

        #Uncomment and set $ou if you want to narrow down the searchbase when finding computers.
        #If you don't set it, we will use the entire domain name for the searchbase.
        #$ou = New-Object System.Object
        #$ou | Add-Member -MemberType NoteProperty -Name DistinguishedName -Value "dc=mpcfl,dc=local"

        #Check if $ou is defined and if not, find domain name.
        if (-not $ou)
        {
            try {
                #Get AD Domain name
                $ou = Get-ADDomain -ErrorAction SilentlyContinue | Select-Object -Expand DistinguishedName
            }
            catch {
                Write-Output "Unable to continue, no domain was found."
                Exit
            }
        }   
        Write-Host "Found domain: " -NoNewLine
        Write-Host $ou

        $activeDays = 30

        $padVal = 22
        $adminLabel = "Admin Users".PadRight($padVal," ")
        $badLabel = "Total invalid Accounts".PadRight($padVal," ")
        $defaultExclude = @("Administrator","Domain Admins","BTAdmin","Remote Support","CM","BTImagePrep","ILAdmin","FTSAdmin","CWAdmin")
        $excludeNames += $defaultExclude
        $onlineLabel = "Online Computers".PadRight($padVal," ")
        $offlineLabel = "Offline Computers".PadRight($padVal," ")
        $activeLabel = "Computers Active Since".PadRight($padVal," ")
        #Write-Output "Local Admin Group membership audit tool. Searches computers for membership of Local Administrators group."
        #Write-Output "Accounts not listed in -excludeNames will be displayed."

        if($ou)
        {
            #$searchLabel = "Searching Computers in".PadRight($padVal," ")
            #Write-Output "$searchLabel : $ou"
            $now  = [datetime]::Now
            $searchDays = $now.AddDays(-$activedays)
            Write-Output "$activeLabel : $searchDays"
        }
        if($pcName)
        {
            #$searchLabel = "Searching Computer".PadRight($padVal," ")
            #Write-Output "$searchLabel : $pcName"
        }
        Write-Output "$adminLabel : $excludeNames"
        $compFoundLabel = "Computers Found".PadRight($padVal, " ")
        $badAccounts = 0
        $onlineComputer = 0
        $offlineComputer = 0
        if($ou)
        {
            $computers = Get-AdComputer -filter {LastLogonDate -ge $searchDays } -searchbase $ou -properties LastLogonDate -SearchScope Subtree | Sort-Object Name -Descending
        }
        if($pcName)
        {
            $computers = Get-AdComputer $pcName -properties *
        }
        $computerCount = ($computers | Measure-Object).Count
        $pCounter = 0
        $report = @()
        Write-Output "$compFoundLabel : $computerCount"
        foreach ($computer in $computers )
        {
            $pCounter++
            #Write-Progress -Activity "Searching.. Please Wait.." -Status $computer.Name -PercentComplete ($pCounter / $computerCount*100)
            try{
                $groupMembers = get-wmiobject win32_groupUser -ComputerName $computer.Name # -ErrorAction Stop
                $onlineComputer++
                $groupMembers = $groupMembers | Where-Object { $_.GroupComponent -like "*Administrators*"}
                foreach ($member in $groupMembers)
                {
                    $name = $member.PartComponent.Split("=")
                    $uName = $name[2].Replace('"',"")
                    $gName = $member.GroupComponent.Split("=")
                    $gName = $gName[2].Replace('"',"")
                    if(($excludeNames) -contains $uName)
                    {
                        # Skip
                    }
                    else
                    {
                        $badAccounts++
                        $comGroupObj = New-Object System.Object
                        $comGroupObj | Add-Member -MemberType NoteProperty -Name Name -Value $computer.Name
                        $comGroupObj | Add-Member -MemberType NoteProperty -Name UserName -Value $uName
                        $comGroupObj | Add-Member -MemberType NoteProperty -Name GroupName -Value $gName
                        $report += $comGroupObj
                        
                    }
                }
            }
            catch{
                $offlineComputer++
            }
        }
        Write-Output ""
        Write-Output "$badLabel : $badAccounts"
        Write-Output "$onlineLabel : $onlineComputer"
        Write-Output "$offlineLabel : $offlineComputer"
        $report | Export-Csv "$dir\LocalAdmins.csv" -NoTypeInformation
    }
}
function BIOSOptions() {
    #$Parameter1 = "true"
    param (
        [parameter (mandatory=$false)]
        [string]$Parameter1
    )
    import-module ActiveDirectory -ErrorAction SilentlyContinue -ErrorVariable ADPSNotFound

    if($ADPSNotFound)
    {
        Write-Output "Active Directory Powershell module is not intalled, cannot continue."
        Exit 
    }   
    else
    {

        #Uncomment and set $ou if you want to narrow down the searchbase when finding computers.
        #If you don't set it, we will use the entire domain name for the searchbase.
        #$ou = New-Object System.Object
        #$ou | Add-Member -MemberType NoteProperty -Name DistinguishedName -Value "dc=mpcfl,dc=local"

        #Check if $ou is defined and if not, find domain name.
        if (-not $ou)
        {
            try {
                #Get AD Domain name
                $ou = Get-ADDomain -ErrorAction SilentlyContinue | Select-Object -Expand DistinguishedName
            }
            catch {
                Write-Output "Unable to continue, no domain was found."
                Exit
            }
        }   
        Write-Host "Found domain: " -NoNewLine
        Write-Host $ou

        #Any computers listed here will be skipped.
        $arrayout = @()
        $defaultExclude = @("YRC-16,XX-XX-XX")
        $excludeNames += $defaultExclude

        #We limit our search for computers to only computers active in the last 30 days
        $activeDays = 30
        $now  = [datetime]::Now
        $searchDays = $now.AddDays(-$activedays)

        #Find computers in AD recursively.
        try {
        $computers = Get-AdComputer -filter {LastLogonDate -ge $searchDays} -searchbase $ou -properties Name -SearchScope Subtree | Select-Object -ExpandProperty Name | Sort-Object Name -Descending
        }
        catch {
            Write-Output "Error getting AD computers."
            Exit
        }

        #Get the amount of computers found
        $computerCount = ($computers | Measure-Object).Count
        Write-Output "Computers found: $computerCount"

        #Loop through all computers
        foreach ($computer in $computers)
        {
        if(($excludeNames) -Match $computer) #Seach the list of computers and see if any match our exlusion list.
            {
            # Skip computer
            Write-Output "$computer is excluded."
            }
            else
            {
                #Gets the manufacturer of the BIOS since settings are unique.
                $manufacturer = Get-WmiObject -Class Win32_BIOS -ComputerName $computer -ErrorAction SilentlyContinue -ErrorVariable ProcessError | Select-Object Manufacturer

                If ($ProcessError) #Computer is offline or has a firewall enabled.
                {
                    Write-Output "$computer is inacessible."
                    #If option R is used, we ouput it to file in the Default switch case.

                }
                switch -Regex ($manufacturer.Manufacturer)
                {

                    "Lenovo"
                    {

                        Write-Output "$computer has a Lenovo BIOS."

                        #If set to false, we dont modify the BIOS settings.
                        if ($Parameter1 -eq "true")
                        {
                            #Some options are different so lets list all of the settings so we can search them when needed.
                            $allsettings = Get-WmiObject -class Lenovo_BiosSetting -namespace root\wmi -Computer $computer | select-object InstanceName,CurrentSetting -ExpandProperty CurrentSetting

                            #Set BIOS settings
                            Write-Output "Setting Power options."
                            $getLenovoBIOS = Get-WmiObject -class Lenovo_SetBiosSetting -namespace root\wmi -ComputerName $computer
                            $getLenovoBIOS.SetBiosSetting("After Power Loss,Power On") | select-object -ExpandProperty return
                            $getLenovoBIOS.SetBiosSetting("Wake Up on Alarm,User Defined") | select-object -ExpandProperty return
                            $getLenovoBIOS.SetBiosSetting("Monday,Enabled") | select-object -ExpandProperty return
                            $getLenovoBIOS.SetBiosSetting("Tuesday,Enabled") | select-object -ExpandProperty return
                            $getLenovoBIOS.SetBiosSetting("Wednesday,Enabled") | select-object -ExpandProperty return
                            $getLenovoBIOS.SetBiosSetting("Thursday,Enabled") | select-object -ExpandProperty return
                            $getLenovoBIOS.SetBiosSetting("Friday,Enabled") | select-object -ExpandProperty return

                            #Search through all settings to see which version of 'User Defined Alarm Time' this BIOS has.
                            if (($allsettings) -match "UserDefinedAlarmTime")
                            {
                                $getLenovoBIOS.SetBiosSetting("UserDefinedAlarmTime,[23:00:00]") | select-object -ExpandProperty return
                            }
                            else
                            {
                                $getLenovoBIOS.SetBiosSetting("User Defined Alarm Time,[23:00:00]") | select-object -ExpandProperty return
                            }

                            $getLenovoBIOS.SetBiosSetting("Smart Power On,Enabled") | select-object -ExpandProperty return
                        
                            #Save BIOS settings
                            Write-Output "Saving BIOS settings."
                            $SaveLenovoBIOS = (Get-WmiObject -class Lenovo_SaveBiosSettings -namespace root\wmi -ComputerName $computer) 
                            $SaveLenovoBIOS.SaveBiosSettings() | select-object -ExpandProperty return

                            $comGroupObj2 = New-Object System.Object
                            $comGroupObj2 | Add-Member -MemberType NoteProperty -Name Computer -Value $computer
                            $comGroupObj2 | Add-Member -MemberType NoteProperty -Name Setting -Value "Settings changed"
                            $arrayout += $comGroupObj2
                        } 
                        else
                        {
                            #Output to file only, no setting changes.
                            $allsettings = Get-WmiObject -class Lenovo_BiosSetting -namespace root\wmi -Computer $computer | select-object InstanceName,CurrentSetting -ExpandProperty CurrentSetting | Where-Object {$_.CurrentSetting -match "Wake Up on Alarm"}
                            Write-Output "$computer : $allsettings"

                            $comGroupObj2 = New-Object System.Object
                            $comGroupObj2 | Add-Member -MemberType NoteProperty -Name Computer -Value $computer
                            $comGroupObj2 | Add-Member -MemberType NoteProperty -Name Setting -Value $allsettings
                            $arrayout += $comGroupObj2
                        }  
                    }
                
                    "Dell"
                    {
                        Write-Output "$computer has a Dell BIOS."
                        $allsettings = "Dell BIOS"
                        $comGroupObj2 = New-Object System.Object
                            $comGroupObj2 | Add-Member -MemberType NoteProperty -Name Computer -Value $computer
                            $comGroupObj2 | Add-Member -MemberType NoteProperty -Name Setting -Value $allsettings
                            $arrayout += $comGroupObj2

                    }
                    "American Megatrends"
                    {
                        Write-Output "$computer has an American Megatrends Inc. BIOS."
                        $allsettings = "American Megatrends Inc. BIOS"
                        $comGroupObj2 = New-Object System.Object
                            $comGroupObj2 | Add-Member -MemberType NoteProperty -Name Computer -Value $computer
                            $comGroupObj2 | Add-Member -MemberType NoteProperty -Name Setting -Value $allsettings
                            $arrayout += $comGroupObj2
                    }
                    "Microsoft Corporation"
                    {
                        Write-Output "$computer is a Hyper-V VM"
                        $allsettings = "Hyper-V VM"
                        $comGroupObj2 = New-Object System.Object
                            $comGroupObj2 | Add-Member -MemberType NoteProperty -Name Computer -Value $computer
                            $comGroupObj2 | Add-Member -MemberType NoteProperty -Name Setting -Value $allsettings
                            $arrayout += $comGroupObj2
                    }
                    Default #Anything we don't have a case for, we output to the screen.
                    {
                        #Computer is accessible and returned a value.
                        if($manufacturer.Manufacturer)
                        {
                        Write-Host "$computer has a(n) " -NoNewLine
                        Write-Host $manufacturer.Manufacturer -NoNewLine
                        Write-Host " BIOS."
                        
                        $comGroupObj2 = New-Object System.Object
                            $comGroupObj2 | Add-Member -MemberType NoteProperty -Name Computer -Value $computer
                            $comGroupObj2 | Add-Member -MemberType NoteProperty -Name Setting -Value $manufacturer.Manufacturer
                            $arrayout += $comGroupObj2
                        }
                        #The computer is inacessible so we output it as inacessible.
                        else {
                            $comGroupObj2 = New-Object System.Object
                            $comGroupObj2 | Add-Member -MemberType NoteProperty -Name Computer -Value $computer
                            $comGroupObj2 | Add-Member -MemberType NoteProperty -Name Setting -Value "Inacessible"
                            $arrayout += $comGroupObj2
                        }
                    }
            
                }
                
                #Export to CSV file
                $arrayout | export-csv $dir\BIOS.csv -NoTypeInformation
            }
        }
    }
}

function WOLDiscovery()
{

#Get IP Address.  Only works if static IP is assigned.
$ip = Get-NetIPAddress| Select-Object IPAddress,PrefixOrigin | Where-Object {$_.PrefixOrigin -like "manual"} | Select-Object -ExpandProperty IPAddress
if (-not $ip)
{
    $subnet = Read-Host "Could not determine subnet.  Type in the first 3 octets of your IP scheme. Ex. 192.168.1"
    
}
else
{
$iparray = $ip.Split("[.]")
$subnet = $iparray[0] + "." + $iparray[1] + "." + $iparray[2]
}
Write-Host "Scanning subnet $subnet"
#Set variables
#$subnet = "192.168.1"
$wake = "false"

#You don't normally have to change these
$dir = "c:\admin\WOL"
$file = "c:\admin\WOL\mac.csv"
$outarray = @()


	#We ping IP addresses for discovery purposes.  Any responses will update our arp cache where get-netneighbor will query.
	if ($subnet -ne "xxx.xxx.xxx")
	{
		For ($i=1; $i -le 254; $i++)
		{
			test-connection "$subnet.$i" -Count 1 -ErrorAction SilentlyContinue
		}
	}
	
	#Gets known MAC addresses from client
	$macs = Get-NetNeighbor | Select-Object LinkLayerAddress

	#Create folder if it doesn't already exist
	if(!(Test-Path -Path $dir ))
	{
		New-Item -ItemType directory -Path $dir
	}

	#Confirm file exists
	if(Test-Path -LiteralPath $file)
	{
		#Checks to see how large the file is and deletes it if it gets too big.
		if ((Get-Item $file).length/1KB -lt 100)
		{
		
			#Import the file that stores known MAC addresses
			$inputFile = Import-CSV $file

			#Combine mac objects from file and Get-NetNeighbor
			$macs = $macs + $inputFile
		}
		else
		{
			Remove-Item $file
		}
	}

	foreach ($mac in $macs)
	{
		if ($mac.LinkLayerAddress)
		{

            #Check MAC address format since it varies depending on versions of Powershell
            if ($mac.LinkLayerAddress -match ":" -or $mac.LinkLayerAddress -match "-")
            {
                $macaddr = $mac.LinkLayerAddress
            }
            else
            {
                $macaddr = $mac.LinkLayerAddress
                #Regex that adds the ":" every 2 characters and then trims the remaining ":" at the end
                $macaddr = $macaddr -replace '(..(?!$))','$1:'
            }
			#Check -wake paramater and send magic packet if set to true
			if ($wake -eq "true")
			{
				Try
				{
					#Sends majic packet			
					$MacByteArray = $macaddr -split "[:-]" | ForEach-Object { [Byte] "0x$_"}
					[Byte[]] $MagicPacket = (,0xFF * 6) + ($MacByteArray  * 16)
					$UdpClient = New-Object System.Net.Sockets.UdpClient
					$UdpClient.Connect(([System.Net.IPAddress]::Broadcast),7)
					$UdpClient.Send($MagicPacket,$MagicPacket.Length)
					$UdpClient.Close()
				}
				catch
				{
					Write-Host "Error building and sending packet"
				}
			}
			#Builds object for CSV file
			$outobject = @{
			LinkLayerAddress = $macaddr}
			$Build = New-Object PSObject -Property $outobject
			$outarray += $Build
			
		}
	   
	}

	#Exports to CSV without duplicates
	$outarray | Sort-Object -Property LinkLayerAddress -Unique | Export-Csv $file -NoTypeInformation -Encoding UTF8
}
function ExportADUserNoExpire{
    import-module ActiveDirectory -ErrorAction SilentlyContinue -ErrorVariable ADPSNotFound

    if($ADPSNotFound)
    {
        Write-Output "Active Directory Powershell module is not intalled, cannot continue."
        Exit 
    }   
    else
    {
        get-aduser -filter * -properties Name, PasswordNeverExpires,Enabled | Where-Object { $_.passwordNeverExpires -eq "true" } |
         Where-Object {$_.enabled -like "true"} | Select-Object DistinguishedName,Name,Enabled |
          Export-csv $dir\PasswordNeverExpires.csv -NoTypeInformation
    }
}
function FirewallStatus{
    import-module ActiveDirectory -ErrorAction SilentlyContinue -ErrorVariable ADPSNotFound

    if($ADPSNotFound)
    {
        Write-Output "Active Directory Powershell module is not intalled, cannot continue."
        Exit 
    }   
    else
    {
         #Check if $ou is defined and if not, find domain name.
         if (-not $ou)
         {
             try {
                 #Get AD Domain name
                 $ou = Get-ADDomain -ErrorAction SilentlyContinue | Select-Object -Expand DistinguishedName
             }
             catch {
                 Write-Output "Unable to continue, no domain was found."
                 Exit
             }
         }   
         Write-Host "Found domain: " -NoNewLine
         Write-Host $ou

        #We limit our search for computers to only computers active in the last 30 days
        $activeDays = 30
        $now  = [datetime]::Now
        $searchDays = $now.AddDays(-$activedays)

        #We need credentials to set up WinRM
        $MyCredential = (get-credential)

        $outarrayfw = @()
        $outarraynp = @()
        #Find computers in AD recursively.
        try {
        $computers = Get-AdComputer -filter {LastLogonDate -ge $searchDays} -searchbase $ou -properties Name -SearchScope Subtree | Select-Object -ExpandProperty Name | Sort-Object Name -Descending
        }
        catch {
            Write-Output "Error getting AD computers."
            Exit
        }
                #Get the amount of computers found
                $computerCount = ($computers | Measure-Object).Count
                Write-Output "Computers found: $computerCount"
        
                #Loop through all computers
                foreach ($computer in $computers)
                {
                    #Windows Remote Management needs to be running to get firewall status.
                    Start-Service -InputObject $(Get-Service -Computer $computer -Name WinRM) #-ErrorAction SilentlyContinue
                    #invoke-command -computername $computer -Credential $MyCredential
                    #winrm create winrm/config/Listener?Address=*+Transport=HTTPS
                    
                    #Starts a session with the remote computer.
                    $cim = New-CimSession -ComputerName $computer -Credential $MyCredential -ErrorAction SilentlyContinue -ErrorVariable ProcessErrorCim 
                    
                    if ($ProcessErrorCim)
                    {
                        $ProcessErrorCim = $null
                        Write-Output "$computer is inacessible."
                    }
                    else 
                    {
                        
                        #$ListWithComputerNames = get-content "c:\list.txt"

                        #foreach($computer in $ListWithComputerNames){
                        #    invoke-command -computername $computer -Credential $MyCredential
                        #        winrm create winrm/config/Listener?Address=*+Transport=HTTPS
                        #    }
                        #}
                        Write-Output "Querying $computer..."

                        #Querys the computer using WinRM to get firewall profile status.
                        $fwstatus = Get-NetFirewallProfile -CimSession $cim -Profile Domain, Public, Private | Select-Object Name, Enabled #Export-csv $dir\PasswordNeverExpires.csv -NoTypeInformation
                        $netprofile = Get-NetAdapter -CimSession $cim -Physical | Get-NetConnectionProfile

                        foreach ($pc in $fwstatus)
                        {
                        #Builds object for firewall CSV file
                                $comGroupObjfw = New-Object System.Object
                                $comGroupObjfw | Add-Member -MemberType NoteProperty -Name Computer -Value $computer
                                $comGroupObjfw | Add-Member -MemberType NoteProperty -Name Name -Value $pc.name
                                $comGroupObjfw | Add-Member -MemberType NoteProperty -Name Enabled -Value $pc.enabled
                                $outarrayfw += $comGroupObjfw

                        

                        #$comGroupObjfw
                        
                        #Stopping this service on a server will generate a ticket so we will leave the service running.
                        #Stop-Service -InputObject $(Get-Service -Computer $computer -Name WinRM) #-ErrorAction SilentlyContinue
                
                        }
                        foreach ($pc in $netprofile)
                        {
                            #Builds object for for network profiles csv file
                            $comGroupObjnp = New-Object System.Object
                            $comGroupObjnp | Add-Member -MemberType NoteProperty -Name Computer -Value $computer
                            $comGroupObjnp | Add-Member -MemberType NoteProperty -Name Name -Value $pc.name
                            $comGroupObjnp | Add-Member -MemberType NoteProperty -Name InterfaceAlias -Value $pc.InterfaceAlias
                            $comGroupObjnp | Add-Member -MemberType NoteProperty -Name NetworkCategory -Value $pc.NetworkCategory
                            $outarraynp += $comGroupObjnp

                        #$comGroupObjnp
                        }
                        
                    }
                    $outarrayfw | export-csv $dir\FWStatus.csv -NoTypeInformation
                    $outarraynp | export-csv $dir\NetworkProfiles.csv -NoTypeInformation
                }
    }

}
function GetOutlookProfileDetail()
{

    function Cleanup()
     {
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
                         $Result2 = $ProfileKey.GetValue("00036649") #Amount cached for greater than 3 months

                         #0003665a does not exist in Outlook 2013
                         if ($ver -eq "16.0")
                         {
                             $Result3 = $ProfileKey.GetValue("0003665a") #Amount cached for less than 3 months
                         }
                         elseif ($ver -eq "15.0")
                         {
                         $Result3 = "00-00-00" 
                         }

                         $Result4 = $ProfileKey.GetValue("001e6750") #Profile name
                         #Email account name stored in Binary so we need to convert it.
                         $Result5 = ($ProfileKey.GetValue("001f3001") | ForEach-Object{ [char]$_ }) -join "" -replace $ascii0   

                         # Convert value to HEX
                         $Result = [System.BitConverter]::ToString($Result)
                         $Result2 = [System.BitConverter]::ToString($Result2)
                         if ($ver -eq "16.0") { $Result3 = [System.BitConverter]::ToString($Result3) }
                         
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
             
         }
         $outarrayopt | export-csv $dir\OutlookProfiles.csv -NoTypeInformation
     }

     import-module ActiveDirectory -ErrorAction SilentlyContinue -ErrorVariable ADPSNotFound

     if($ADPSNotFound)
     {
         Write-Output "Active Directory Powershell module is not intalled, cannot continue."
         Exit 
     }   
     else
     {
          #Check if $ou is defined and if not, find domain name.
          if (-not $ou)
          {
              try
              {
                  #Get AD Domain name
                  $ou = Get-ADDomain -ErrorAction SilentlyContinue | Select-Object -Expand DistinguishedName
              }
              catch {
                  Write-Output "Unable to continue, no domain was found."
                  Exit
              }
          }   
          Write-Host "Found domain: " -NoNewLine
          Write-Host $ou
     }

     #Gets computername
     #$computer = $env:computername
     

     #We limit our search for computers to only computers active in the last 30 days
     $activeDays = 30
     $now  = [datetime]::Now
     $searchDays = $now.AddDays(-$activedays)

       #Find computers in AD recursively.
       try
        {
            $computers = Get-AdComputer -filter {LastLogonDate -ge $searchDays} -searchbase $ou -properties Name -SearchScope Subtree | Select-Object -ExpandProperty Name | Sort-Object Name -Descending
        }
        catch 
        {
            Write-Output "Error getting AD computers."
            Exit
        }

        #Loop through all computers
        foreach ($Computer in $computers)
        {

            try 
            {
                           
                #Used to get Outlook version
                $HKEY_LM = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("LocalMachine",$Computer) 
                $OutKey = $HKEY_LM.OpenSubkey("SOFTWARE\Classes\Outlook.Application\CurVer")
                $version = $OutKey.GetValue("") #(Default)
                $HKEY_LM.Close()

            }
            catch
            {
                Write-Host "Unable to open Remote Registry for $Computer"
                $version = "Not Found"
                #$HKEY_LM.Close()
            }

            #Used to get user SID's
            $HKEY_Users = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("Users",$Computer)
        
            #used to open Outlook Profile registry keys
            $remoteCURegKey = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey("Users",$Computer)

            # Get list of SIDs
            $SIDs = $HKEY_Users.GetSubKeyNames() | Where-Object { ($_ -like "S-1-5-21*") -and ($_ -notlike "*_Classes") }
    
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
                    
                }
                elseif ($version -eq "Outlook.Application.12")
                {
                    $regKey = $remoteCURegKey.OpenSubkey("$UserSID\SOFTWARE\Microsoft\Office\15.0\Outlook\Profiles") 
                    $ver ="12.0"
                    Write-Host "Outlook version 2007 does not use an offline cache"
                    Cleanup
                    
                }
                elseif ($version -eq "Not Found")
                {   
                    #skip looping through outlook profiles.
                }
            }    

        }

    Cleanup

}
#Prompts
do { $myInput = (Read-Host 'Export all enabled AD User Accounts? (Y/N)').ToLower() } while ($myInput -notin @('y','n'))
if ($myInput -eq 'y') {
ExportADUsers
} else {
Write-Output "Skipped"
}
do { $myInput = (Read-Host 'Export all enabled AD Computer Accounts? (Y/N)').ToLower() } while ($myInput -notin @('y','n'))
if ($myInput -eq 'y') {
ExportComputerAccounts
} else {
    Write-Output "Skipped"
}
do { $myInput = (Read-Host 'Export all enabled AD User Accounts that have password set to never expire? (Y/N)').ToLower() } while ($myInput -notin @('y','n'))
if ($myInput -eq 'y') {
ExportADUserNoExpire
} else {
    Write-Output "Skipped"
}
do { $myInput = (Read-Host 'Export replication type? (Y/N)').ToLower() } while ($myInput -notin @('y','n'))
if ($myInput -eq 'y') {
ExportReplicationType
} else {
    Write-Output "Skipped"
}
do { $myInput = (Read-Host 'Export Schema,Domain and Forest versions? (Y/N)').ToLower() } while ($myInput -notin @('y','n'))
if ($myInput -eq 'y') {
ExportADVersion
} else {
    Write-Output "Skipped"
}
do { $myInput = (Read-Host 'Export all local admins? (Y/N)').ToLower() } while ($myInput -notin @('y','n'))
if ($myInput -eq 'y') {
ExportLocalAdmins
} else {
    Write-Output "Skipped"
}
do { $myInput = (Read-Host 'Set all BIOS power options for all AD computers?  (F) will output to file only. (Y/N/F)').ToLower() } while ($myInput -notin @('y','n','f'))
if ($myInput -eq 'y') {
BIOSOptions("true")
} elseif ($myInput -eq 'f') {
    Write-Output "Outputting current Wake Up on Alarm setting to file only, no setting changes."
    BIOSOptions("false")
} else {
    Write-Output "Skipped"
}
do { $myInput = (Read-Host 'Perform WOL Discovery, will take a very long time? (Y/N)').ToLower() } while ($myInput -notin @('y','n'))
if ($myInput -eq 'y') {
WOLDiscovery
} else {
    Write-Output "Skipped"
}
do { $myInput = (Read-Host 'Get Outlook Profile Cache settings for all domain computers? (Y/N)').ToLower() } while ($myInput -notin @('y','n'))
if ($myInput -eq 'y') {
GetOutlookProfileDetail
} else {
    Write-Output "Skipped"
}

