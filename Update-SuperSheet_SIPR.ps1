#If we don't have enough RAM, we're going to have to run these computations on a server that does have RAM
#Therefore, we need to see if we should be running this as Admin
$CompRAM = (Get-WmiObject Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum | select -ExpandProperty sum) / 1GB
if ($compRAM -lt 20.0) {
    $myWindowsID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
    $myWindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($myWindowsID)
    $adminRole=[System.Security.Principal.WindowsBuiltInRole]::Administrator
    if (-not $myWindowsPrincipal.IsInRole($adminRole)) {
        $scriptpath = $MyInvocation.MyCommand.Definition
        $scriptpaths = "'$scriptPath'"
        Start-Process -FilePath PowerShell.exe -Verb runAs -ArgumentList "& $scriptPaths"
        exit
        }
    }
    
$pshost = get-host
$pswindow = $pshost.ui.rawui

$newsize = $pswindow.buffersize
$newsize.width = 200
$pswindow.buffersize = $newsize

#If you don't provide output, it will just change the extension and keep the name
Function ExcelCSV ($File,$output) {
    if ($File.fullname -eq "" -or $File.fullname -eq $null -or !(Test-Path $file.fullname)) {return}
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    $wb = $Excel.Workbooks.Open($File.fullname)
    #Make it a special case, the 83 NOS report has more than one sheet
    #Hard coded.  Supersheet has more than one worksheet, but we only want the ACAS from it.  
    #Some bases have their reports in xlsx as well, with mult sheets.
    if ($file.name -match "Supersheet") {
        $ws = $wb.Worksheets | select -First 1 -skip 2
        if (!$output) {$output = ($file.fullname -replace ".xlsx","_ACAS.csv")}
        $ws.SaveAs($output, 6)
        }
    else {#if ($file.name -match "SIPR Technical Vulnerability Report") {
        $ws = $wb.worksheets | select -First 1
        if (!$output) {$output = ($file.fullname -replace ".xlsx",".csv")}
        $ws.SaveAs($output, 6)
        }
    <#
    else {
        foreach ($ws in $wb.Worksheets) {
            $name = $ws.name
            if (!$output) {$output = ($file -replace ".xlsx","_$name.csv")}
            $ws.SaveAs($output, 6)
            }
        }#>
    $Excel.Quit()
}

#Simple enough unzip function
Function Unzip {
    param([string]$zipfile, [string]$outpath)
    $shell = New-Object -com shell.application
    $Zip = $shell.NameSpace($zipfile)
    
    if (!(Test-Path $outpath)) {New-Item -Path $outpath -ItemType Directory | Out-Null}
    foreach ($item in $zip.items()) {
        $shell.Namespace($outpath).copyhere($item)
        }
    }

#Add the given file into the specified zip file, creating said zip file if it doesn't exist
Function ZIP {
    param($ZipFullName,$FileToAdd)
    
    #$OutDir = $ZipFullName.substring(0,$ZipFullName.lastindexof("\"))
    $fileName = $FileToAdd.substring($FileToAdd.lastindexof("\") + 1)
    
    if(!(Test-Path $ZipFullName)) {
        Set-Content $ZipFullName ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
        (dir $ZipFullName).IsReadOnly = $false
        }
    
    $shell = New-Object -com shell.application
    $ZipPackage = $shell.namespace($ZipFullName)
    
    $ZipPackage.CopyHere($FileToAdd,0x14)
    
    while ($ZipPackage.Items().Item($fileName) -eq $null) {Start-Sleep -Seconds 1}
    }

#A little bit of initialization
#Working Directories
$CCRIFolder = "\\MUHJ-FS-001\cyod\11--Cyber 365\"
$WorkingDirectory = "$env:USERPROFILE\Supersheet\"

#Give us a clean working directory
Remove-Item "$WorkingDirectory*" -Force -Recurse -Confirm:$false -EA SilentlyContinue

#Get our list of IPs to collect results for
$AssetsFile = $WorkingDirectory + "83NOS_Assets.csv"

#Get our ACAS zips from here
$ScanRepository = "\\$ip\G\Scan Data\"

Foreach ($domain in @("ACC","AFMC")) {
    
    #Since we delete the directory after each domain, make sure we rebuild it
    while (!(Test-Path $WorkingDirectory)) {New-Item -Path $WorkingDirectory -ItemType Directory | out-null}

    #List of bases to get our scans for
    #We need to specify our filters to match as few as possible.  If we match multiple baes in one line, but theyre not scanned at the same time, we miss one of them
    if ($domain -eq "ACC") {
        #Do we care about Curacao?
        $Bases = @"
83 NOS
Barksdale
Beale
Creech
Davis Monthan
Dyess
Ellsworth
Holloman
Langley
Minot
Mt. Home
Nellis
Offut
Seymour Johnson
Tyndall
"@.split("`n") | Foreach {$_.trim()}
        }
    else {
        #Why does Greenville have a rescan?
        $Bases = @"
Arnold
Edwards
Eglin
Greenville
Gunter
Hanscom
Hill
Kirtland
Robins
Rome Labs
Tinker
Wright-Patt
"@.split("`n") | Foreach {$_.trim()}
        }
    
    #Our base ACAS File to work from, if we don't already have a .csv available from a previous run
    $ACASRawFile = (Get-ChildItem $CCRIFolder "SUPERSHEET-$domain.xlsx" -Recurse | Select -First 1) | Select Name,fullname

    #Advanced method for getting latest year
    #This Where ensures we only get folders with numeric names
    $Years = Get-ChildItem $ScanRepository | Where {$_.PSIsContainer -and $_.Name -match "^\d+(\D{0})$"} | foreach {
        New-Object PSObject -Property @{
            Year = [convert]::ToInt32($_.Name,10)
            Fullname = $_.FullName
            }
        }
    $ScanRepositoryYear = $Years | Sort-Object Year -Descending | select -First 1 -ExpandProperty Fullname

    #Advanced method for sorting months by their number
    Function Sort-Months {
        param ($year)
        return (Get-ChildItem $ScanRepositoryYear -Filter "* - *" | Where {$_.PSIsContainer} | foreach {
            New-Object PSObject -Property @{
                MonthNum = [convert]::ToInt32($_.Name.split(".")[0].trim(),10)
                Fullname = $_.FullName
                }
            } | sort MonthNum -Descending)
        }

    $Months = Sort-Months $ScanRepositoryYear
    $YearsSkip = 0

    #What if they made the year folder but didn't put any months in yet?
    if ($Months.count -eq 0) {
        $Years | Sort-Object Year -Descending | select -Skip 1 -First 1 -ExpandProperty Fullname
        $YearsSkip++
        }

    $ScanRepositoryMonth = $Months | select -First 1 -ExpandProperty Fullname

    #What if its January?
    if ($months.count -lt 2) {
        $ScanRepositoryLastYear = $Years | Sort-Object Year -Descending | select -Skip ($YearsSkip + 1) -First 1 -ExpandProperty Fullname
        $LastMonths = Sort-Months $ScanRepositoryLastYear
        $ScanRepositoryPastMonth = $LastMonths | select -First 1 -ExpandProperty Fullname
        }
    else {$ScanRepositoryPastMonth = $Months | select -Skip 1 -First 1 -ExpandProperty Fullname}

    Write-Host (Get-Date -Format "yyyy-MM-dd_hh-mm-ss") "Copying ACAS results to $WorkingDirectory" -ForegroundColor Cyan
    Foreach ($basename in $Bases) {
        #Special finagaling for multiple sites with the same name (WP, Scott, Andrews, etc)
        $subBases = @()
        if ($basename -eq "Greenville") {
            $subBases += "Greenville"
            }
        else {
            $subBases += $basename
            }

        #Our main filter will be the same, but we choose APC/not based on include and exclude filters
        $filter = "*" + $basename + "*.zip"
        foreach ($base in $subBases) {
            #If we are looking for an APC, our include filter will ensure we only get APC scans for the base
            #Otherwise, the exclude filter will get us just the base scans
            #For Greenville, only get the rescan, if it exists 
            if ($base -eq "Greenville" -and ((Get-ChildItem "$ScanRepositoryMonth\$domain\*" -Filter "*Rescan*") -or (Get-ChildItem "$ScanRepositoryPastMonth\$domain\*" -Filter "*Rescan*"))) {
                Remove-Variable excludeFilter -EA SilentlyContinue
                $includeFilter = "*Rescan*"
                }
            else {
                $excludeFilter = "*Rescan*"
                Remove-Variable includeFilter -EA SilentlyContinue
                }

            #Check the latest Month for a ACAS scan for the base.
            $ZippityDooDah = Get-ChildItem "$ScanRepositoryMonth\$domain\*" -Filter $filter -Include $includeFilter -Exclude $excludeFilter -Recurse | Sort-Object name -Descending | Select Name,Fullname
            #If it's not there, check last month's.  We're assuming it's guaranteed to be there.
            if (!$ZippityDooDah) {
                $ZippityDooDah = Get-ChildItem "$ScanRepositoryPastMonth\$domain\*" -Filter $filter -Include $includeFilter -Exclude $excludeFilter -Recurse | Sort-Object name -Descending | Select Name,Fullname
                }
            #If we STILL couldn't find any scans, screw it and skip.
            if ($ZippityDooDah -eq $null -or $ZippityDooDah -eq "") {
                Write-Host "Error: Could not find any ACAS results for $basename.  We will be skipping this base"
                continue
                }

            #Error checking Loop, in case the repository has a communication hiccup
            #This WILL loop forever if the server goes down or the repository moves
            do {
                try {
                    Write-Host -ForegroundColor green ("Copying " + $ZippityDooDah.FullName)
                    Copy-Item -LiteralPath $ZippityDooDah.FullName -Destination ($WorkingDirectory + $ZippityDooDah.Name) -EA Stop
                    $success = $true
                    }
                #If it fails, delete whatever it copied and try again
                catch { 
                    Remove-Item ($WorkingDirectory + $ZippityDooDah.Name) -Force -Confirm:$false -EA SilentlyContinue
                    $success = $False
                    }
                } until ($success)
            }
        }

    #Make sure our assets file is in readable format
    #We want to force a new copy, just in case it has been updated
    $sourceXLSX = (Get-ChildItem $CCRIFolder -Filter "83NOS_Assets.xlsx" -Recurse | Select -First 1) | select Name,fullname
    ExcelCSV $sourceXLSX $AssetsFile

    #Get our current supersheet to update
    $BaseACASFile = (Get-ChildItem $CCRIFolder -Filter SuperSheet-$domain*.csv | sort name -Descending | Select -First 1 name,fullname)
    $Timestamp = Get-Date -Format "yyy-MM-dd_hh-mm-ss"
    #Copy it locally if found
    if ($BaseACASFile) {
        $BaseACASFileName = $BaseACASFile.name
        $BaseACASFile = $BaseACASFile.fullname
        }
    #If we don't already have a csv to work with, make one locally
    else {
    Write-Host (Get-Date -Format "yyyy-MM-dd_hh-mm-ss") "Converting our Supersheet" -ForegroundColor Cyan
        $BaseACASFileName = "SUPERSHEET-$domain" + "_$timestamp.csv"
        $BaseACASFile = ($WorkingDirectory + $BaseACASFileName)
        if (!(Test-Path $BaseACASFile)) {ExcelCSV $ACASRawFile $BaseACASFile}
        }
    

    #Add a pause so we can manually add ZIPs
    $AppendACASzips = (Get-ChildItem $WorkingDirectory -Filter "*.zip") | foreach {$_.Fullname}
    Write-Host -ForegroundColor Green "Take this time to add any zips you want to include and/or were missed by this script"
    Write-Host -ForegroundColor Green "There are"($AppendACASzips.Count)"Zips currently."
    Write-Host -ForegroundColor Green "Copy your files into $WorkingDirectory"
    Write-Host -ForegroundColor Green "Press enter when you are ready to continue..." -NoNewline
    Read-Host

    #Collect the zips in our working directory
    $AppendACASzips = (Get-ChildItem $WorkingDirectory -Filter "*.zip") | foreach {$_.Fullname}
    if (!$AppendACASzips) {
        Write-Host "There's no new zip to update the supersheet with"
        Read-Host -Prompt "Exit the script at this time"
        exit
        }

    #Unzip them
    Write-Host (Get-Date -Format "yyyy-MM-dd_hh-mm-ss") "Unzipping"($AppendACASzips.count)"Files" -ForegroundColor Cyan
    foreach ($zip in $AppendACASzips) {
        $newFldr = $zip.trimend(".zip")
        if (!(Test-path $newFldr)) {
            Unzip $zip $newFldr
            }
        }
        
    #Why they gotta put AFMC files in a zip in a zip, i dunno, but its stupid.
    if ($Domain -eq "AFMC") {
        Get-ChildItem $WorkingDirectory | where {$_.PSIsContainer} | foreach {
            Remove-Variable subzip -EA SilentlyContinue
            $subZip = Get-ChildItem ($_.Fullname) -filter *.zip | select -ExpandProperty fullname
            if ($subZip) {
                $newFldr = $subZip.trimend(".zip")
                Unzip $subZip $newFldr
                }
            }
        }

    #convert our ACAS scans results to CSV
    $AppendACASRaw = @()
    $AppendACASRaw += (Get-ChildItem $WorkingDirectory -Recurse -Filter "*Report*.xlsx" -exclude "*30*","*failed*") | Where {!($_.PSIsContainer)} | select Name,Fullname
    Write-Host (Get-Date -Format "yyyy-MM-dd_hh-mm-ss") "Converting"($AppendACASRaw.count)".xlsx file(s) to .csv" -ForegroundColor Cyan #todo
    foreach ($xlsx in $AppendACASRaw) {
        if (!(Test-Path ($xlsx.FullName -replace ".xlsx",".csv"))) {
            ExcelCSV $xlsx
            }
        }
    $AppendACASFiles = (Get-ChildItem $WorkingDirectory -Recurse -Filter "*.csv" -exclude "*asset*","*supersheet*","*30*","DO_NOT_OPEN*","*failed*") | Where {!($_.PSIsContainer)} | foreach {$_.Fullname}
    
        
    #Free up some memory
    Remove-variable AppendACASRaw,AppendACASzips,sourceXLSX,ZippityDooDah,subBases,base,basename,bases,ScanRepositoryPastMonth,ScanRepositoryLastYear,ScanRepositoryMonth,ScanRepositoryYear,Years,months,ACASRawFile -EA silentlycontinue
    [GC]::Collect()
        
    #We might have to execute this code elsewhere due to RAM, so put it in a script block
    $ScriptBlock = {
        $AppendACASFiles = (Get-ChildItem $WorkingDirectory -Recurse -Filter "*.csv" -exclude "*asset*","*supersheet*","*30*","DO_NOT_OPEN*","*failed*") | Where {!($_.PSIsContainer)} | foreach {$_.Fullname}
    
        #Finally compare ACAS's
        #BaseACAS will be used to fillw in any gaps in missing assets, after we compile all the new ACAS results
        $Assets = Import-Csv $AssetsFile | Select 'DNS Name','IP Address','MAJCOM','Base'
        $BaseACAS = Import-Csv $BaseACASFile

        #Get the ACAS results for our assets
        $NewACAS = @()

        #Start of with new results for our assets; we know we'll want these
        Write-Host (Get-Date -Format "yyyy-MM-dd_hh-mm-ss") "We will be gathering ACAS scans from"($AppendACASFiles.count)"reports" -ForegroundColor Cyan
        foreach ($AppendACASFile in $AppendACASFiles){
        Write-Host (Get-Date -Format "yyyy-MM-dd_hh-mm-ss") "Parsing results from $AppendACASFile" -ForegroundColor Cyan
            do { #This is  an error checking loop, just in case there's a communication hiccup
                try {
                    #Checking to make sure they didn't label the document's classification
                    $AppendACAS = Get-Content $AppendACASFile
                    if ($AppendACAS[0] -match "UNCLASSIFIED") {
                        $newAppendAcas = $AppendACAS[1..($AppendACAS.count - 2)]
                        Out-File -InputObject $newAppendAcas -FilePath $AppendACASFile -Encoding default
                        }

                    #Change any instances of "Plugin Text" column header to "Plugin Output"
                    #Why, WHY do they have reports with both headers?
                    $AppendACAS = Get-Content $AppendACASFile
                    if ($AppendACAS[0] -like "*Plugin Text*") {
                        if ($AppendACAS[0] -like "*Plugin Output*") {$AppendACAS[0] = $AppendACAS[0] -replace "Plugin Output","Garbage"}
                        $AppendACAS[0] = $AppendACAS[0] -replace "Plugin Text","Plugin Output"
                        Out-File -InputObject $AppendACAS -FilePath $AppendACASFile -Encoding default
                        }

                    #Just make sure there's no reading errors
                    $AppendACAS = Import-Csv $AppendACASFile -EA Stop
                    #Separate our ACAS entries by unique pairs of hostname and IP
                    $List = $AppendACAS | Group-Object -Property 'IP Address','DNS Name'
                    foreach ($item in $List) {
                        $ip = $item.name.split(",")[0].trim()
                        $name = $item.name.split(",")[1].trim().trimend(".")
                        #Make sure it's one of our assets
                        $matched = $Assets | Where {$_.'IP Address' -eq $IP} | select -first 1
                        if ($matched) {
                            foreach ($subitem in $item.Group) {
                                #Set the BASE, MAJCOM, and hostname(if blank) for each finding
                                $subitem | Add-Member -MemberType NoteProperty -Name "BASE" -Value ($matched.Base) -Force
                                $subitem | Add-Member -MemberType NoteProperty -Name "MAJCOM" -Value ($matched.MAJCOM) -Force
                                if ($name -eq "") {$subitem.'DNS Name' = $matched.'DNS Name'}
                                $NewACAS += $subitem
                                }
                            continue
                            }
                        <#
                        #If we didn't find it by IP, check the hostname against our list of assets; just in case an IP updated or we did not record it.
                        #Maybe we should get rid of this part
                        $matched = $Assets | Where {$name -ne "" -and $_.'DNS Name' -match $name}
                        if ($matched) {
                            foreach ($subitem in $item.Group) {
                                #Set the BASE and MAJCOM for each finding
                                $subitem | Add-Member -MemberType NoteProperty -Name "BASE" -Value ($matched.Base)
                                $subitem | Add-Member -MemberType NoteProperty -Name "MAJCOM" -Value ($matched.MAJCOM)
                                $NewACAS += $subitem
                                }
                            continue
                            }
                        #>
                        }
                    $success = $true
                    }
                catch {
                    $success = $false
                    Write-Host -ForegroundColor Magenta "Retrying $AppendACASFile :"(get-date -Format "yyyy-MM-dd_hh-mm-ss")
                    }
                } until ($success)
            }
        
        #Free up some memory
        Remove-variable AppendACAS,List,newAppendAcas,success,item,ip,name -EA SilentlyContinue
        [GC]::Collect()

        #Now get the old ACAS results of our assets that aren't in the new list.
        #Separate our ACAS entries by unique pairs of hostname and IP
        $List = $BaseACAS | Group-Object -Property 'IP Address','DNS Name'
        Write-Host (Get-Date -Format "yyyy-MM-dd_hh-mm-ss") "Grabbing old ACAS results" -ForegroundColor Cyan
        foreach ($item in $List) {
            $ip = $item.name.split(",")[0].trim()
            if ($ip -eq "") {continue}
            $name = $item.name.split(",")[1].trim().trimend(".")
            #If the IP already has ACAS entries from our newest ACAN scans, ignore it
            if ($NewACAS | Where {$_.'IP Address' -eq $IP}) {continue}
            #We can assume that they are our assets, since this is a previously generated supersheet
            $matched = $Assets | Where {$_.'IP Address' -eq $IP}
            foreach ($subitem in $item.Group) {
                #Set the BASE, MAJCOM, and hostname for each finding, if they are missing
                if ($subitem.'DNS Name' -eq "" -or $subitem.'DNS Name' -eq $null) {$subitem.'DNS Name' = $matched.'DNS Name'}
                if ($subitem.BASE -eq "" -or $subitem.BASE -eq $null) {$subitem | Add-Member -MemberType NoteProperty -Name "BASE" -Value ($matched.Base) -Force}
                if ($subitem.MAJCOM -eq "" -or $subitem.MAJCOM -eq $null) {$subitem | Add-Member -MemberType NoteProperty -Name "MAJCOM" -Value ($matched.MAJCOM) -Force}
                $NewACAS += $subitem
                }
            }
        
        #Free up some memory
        Remove-variable List,BaseACAS,item,subitem,ip,name,matched,subitem -EA SilentlyContinue
        [GC]::Collect()
            
        #Write the finalized report locally first, then copy it to the share.
        Write-Host (Get-Date -Format "yyyy-MM-dd_hh-mm-ss") "Exporting Combined ACAS" -ForegroundColor Cyan
        $tempLocalLocation = $WorkingDirectory + "DO_NOT_OPEN.csv"
        $NewACAS | Select "Plugin","Plugin Name","Severity","IP Address","DNS Name","Plugin Output","Synopsis","Solution","Last Observed","MAJCOM","BASE" | Export-Csv -NoTypeInformation $tempLocalLocation
        #stop here for testing
        #exit
        #pause
        }#End Script Block
    
    #Stop here for testing
    #exit
    
    #We need lots of RAM, so we need to make sure we don't stall out the computer with this script.
    if ($compRAM -lt 20.0) {
        Write-host -f Red "This sytem does not have enough RAM.  Setting up files to execute code on a remote server"
    
        #This is WIP
        $executeRemotely = $true
        If ($executeRemotely) {
            #Remote Server/Comp to do our computations
            $ComputerName = "muhj-dc-002"
            
            ##Copy Files over to server to run the RAM intensive part of this script
            #Create our remote Directory
            $remoteDir = "\\$ComputerName\c$\supersheet\$domain\"
            while (Test-Path "\\$ComputerName\c$\supersheet") {Remove-Item "\\$ComputerName\c$\supersheet" -Force -Recurse -EA silentlycontinue}
            while (!(Test-Path $remoteDir)) {New-Item $remoteDir -ItemType directory | out-null}
            
            #Copy over our needed files
            Copy-Item $BaseACASFile $remoteDir
            Copy-Item $AssetsFile $remoteDir
            foreach ($csv in $AppendACASFiles) {
                Copy-Item $csv $remoteDir
                }
                
            #We can either output our scriptblock into a remote file, to be executed, or we can try invoke-command on the script block and pass params
            #Make our scriptblock accept parameters
            $FinalBlock = [scriptblock]::Create("param(`$domain,`$timestamp,`$WorkingDirectory,`$CCRIFolder,`$AssetsFile,`$BaseACASFile)" + $Scriptblock.toString())
            $RemoteLocalWorkDir = "C:\supersheet\$domain\"

            write-host -f green "Executing code on $Computername"
            Invoke-Command -ComputerName $ComputerName -ScriptBlock $FinalBlock -ArgumentList $domain,$timestamp,$RemoteLocalWorkDir,$CCRIFolder,"$RemoteLocalWorkDir\83NOS_Assets.csv","$RemoteLocalWorkDir\$BaseACASFileName"
            
            #Since we can't pass our admin token on remotely, we have to copy the file to the share using our local PS token
            #Yes this is overly complicated, but I'm afraid of some idjit opening it up while we're moving it to the share.
            $newFileName = "SuperSheet-$domain" + "_" + (Get-Date -Format "yyyy-MM-dd_hh-mm-ss") + ".csv"
            do {
                Remove-Item ($CCRIFolder + "DO_NOT_OPEN.csv") -Force -ErrorAction SilentlyContinue
                Copy-Item -LiteralPath ($remoteDir + "DO_NOT_OPEN.csv") -Destination ($CCRIFolder + "DO_NOT_OPEN.csv")
                } until (Test-Path ($CCRIFolder + "DO_NOT_OPEN.csv"))
            Rename-Item ($CCRIFolder + "DO_NOT_OPEN.csv") -NewName $newFileName
            
            #Delete our remote workign directory
            while (Test-Path "\\$ComputerName\c$\supersheet") {Remove-Item "\\$ComputerName\c$\supersheet" -Force -Recurse -EA silentlycontinue}
            }
        Else { #Copy everything into a ZIP so we can just copy it over to a DC and run it
            #Bad idea unless we can get into a Virt Console
            #Terminal session timeouts suck ass
            #Our folder to contain everything and ZIP up
            $ZipMe = $WorkingDirectory + "SuperSheet2B-" + $domain + "_$timestamp.zip"
            #New-Item $zipme -ItemType directory | out-null
            
            #Add our CSVs to work off of
            ZIP $ZipMe $BaseACASFile
            ZIP $ZipMe $AssetsFile
            foreach ($csv in $AppendACASFiles) {
                ZIP $ZipMe $csv
                }
                
            #Since we have to execute the code on a different machine, we need to transfer our variables with depenedencies into our script block.
            $FinalBlock = [scriptblock]::Create("`$domain = `"$domain`"`r`n" +
                "`$timestamp = `"$timestamp`"`r`n" +
                "`$temp = `$MyInvocation.MyCommand.Definition`r`n" +
                "`$WorkingDirectory = `$temp.substring(0,`$temp.lastindexof(`"\`")) + `"\`"`r`n" +
                "`$CCRIFolder = `"$CCRIFolder\`"`r`n" +
                "`$AssetsFile = `$WorkingDirectory + `"83NOS_Assets.csv`"`r`n" +
                "`$BaseACASFile = `$WorkingDirectory + `"$BaseACASFileName`"`r`n" +
                $Scriptblock.toString())
                
            #Create our script to run, and then add it to the ZIP
            #The script should just be run in the folder it was exported into
            Out-File -FilePath ($WorkingDirectory + "CombineACAS.ps1") -Encoding default -InputObject ($FinalBlock.toString())
            ZIP $ZipME ($WorkingDirectory + "CombineACAS.ps1")
            
            #Copy ZIP to share drive CCRI folder for ease of transfer to server
            Copy-Item $ZipMe $CCRIFolder
            Write-Host  -NoNewline "Your drop package is located at "
            Write-Host -ForegroundColor Green ($CCRIFolder + "SuperSheet2B-" + $domain + "_$timestamp.zip")
            }
        }
    else {
        &$Scriptblock
        
        #Yes this is overly complicated, but I'm afraid of some idjit opening it up while we're moving it to the share.
        $newFileName = "SuperSheet-$domain" + "_" + (Get-Date -Format "yyyy-MM-dd_hh-mm-ss") + ".csv"
        do {
            Remove-Item ($CCRIFolder + "DO_NOT_OPEN.csv") -Force -ErrorAction SilentlyContinue
            Copy-Item -LiteralPath $tempLocalLocation -Destination ($CCRIFolder + "DO_NOT_OPEN.csv")
            } until (Test-Path ($CCRIFolder + "DO_NOT_OPEN.csv"))
        Rename-Item ($CCRIFolder + "DO_NOT_OPEN.csv") -NewName $newFileName
        }
    
    #Delete our old ACAS files and zips
    while (Test-Path $WorkingDirectory) {Remove-Item $WorkingDirectory -Force -Recurse -Confirm:$false -EA SilentlyContinue}
    }
Read-Host -Prompt "Script complete.  Press Enter to close window."
