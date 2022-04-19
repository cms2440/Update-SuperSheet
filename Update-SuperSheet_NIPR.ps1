
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
        $ws = $wb.Worksheets[3]
        if (!$output) {$output = ($file.fullname -replace ".xlsx","_ACAS.csv")}
        $ws.SaveAs($output, 6)
        }
    else {#if ($file.name -match "NIPR Technical Vulnerability Report") {
        $ws = $wb.worksheets[1]
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
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Unzip
{
    param([string]$zipfile, [string]$outpath)
    $ErrorActionPreference = "SilentlyContinue"
    [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
    $ErrorActionPreference = "Continue" 
}

#A little bit of initialization
#Working Directories
$CCRIFolder = "\\apca-fs-001v\cyod\07--Cyber 365\"
$WorkingDirectory = "$env:USERPROFILE\Supersheet\"
if ($env:COMPUTERNAME -eq "MUHJW-431MFQ") {$WorkingDirectory = "C:\Users\1456084571E\Documents\PS Workshop\CCRI\Supersheet\"}
if (!(Test-Path $WorkingDirectory)) {New-Item -Path $WorkingDirectory -ItemType Directory | out-null}

#Data to work with
$ACASRawFile = (Get-ChildItem $CCRIFolder -File "SUPERSHEETv2.xlsx" -Recurse | Select -First 1) | Select Name,fullname
$AssetsFile = $WorkingDirectory + "83NOS_Assets.csv"

#Get our ACAS zips from here
$ScanRepository = "\\$ip\Scan Data\"

#List of bases to get our scans for
#We need to specify our filters to match as few as possible.  If WPAFB and WPAPC aren't done at the same time, we will miss one
#Only bases this is an issue, WP and Andrews
#Barksdale isnt on the repository for some reason.
$Bases = @"
83 NOS
AFOSR
Andrews
Arnold
Asheville
Beale
Bolling
Cannon
Curacao
Davis Monthan
Dyess
Edwards
Eglin
Ellsworth
Ft Detrick
Ft Meade
Ft Sam Houston
Greenville
Gunter
Hanscom
Hurlburt
Kirtland
Langley
Maui
Moody
Mountain Home
Nellis
Offutt
Pentagon
Robins AFRC
Rome Labs
Seymour Johnson
Shaw
Tinker
Wright Pat
"@.split("`n") | Foreach {$_.trim()}

#Get the latest ZIP for each base and copy to our work folder
$ScanRepository = "\\132.54.9.71\Scan Data\"

#This Where ensures we only get folders with numeric names
#Advanced method for getting latest year
$Years = Get-ChildItem $ScanRepository -Directory | Where {$_.Name -match "^\d+(\D{0})$"} | foreach {# -Filter "20*" 
    New-Object PSObject -Property @{
        Year = [convert]::ToInt32($_.Name,10)
        Fullname = $_.FullName
        }
    }
$ScanRepositoryYears = $Years | Sort-Object Year -Descending | select -ExpandProperty fullname 

Write-Host (Get-Date -Format "yyyy-MM-dd_hh-mm-ss") "Copying ACAS results to $WorkingDirectory" -ForegroundColor Cyan
Foreach ($basename in $Bases) {
    #Special finagaling for WPAPC and Andrews
    $subBases = @()
    if ($basename -eq "Andrews") {
        $subBases += "Andrews"
        $subBases += "Andrews APC"
        }
    elseif ($basename -eq "Wright Pat") {
        $subBases += "Wright Pat"
        $subBases += "Wright Pat APC"
        }
    else {
        $subBases += $basename
        }

    #Our main filter will be the same, but we choose APC/not based on include and exclude filters
    $filter = "*" + $basename + "*.zip"

    foreach ($base in $subBases) {
        #If we are looking for an APC, our include filter will ensure we only get APC scans for the base
        #Otherwise, the exclude filter will get us just the base scans
        if ($base -eq "Andrews APC" -or $base -eq "Wright Pat APC") {
            Remove-Variable excludeFilter -EA SilentlyContinue
            $includeFilter = "*APC*"
            }
        else {
            $excludeFilter = "*APC*"
            Remove-Variable includeFilter -EA SilentlyContinue
            }

        #Try to find our newest ACAS scan, counting down by months in years
        :main foreach ($ScanRepositoryYear in $ScanRepositoryYears) {
            #Advanced method for sorting months by their number
            $Months = Get-ChildItem $ScanRepositoryYear -Directory -Filter "* - *" | foreach {
                New-Object PSObject -Property @{
                    MonthNum = [convert]::ToInt32($_.Name.split("-")[0].trim(),10)
                    Fullname = $_.FullName
                    }
                }
            $ScanRepositoryMonths = $Months | Sort-Object MonthNum -Descending | select -ExpandProperty fullname
            #Parse through each month for the given year
            foreach ($ScanRepositoryMonth in $ScanRepositoryMonths) {
                $ZippityDooDah = Get-ChildItem $ScanRepositoryMonth -Filter $filter -Recurse | Sort-Object name -Descending | Select Name,Fullname
                if ($ZippityDooDah -ne $null -and $ZippityDooDah -ne "") {break main}
                }
            }

        #If we STILL couldn't find any scans, screw it and skip.
        if ($ZippityDooDah -eq $null -or $ZippityDooDah -eq "") {
            Write-Host "Error: Could not find any ACAS results for $basename.  We will be skipping this base"
            continue
            }

        #Error checking loop, in case the repository has a communication hiccup
        #This WILL loop forever if the server goes down or the repository moves
        do {
            try {
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
$sourceXLSX = (Get-ChildItem $CCRIFolder -Filter "Server-IP for ACAS Updated.xlsx" -Recurse | Select -First 1) | select Name,fullname
ExcelCSV $sourceXLSX $AssetsFile

#Get our current supersheet to update
$BaseACASFile = (Get-ChildItem $CCRIFolder -Filter SuperSheet*.csv | sort name -Descending | Select -First 1).fullname
#If we don't already have a csv to work with, make one locally
if (!$BaseACASFile) {
Write-Host (Get-Date -Format "yyyy-MM-dd_hh-mm-ss") "Converting our Supersheet" -ForegroundColor Cyan
    $timestamp = get-date -Format "yyyy-MM-dd_hh-mm-ss"
    $BaseACASFile = ($WorkingDirectory + "SUPERSHEET_$timestamp.csv")
    if (!(Test-Path $BaseACASFile)) {ExcelCSV $ACASRawFile $BaseACASFile}
    }

#Add a pause so we can manually add ZIPs
$AppendACASzips = (Get-ChildItem $WorkingDirectory -Filter "*.zip") | foreach {$_.Fullname}
Write-Host -ForegroundColor Green "Take this time to add any zips you want to include and/or were missed by this script *cough*Barksdale*cough*"
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

#convert our ACAS scans results to CSV
$AppendACASRaw = @()
$AppendACASRaw += (Get-ChildItem $WorkingDirectory -Recurse -Filter "*NIPR Technical Vulnerability Report*.xlsx" -exclude "*30*") | select Name,Fullname
Write-Host (Get-Date -Format "yyyy-MM-dd_hh-mm-ss") "Converting"($AppendACASRaw.count)".xlsx file(s) to .csv" -ForegroundColor Cyan #todo
foreach ($xlsx in $AppendACASRaw) {
    if (!(Test-Path ($xlsx.FullName -replace ".xlsx",".csv"))) {
        ExcelCSV $xlsx
        }
    }
#$AppendACASFiles = $AppendACASRaw | foreach {$_ -replace "xlsx","csv"}
$AppendACASFiles = (Get-ChildItem $WorkingDirectory -Recurse -Filter "*.csv" -exclude "*asset*","*supersheet*","*30*","DO_NOT_OPEN*") | foreach {$_.Fullname}

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
            $AppendACAS = Get-Content $AppendACASFile
            if ($AppendACAS[0] -like "*Plugin Text*") {
                $AppendACAS[0] = $AppendACAS[0] -replace "Plugin Text","Plugin Output"
                Out-File -InputObject $AppendACAS -FilePath $AppendACASFile -Encoding default
                }

            #Just make sure there's no reading errors
            $AppendACAS = Import-Csv $AppendACASFile -EA Stop
            #Separate our ACAS entries by unique pairs of hostname and IP
            $List = $AppendACAS | Group-Object -Property 'IP Address','DNS Name'
            foreach ($item in $list) {
                $ip = $item.name.split(",")[0].trim()
                $name = $item.name.split(",")[1].trim().trimend(".")
                #Make sure it's one of our assets
                $matched = $Assets | Where {$_.'IP Address' -eq $IP}
                if ($matched) {
                    foreach ($subitem in $item.Group) {
                        #Set the BASE, MAJCOM, and hostname(if blank) for each finding
                        $subitem | Add-Member -MemberType NoteProperty -Name "BASE" -Value ($matched.Base)
                        $subitem | Add-Member -MemberType NoteProperty -Name "MAJCOM" -Value ($matched.MAJCOM)
                        if ($name -eq "") {$subitem.'DNS Name' = $matched.'DNS Name'}
                        $NewACAS += $subitem
                        }
                    continue
                    }
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
                }
            $success = $true
            }
        catch {
            $success = $false
            Write-Host -ForegroundColor Magenta "Retrying $AppendACASFile :"(get-date -Format "yyyy-MM-dd_hh-mm-ss")
            }
        } until ($success)
    }

#Now get the old ACAS results of our assets that aren't in the new list.
#Separate our ACAS entries by unique pairs of hostname and IP
$List = $BaseACAS | Group-Object -Property 'IP Address','DNS Name'
Write-Host (Get-Date -Format "yyyy-MM-dd_hh-mm-ss") "Grabbing old ACAS results" -ForegroundColor Cyan
foreach ($item in $list) {
    $ip = $item.name.split(",")[0].trim()
    $name = $item.name.split(",")[1].trim().trimend(".")
    #If the IP already has ACAS entries from our newest ACAN scans, ignore it
    if ($NewACAS | Where {$_.'IP Address' -eq $IP}) {continue}
    #We can assume that they are our assets, since this is a previously generated supersheet
    $matched = $Assets | Where {$_.'IP Address' -eq $IP}
    foreach ($subitem in $item.Group) {
        #Set the BASE, MAJCOM, and hostname for each finding, if they are missing
        if ($subitem.'DNS Name' -eq "" -or $subitem.'DNS Name' -eq $null) {$subitem.'DNS Name' = $matched.'DNS Name'}
        if ($subitem.BASE -eq "" -or $subitem.BASE -eq $null) {$subitem.BASE = $matched.BASE}
        if ($subitem.MAJCOM -eq "" -or $subitem.MAJCOM -eq $null) {$subitem.MAJCOM = $matched.MAJCOM}
        $NewACAS += $subitem
        }
    }
    
#Write the finalized report locally first, then copy it to the share.
Write-Host (Get-Date -Format "yyyy-MM-dd_hh-mm-ss") "Exporting Combined ACAS" -ForegroundColor Cyan
$newFileName = "SuperSheetv2_" + (Get-Date -Format "yyyy-MM-dd_hh-mm-ss") + ".csv"
$tempLocalLocation = $WorkingDirectory + "DO_NOT_OPEN.csv"
if ($env:COMPUTERNAME = "MUHJW-431MFQ") {$tempLocalLocation = "C:\Users\1456084571E\Documents\PS Workshop\CCRI\Supersheet\DO_NOT_OPEN.csv"}
$NewACAS | Select "Plugin","Plugin Name","Severity","IP Address","DNS Name","Plugin Output","Synopsis","Solution","Last Observed","MAJCOM","BASE" | Export-Csv -NoTypeInformation $tempLocalLocation
#stop here for testing
#exit
#pause

#Yes this is overly complicated, but I'm afraid of some idjit opening it up while we're moving it to the share.
do {
    Remove-Item ($CCRIFolder + "DO_NOT_OPEN.csv") -Force -ErrorAction SilentlyContinue
    Copy-Item -LiteralPath $tempLocalLocation -Destination ($CCRIFolder + "DO_NOT_OPEN.csv")
    } until (Test-Path ($CCRIFolder + "DO_NOT_OPEN.csv"))
Rename-Item ($CCRIFolder + "DO_NOT_OPEN.csv") -NewName $newFileName

#Delete our old ACAS files and zips
Remove-Item $WorkingDirectory -Force -Recurse -Confirm:$false
Read-Host -Prompt "Script complete.  Press Enter to close window."