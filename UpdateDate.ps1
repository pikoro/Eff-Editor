#  Script to extract and modify EFF files to update for the current date
Add-Type -AssemblyName System.IO.Compression.Filesystem

$global:debug = 0
$global:TempDir = "C:\WIP\"
$global:TempDatDir = "DAT\"
$global:TempEFFDir = "EFF\"
$global:FlySmartVersion = "12AUG2018"


function NormalizeName{
	param([string]$effName)
#	if($debug){"Input Name: $effName"}
	$effName = $effName.TrimStart(".\")
#	if($debug){"Trimmed String:  $effName"}
	$effName = $effName.TrimEnd(".eff")
#	if($debug){"Fully Trimmed String: $effName"}
	return $effName
}

function UnpackEff{
	param([string]$zipfile)
	""
	"Normalizing Name"
	""
	$zipfile = NormalizeName($zipfile)
	"Unpacking $zipfile"
	if($debug){"EFF Name:  $zipfile"}
	$zipfile = $TempDir + $zipfile + ".eff"
	if($debug){"File Name: $zipfile"}
	$outpath = $TempDir
	if($debug){"OutPath:  $outpath"}
	[System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
#	try {[System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)}
#	catch { "EFF Already Unpacked"}
	if($debug){"-------------------------------"}
}

function ExtractDat{
	"Extracting EFF Contents"
	$zipfile = Get-ChildItem -Filter "*.dat" -Name
	if($debug){"DAT Name:  $zipfile"}
	$zipfile = $TempDir + $zipfile
	if($debug){"File Name: $zipfile"}
	$outpath = $TempDir + $TempDatDir
	if($debug){"OutPath:  $outpath"}
	try {[System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)}
	catch {"! Contents already extracted"}
	if($debug){"-------------------------------"}
}

function ReplaceEffDate{
	"Changing EFF Date"
	$xmlFile = $TempDatDir + "eff.xml"
	$xmlFile = Get-ChildItem -Filter $xmlFile -Name
	$xmlFile = $TempDir + $TempDatDir + $xmlFile
	if($debug){"EFF XML: $xmlFile"}
	[xml]$XmlDocument = Get-Content -Path $xmlFile
	$DateStamp = $XmlDocument.EFUSUB.SubFolder.SubFolder.Document
	foreach($element in $DateStamp){
		try {$id = $element.id.ToString()} catch {"! Element does not exist"}
		try {if($debug){"Date Before: " + $element.UpdateDateTime.ToString()}} catch {"! Could not convert null value"}
		$newDate = Get-Date -Format yyyy-MM-ddThh:mm:ss.fff
		try{$element.updateDateTime = $newDate.ToString()}
		catch{"! This node does not have a date"}
		if($debug){"Updated $id to $newDate"}
	}
	$xmlDocument.EFUSUB.M633Header.timestamp = $newDate.ToString()
	try{$XmlDocument.Save($xmlFile)}
	catch{"! Unable to Save $xmlfile"}
	if($debug){"Date set to $newDate"}
	if($debug){"-------------------------------"}
}

function updateOFP{
    
    "Updating Operational Flight Plan"
    ""
    $xmlFile = $TempDatDir + "OperationalFlightPlan.xml"
    $xmlFile = Get-ChildItem -Filter $xmlFile -Name
    $xmlFile = $TempDir + $TempDatDir + $xmlFile
    if($debug){"OFP XML: $xmlfile"}
    [xml]$XmlDocument = Get-Content -Path $xmlFile
    " + Updating Computed Time"
    $XmlDocument.FlightPlan.ComputedTime = $newDate = (Get-Date -Format yyyy-MM-ddThh:mm:ssZ).ToString()
    " + Updating OFP Timestamp"
    $XmlDocument.FlightPlan.M633Header.timestamp = $newDate = (Get-Date -Format yyyy-MM-ddThh:mm:ssZ).ToString()
    
    " + Updating Scheduled Time of Departure"
    $dateArray = $XmlDocument.FlightPlan.M633SupplementaryHeader.Flight
    foreach($item in $dateArray){
        $tmpSplit = $item.scheduledTimeOfDeparture.ToString()
        $timeArray = $tmpSplit.split("T")
        $tmpDate = $timeArray[0]
        if($debug){"tmpDate:  $tmpDate"}
        $tmpTime = $timeArray[1]
        if($debug){"tmpTime:  $tmpTime"}
        $newDate = (Get-Date -Format yyyy-MM-dd).ToString()
        $tmpItem = $newDate + "T" + $tmpTime
        $item.flightOriginDate = $newDate
        $item.scheduledTimeOfDeparture = $tmpItem.ToString()
    }

    " + Updating Estimated Time over Waypoint"
    $dateArray = $xmldocument.FlightPlan.Waypoints.Waypoint.Timeoverwaypoint.estimatedtime
    foreach($item in $dateArray){
        $tmpSplit = $item.value.ToString()
        $timeArray = $tmpSplit.split("T")
        $tmpDate = $timeArray[0]
        if($debug){"tmpDate:  $tmpDate"}
        $tmpTime = $timeArray[1]
        if($debug){"tmpTime:  $tmpTime"}
        $newDate = (Get-Date -Format yyyy-MM-dd).ToString()
        $tmpItem = $newDate + "T" + $tmpTime
        $item.value = $tmpItem.ToString()

    }

    " + Updating Suitable To/Until Period"
    $dateArray = $xmldocument.FlightPlan.AirportDataList.AirportData.SuitablePeriod
    foreach($item in $dateArray){
# Update From Time
        $tmpSplit = $item.from.ToString()
        $timeArray = $tmpSplit.split("T")
        $tmpDate = $timeArray[0]
        if($debug){"tmpDate:  $tmpDate"}
        $tmpTime = $timeArray[1]
        if($debug){"tmpTime:  $tmpTime"}
        $newDate = (Get-Date -Format yyyy-MM-dd).ToString()
        $tmpItem = $newDate + "T" + $tmpTime
        $item.from = $tmpItem.ToString()
# Update Until Time
        $tmpSplit = $item.until.ToString()
        $timeArray = $tmpSplit.split("T")
        $tmpDate = $timeArray[0]
        if($debug){"tmpDate:  $tmpDate"}
        $tmpTime = $timeArray[1]
        if($debug){"tmpTime:  $tmpTime"}
        $newDate = (Get-Date -Format yyyy-MM-dd).ToString()
        $tmpItem = $newDate + "T" + $tmpTime
        $item.until = $tmpItem.ToString()

    }
    " + Updating Scheduled Time of Arrival"
    $dateArray = $XmlDocument.FlightPlan.FlightPlanSummary
    foreach($item in $dateArray){
        $tmpSplit = $item.ScheduledTimeOfArrival.ToString()
        $timeArray = $tmpSplit.split("T")
        $tmpDate = $timeArray[0]
        if($debug){"tmpDate:  $tmpDate"}
        $tmpTime = $timeArray[1]
        if($debug){"tmpTime:  $tmpTime"}
        $newDate = (Get-Date -Format yyyy-MM-dd).ToString()
        $tmpItem = $newDate + "T" + $tmpTime
        $item.ScheduledTimeOfArrival = $tmpItem.ToString()
    }

    " + Updating Estimated Out Time"
    $dateArray = $XmlDocument.FlightPlan.FlightPlanSummary.OutTime.EstimatedTime
    foreach($item in $dateArray){
        $tmpSplit = $item.value.ToString()
        $timeArray = $tmpSplit.split("T")
        $tmpDate = $timeArray[0]
        if($debug){"tmpDate:  $tmpDate"}
        $tmpTime = $timeArray[1]
        if($debug){"tmpTime:  $tmpTime"}
        $newDate = (Get-Date -Format yyyy-MM-dd).ToString()
        $tmpItem = $newDate + "T" + $tmpTime
        $item.value = $tmpItem.ToString()
    }

    " + Updating Estimated Off Time"
    $dateArray = $XmlDocument.FlightPlan.FlightPlanSummary.OffTime.EstimatedTime
    foreach($item in $dateArray){
        $tmpSplit = $item.value.ToString()
        $timeArray = $tmpSplit.split("T")
        $tmpDate = $timeArray[0]
        if($debug){"tmpDate:  $tmpDate"}
        $tmpTime = $timeArray[1]
        if($debug){"tmpTime:  $tmpTime"}
        $newDate = (Get-Date -Format yyyy-MM-dd).ToString()
        $tmpItem = $newDate + "T" + $tmpTime
        $item.value = $tmpItem.ToString()
    }

    " + Updating Estimated On Time"
    $dateArray = $XmlDocument.FlightPlan.FlightPlanSummary.OnTime.EstimatedTime
    foreach($item in $dateArray){
        $tmpSplit = $item.value.ToString()
        $timeArray = $tmpSplit.split("T")
        $tmpDate = $timeArray[0]
        if($debug){"tmpDate:  $tmpDate"}
        $tmpTime = $timeArray[1]
        if($debug){"tmpTime:  $tmpTime"}
        $newDate = (Get-Date -Format yyyy-MM-dd).ToString()
        $tmpItem = $newDate + "T" + $tmpTime
        $item.value = $tmpItem.ToString()
    }

    " + Updating Estimated In Time"
    $dateArray = $XmlDocument.FlightPlan.FlightPlanSummary.InTime.EstimatedTime
    foreach($item in $dateArray){
        $tmpSplit = $item.value.ToString()
        $timeArray = $tmpSplit.split("T")
        $tmpDate = $timeArray[0]
        if($debug){"tmpDate:  $tmpDate"}
        $tmpTime = $timeArray[1]
        if($debug){"tmpTime:  $tmpTime"}
        $newDate = (Get-Date -Format yyyy-MM-dd).ToString()
        $tmpItem = $newDate + "T" + $tmpTime
        $item.value = $tmpItem.ToString()
    }

    " + Updating Departure Date/Time"
    $dateArray = $XmlDocument.FlightPlan.CustomerExtensions.Departure
    foreach($item in $dateArray){
        $tmpSplit = $item.scheduledLocalTime.ToString()
        $timeArray = $tmpSplit.split("T")
        $tmpDate = $timeArray[0]
        if($debug){"tmpDate:  $tmpDate"}
        $tmpTime = $timeArray[1]
        if($debug){"tmpTime:  $tmpTime"}
        $newDate = (Get-Date -Format yyyy-MM-dd).ToString()
        $tmpItem = $newDate + "T" + $tmpTime
        $item.LocalDate = $newDate
        $item.scheduledLocalTime = $tmpItem.ToString()
        $item.EstimatedLocalTime = $tmpItem.ToString()
    }

    " + Updating Take Off Date/Time"
    $dateArray = $XmlDocument.FlightPlan.CustomerExtensions.Takeoff
    foreach($item in $dateArray){
        $tmpSplit = $item.EstimatedLocalTime.ToString()
        $timeArray = $tmpSplit.split("T")
        $tmpDate = $timeArray[0]
        if($debug){"tmpDate:  $tmpDate"}
        $tmpTime = $timeArray[1]
        if($debug){"tmpTime:  $tmpTime"}
        $newDate = (Get-Date -Format yyyy-MM-dd).ToString()
        $tmpItem = $newDate + "T" + $tmpTime
        $item.LocalDate = $newDate
        $item.EstimatedLocalTime = $tmpItem.ToString()
    }

    " + Updating Landing Date/Time"
    $dateArray = $XmlDocument.FlightPlan.CustomerExtensions.Landing
    foreach($item in $dateArray){
        $tmpSplit = $item.EstimatedLocalTime.ToString()
        $timeArray = $tmpSplit.split("T")
        $tmpDate = $timeArray[0]
        if($debug){"tmpDate:  $tmpDate"}
        $tmpTime = $timeArray[1]
        if($debug){"tmpTime:  $tmpTime"}
        $newDate = (Get-Date -Format yyyy-MM-dd).ToString()
        $tmpItem = $newDate + "T" + $tmpTime
        $item.LocalDate = $newDate
        $item.EstimatedLocalTime = $tmpItem.ToString()
        $item.LandingWindow.from = $tmpItem.ToString()
        $item.Landingwindow.until = $tmpItem.ToString()
    }

    " + Updating Arrival Date/Time"
    $dateArray = $XmlDocument.FlightPlan.CustomerExtensions.Arrival
    foreach($item in $dateArray){
        $tmpSplit = $item.scheduledLocalTime.ToString()
        $timeArray = $tmpSplit.split("T")
        $tmpDate = $timeArray[0]
        if($debug){"tmpDate:  $tmpDate"}
        $tmpTime = $timeArray[1]
        if($debug){"tmpTime:  $tmpTime"}
        $newDate = (Get-Date -Format yyyy-MM-dd).ToString()
        $tmpItem = $newDate + "T" + $tmpTime
        $item.LocalDate = $newDate
        $item.scheduledLocalTime = $tmpItem.ToString()
        $item.EstimatedLocalTime = $tmpItem.ToString()
    }

    try{$XmlDocument.Save($xmlFile)}
	catch{"! Unable to Save $xmlfile"}

}

function UpdateTP{
    "Updating TP"
    $xmlFile = $TempDatDir + "TPs.xml"
    $xmlFile = Get-ChildItem -Filter $xmlFile -Name
    $xmlFile = $TempDir + $TempDatDir + $xmlFile
    if($debug){"OFP XML: $xmlfile"}
    [xml]$XmlDocument = Get-Content -Path $xmlFile
    " + Updating Timestamp"
    $XmlDocument.HazardBriefing.M633Header.timestamp = $newDate = (Get-Date -Format yyyy-MM-ddThh:mm:ssZ).ToString()
    try{$XmlDocument.Save($xmlFile)}
	catch{"! Unable to Save $xmlfile"}
}

function UpdateWeather{
    "Updating METAR"
    $xmlFile = $TempDatDir + "AirportWeather.xml"
    $xmlFile = Get-ChildItem -Filter $xmlFile -Name
    $xmlFile = $TempDir + $TempDatDir + $xmlFile
    if($debug){"OFP XML: $xmlfile"}
    [xml]$XmlDocument = Get-Content -Path $xmlFile
    " + Updating Timestamp"
    $XmlDocument.AirportWeather.M633Header.timestamp = $newDate = (Get-Date -Format yyyy-MM-ddThh:mm:ssZ).ToString()
    
    $dateArray = $XmlDocument.AirportWeather.WeatherBulletins.WeatherBulletin.Observation
    " + Updating Observation Date"
    foreach($item in $dateArray){
        try {
        $tmpSplit = $item.observationTime.ToString()
        $timeArray = $tmpSplit.split("T")
        $tmpDate = $timeArray[0]
        if($debug){"tmpDate:  $tmpDate"}
        $tmpTime = $timeArray[1]
        if($debug){"tmpTime:  $tmpTime"}
        $newDate = (Get-Date -Format yyyy-MM-dd).ToString()
        $tmpItem = $newDate + "T" + $tmpTime
        $item.observationTime = $tmpItem.ToString()
        }
        catch{" - No Update Needed"}
    }

    " + Updating Forecast Date"
    $dateArray = $XmlDocument.AirportWeather.WeatherBulletins.WeatherBulletin.Forecast
    foreach($item in $dateArray){
        $tmpSplit1 = $item.forecastTime.ToString()
        $tmpSplit2 = $item.forecastStartTime.ToString()
        $tmpSplit3 = $item.forecastEndTime.ToString()
        
        $timeArray1 = $tmpSplit1.split("T")
        $timeArray2 = $tmpSplit2.split("T")
        $timeArray3 = $tmpSplit3.split("T")
        
        $tmpDate1 = $timeArray1[0]
        $tmpDate2 = $timeArray2[0]
        $tmpDate3 = $timeArray3[0]
                
        $tmpTime1 = $timeArray1[1]
        $tmpTime1 = $timeArray2[1]
        $tmpTime1 = $timeArray3[1]
        
        $newDate = (Get-Date -Format yyyy-MM-dd).ToString()
        
        $tmpItem1 = $newDate + "T" + $tmpTime1
        $tmpItem2 = $newDate + "T" + $tmpTime2
        $tmpItem3 = $newDate + "T" + $tmpTime3
                
        $item.forecastTime = $tmpItem1.ToString()
        $item.forecastStartTime = $tmpItem2.ToString()
        $item.forecastEndTime = $tmpItem3.ToString()

    }

    try{$XmlDocument.Save($xmlFile)}
	catch{"! Unable to Save $xmlfile"}
}

function UpdateCrewList{
    "Updating Crew List"
}

function UpdateWinds{
    "Updating Forecast Winds"
}

function updateCBP{
    # change date at top to current
}

function updateADR{
    # replace TODAY with today's date
    (Get-Content $TempDir + $TempDatDir + AircraftDiscrepancyReport.txt ).replace('[MYID]', 'MyValue') | Set-Content c:\temp\test.txt
}


function GetHash{
	param([string]$filename)
	$md5 = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
	$hash = $md5.ComputeHash([System.IO.File]::ReadAllBytes($filename))
	$res = [System.Convert]::ToBase64String($hash)
	return $res
}


function RepackDat{
	"Repacking DAT File"
	$zipfile = Get-ChildItem -Filter "*.dat" -Name
	$zipfile = $TempDir + $zipfile
	if($debug){"Dat File Name:  $zipfile"}
	$foldername = $TempDir + $TempDatDir
	if($debug){"Folder Name:  $foldername"}
	if($debug){"Removing Old DAT File"}
	Remove-Item($TempDir + "*.dat")
	[System.IO.Compression.ZipFile]::CreateFromDirectory($foldername, $zipfile)
	if($debug){"-------------------------------"}
}

function ComputeHashes{
	"Calculating Hashes"
	""
	$path = $TempDatDir + "*"
	$FileList = Get-ChildItem -Filter $path	-Name
	foreach($file in $FileList){
		Gethash($TempDir + $TempDatDir + $file)
	}
	if($debug){"-------------------------------"}
	""
}

function UpdateListFile{
	"Updating Checksum List"
	$xmlFile = Get-ChildItem -Filter "*.lst" -Name
	$xmlFile = $TempDir + $xmlFile
	if($debug){"HashFile: $xmlFile"}
	[xml]$XmlDocument = Get-Content -Path $xmlFile
	$list = $XmlDocument.hashfilelist
	foreach($item in $list.hashedfile){
		$dataname = $item.href
#		$dataname.ToString()
		if($debug){"Updating $dataname"}
		$newhash = GetHash($TempDir + $TempDatDir + $dataname)
		if($debug){"New Hash:  $newhash"}
		$curelement = $item.InnerText
		if($debug){"Old Hash:  $curelement"}
		$item.InnerText = $newhash
	}
	try{$XmlDocument.Save($xmlFile)}
	catch{"! Unable to Save $xmlfile"}
	
}

function SignDatFile{
	"Signing DAT File"
	$xmlFile = Get-ChildItem -Filter "*.lst" -Name
	$xmlFile = $TempDir + $xmlFile
	$datfile = Get-ChildItem -Filter "*.dat" -Name
	$datFile = $TempDir + $datFile
	if($debug){"HashFile: $xmlFile"}
	if($debug){"DATFile: $datfile"}
	[xml]$XmlDocument = Get-Content -Path $xmlFile
	$checkcode = $XmlDocument.hashfilelist.checkcode
	if($debug){"Old Hash: $checkcode"}
	$newhash = GetHash($datfile)
	if($debug){"New Hash:  $newhash"}
	$xmlDocument.hashfilelist.checkcode = $newhash
	
	try{$XmlDocument.Save($xmlFile)}
	catch{"Unable to Save $xmlfile"}

}

function GenerateEFF{
	"Generating New EFF File"
#	$efffile = NewName($oldname)
	$efffile = Get-ChildItem -Filter "*.eff" -Name
	$newname = NewName($efffile)
	$newname = $TempDir + $newname
	$efffile = $TempDir + $efffile
	if($debug){"Original EFF File Name:  $efffile"}
	if($debug){"New EFF File Name:  $newname"}
	$foldername = $TempDir + $TempEFFDir
	if($debug){"Folder Name:  $foldername"}
	if($debug){"Removing Old EFF File"}
	Remove-Item($TempDir + "*.eff")
	New-Item $foldername -itemtype directory | Out-Null
	if($debug){"Moving files to working directory"}
	$tmpdat = $TempDir + "*.dat"
	$tmplst = $TempDir + "*.lst"
	$tmpeff = $TempDir + $TempEFFDir
	Move-Item -Path $tmpdat -Destination $tmpeff
	Move-Item -Path $tmplst -Destination $tmpeff
	[System.IO.Compression.ZipFile]::CreateFromDirectory($foldername, $newname)
	if($debug){"-------------------------------"}
}

function CleanUp{
	try {
        Remove-Item($TempDir + $TempDatDir) -Recurse
        } catch { "! DAT folder not found" }
    try {
	    Remove-Item($TempDir + "*.dat")
	    Remove-Item($TempDir + "*.lst")
        } catch { "! DAT and LST files not found" }
    try {
	    Remove-Item($TempDir + $TempEFFDir) -Recurse
        } catch { "! EFF Folder not found" }
}

function NewName($oldname){
	$newDate = Get-Date -Format yyMMdd
	$newDate = $newDate.ToString()
	$newTime = Get-Date -Format hhmmss
	$newtime = $newtime.ToString()
	$tmpName = $oldname.split("_")
	$tailNum = $tmpName[0]
	$flgtNum = $tmpName[1]
	$effdate = $newDate
	$cityOne = $tmpName[3]
	$cityTwo = $tmpName[4]
	$effdatetime = "$newDate$newTime"
	$newname = $tailNum + "_" + $flgtNum + "_" + $effdate + "_" + $cityOne + "_" + $cityTwo + "_F" + $effdatetime + ".eff"
	return $newname
}

if(!$args){
    "EFF Updater v1.0"
    "Author:  Aaron Anderson <aaron.d.anderson@delta.com>"
    "July 26, 2018"
    ""
    "This program has been written to facilitate updating"
    "Electronic Flight Folder files for use in the A350 Flight Simulator"
    ""
    "Usage:  UpdateDate.ps1 <effname> [options]"
    ""
    "Command Line Options:"
    "-unpack   Unpacks the EFF file for editing"
    "-repack   Repacks edited EFF back into usable .eff file"
    "-clean    Cleans up unpacked files"

} else {

    switch($args[1]){
        "-unpack"{
            UnpackEff($args[0])
            ExtractDat
        }
        "-repack"{
            ComputeHashes
            UpdateListFile
            RepackDat
            SignDatFile
            GenerateEFF
            CleanUp
        }
        "-clean"{
            Cleanup
        }
        "-test"{
            # put test functions here
            UpdateWeather
        }

        "" {
            UnpackEff($args[0])
            ExtractDat
            ReplaceEffDate
            UpdateOFP
            UpdateTP
            ComputeHashes
            UpdateListFile
            RepackDat
            SignDatFile
            GenerateEFF
            CleanUp
        }
    }
    
}