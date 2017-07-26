$tempDirectory = $HOME + '\Scripts\Temp\'
$resourcesDirectory = $HOME + '\Scripts\Resources\'

<# Get the list of highest priority divisions #>
$hiPriDivsFile = $resourcesDirectory + 'priority.csv'
$hiPriDivs = Import-Csv -Path $hiPriDivsFile

<# Get a temporary copy of the master Assignment Tracker form the SharePoint Server #>
$trackerTempSave = $tempDirectory + 'master_tracker.xlsx'
Invoke-WebRequest -UseDefaultCredentials -URI "http://compass.repsrv.com/KMA Team Division Documents/Division Trackers/Assignment Tracker by Wave.xlsx" -OutFile $trackerTempSave

<# Get a list of the worksheet names #>
$worksheetList = Get-ExcelSheetInfo $trackerTempSave | Select-Object "Name"

<# Build new master table with all documents #>

$masterTable = @()

$worksheetList | ForEach-Object {
    $currentObject = $_
    $tempTable = Import-Excel $trackerTempSave -WorkSheetname $currentObject.Name
    $masterTable += $tempTable
}

$masterTable | ForEach-Object {
    $date = $_.'Date Ready for HTML' -as [DateTime]
    if (!$date) { $_.'Date Ready for HTML' = $null }
}

$masterIndexFilePath = $tempDirectory + 'master-index.xlsx'
$masterTable | Export-Excel -Path $masterIndexFilePath -WorkSheetname "Master"

<# Find any unassigned documents #>

$unassignedBuild = @()


Function GrabOpenAssignments ([Int]$numToGrab) {
    if ($numToGrab -ne $null) {
        $tempTable = $masterTable | Where-Object { $_."HTML Build Assigned" -eq $null } | Where-Object { $_."Date Ready for HTML" -ne $null } | Select-Object -First $numToGrab
        $unassignedBuild += $tempTable
    } else {
        $hiPriDivs | ForEach-Object {
            $currentObject = $_."Priority Divisions"
            $tempTable = $masterTable | Where-Object { $_."HTML Build Assigned" -eq $null } | Where-Object { $_."Date Ready for HTML" -ne $null } | Where-Object { $_.Division -eq $currentObject }
            $unassignedBuild += $tempTable
        }
    }
    $outputFileLocation = $tempDirectory + 'Assignments to Grab.xlsx'
    $outputTable = $unassignedBuild | Select-Object "Division_Documents","Waves","Area","Division","Lawson","SME Scrub Assigned","Date Ready for HTML"
    $outputTable | Export-Excel $outputFileLocation -WorkSheetname "Personal Tracker"
    start $outputFileLocation
}

Function RefreshMasterTable {
    $trackerTempSave = $tempDirectory + 'master_tracker.xlsx'
    Invoke-WebRequest -UseDefaultCredentials -URI "http://compass.repsrv.com/KMA Team Division Documents/Division Trackers/Assignment Tracker by Wave.xlsx" -OutFile $trackerTempSave
    
    <# Get a list of the worksheet names #>
    $worksheetList = Get-ExcelSheetInfo $trackerTempSave | Select-Object "Name"
    
    <# Build new master table with all documents #>
    
    $masterTable = @()
    
    $worksheetList | ForEach-Object {
        $currentObject = $_
        $tempTable = Import-Excel $trackerTempSave -WorkSheetname $currentObject.Name
        $masterTable += $tempTable
    }
    $masterIndexFilePath = $tempDirectory + 'master-index.xlsx'
    $masterTable | Export-Excel -Path $masterIndexFilePath -WorkSheetname "Master"
}
