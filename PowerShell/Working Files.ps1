$workingFile = $HOME + '\Documents\Personal Build Log Tracker.xlsx'
$templateFolder = $HOME + '\Scripts\Resources\Doc_Templates\'
$defaultTemplate = $templateFolder + 'Residential_Contract_or_OM.html'
$filesLocation = '\\repsharedfs\share\Customer Experience\Compass\CE Analyst Team Files\Site Visit Conversions\'

<# Create a list of the available templates #>
$templateList = @(
"Billing_and_Collection_Information",
"Commercial_Pricing",
"Commercial_Service_Frequency_By_City_Town",
"Data_Entry_Information",
"Disposal_Notes",
"High_Priority_Accounts",
"Industrial_Pricing",
"Perm_Container_Information",
"Residential_Contract_or_OM",
"Residential_Contract_or_OM_2Column",
"Residential_Contract_or_OM_3Column",
"Residential_Contract_or_OM_4Column",
"Service_Commitments",
"Site_Information_Solid_Waste_Districts",
"Temp_Container_Information"
)

$workTable = Import-Excel -Path $workingFile -WorkSheetname 'Master'

$workTable | Add-Member -MemberType NoteProperty -Name "Excel_File" -Value $null
$workTable | Add-Member -MemberType NoteProperty -Name "HTML_File" -Value $null
$workTable | Add-Member -MemberType NoteProperty -Name "DIV_Directory" -Value $null
$workTable | Add-Member -MemberType NoteProperty -Name "HTML_Destination" -Value $null
$workTable | Add-Member -MemberType NoteProperty -Name "Excel_Destination" -Value $null
$workTable | Add-Member -MemberType NoteProperty -Name "Error" -Value $null

$workTable | ForEach-Object {
    $currentObject = $_

    $lawsonFilter = '*' + $currentObject.Lawson + '*'
    $divFilter = '*' + $currentObject.Division + '*'
    try { $areaSearch = $currentObject.Area.Substring(0,3) + '*' }
    catch {
        if ( $Error[0].Exception.Message.Contains("null-valued expression")) {
            $currentObject.Error
        }
    }
    $areaDirectory = Get-ChildItem $filesLocation -Filter $areaSearch -Directory | Select-Object Name | Select-Object -ExpandProperty "Name"
    $areaSearchDirectory = $filesLocation + $areaDirectory + '\'
    $divSearch = '*' + $currentObject.Division + '*'
    $divDirectory = Get-ChildItem $areaSearchDirectory -Filter $divSearch -Directory | Select-Object Name | Select-Object -ExpandProperty "Name"
    if ( $divDirectory.Count -ne 1 ) {
        $divDirectory = $divDirectory -like $lawsonFilter
    }
    $divSearchDirectory = $areaSearchDirectory + $divDirectory + '\'
    $fileSearch = $currentObject.Document_Name + '*'
    try { $fileSearchResults = Get-ChildItem $divSearchDirectory -Filter $fileSearch -Recurse -ErrorAction Stop | Where-Object { $_.Directory -notmatch 'archive' } }
    catch {
    if ($Error[0].Exception.Message.Contains("Cannot find path")) {
        $currentObject.Error = "Incorrect path"
        }
    }
    $excelFileVerbose = $fileSearchResults | Where-Object { [System.IO.Path]::GetExtension($_) -eq '.xlsx' }
    $excelFileVerbose = $excelFileVerbose | Where-Object { $_.Directory -match 'build' }
    $excelFile = $excelFileVerbose.FullName
    if ( $excelFile -eq $null ) {
        $currentObject.Excel_File = 'MISSING/MIS-NAMED EXCEL FILE'
    } else {
        $testExcel = Test-Path $excelFile
        if ($testExcel -eq 'True' ) {
            $currentObject.Excel_File = $excelFile
        } else {
            $currentObject.Excel_File = 'MISSING/MIS-NAMED EXCEL FILE'
        }
    }
    $currentObject.DIV_Directory = $divSearchDirectory
    $htmlDirectoryPath = $divSearchDirectory + '3 - Peer Review\'
    $htmlDirectoryPathTest = Test-Path $htmlDirectoryPath
    if ($htmlDirectoryPathTest -eq 'True') {
        $currentObject.HTML_Destination = $htmlDirectoryPath
    } else {
        New-Item $htmlDirectoryPath -ItemType Directory
        $currentObject.HTML_Destination = $htmlDirectoryPath
    }
    $currentObject.HTML_File = $htmlDirectoryPath + $currentObject.Document_Name + '.html'
    $excelMoveAfterBuildPath = $divSearchDirectory + '2 - HTML Build\Built to HTML\'
    $excelMoveAfterBuildPathTest = Test-Path $excelMoveAfterBuildPath
    $excelMoveFile = $excelMoveAfterBuildPath + $currentObject.Document_Name + '.xlsx'
    if ($excelMoveAfterBuildPathTest -eq 'Yes') {
        $currentObject.Excel_Destination = $excelMoveFile
    } else {
        New-Item $excelMoveAfterBuildPath -ItemType Directory
        $currentObject.Excel_Destination = $excelMoveFile
    }
    if ( $currentObject.Type -eq 'PDF' ) {
        $currentObject.Excel_File = 'N/A'
        $currentObject.DIV_Directory = 'N/A'
        $currentObject.HTML_Destination = 'N/A'
        $currentObject.Excel_Destination = 'N/A'
    }
}

Function PrepLawson ($lawson) {
    $workTable | Where-Object { $_.Lawson -eq $lawson } | Where-Object { $_.Type -eq 'HTML' } | Where-Object { $_.Excel_File -ne 'MISSING/MIS-NAMED EXCEL FILE' } | ForEach-Object {
        start $_.Excel_File
        $templateToUse = $templateFolder + $_.Template + '.html'
        Copy-Item $templateToUse $_.HTML_File
        start $_.HTML_File
    }
}

