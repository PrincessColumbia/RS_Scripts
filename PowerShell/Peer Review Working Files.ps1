$peerReviewWorkingFile = $HOME + '\Documents\Personal Change Log Tracking.xlsx'
$templateFolder = $HOME + '\Scripts\Resources\Doc_Templates\'
$defaultTemplate = $templateFolder + 'Residential_Contract_or_OM.html'
#$filesLocation = '\\repsharedfs\share\Customer Experience\Compass\CE Analyst Team Files\Site Visit Conversions\'
$filesLocation = 'Z:\Site Visit Conversions\'
$tempFileList = $HOME + '\Scripts\Temp\temp-file-list.csv'

Add-Type -assembly System.IO.Compression.FileSystem

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

$peerReviewTable = Import-Excel $peerReviewWorkingFile -WorkSheetname 'Peer Review Tracker'


$peerReviewTable | Add-Member -MemberType NoteProperty -Name "Current_Location" -Value $null
$peerReviewTable | Add-Member -MemberType NoteProperty -Name "Excel_File" -Value $null
$peerReviewTable | Add-Member -MemberType NoteProperty -Name "HTML_File" -Value $null
$peerReviewTable | Add-Member -MemberType NoteProperty -Name "DIV_Directory" -Value $null
$peerReviewTable | Add-Member -MemberType NoteProperty -Name "HTML_Destination" -Value $null
$peerReviewTable | Add-Member -MemberType NoteProperty -Name "Built_Destination" -Value $null
$peerReviewTable | Add-Member -MemberType NoteProperty -Name "Excel_Destination" -Value $null
$peerReviewTable | Add-Member -MemberType NoteProperty -Name "Excel_Moved" -Value $null
$peerReviewTable | Add-Member -MemberType NoteProperty -Name "Error" -Value $null

$peerReviewTable | ForEach-Object {
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
    $excelFileVerbose = $excelFileVerbose | Where-Object { $_.Directory -match 'Peer review' }
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
    $pdfFileVerbose = $fileSearchResults | Where-Object { [System.IO.Path]::GetExtension($_) -eq '.pdf' }
    $pdfFile = $pdfFileVerbose.FullName
    if ( $excelFile -eq $null ) {
        $currentObject.Current_Location = $pdfFile
    } else {
        $currentObject.Current_Location = $excelFile
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
        $currentObject.Built_Destination = $excelMoveAfterBuildPath
        $currentObject.Excel_Destination = $excelMoveFile
    } else {
        New-Item $excelMoveAfterBuildPath -ItemType Directory
        $currentObject.Built_Destination = $excelMoveAfterBuildPath
        $currentObject.Excel_Destination = $excelMoveFile
    }
    if ( $currentObject.Type -eq 'PDF' ) {
        $currentObject.Excel_File = 'N/A'
        $currentObject.HTML_File = 'N/A'
        $currentObject.Excel_Destination = 'N/A'
    }
    $testCurrentPath = $divSearchDirectory + '2 - HTML Build\' + $currentObject.Document_Name + '.xlsx'
    $testResult = $currentObject.Current_Location -eq $testCurrentPath
    if ( $testResult -eq 'True' ) {
        $currentObject.Excel_Moved = 'Moved'
    }
    if ($currentObject.Type -eq 'PDF') {
        $currentObject.Excel_Moved = 'PDF, N/A'
    }
}


$openPeerReview = $peerReviewTable | Where-Object { $_.'Peer review complete' -eq 'Incomplete' }

Function PrepLawson ($lawson) {
    $workTable | Where-Object { $_.Lawson -eq $lawson } | Where-Object { $_.Type -eq 'HTML' } | Where-Object { $_.Excel_File -ne 'MISSING/MIS-NAMED EXCEL FILE' } | Where-Object { $_."Document Built" -eq 'Incomplete' } | ForEach-Object {
        $currentObject = $_
        start $currentObject.Excel_File
        $templateToUse = $templateFolder + $currentObject.Template + '.html'
        $fileLocationPath = $currentObject.HTML_File
        $fileExistsTest = Test-Path $fileLocationPath
        if ( $fileExistsTest -ne 'True' ) {
            Copy-Item $templateToUse $currentObject.HTML_File
            $file = Get-Item $currentObject.HTML_File
            $file.LastWriteTime = (Get-Date)
        }
        start $_.HTML_File
    }
}
