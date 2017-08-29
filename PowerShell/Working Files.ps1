﻿$workingFile = $HOME + '\Documents\Personal Build Log Tracker.xlsx'
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

$workTable = Import-Excel -Path $workingFile -WorkSheetname 'Master'

$workTable | Add-Member -MemberType NoteProperty -Name "Current_Location" -Value $null
$workTable | Add-Member -MemberType NoteProperty -Name "Excel_File" -Value $null
$workTable | Add-Member -MemberType NoteProperty -Name "HTML_File" -Value $null
$workTable | Add-Member -MemberType NoteProperty -Name "DIV_Directory" -Value $null
$workTable | Add-Member -MemberType NoteProperty -Name "HTML_Destination" -Value $null
$workTable | Add-Member -MemberType NoteProperty -Name "Built_Destination" -Value $null
$workTable | Add-Member -MemberType NoteProperty -Name "Excel_Destination" -Value $null
$workTable | Add-Member -MemberType NoteProperty -Name "Excel_Moved" -Value $null
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

Function GetFileTypes {
    $fileCheckList = $workTable | Where-Object { $_.Type -eq $null } | Select-Object "Document_Name" | ForEach-Object { $_."Document_Name" }
    $fileCheckList | ForEach-Object {
        $currentObject = $_
        $directoryToSearch = $workTable | Where-Object { $_.Document_Name -eq $currentObject } | Select-Object -ExpandProperty "Div_Directory"
        $filenameSearch = $currentObject + '*'
        $searchResults = Get-ChildItem -Path $directoryToSearch -Filter $filenameSearch -Recurse
    }
}

Function FastHTMLOpen {
    $tempOpen = Import-Csv -Path $tempFileList
    $tempOpen.Document_Name | ForEach-Object {
        $currentObject = $_
        $fileToOpen = $workTable | Where-Object { $_.Document_Name -eq $currentObject } | Select-Object -ExpandProperty "HTML_File"
        try { start $fileToOpen }
        catch {
            if ( $Error[0].Exception.Message.Contains('Cannot validate argument') -or $Error[0].Exception.Message.Contains('cannot find the file') ) {
                Write-Host 'Cannot open the HTML file for' $currentObject
            }
        }
    }
}

Function FastExcelOpen {
    $tempOpen = Import-Csv -Path $tempFileList
    $tempOpen.Document_Name | ForEach-Object {
        $currentObject = $_
        $excelTest = $currentObject.Excel_File
        if ( $excelTest -eq "MISSING/MIS-NAMED EXCEL FILE" ) {
            Write-Host 'Unable to open Excel file for' $currentObject
        } else {
            $fileToOpen = $workTable | Where-Object { $_.Document_Name -eq $currentObject } | Select-Object -ExpandProperty "Excel_File"
            try { start $fileToOpen }
            catch {
                if ( $Error[0].Exception.Message.Contains('Cannot validate argument') -eq 'True' ) {
                    Write-Host 'Unable to open Excel file for' $currentObject
                }
            }
        }
    }
}

Function BuildOpen ($docName) {
    $currentObject = $workTable | Where-Object { $_.Document_Name -eq $docName }
    $excelTest = $currentObject.Excel_File
    if ( $excelTest -eq "MISSING/MIS-NAMED EXCEL FILE" ) {
            Write-Host 'Unable to open Excel file for' $docName
        } else {
            $fileToOpen = $currentObject.Excel_File
            try { start $fileToOpen }
            catch {
            if ( $Error[0].Exception.Message.Contains('Cannot validate argument') -eq 'True' ) {
                Write-Host 'Unable to open Excel file for' $currentObject
            }
       }
    }
    $fileToOpen = $currentObject.HTML_File
    try { start $fileToOpen }
        catch {
        if ( $Error[0].Exception.Message.Contains('Cannot validate argument') -eq 'True' ) {
            Write-Host 'Cannot open the HTML file for' $currentObject
        }
    }
}

Function SeeDetails ($docName) {
    $workTable | Where-Object { $_.Document_Name -eq $docName }
}

Function CreateHTMLFile ($docName) {
    $workTable | Where-Object { $_.Document_Name -eq $docName } | ForEach-Object {
        $currentObject = $_
        $templateToUse = $templateFolder + $currentObject.Template + '.html'
        $fileLocationPath = $currentObject.HTML_File
        $fileExistsTest = Test-Path $fileLocationPath
        if ( $fileExistsTest -ne 'True' ) {
            Copy-Item $templateToUse $currentObject.HTML_File
        }
        start $currentObject.HTML_File
    }
}

Function CopyPDFs {
    $workTable | Where-Object { $_."Date ready for Peer Review" -eq $null -and $_.Type -eq 'PDF' } | Where-Object { $_."Feedback Needed" -eq $null } | ForEach-Object {
        $currentObject = $_
        $currentFileName = $currentObject.Document_Name + '.pdf'
        $targetDirectoryBuilt = 
        if ( $currentObject.HTML_Destination -eq 'N/A' -or $currentObject.Current_Location -eq $null ) {
            Write-Host 'Unable to move' $currentFileName -BackgroundColor Red -ForegroundColor Yellow
        } else {
            Copy-Item $currentObject.Current_Location $currentObject.HTML_Destination
            Move-Item $currentObject.Current_Location $currentObject.Built_Destination
        }
    }
}

Function ChangeToLawsonDirectory ($lawNum) {
    $targetDir = ($workTable | Where-Object { $_.Lawson -eq $lawNum })[0].DIV_Directory
    cd $targetDir
}


Function ListExcelToMove ($lawNum) {
    $listSet = $workTable | Where-Object { $_.Lawson -eq $lawNum } | Where-Object { $_.Excel_Moved -eq $null } | Select-Object Document_Name | Sort
    $listSet
    $listSet.Count
}

Function PostBuildMoveTest ($lawNum) {
    $listSet = $workTable | Where-Object { $_.Lawson -eq $lawNum } | Where-Object { $_."Document Built" -eq "Built" }
    $listSet | ForEach-Object {
        $currentObject = $_
        $isFileMoved = Test-Path $currentObject.Excel_Destination
        if (!$isFileMoved) {
            $currentObject.Document_Name
        }
    }
}

Function PostBuildMove {
    $listSet = $workTable | Where-Object { $_."Document Built" -eq "Built" }
    $listSet | ForEach-Object {
        $currentObject = $_
        $isFileMoved = Test-Path $currentObject.Excel_Destination
        if (!$isFileMoved) {
            Move-Item $currentObject.Current_Location $currentObject.Excel_Destination
        }
    }
}

Function TestPreBuildMove ($lawNum) {
    $listSet = $workTable | Where-Object { $_."Document Built" -ne "Built" } | Where-Object { $_.Lawson -eq $lawNum }
    $listSet | Where-Object { $_.Excel_Moved -eq $null } | Select-Object Document_Name,Current_Location
}


<#
Function BrowserHTMLTest ($lawnum) {
    $listSet = $workTable | Where-Object { $_."Document Built" -ne "Built" } | Where-Object { $_.Lawson -eq $lawNum }
    $listSet | Where-Object { $_.Excel_Moved -eq $null } | Select-Object Document_Name,Current_Location

}
#>

Function ExcelOpen ($docName) {
    $currentObject = $workTable | Where-Object { $_.Document_Name -eq $docName }
    $excelTest = $currentObject.Excel_File
    if ( $excelTest -eq "MISSING/MIS-NAMED EXCEL FILE" ) {
            Write-Host 'Unable to open Excel file for' $docName
        } else {
            $fileToOpen = $currentObject.Excel_File
            try { start $fileToOpen }
            catch {
            if ( $Error[0].Exception.Message.Contains('Cannot validate argument') -eq 'True' ) {
                Write-Host 'Unable to open Excel file for' $currentObject
            }
       }
    }
}

Function HTMLOpen ($docName) {
    $currentObject = $workTable | Where-Object { $_.Document_Name -eq $docName }
    $fileToOpen = $currentObject.HTML_File
    try { start $fileToOpen }
        catch {
        if ( $Error[0].Exception.Message.Contains('Cannot validate argument') -eq 'True' ) {
            Write-Host 'Cannot open the HTML file for' $currentObject
        }
    }
}

Function QuickStart ($fileName) {
    $fullFileName = $fileName + '.xlsx'
    start $fullFileName
}

Function BrowserTest {
    $tempOpen = Import-Csv -Path $tempFileList
    $tempOpen.Document_Name | ForEach-Object {
        $currentObject = $_
        $fileToOpen = $workTable | Where-Object { $_.Document_Name -eq $currentObject } | Select-Object -ExpandProperty "HTML_File"
        $fixedFilePath = '"' + $fileToOpen + '"'
        $chromeParams = " `'--profile-directory=`"Default`" --start-maximized " + $fixedFilePath + "`'"
        try { Start-Process "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" '--profile-directory="Default" --start-maximized $fileToOpen' $fileToOpen }
        catch {
            if ( $Error[0].Exception.Message.Contains('Cannot validate argument') -or $Error[0].Exception.Message.Contains('cannot find the file') ) {
                Write-Host 'Cannot open the HTML file for' $currentObject
            }
        }
    }
}

Function ExcelPath ($fileName) {
    $fileList = Get-ChildItem -Recurse | Where-Object { $_.Name -match $fileName }
    $fileList = $fileList | Where-Object { $_.DirectoryName -notmatch 'Archive' }
    $result = $fileList | Where-Object { $_.Extension -match '.xlsx' }
    $result = $result.FullName -replace 'z:','\\repsharedfs\share\Customer Experience\Compass\CE Analyst Team Files'
    Write-Host 'Excel file:' $result
}