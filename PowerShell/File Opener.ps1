#$dreamweaver = "C:\Program Files\Adobe\Adobe Dreamweaver CC 2015\Dreamweaver.exe"
$indexFile = 'C:\Users\mcfarch\Documents\Personal Change Log Tracking.xlsx'
$filesLocation = 'Z:\Site Visit Conversions\'
$moveToDirectory = '~HTML Ready for Upload\'
#$todaysDate = 
$dataEntryPerson = "Chris M."

$personalAssignmentsTracker = Import-Excel -Path $indexFile -WorkSheetname "Peer Review Tracker"
$personalAssignmentsTracker | Add-Member -MemberType NoteProperty -Name "Area_Directory" -Value $null
$personalAssignmentsTracker | Add-Member -MemberType NoteProperty -Name "Div_Directory" -Value $null
$personalAssignmentsTracker | Add-Member -MemberType NoteProperty -Name "Excel_File" -Value $null
$personalAssignmentsTracker | Add-Member -MemberType NoteProperty -Name "HTML_File" -Value $null
$personalAssignmentsTracker | Add-Member -MemberType NoteProperty -Name "Upload_Directory" -Value $null
$personalAssignmentsTracker | Add-Member -MemberType NoteProperty -Name "Error" -Value $null

$personalAssignmentsTracker | ForEach-Object {
    
    $currentObject = $_

    if ( $_.Document_Name -eq $null ) {
        $_."Peer review complete" = "Complete"
    }


    if ( $_."Peer review complete" -eq 'Complete' ) {
        $doneToken = 1
    } else {
        $doneToken = 0
    }
    $lawsonFilter = '*' + $_.Lawson + '*'
    $nameTest = $_.Document_Name.Length + 5
    try { $areaSearch = $_.Area.Substring(0,3) + '*' }
    catch {
        if ( $Error[0].Exception.Message.Contains("null-valued expression")) {
            $currentObject.Error
        }
    }
    $areaDirectory = Get-ChildItem $filesLocation -Filter $areaSearch -Directory | Select-Object Name | Select-Object -ExpandProperty "Name"
    $_.Area_Directory = $areaDirectory
    $areaSearchDirectory = $filesLocation + $areaDirectory + '\'
    $divSearch = '*' + $_.Division + '*'
    $divDirectory = Get-ChildItem $areaSearchDirectory -Filter $divSearch -Directory | Select-Object Name | Select-Object -ExpandProperty "Name"
    $_.Div_Directory = $divDirectory
    if ( $_.Div_Directory.Count -ne 1 ) {
            $_.Div_Directory = $_.Div_Directory -like $lawsonFilter
        } else {
            $_.Div_Directory = $_.Div_Directory
        }
    $divSearchDirectory = $areaSearchDirectory + $_.Div_Directory + '\'
    $fileSearch = $_.Document_Name + '*'
    try { $fileSearchResults = Get-ChildItem $divSearchDirectory -Filter $fileSearch -Recurse -ErrorAction Stop | Where-Object { ($_.Directory -notmatch 'archive') -and ($_.Name.Length -eq $nameTest) } }
    catch { 
        if ($Error[0].Exception.Message.Contains("Cannot find path")) {
            $currentObject.Error = "Incorrect path"
        }
    }
    $excelFileVerbose = $fileSearchResults | Where-Object { [System.IO.Path]::GetExtension($_) -eq '.xlsx' }
    $excelFileVerbose = $excelFileVerbose | Where-Object { $_.Directory -match 'built' }
    $excelFile = $excelFileVerbose.FullName
    if ( $excelFile -eq $null ) {
        $_.Excel_File = 'Missing/Mis-named Excel File'
    } else {
        $_.Excel_File = $excelFile
    }

    if ( $doneToken -eq 0 ) {
        $HTMLFileVerbose = $fileSearchResults | Where-Object { [System.IO.Path]::GetExtension($_) -eq '.html' }
        $HTMLFileVerbose = $HTMLFileVerbose | Where-Object { $_.Directory -match 'peer' }
        $HTMLFile = $HTMLFileVerbose.FullName
        if ( $HTMLFile -eq $null ) {
            $_.HTML_File = 'File moved to Upload'
        } else {
            $_.HTML_File = $HTMLFile
        }
    } else {
        $HTMLFileVerbose = $fileSearchResults | Where-Object { [System.IO.Path]::GetExtension($_) -eq '.html' }
        $HTMLFileVerbose = $HTMLFileVerbose | Where-Object { $_.Directory -match 'peer' }
        $HTMLFile = $HTMLFileVerbose.FullName
        if ( $HTMLFile -eq $null ) {
            $_.HTML_File = 'File missing or has a bad file name'
        } else {
            $_.HTML_File = $HTMLFile
        }
    }

    $moveToLocation = $filesLocation + $_.Area_Directory + '\' + $moveToDirectory
    $_.Upload_Directory = $moveToLocation
}

$badPath = $personalAssignmentsTracker | Where-Object { $_.Error -eq "Incorrect Path" }

$openAssignments = $personalAssignmentsTracker | Where-Object { $_."Peer review complete" -ne 'Complete' }

$finishedAssignments = $personalAssignmentsTracker | Where-Object { $_."Checked In" -eq 'No' }

Function OpenLawson ($lawsonToOpen) {
    $openAssignments | Where-Object { $_.Lawson -eq $lawsonToOpen } | ForEach-Object {
        $currentObject = $_
        try { Start-Process $currentObject.Excel_File }
        catch {
            if ( $Error[0].Exception.Message.Contains("The system cannot find the file specified.")) {
                $excelName = $currentObject.Document_Name + '.xlsx'
                Write-Host $excelName '- Unable to open the file'
            }
        }
        try { Start-Process $currentObject.HTML_File }
        catch {
            if ( $Error[0].Exception.Message.Contains("The system cannot find the file specified.")) {
                $htmlName = $currentObject.Document_Name + '.html'
                Write-Host $htmlName '- Unable to open the file'
            }
        }
    }
}

Function OpenLawsonHTML ($lawsonToOpen) {
    $openAssignments | Where-Object { $_.Lawson -eq $lawsonToOpen } | ForEach-Object {
        $currentObject = $_
        try { Start-Process $currentObject.HTML_File }
        catch {
            if ( $Error[0].Exception.Message.Contains("The system cannot find the file specified.")) {
                $htmlName = $currentObject.Document_Name + '.html'
                Write-Host $htmlName '- Unable to open the file'
            }
        }
    }
}


Function PROpen ($docToOpen) {
    $openAssignments | Where-Object { $_.Document_Name -eq $docToOpen } | ForEach-Object {
        Start-Process $_.Excel_File
        Start-Process $_.HTML_File
    }
}

Function RefreshMenu {
    $menuBuild = Import-Excel 'C:\Users\mcfarch\Documents\Nave Menu Builder for Commercial Pricing.xlsx' -WorkSheetname "Variable Builder"
    $menuList = $menuBuild."Menu Builder" | Select-Object -Unique
    $menuList
    $menuList | ForEach-Object {
        $menuItem = $_
        $menuItem
        $menuBuild | Where-Object { $_."Menu Builder" -eq $menuItem } | Select-Object "Nav Builder" -Unique
    }
}

Function MoveCompleted {
    $finishedAssignments | ForEach-Object {
        $currentObject = $_
        $uploadReadyPath = $currentObject.Upload_Directory + $currentObject.Document_Name + '.html'
        $uploadReadyCheck = Test-Path $uploadReadyPath
        if ( $uploadReadyCheck ) {
            Write-Host $currentObject.Document_Name 'already in Ready for Upload'
        } else {
            Move-Item $_.HTML_File $_.Upload_Directory
        }
    }
}

Function ShowMissingExcel ($divNum) {
    $missingExcelTable = $openAssignments | Where-Object { $_.Excel_File -eq $null } | Select-Object "Document_Name","Area","Division","Lawson"
    if ($divNum -ne $null) {
        $missingExcelTable | Where-Object { $_.Division -eq $divNum }
    } else {
        $missingExcelTable
    }
}

Function ExcelFilenameFixer {
    $excludeDirectory = Get-ChildItem -Directory -Recurse ~*
    $excludeShortcuts = Get-ChildItem -Recurse *.lnk
    $allExclude = @($excludeDirectory) + @($excludeShortcuts)
    $exclude = $allExclude | ForEach-Object { $_.Name } | Select-Object -Unique
    Get-ChildItem -Exclude $exclude -Recurse ~* | Rename-Item -NewName { $_.Name -replace "~","" }
}

Function BrowserPreview ($openValue) {
    if ( $openValue.Length -eq 0 ) {
        $openValue = Read-Host "Enter the Lawson number or name of the document(s) to open in the browser"
    }
    if ( ($openValue).tostring().length -eq 4 ) {
        $openAssignments | Where-Object { $_.Lawson -eq $openValue } | ForEach-Object {
            $fileToOpen = '"' + $_.HTML_File + '"'
            try { Start-Process chrome $fileToOpen }
            catch {
                if ( $Error[0].Exception.Message.Contains("The System cannot find the file specified.")) {
                    Write-Host $_.HTML_File 'Unable to open file, may be missing or mislabelled.'
                }
                }
        }
    } else {
        $openAssignments | Where-Object { $_.Document_Name -eq $openValue } | ForEach-Object {
            $fileToOpen = '"' + $_.HTML_File + '"'
            Start-Process chrome $fileToOpen

        }
    }
}

Function OpenHTMLDoc ($docName) {
    $pathToOpen = $openAssignments | Where-Object { $_.Document_Name -eq $docName }
    Start-Process $pathToOpen.HTML_File
}


$completedAssignments = @()

Function ExportFinishedPeerReview {

    $infoGatherTable = $personalAssignmentsTracker | Where-Object { $_."Peer review complete" -eq "Complete" } | Where-Object { $_."Checked In" -eq "No" }

    $infoGatherTable | Where-Object { $_.Document_Name -ne $null } | ForEach-Object {

        # Gather Data
        
        $divisionDocuments = $_.Document_Name
        $waveItem = $_.Wave
        $areaItem = $_.Area
        $divItem = $_.Division
        $lawsonItem = $_.Lawson
        $peerReviewNotes = $_."Peer Review Notes"
        if ($_."Peer Review Notes" -eq $null) {
            $htmlCorrection = "No"
        } else {
            $htmlCorrection = "Yes"
        }
        $peerReviewDate = $_."Ready for/Moved to Upload"

        # Assign data to a new table item, any lines that are commented out are not useful at this time
        $compListObject = New-Object PSObject
        #$compListObject | Add-Member -MemberType NoteProperty -Name "CRC" -Value $null
        $compListObject | Add-Member -MemberType NoteProperty -Name "Waves" -Value $waveItem
        #$compListObject | Add-Member -MemberType NoteProperty -Name "Phase Date" -Value $null
        #$compListObject | Add-Member -MemberType NoteProperty -Name "Area" -Value $areaItem
        #$compListObject | Add-Member -MemberType NoteProperty -Name "P8/Non P8" -Value $null
        $compListObject | Add-Member -MemberType NoteProperty -Name "Division" -Value $divItem
        $compListObject | Add-Member -MemberType NoteProperty -Name "Lawson" -Value $lawsonItem
        $compListObject | Add-Member -MemberType NoteProperty -Name "Division_Documents" -Value $divisionDocuments
        #$compListObject | Add-Member -MemberType NoteProperty -Name "Date Reconciled" -Value $null
        #$compListObject | Add-Member -MemberType NoteProperty -Name "SME Scrub Assigned" -Value $null
        #$compListObject | Add-Member -MemberType NoteProperty -Name "SME Scrub Date" -Value $null
        #$compListObject | Add-Member -MemberType NoteProperty -Name "SME Review Notes" -Value $null
        #$compListObject | Add-Member -MemberType NoteProperty -Name "SME Peer Review Assigned" -Value $null
        #$compListObject | Add-Member -MemberType NoteProperty -Name "Date Ready for HTML" -Value $null
        #$compListObject | Add-Member -MemberType NoteProperty -Name "HTML Build Assigned" -Value $null
        #$compListObject | Add-Member -MemberType NoteProperty -Name "HTML Build Date" -Value $null
        $compListObject | Add-Member -MemberType NoteProperty -Name "Peer Review Assigned" -Value $dataEntryPerson
        $compListObject | Add-Member -MemberType NoteProperty -Name "Peer Review Date" -Value $peerReviewDate
        $compListObject | Add-Member -MemberType NoteProperty -Name "HTML Correction" -Value $htmlCorrection
        $compListObject | Add-Member -MemberType NoteProperty -Name "Correction Needed" -Value $peerReviewNotes
        #$compListObject | Add-Member -MemberType NoteProperty -Name "Compass Upload Assigned" -Value $null
        #$compListObject | Add-Member -MemberType NoteProperty -Name "Compass Upload Date" -Value $null
        #$compListObject | Add-Member -MemberType NoteProperty -Name "Compass Review Date" -Value $null
        #$compListObject | Add-Member -MemberType NoteProperty -Name "Communication Assigned" -Value $null
        #$compListObject | Add-Member -MemberType NoteProperty -Name "Email sent to DIV for Review" -Value $null
        #$compListObject | Add-Member -MemberType NoteProperty -Name "CRC Location" -Value $null

        # Add the new table item to the table
        $completedAssignments += $compListObject
    }see

    $destinationPath = $HOME + '\Documents\TempCheckIn.csv'
    $completedAssignments | Export-Csv -NoTypeInformation -Path $destinationPath
}

Function Seed ($docName) {
    $openAssignments | Where-Object { $_.Document_Name -eq $docName }
}
