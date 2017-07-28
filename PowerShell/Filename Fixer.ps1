Function FixBadFilename ($badChar,$repChars) {
    if ($repChars -eq $null) {
        $repChars = "_"
    }

    Get-ChildItem * -Recurse | Where-Object { ! $_.PSIsContainer } | Rename-Item -NewName { $_.Name -replace $badChar,$repChars }
}