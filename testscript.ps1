$ErrorActionPreference = "Stop"
try {
    Write-Host "Adding PowerShell Snap-in" -ForegroundColor Green
    # Try to get the PowerShell Snappin.  If not, then adding the PowerShell snappin on the Catch Block
    Get-PSSnapin "Microsoft.SharePoint.PowerShell"
}
catch {
    if($Error[0].Exception.Message.Contains("No Windows PowerShell snap-ins have been registered for Windows PowerShell version 5."))
    {
        Add-PSSnapin "Microsoft.SharePoint.PowerShell"
    }
}
Write-Host "Finished Adding PowerShell Snap-in" -ForegroundColor Green