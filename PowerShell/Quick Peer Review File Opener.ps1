$errorLog = 'C:\Users\mcfarch\Scripts\Temp\error.log'
$startingSeparator = "

============================================================"
$closingSeparator = "============================================================

"


$filesToOpen = (

#paste all filenames to be opened below and ADD COMMAS at the end of each line
"Z:\Site Visit Conversions\A02 West\210 - Antioch Pacheco, CA\3 - Peer Review\DIV_210_FRAN_Commercial_Pricing_Walnut_Creek_CA.html",
"Z:\Site Visit Conversions\A02 West\210 - Antioch Pacheco, CA\3 - Peer Review\DIV_210_Temp_Container_Pricing_Dialo_CA_Metal.html",
"Z:\Site Visit Conversions\A02 West\210 - Antioch Pacheco, CA\3 - Peer Review\DIV_210_Temp_Container_Pricing_Lafayette_CA_Solid_Waste_Compactors.html",
"Z:\Site Visit Conversions\A02 West\210 - Antioch Pacheco, CA\3 - Peer Review\DIV_210_Temp_Container_Pricing_Lafayette_CA_Special_Handled_Materials.html",
"Z:\Site Visit Conversions\A02 West\210 - Antioch Pacheco, CA\2 - HTML Build\DIV_210_FRAN_Commercial_Pricing_Walnut_Creek_CA.xlsx",
"Z:\Site Visit Conversions\A02 West\210 - Antioch Pacheco, CA\2 - HTML Build\DIV_210_Temp_Container_Pricing_Dialo_CA_Metal.xlsx",
"Z:\Site Visit Conversions\A02 West\210 - Antioch Pacheco, CA\2 - HTML Build\DIV_210_Temp_Container_Pricing_Lafayette_CA_Solid_Waste_Compactors.xlsx",
"Z:\Site Visit Conversions\A02 West\210 - Antioch Pacheco, CA\2 - HTML Build\DIV_210_Temp_Container_Pricing_Lafayette_CA_Special_Handled_Materials.xlsx"
#paste all filenames to be opened above and ADD COMMAS at the end of each line

)

$filesToOpen | ForEach-Object {
    
    $currentObject = $_
    try { start $currentObject }
    catch {
        if ( $Error[0].Exception.Message.Contains("The system cannot find the file specified.")) {
            $startingSeparator | Out-File -FilePath $errorLog -Append
            Get-Date | Out-File -FilePath $errorLog -Append
            $currentObject | Out-File -FilePath $errorLog -Append
            $Error[0] | Out-File -FilePath $errorLog -Append
            $closingSeparator | Out-File -FilePath $errorLog -Append
        }
    }
}

Function BrowserOpen {
    $filesToOpen | ForEach-Object {
        
        $currentObject = $_
        try { start chrome $currentObject }
        catch {
            if ( $Error[0].Exception.Message.Contains("The system cannot find the file specified.")) {
                $startingSeparator | Out-File -FilePath $errorLog -Append
                Get-Date | Out-File -FilePath $errorLog -Append
                $currentObject | Out-File -FilePath $errorLog -Append
                $Error[0] | Out-File -FilePath $errorLog -Append
                $closingSeparator | Out-File -FilePath $errorLog -Append
            }
        }
    }
}