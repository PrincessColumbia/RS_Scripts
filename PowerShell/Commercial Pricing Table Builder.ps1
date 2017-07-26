$menuBuild = Import-Excel 'C:\Users\mcfarch\Documents\Nave Menu Builder for Commercial Pricing.xlsx' -WorkSheetname "Variable Builder"
$tempDirectory = 'C:\Users\mcfarch\Scripts\Temp\'
$todaysDate = Get-Date
$currentYear = $todaysDate.Year
$currentMonth = Get-Date -Format MM
$currentDay = Get-Date -Format dd
$dateString = [string]$currentYear + [string]$currentMonth + [string]$currentDay

<# Legacy code
Function RefreshMenu {
    $menuList = $menuBuild."Menu Builder" | Select-Object -Unique
    $menuList
    $menuList | ForEach-Object {
        $menuItem = $_
        $menuItem
        $menuBuild | Where-Object { $_."Menu Builder" -eq $menuItem } | Select-Object "Nav Builder" -Unique
    }
}
#>

Function CreateNav {
    Write-Host -ForegroundColor Yellow -BackgroundColor Red ====================================================================================
    $menuBuild | Select-Object "Menu Builder", "Menu Div ID" -Unique | ForEach-Object {
        $currentObject = $_
        $currentMenu = $currentObject."Menu Builder"
        $divPrefix = '<div id="'
        $divID = $_."Menu Div ID"
        $divSuffix = '">'
        $fullDivString = $divPrefix + $divID + $divSuffix
        Write-Host $fullDivString
        $menuPrefix = '<p>'
        $menuText = $_."Menu Builder"
        $menuSuffix = '</p>'
        $wholeMenuItem = $menuPrefix + $menuText + $menuSuffix
        Write-Host $wholeMenuItem
        Write-Host '<ul>'
        $menuBuild | Where-Object { $_."Menu Builder" -eq $currentMenu } | Select-Object "Nav Builder", "href builder" | ForEach-Object {
            $liPrefix = '<li><a href="'
            $liHrefItem = $_."href builder"
            $liPostHref = '">'
            $liMenuItem = $_."Nav Builder"
            $liSuffix = '</a></li>'
            $fullListItem = $liPrefix + $liHrefItem + $liPostHref + $liMenuItem + $liSuffix
            Write-Host $fullListItem
        }
        Write-Host '</ul>'
        Write-Host '</div>'
    }
    Write-Host -ForegroundColor Yellow -BackgroundColor Red ====================================================================================
}

Function TablesBuilder {

    $appendPath = $tempDirectory + 'Table-Builder-Output.txt'

    $doesTableFileExist = Test-Path $appendPath
    if ($doesTableFileExist) {
        Remove-Item -Path $appendPath -Force
    }

    New-Item -Path $tempDirectory -Name Table-Builder-Output.txt -ItemType file

    $menuBuild | ForEach-Object {

        $tablePart1 = '<div id="'
        $tableDivID = $_."Table Div ID"
        $tablePart2 = '"><a name="'
        $anchorName = $_."Anchor Builder"
        $tablePart3 = '"></a>
          <table width="100%" border="1" cellpadding="2" cellspacing="2" class="tabletext">
            <colgroup>
            <col>
            </colgroup>
            <colgroup>
            <col>
            <col>
            <col>
            <col>
            <col>
            <col>
            <col>
            <col>
            <col>
            <col>
            </colgroup>
            <thead>
              <tr class="tableheader">
                <th colspan="11" scope="col">'
        $tableCaption = $_."Table Caption Builder"
        $tablePart4 = '</th>
              </tr>
              <tr class="tablelabels">
                <th scope="col">&nbsp;</th>
                <th width="8%" scope="col">1 pick up per week</th>
                <th width="8%" scope="col">2 pick ups per week</th>
                <th width="8%" scope="col">3 pick ups per week</th>
                <th width="8%" scope="col">4 pick ups per week</th>
                <th width="8%" scope="col">5 pick ups per week</th>
                <th width="8%" scope="col">6 pick ups per week</th>
                <th width="8%" scope="col">7 pick ups per week</th>
                <th width="8%" scope="col">On Call</th>
                <th width="8%" scope="col">Every other week</th>
                <th scope="col">1 pick up per month</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <th width="21%" class="tablelabels2" scope="row" data-section="row-header">A - List Price</th>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <th width="21%" class="tablelabels2" scope="row" data-section="row-header">A - List Price (2nd container)</th>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <th width="21%" class="tablelabels2" scope="row" data-section="row-header">Notes</th>
                <td colspan="10">&nbsp;</td>
              </tr>
            </tbody>
          </table>
          <p><a href="#Top">Back to top</a></p>
        </div>'
        $newTable = $tablePart1 + $tableDivID + $tablePart2 + $anchorName + $tablePart3 + $tableCaption + $tablePart4
        Add-Content -Value $newTable -Path $appendPath
    }
}

Function RefreshMenuTable {
    $menuBuild = Import-Excel 'C:\Users\mcfarch\Documents\Nave Menu Builder for Commercial Pricing.xlsx' -WorkSheetname "Variable Builder"
}