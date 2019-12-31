#Chrome and firefox or on Windows and Linux. Get the OS/PS version info for later.
#If we're on Windows add IE  as a browser. If PS 6 (not 7) add WinPowershell to the module path
$BrowserList    = 'Chrome' , 'FireFox'
$Platform       = ([environment]::OSVersion.Platform).ToString() + ' PS' + $PSVersionTable.PSVersion.Major
if ($Platform -like 'win*') {
    $BrowserList+= 'IE'
    if ($Platform -like '*6') {
        $env:PSModulePath -split ';' | Where-Object {$_ -match "\w:\\Prog.*PowerShell\\modules"} | ForEach-Object {
            $env:PSModulePath = ($_ -replace "PowerShell","WindowsPowerShell") + ";" +  $env:PSModulePath
        }
    }
}
else {$env:AlwaysHeadless = $true}

#Make sure we have the modules we need
Import-Module .\Selenium.psd1 -Force -ErrorAction Stop
if (-not (Get-Module -ListAvailable ImportExcel)) {
    Write-Verbose -Verbose 'Installing ImportExcel'
    Install-Module ImportExcel -Force -SkipPublisherCheck
}
if (-not (Get-Module -ListAvailable Pester| Where-Object {$_.version.major -ge 4 -and $_.version.minor -ge 4})) {
    Write-Verbose -Verbose 'Installing Pester'
    Install-Module Pester -Force -SkipPublisherCheck
}

#Run the test and results export to an Excel file for current OS - Test picks up the selected browser from an environment variable.
$RunParameters  = @{
    XLFile     = '{0}/results/Results-{1}.xlsx' -f $env:BUILD_ARTIFACTSTAGINGDIRECTORY, [environment]::OSVersion.Platform.ToString()
    Script     = Join-Path -Path (Join-Path $pwd 'Examples') -ChildPath 'Combined.tests.ps1'
}
foreach ( $b   in $BrowserList) {
    $env:DefaultBrowser = $b
    $RunParameters['OutputFile']    = Join-Path $pwd "testresults-$platform$b.xml"
    $RunParameters['WorkSheetName'] =  "$B $Platform"
    $RunParameters | Out-Host
    & "$PSScriptRoot\Pester-To-XLSx.ps1"  @RunParameters
}

#Merge the results sheets into a sheet named 'combined'.
$excel          = Open-ExcelPackage $RunParameters.XLFile
$wslist         = $excel.Workbook.Worksheets.name
Close-ExcelPackage -NoSave $excel
Write-Host ("Merging sheets" + ($wslist -join ',') + "in $($RunParameters.XLFile)")
Merge-MultipleSheets -path  $RunParameters.XLFile -WorksheetName $wslist -OutputSheetName combined -OutputFile $RunParameters.XLFile -HideRowNumbers -Property name,result

#Hide everything on 'combined' except test name, results for each browser, and test group, Set column widths, tweak titles, apply conditional formatting.
$excel          = Open-ExcelPackage $RunParameters.XLFile
$ws             = $excel.combined
2..$ws.Dimension.end.Column | ForEach-Object {
    if ($ws.Cells[1,$_].value -notmatch '^Name|Result$|PS\dGroup$') {
        Set-ExcelColumn -Worksheet $ws -Column $_ -Hid
    }
    elseif ($ws.Cells[1,$_].value -match 'Result$' )  {
        Set-ExcelColumn -Worksheet $ws -Column $_ -Width 17
        Set-ExcelRange $ws.Cells[1,$_] -WrapText
    }
    if ($ws.cells[1,$_].value -match 'PS\dGroup$') {
            Set-ExcelRange $ws.Cells[1,$_] -WrapText -Value 'Group'
    }
    if ($ws.cells[1,$_].value -match '^Name|PS\dGroup$' -and ($ws.Column($_).Width -gt 80)) {
        $ws.Column($_).Width = 80
    }
}
Set-ExcelRow -Worksheet $ws -Height 28.5
$cfRange        = [OfficeOpenXml.ExcelAddress]::new(2,3,$ws.Dimension.end.Row,  (3*$wslist.count -2)).Address
Add-ConditionalFormatting -WorkSheet $ws -range $cfRange -RuleType ContainsText -ConditionValue "Failure" -BackgroundPattern None -ForegroundColor Red   -Bold
Add-ConditionalFormatting -WorkSheet $ws -range $cfRange -RuleType ContainsText -ConditionValue "Success" -BackgroundPattern None -ForeGroundColor Green
Write-Host ("Saving $($excel.File.FullName)")
Close-ExcelPackage $excel