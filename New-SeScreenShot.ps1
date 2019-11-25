function New-SeScreenshot {
    [cmdletbinding(DefaultParameterSetName='Path')]
    param(
        [Parameter(ParameterSetName='Path',Mandatory=$true,Position=0)]
        [Parameter(ParameterSetName='PassThru',Position=0)]
        $Path,

        [Parameter(ParameterSetName='Path',Position=1)]
        [Parameter(ParameterSetName='PassThru',Position=1)]
        [OpenQA.Selenium.ScreenshotImageFormat]$ImageFormat = [OpenQA.Selenium.ScreenshotImageFormat]::Png,
        
        [Alias("Driver")]
        $Target = $Global:SeDriver ,
        
        [Parameter(ParameterSetName='Base64',Mandatory=$true)]
        [Switch]$AsBase64EncodedString,   
        
        [Parameter(ParameterSetName='PassThru',Mandatory=$true)]     
        [Alias('PT')]
        [Switch]$PassThru  
    )
    if (-not $Target -is [OpenQA.Selenium.Remote.RemoteWebDriver]) {
        throw "No valid driver was provided" ; return}

    $Screenshot = [OpenQA.Selenium.Support.Extensions.WebDriverExtensions]::TakeScreenshot($Target)
    if ($AsBase64EncodedString) {$Screenshot.AsBase64EncodedString}
    elseif ($Path)              {
        $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
        $Screenshot.SaveAsFile($Path, $ImageFormat) }
    if ($Passthru)              {$Screenshot}
}
