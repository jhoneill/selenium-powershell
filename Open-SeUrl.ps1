function Open-SeUrl {
    param(
        [Parameter(Mandatory=$true, position=0)]    
        [string]$Url,
        [Alias("Driver")]
        $Target = $Global:SeDriver 
    )
    if (-not $Target -is [OpenQA.Selenium.Remote.RemoteWebDriver]) {
        throw "No valid driver was provided. "
    }
    else {$Target.Navigate().GoToUrl($Url) }
}
