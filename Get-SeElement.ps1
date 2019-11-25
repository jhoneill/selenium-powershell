function Get-SeElement {
    param(
        [Parameter(Position=0)]
        [ValidateSet("CssSelector", "Name", "Id", "ClassName", "LinkText", "PartialLinkText", "TagName", "XPath")]
        [string]$By = "XPath",
        
        [Parameter(Position=1,Mandatory=$true)]
        [string]$Selection,

        [Parameter(Position=2)]
        [Int]$Timeout = 0,
        
        [Parameter(Position=3,ValueFromPipeline=$true)]
        [Alias('Element','Driver')]
        $Target = $Global:SeDriver    
    )
    process {
        if($TimeOut -and $Target -is [OpenQA.Selenium.Remote.RemoteWebDriver]) { 
            $TargetElement = [OpenQA.Selenium.By]::$By($Selection)
            $WebDriverWait = New-Object -TypeName OpenQA.Selenium.Support.UI.WebDriverWait($Driver, (New-TimeSpan -Seconds $Timeout))
            $Condition     = [OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementExists($TargetElement)
            $WebDriverWait.Until($Condition)
        }
        elseif ($Target -is [OpenQA.Selenium.Remote.RemoteWebElement] -or 
                $Target -is [OpenQA.Selenium.Remote.RemoteWebDriver]) {
            if ($Timeout) {Write-warning}
            $Target.FindElements([OpenQA.Selenium.By]::$By($Selection))
        }
        else {throw "No valid target was provided."}
    }
}
