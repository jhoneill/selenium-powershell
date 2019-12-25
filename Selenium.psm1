$Script:SeKeys = [OpenQA.Selenium.Keys] | Get-Member -MemberType Property -Static |
        Select-Object -Property Name, @{N = "ObjectString"; E = { "[OpenQA.Selenium.Keys]::$($_.Name)" } }

#region Set path to assemblies on Linux and MacOS and Grant Execution permissions on them
if($IsLinux){
    $AssembliesPath = "$PSScriptRoot/assemblies/linux"
}
elseif($IsMacOS){
    $AssembliesPath = "$PSScriptRoot/assemblies/macos"
}

# Grant Execution permission to assemblies on Linux and MacOS
if($AssembliesPath){
    # Check if powershell is NOT running as root
    Get-Item -Path "$AssembliesPath/chromedriver", "$AssembliesPath/chromedriver" | ForEach-Object {
        if($IsLinux)    {$FileMod          = stat -c "%a" $_.fullname }
        elseif($IsMacOS){$FileMod = /usr/bin/stat -f "%A" $_.fullname}
        if($FileMod[2] -ne '5' -and $FileMod[2] -ne '7' ){
            Write-Host "Granting $($AssemblieFile.fullname) Execution Permissions ..."
            chmod +x $_.fullname
        }
    }
}

#endregion
function ValidateURL {
    [Alias("Validate-Url")]
    param(
        [Parameter(Mandatory=$true)]$URL
    )
    $Out = $null
    [uri]::TryCreate($URL,[System.UriKind]::Absolute, [ref]$Out)
}

function Start-SeNewEdge {
    [cmdletbinding(DefaultParameterSetName='default')]
    [Alias('CrEdge','NewEdge')]
    param(
        [ValidateURI()]
        [Parameter(Position=0)]
        [string]$StartURL,
        [switch]$HideVersionHint,
        [switch]$Minimized,
        $BinaryPath = "C:\Program Files (x86)\Microsoft\Edge Dev\Application\msedge.exe",
        $ProfileDirectoryPath,
        $DefaultDownloadPath,
        [switch]$AsDefaultDriver,
        [switch]$Headless,
        [Alias('Incognito')]
        [switch]$PrivateBrowsing,
        $WebDriverDirectory = "$PSScriptRoot\Assemblies\"
    )
    if(!$HideVersionHint){
        Write-Verbose "Download the right webdriver from 'https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/'"
    }

    $Service = [OpenQA.Selenium.Chrome.ChromeDriverService]::CreateDefaultService($WebDriverDirectory, 'msedgedriver.exe')
    $options = New-Object -TypeName OpenQA.Selenium.Chrome.ChromeOptions -Property  @{BinaryLocation = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($BinaryPath)}
    if ($PrivateBrowsing)       {$options.AddArguments('InPrivate')}
    if ($Headless)              {$options.AddArguments('headless')}
    if ($ProfileDirectoryPath)  {
        $ProfileDirectoryPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ProfileDirectoryPath)
        Write-Verbose "Setting Profile directory: $ProfileDirectoryPath"
        $options.AddArgument("user-data-dir=$ProfileDirectoryPath")
    }
    if($DefaultDownloadPath){
        $DefaultDownloadPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DefaultDownloadPath)
        Write-Verbose "Setting Default Download directory: $DefaultDownloadPath"
        $Options.AddUserProfilePreference('download', @{'default_directory' = $DefaultDownloadPath; 'prompt_for_download' = $false; })
    }
    $Driver  = New-Object -TypeName OpenQA.Selenium.Chrome.ChromeDriver  -ArgumentList $Service, $options
    if(-not $Driver) {Write-Warning "Web driver was not created"; return}

    if($StartURL) {$Driver.Navigate().GoToUrl($StartURL)}

    if($Minimized){
        $Driver.Manage().Window.Minimize();
    }

    if($AsDefaultDriver) {
        if($Global:SeDriver) {$Global:SeDriver.Dispose()}
        $Global:SeDriver = $Driver
    }
    else {$Driver}
}

function Start-SeChrome {
    [cmdletbinding(DefaultParameterSetName='default')]
    [Alias('Chrome')]
    param(
        [ValidateURI()]
        [Parameter(Position=0)]
        [string]$StartURL,
        [Parameter(Mandatory = $false)]
        [array]$Arguments,
        [switch]$HideVersionHint,
        [System.IO.FileInfo]$DefaultDownloadPath,
        [System.IO.FileInfo]$ProfileDirectoryPath,
        [Parameter(DontShow)]
        [bool]$DisableBuiltInPDFViewer=$true,
        [switch]$EnablePDFViewer,
        [switch]$Headless,
        [Alias('PrivateBrowsing')]
        [switch]$Incognito,
        [parameter(ParameterSetName='Min',Mandatory=$true)]
        [switch]$Maximized,
        [parameter(ParameterSetName='Max',Mandatory=$true)]
        [switch]$Minimized,
        [parameter(ParameterSetName='Ful',Mandatory=$true)]
        [switch]$Fullscreen,
        [System.IO.FileInfo]$ChromeBinaryPath,
        [switch]$AsDefaultDriver
    )

    process {
        #region chrome set-up options
        $Chrome_Options = New-Object -TypeName "OpenQA.Selenium.Chrome.ChromeOptions"

        if($DefaultDownloadPath){
            Write-Verbose "Setting Default Download directory: $DefaultDownloadPath"
            $Chrome_Options.AddUserProfilePreference('download', @{'default_directory' = $($DefaultDownloadPath.FullName); 'prompt_for_download' = $false; })
        }

        if($ProfileDirectoryPath){
            Write-Verbose "Setting Profile directory: $ProfileDirectoryPath"
            $Chrome_Options.AddArgument("user-data-dir=$ProfileDirectoryPath")
        }

        if($ChromeBinaryPath){
            Write-Verbose "Setting Chrome Binary directory: $ChromeBinaryPath"
            $Chrome_Options.BinaryLocation ="$ChromeBinaryPath"
        }

        if($DisableBuiltInPDFViewer -and -not $EnablePDFViewer){
            $Chrome_Options.AddUserProfilePreference('plugins', @{'always_open_pdf_externally' =  $true;})
        }

        if($Headless){
            $Chrome_Options.AddArguments('headless')
        }

        if($Incognito){
            $Chrome_Options.AddArguments('Incognito')
        }

        if($Maximized){
            $Chrome_Options.AddArguments('start-maximized')
        }

        if($Fullscreen){
            $Chrome_Options.AddArguments('start-fullscreen')
        }

        if($Arguments){
            foreach ($Argument in $Arguments){
                $Chrome_Options.AddArguments($Argument)
            }
        }
        #endregion
        if(!$HideVersionHint){
            Write-Verbose "Download the right chromedriver from 'http://chromedriver.chromium.org/downloads'"
        }

        if($AssembliesPath){ $arglist = @( $AssembliesPath,$Chrome_Options) }
        else               { $arglist = @( $Chrome_Options)}
        $Driver = New-Object -TypeName "OpenQA.Selenium.Chrome.ChromeDriver" -ArgumentList $arglist
        if(-not $Driver) {Write-Warning "Web driver was not created"; return}

        #region post start options
        if($Minimized){
            $Driver.Manage().Window.Minimize();
        }

        if($Headless -and $DefaultDownloadPath){
            $HeadlessDownloadParams = New-Object 'system.collections.generic.dictionary[[System.String],[System.Object]]]'
            $HeadlessDownloadParams.Add('behavior', 'allow')
            $HeadlessDownloadParams.Add('downloadPath', $DefaultDownloadPath.FullName)
            $Driver.ExecuteChromeCommand('Page.setDownloadBehavior', $HeadlessDownloadParams)
        }

        if($StartURL) {$Driver.Navigate().GoToUrl($StartURL)}
        #endregion

        if($AsDefaultDriver) {
            if($Global:SeDriver) {$Global:SeDriver.Dispose()}
            $Global:SeDriver = $Driver
        }
        else {$Driver}
    }
}

function Start-SeInternetExplorer {
    [Alias('InternetExplorer','IE')]
    param(
        [ValidateURI()]
        [Parameter(Position=0)]
        [string]$StartURL,
        [switch]$AsDefaultDriver
    )
    $InternetExplorer_Options = New-Object -TypeName "OpenQA.Selenium.IE.InternetExplorerOptions"
    $InternetExplorer_Options.IgnoreZoomLevel = $true
    $Driver = New-Object -TypeName "OpenQA.Selenium.IE.InternetExplorerDriver" -ArgumentList $InternetExplorer_Options

    if(-not $Driver) {Write-Warning "Web driver was not created"; return}

    $Driver.Manage().Timeouts().ImplicitWait = [TimeSpan]::FromSeconds(10)
    if($StartURL)  {$Driver.Navigate().GoToUrl($StartURL) }

    if($AsDefaultDriver) {
        if($Global:SeDriver) {$Global:SeDriver.Dispose()}
        $Global:SeDriver = $Driver
    }
    else {$Driver}
}

function Start-SeEdge {
    [cmdletbinding(DefaultParameterSetName='default')]
    [Alias('MSEdge')]
    param(
        [ValidateURI()]
        [Parameter(Position=0)]
        [string]$StartURL,
        [parameter(ParameterSetName='Min',Mandatory=$true)]
        [switch]$Maximized,
        [parameter(ParameterSetName='Max',Mandatory=$true)]
        [switch]$Minimized,
        [switch]$AsDefaultDriver
    )
    $Driver = New-Object -TypeName "OpenQA.Selenium.Edge.EdgeDriver"
    if(-not $Driver) {Write-Warning "Web driver was not created"; return}

    #region post creation options
    $Driver.Manage().Timeouts().ImplicitWait = [TimeSpan]::FromSeconds(10)
    if($Minimized) {$Driver.Manage().Window.Minimize()    }
    if($Maximized) {$Driver.Manage().Window.Maximize()    }
    if($StartURL)  {$Driver.Navigate().GoToUrl($StartURL) }
    #endregion

    if($AsDefaultDriver) {
        if($Global:SeDriver) {$Global:SeDriver.Dispose()}
        $Global:SeDriver = $Driver
    }
    else {$Driver}
}

function Start-SeFirefox {
    [cmdletbinding(DefaultParameterSetName='default')]
    [Alias('Firefox')]
    param(
        [ValidateURI()]
        [Parameter(Position=0)]
        [string]$StartURL,
        [array]$Arguments,
        [System.IO.FileInfo]$DefaultDownloadPath,
        [switch]$Headless,
        [alias('Incognito')]
        [switch]$PrivateBrowsing,
        [parameter(ParameterSetName='Min',Mandatory=$true)]
        [switch]$Maximized,
        [parameter(ParameterSetName='Max',Mandatory=$true)]
        [switch]$Minimized,
        [parameter(ParameterSetName='Ful',Mandatory=$true)]
        [switch]$Fullscreen,
        [switch]$SuppressLogging,
        [switch]$AsDefaultDriver
    )
    process {
        #region firefox set-up options
        $Firefox_Options = New-Object -TypeName "OpenQA.Selenium.Firefox.FirefoxOptions"

        if($Headless){
            $Firefox_Options.AddArguments('-headless')
        }

        if($DefaultDownloadPath){
            Write-Verbose "Setting Default Download directory: $DefaultDownloadPath"
            $Firefox_Options.setPreference("browser.download.folderList",2);
            $Firefox_Options.SetPreference("browser.download.dir", "$DefaultDownloadPath");
        }

        if($PrivateBrowsing){
            $Firefox_Options.SetPreference("browser.privatebrowsing.autostart", $true)
        }

        if($Arguments){
            foreach ($Argument in $Arguments){
                $Firefox_Options.AddArguments($Argument)
            }
        }

        if($SuppressLogging){
            # Sets GeckoDriver log level to Fatal.
            $Firefox_Options.LogLevel = 6
        }
        #endregion

        if($AssembliesPath){ $arglist = @( $AssembliesPath,$Firefox_Options) }
        else               { $arglist = @( $Firefox_Options)}
        $Driver = New-Object -TypeName "OpenQA.Selenium.Firefox.FirefoxDriver"-ArgumentList $arglist
        if(-not $Driver) {Write-Warning "Web driver was not created"; return}

        #region post creation options
        $Driver.Manage().Timeouts().ImplicitWait = [TimeSpan]::FromSeconds(10)
        if($Minimized) {$Driver.Manage().Window.Minimize()    }
        if($Maximized) {$Driver.Manage().Window.Maximize()    }
        if($Fullscreen){$Driver.Manage().Window.FullScreen()  }
        if($StartURL)  {$Driver.Navigate().GoToUrl($StartURL) }
        #endregion

        if($AsDefaultDriver) {
            if($Global:SeDriver) {$Global:SeDriver.Dispose()}
            $Global:SeDriver = $Driver
        }
        else {$Driver}
    }
}

function Stop-SeDriver {
    [alias('SeClose')]
    param(
        [Parameter(ValueFromPipeline=$true, position=0,ParameterSetName='Driver')]
        [ValidateIsWebDriver()]
        $Driver,
        [Parameter(Mandatory=$true, ParameterSetName='Default')]
        [switch]$Default
    )
    if(-not $PSBoundParameters.ContainsKey('Driver') -and ($Default -or $MyInvocation.InvocationName -eq 'SeClose')) {
        Write-Verbose -Message "Closing $($Global:SeDriver.Capabilities.browsername)..."
        $Global:SeDriver.Close()
        $Global:SeDriver.Dispose()
        Remove-Variable -Name SeDriver -Scope global
    }
    elseif ($Driver) {
        $Driver.Close()
        $Driver.Dispose()
    }
    else {Write-Warning -Message 'No Driver Specified'}
}

<#function Enter-SeUrl {
    param($Driver, $Url)

    $Driver.Navigate().GoToUrl($Url)
}
#>
function Open-SeUrl {
    [Alias('SeNavigate',"Enter-SeUrl")]
    param(
        [Parameter(Mandatory=$true, position=0)]
        [string]$Url,
        [Alias("Driver")]
        [ValidateIsWebDriver()]
        $Target = $Global:SeDriver
    )
    $Target.Navigate().GoToUrl($Url)
}

<#function Find-SeElement {
    param(
        [Parameter()]
        $Driver,
        [Parameter()]
        $Element,
        [Parameter()][Switch]$Wait,
        [Parameter()]$Timeout = 30,
        [Parameter(ParameterSetName = "ByCss")]
        $Css,
        [Parameter(ParameterSetName = "ByName")]
        $Name,
        [Parameter(ParameterSetName = "ById")]
        $Id,
        [Parameter(ParameterSetName = "ByClassName")]
        $ClassName,
        [Parameter(ParameterSetName = "ByLinkText")]
        $LinkText,
        [Parameter(ParameterSetName = "ByPartialLinkText")]
        $PartialLinkText,
        [Parameter(ParameterSetName = "ByTagName")]
        $TagName,
        [Parameter(ParameterSetName = "ByXPath")]
        $XPath
        )


    process {

        if($null -ne $Driver  -and $null -ne  $Element) {
            throw "Driver and Element may not be specified together."
        }
        elseif($null -ne $Driver) {
            $Target = $Driver
        }
        elseif(-ne $Null $Element) {
            $Target = $Element
        }
        else {
            "Driver or element must be specified"
        }

        if($Wait){
            if($PSCmdlet.ParameterSetName -eq "ByName") {
                $TargetElement = [OpenQA.Selenium.By]::Name($Name)
            }

            if($PSCmdlet.ParameterSetName -eq "ById") {
                $TargetElement = [OpenQA.Selenium.By]::Id($Id)
            }

            if($PSCmdlet.ParameterSetName -eq "ByLinkText") {
                $TargetElement = [OpenQA.Selenium.By]::LinkText($LinkText)
            }

            if($PSCmdlet.ParameterSetName -eq "ByPartialLinkText") {
                $TargetElement = [OpenQA.Selenium.By]::PartialLinkText($PartialLinkText)
            }

            if($PSCmdlet.ParameterSetName -eq "ByClassName") {
                $TargetElement = [OpenQA.Selenium.By]::ClassName($ClassName)
            }

            if($PSCmdlet.ParameterSetName -eq "ByTagName") {
                $TargetElement = [OpenQA.Selenium.By]::TagName($TagName)
            }

            if($PSCmdlet.ParameterSetName -eq "ByXPath") {
                $TargetElement = [OpenQA.Selenium.By]::XPath($XPath)
            }

            if($PSCmdlet.ParameterSetName -eq "ByCss") {
                $TargetElement = [OpenQA.Selenium.By]::CssSelector($Css)
            }

            $WebDriverWait = New-Object -TypeName OpenQA.Selenium.Support.UI.WebDriverWait($Driver, (New-TimeSpan -Seconds $Timeout))
            $Condition = [OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementExists($TargetElement)
            $WebDriverWait.Until($Condition)
        }
        else{
            if($PSCmdlet.ParameterSetName -eq "ByName") {
                $Target.FindElements([OpenQA.Selenium.By]::Name($Name))
            }

            if($PSCmdlet.ParameterSetName -eq "ById") {
                $Target.FindElements([OpenQA.Selenium.By]::Id($Id))
            }

            if($PSCmdlet.ParameterSetName -eq "ByLinkText") {
                $Target.FindElements([OpenQA.Selenium.By]::LinkText($LinkText))
            }

            if($PSCmdlet.ParameterSetName -eq "ByPartialLinkText") {
                $Target.FindElements([OpenQA.Selenium.By]::PartialLinkText($PartialLinkText))
            }

            if($PSCmdlet.ParameterSetName -eq "ByClassName") {
                $Target.FindElements([OpenQA.Selenium.By]::ClassName($ClassName))
            }

            if($PSCmdlet.ParameterSetName -eq "ByTagName") {
                $Target.FindElements([OpenQA.Selenium.By]::TagName($TagName))
            }

            if($PSCmdlet.ParameterSetName -eq "ByXPath") {
                $Target.FindElements([OpenQA.Selenium.By]::XPath($XPath))
            }

            if($PSCmdlet.ParameterSetName -eq "ByCss") {
                $Target.FindElements([OpenQA.Selenium.By]::CssSelector($Css))
            }
        }
    }
}
#>
function Get-SeElement {
    [Alias('Find-SeElement','SeElement')]
    param(
        #Specifies whether the selction text is to select by name, ID, Xpath etc
        [ValidateSet("CssSelector", "Name", "Id", "ClassName", "LinkText", "PartialLinkText", "TagName", "XPath")]
        [string]$By = "XPath",
        #Text to select on
        [Alias("Name", "Id", "ClassName","LinkText", "PartialLinkText", "TagName","XPath")]
        [Parameter(Position=1,Mandatory=$true)]
        [string]$Selection,
        #Specifies a time out
        [Parameter(Position=2)]
        [Int]$Timeout = 0,
        #The driver or Element where the search should be performed.
        [Parameter(Position=3,ValueFromPipeline=$true)]
        [Alias('Element','Driver')]
        $Target = $Global:SeDriver,

        [parameter(DontShow)]
        [Switch]$Wait

    )
    process {
        #if one of the old parameter names was used and BY was NIT specified, look for
        # <cmd/alias name> [anything which doesn't mean end of command]  -Param
        # capture Param and set it as the value for by
        $mi = $MyInvocation.InvocationName
        if(-not $PSBoundParameters.ContainsKey("By") -and
          ($MyInvocation.Line -match  "$mi[^>\|;]*-(Name|Id|ClassName|LinkText|PartialLinkText|TagName|XPath)")) {
                $By = $Matches[1]
        }
        if($wait -and $Timeout -eq 0) {$Timeout = 30 }

        if($TimeOut -and $Target -is [OpenQA.Selenium.Remote.RemoteWebDriver]) {
            $TargetElement = [OpenQA.Selenium.By]::$By($Selection)
            $WebDriverWait = New-Object -TypeName OpenQA.Selenium.Support.UI.WebDriverWait($Driver, (New-TimeSpan -Seconds $Timeout))
            $Condition     = [OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementExists($TargetElement)
            $WebDriverWait.Until($Condition)
        }
        elseif($Target -is [OpenQA.Selenium.Remote.RemoteWebElement] -or
               $Target -is [OpenQA.Selenium.Remote.RemoteWebDriver]) {
            if($Timeout) {Write-Warning "Timeout does not apply when searching an Element"}
            $Target.FindElements([OpenQA.Selenium.By]::$By($Selection))
        }
        else {throw "No valid target was provided."}
    }
}

function Invoke-SeClick {
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [OpenQA.Selenium.IWebElement]$Element,
        [Parameter()]
        [Switch]$JavaScriptClick,
        [Parameter()]
        $Driver
    )

    if($JavaScriptClick) {
        $Driver.ExecuteScript("arguments[0].click()", $Element)
    }
    else {
        $Element.Click()
    }

}

function Send-SeClick {
    [alias('SeClick')]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true,Position=0)]
        [OpenQA.Selenium.IWebElement]$Element,
        [Alias('JS')]
        [Switch]$JavaScriptClick,
        $SleepSeconds = 0 ,
        [Parameter(DontShow)]
        $Driver,
        [Alias('PT')]
        [switch]$PassThru
    )
    Process {
        if($JavaScriptClick) { $Element.WrappedDriver.ExecuteScript("arguments[0].click()", $Element) }
        else                 { $Element.Click() }
        if($SleepSeconds)    { Start-Sleep -Seconds $SleepSeconds}
        if($PassThru)        { $Element}
    }
}

function Get-SeKeys {
    [OpenQA.Selenium.Keys] | Get-Member -MemberType Property -Static | Select-Object -Property Name, @{N = "ObjectString"; E = { "[OpenQA.Selenium.Keys]::$($_.Name)" } }
}

function Send-SeKeys {
    param(
        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true)]
        [OpenQA.Selenium.IWebElement]$Element,
        [Parameter(Mandatory=$true,Position=1)]
        [string]$Keys
    )
    foreach ($Key in $Script:SeKeys.Name) {
        $Keys = $Keys -replace "{{$Key}}", [OpenQA.Selenium.Keys]::$Key
    }
    $Element.SendKeys($Keys)
}

function Get-SeCookie {
    param(
        [Alias("Driver")]
        [ValidateIsWebDriver()]
        $Target = $Global:SeDriver
    )
    $Target.Manage().Cookies.AllCookies.GetEnumerator()
}

function Remove-SeCookie {
    param(
        $Driver,
        [switch]$DeleteAllCookies,
        [string]$Name
    )

    if($DeleteAllCookies){
        $Driver.Manage().Cookies.DeleteAllCookies()
    }
    else{
        $Driver.Manage().Cookies.DeleteCookieNamed($Name)
    }
}

function Set-SeCookie {
    [cmdletbinding()]
    param(
        [string]$Name,
        [string]$Value,
        [string]$Path,
        [string]$Domain,
        $ExpiryDate,
        [Alias("Driver")]
        [ValidateIsWebDriver()]
        $Target = $Global:SeDriver
    )

    <# Selenium Cookie Information
    Cookie(String, String)
    Initializes a new instance of the Cookie class with a specific name and value.
    Cookie(String, String, String)
    Initializes a new instance of the Cookie class with a specific name, value, and path.
    Cookie(String, String, String, Nullable<DateTime>)
    Initializes a new instance of the Cookie class with a specific name, value, path and expiration date.
    Cookie(String, String, String, String, Nullable<DateTime>)
    Initializes a new instance of the Cookie class with a specific name, value, domain, path and expiration date.
    #>
    begin {
        if($null -ne $ExpiryDate -and $ExpiryDate.GetType().Name -ne 'DateTime'){
            throw '$ExpiryDate can only be $null or TypeName: System.DateTime'
        }
    }

    process {
        if($Name -and $Value -and (!$Path -and !$Domain -and !$ExpiryDate)){
            $cookie = New-Object -TypeName OpenQA.Selenium.Cookie -ArgumentList $Name,$Value
        }
        Elseif($Name -and $Value -and $Path -and (!$Domain -and !$ExpiryDate)){
            $cookie = New-Object -TypeName OpenQA.Selenium.Cookie -ArgumentList $Name,$Value,$Path
        }
        Elseif($Name -and $Value -and $Path -and $ExpiryDate -and !$Domain){
            $cookie = New-Object -TypeName OpenQA.Selenium.Cookie -ArgumentList $Name,$Value,$Path,$ExpiryDate
        }
        Elseif($Name -and $Value -and $Path -and $Domain -and (!$ExpiryDate -or $ExpiryDate)){
            if($Driver.Url -match $Domain){
                $cookie = New-Object -TypeName OpenQA.Selenium.Cookie -ArgumentList $Name,$Value,$Domain,$Path,$ExpiryDate
            }
            else{
                Throw 'In order to set the cookie the browser needs to be on the cookie domain URL'
            }
        }
        else{
            Throw "Incorrect Cookie Layout:
            Cookie(String, String)
            Initializes a new instance of the Cookie class with a specific name and value.
            Cookie(String, String, String)
            Initializes a new instance of the Cookie class with a specific name, value, and path.
            Cookie(String, String, String, Nullable<DateTime>)
            Initializes a new instance of the Cookie class with a specific name, value, path and expiration date.
            Cookie(String, String, String, String, Nullable<DateTime>)
            Initializes a new instance of the Cookie class with a specific name, value, domain, path and expiration date."
        }

        $Driver.Manage().Cookies.AddCookie($cookie)
    }
}

function Get-SeElementAttribute {
    param(
        [Parameter(ValueFromPipeline = $true, Mandatory = $true)]
        [OpenQA.Selenium.IWebElement]$Element,
        [Parameter(Mandatory = $true)]
        [string]$Attribute
    )
    process {
        $Element.GetAttribute($Attribute)
    }
}

function Invoke-SeScreenshot {
    param(
        [Alias("Driver")]
        [ValidateIsWebDriver()]
        $Target = $Global:SeDriver,
        [Switch]$AsBase64EncodedString
    )
    $Screenshot = [OpenQA.Selenium.Support.Extensions.WebDriverExtensions]::TakeScreenshot($Target)
    if($AsBase64EncodedString) {
        $Screenshot.AsBase64EncodedString
    }
    else {
        $Screenshot
    }
}

function Save-SeScreenshot {
    param(
        [Parameter(ValueFromPipeline = $true, Mandatory = $true)]
        [OpenQA.Selenium.Screenshot]$Screenshot,
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter()]
        [OpenQA.Selenium.ScreenshotImageFormat]$ImageFormat = [OpenQA.Selenium.ScreenshotImageFormat]::Png)

    process {
        $Screenshot.SaveAsFile($Path, $ImageFormat)
    }
}

function New-SeScreenshot {
    [Alias('SeScreenshot')]
    [cmdletbinding(DefaultParameterSetName='Path')]
    param(
        [Parameter(ParameterSetName='Path'    ,Position=0,Mandatory=$true)]
        [Parameter(ParameterSetName='PassThru',Position=0)]
        $Path,

        [Parameter(ParameterSetName='Path',    Position=1)]
        [Parameter(ParameterSetName='PassThru',Position=1)]
        [OpenQA.Selenium.ScreenshotImageFormat]$ImageFormat = [OpenQA.Selenium.ScreenshotImageFormat]::Png,

        [Alias("Driver")]
        [ValidateIsWebDriver()]
        $Target = $Global:SeDriver ,

        [Parameter(ParameterSetName='Base64',  Mandatory=$true)]
        [Switch]$AsBase64EncodedString,

        [Parameter(ParameterSetName='PassThru',Mandatory=$true)]
        [Alias('PT')]
        [Switch]$PassThru
    )
    $Screenshot = [OpenQA.Selenium.Support.Extensions.WebDriverExtensions]::TakeScreenshot($Target)
    if($AsBase64EncodedString) {$Screenshot.AsBase64EncodedString}
    elseif($Path)              {
        $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
        $Screenshot.SaveAsFile($Path, $ImageFormat) }
    if($Passthru)              {$Screenshot}
}

function Get-SeWindow {
    param(
        [Parameter(Mandatory = $true)][OpenQA.Selenium.IWebDriver]$Driver
    )

    process {
        $Driver.WindowHandles
    }
}

function Switch-SeWindow {
    param(
        [Parameter(Mandatory = $true)][OpenQA.Selenium.IWebDriver]$Driver,
        [Parameter(Mandatory = $true)]$Window
    )

    process {
        $Driver.SwitchTo().Window($Window)|Out-Null
    }
}

function Switch-SeFrame {
    [Alias('SeFrame')]
    param (
        $Frame,
        [switch]$Parent,
        [ValidateIsWebDriver()]
        $Target = $Global:SeDriver
    )
    if     ($frame)  {[void]$Target.SwitchTo().Frame($Frame) }
    elseif ($Parent) {[void]$Target.SwitchTo().ParentFrame()}
}

function Clear-SeAlert {
    [Alias('SeAccept','SeDismiss')]
    param (
        [parameter(ParameterSetName='Alert', Position=0,ValueFromPipeline=$true)]
        $Alert,
        [parameter(ParameterSetName='Driver')]
        [ValidateIsWebDriver()]
        [Alias("Driver")]
        $Target = $Global:SeDriver,
        [ValidateSet('Accept','Dismiss')]
        $Action = 'Dismiss',
        [Alias('PT')]
        [switch]$PassThru
    )
    if ($Target) {
        try   {$Alert = $Target.SwitchTo().alert() }
        catch {Write-warning 'No alert was displayed'; return}
    }
    if (-not $PSBoundParameters.ContainsKey('Action') -and
        $MyInvocation.InvocationName -match 'Accept') {$Action = 'Accept'}
    if ($Alert) {$alert.$action() }
    if ($PassThru) {$Alert}
}

function SeOpen {
    [CmdletBinding()]
    Param(
        [ValidateSet('Chrome','CrEdge','FireFox','InternetExplorer','IE','MSEdge','NewEdge')]
        [Parameter(Mandatory=$true ,Position=0)]
        $In,
        [ValidateURI()]
        [Parameter(Mandatory=$False,Position=1)]
        $URL,
        [hashtable]$Options =@{},
        [int]$SleepSeconds
    )
    $Options['AsDefaultDriver']     = $true;
    $Options['ErrorAction']         = 'Stop'
    if ($url) {$Options['StartUrl'] = $url}
    switch -regex ($In) {
        'Chrome'   {Start-SeChrome           @Options; continue}
        'FireFox'  {Start-SeFirefox          @Options; continue}
        'MSEdge'   {Start-SeEdge             @Options; continue}
        'Edge$'    {Start-SeNewEdge          @Options; continue}
        '^I'       {Start-SeInternetExplorer @Options; continue}
    }
    Write-Verbose -Message "Opened $($Global:SeDriver.Capabilities.browsername)"
    if ($SleepSeconds) {Start-Sleep -Seconds $SleepSeconds}
}

function SeType {
    param(
        [Parameter(Mandatory=$true,Position=0)]
        [string]$Keys,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [OpenQA.Selenium.IWebElement]$Element,
        [switch]$ClearFirst,
        [Alias('PT')]
        [switch]$PassThru
    )
    begin{
        foreach ($Key in $Script:SeKeys.Name) {
            $Keys = $Keys -replace "{{$Key}}", [OpenQA.Selenium.Keys]::$Key
        }
    }
    process {
        if ($ClearFirst) {$Element.Clear()}
        $Element.SendKeys($Keys)
        if ($PassThru) {$Element}
    }
}

function SeShouldHave {
    [cmdletbinding(DefaultParameterSetName='DefaultPS')]
    param(
        [Parameter(ParameterSetName='DefaultPS', Mandatory=$true , Position=0, ValueFromPipeline=$true)]
        [Parameter(ParameterSetName='Element'  , Mandatory=$true , Position=0, ValueFromPipeline=$true)]
        [string[]]$Selection,

        [Parameter(ParameterSetName='DefaultPS', Mandatory=$false)]
        [Parameter(ParameterSetName='Element'  , Mandatory=$false)]
        [ValidateSet('CssSelector', 'Name', 'Id', 'ClassName', 'LinkText', 'PartialLinkText', 'TagName', 'XPath')]
        [string]$By = 'XPath',

        [Parameter(ParameterSetName='Element'  , Mandatory=$true , Position=1)]
        [string]$With,

        [Parameter(ParameterSetName='Alert'    , Mandatory=$true )]
        [switch]$Alert,
        [Parameter(ParameterSetName='NoAlert'  , Mandatory=$true )]
        [switch]$NoAlert,
        [Parameter(ParameterSetName='Title'    , Mandatory=$true )]
        [switch]$Title,
        [Parameter(ParameterSetName='URL'      , Mandatory=$true )]
        [Alias('URI')]
        [switch]$URL,

        [Parameter(ParameterSetName='Element'  , Mandatory=$false, Position=3)]
        [Parameter(ParameterSetName='Alert'    , Mandatory=$false, Position=3)]
        [Parameter(ParameterSetName='Title'    , Mandatory=$false, Position=3)]
        [Parameter(ParameterSetName='URL'      , Mandatory=$false, Position=3)]
        [ValidateSet('like', 'notlike', 'match',
                         'notmatch', 'eq', 'ne', 'gt', 'lt')]
        [String]$Operator = 'like',

        [Parameter(ParameterSetName='Element'  , Mandatory=$false, Position=4)]
        [Parameter(ParameterSetName='Alert'    , Mandatory=$false, Position=4)]
        [Parameter(ParameterSetName='Title'    , Mandatory=$true , Position=4)]
        [Parameter(ParameterSetName='URL'      , Mandatory=$true , Position=4)]
        [Alias('like', 'notlike', 'match', 'notmatch', 'eq', 'ne', 'gt', 'lt')]
        [AllowEmptyString()]
        $Value,

        [Parameter(ParameterSetName='DefaultPS')]
        [Parameter(ParameterSetName='Element'  )]
        [Parameter(ParameterSetName='Alert'    )]
        [Alias('PT')]
        [switch]$PassThru,

        [Int]$Timeout = 0
    )
    begin {
        $lineText = $MyInvocation.Line.TrimEnd("$([System.Environment]::NewLine)")
        $lineNo   = $MyInvocation.ScriptLineNumber
        $file     = $MyInvocation.ScriptName
        Function expandErr {
            param ($message)
            $ex       = New-Object exception $message
            $id       = 'PesterAssertionFailed';
            $cat      = [Management.Automation.ErrorCategory]::InvalidResult ;
            New-Object Management.Automation.ErrorRecord $ex, $id, $cat,
                @{Message = $message; File = $file; Line=$lineNo; Linetext=$lineText}
        }
    }
    process {
        $testitems = @()
        #If we have been asked to check URL or title get them from the driver. Otherwise call Get-SEElement.
        if     ($URL)   {$testitems += $Global:SeDriver.Url   ; $With = 'URL' }
        elseif ($Title) {$testitems += $Global:SeDriver.Title ; $With = 'Title'}
        elseif ($Alert -or $NoAlert) {
            try  {
                $a =$Global:SeDriver.SwitchTo().alert()
                if($Alert) {$testitems += $a.Text ; $with = 'Alert' };
            }
            catch {
                if($Alert) {throw (expandErr  "Expected an alert but but none was displayed") }
            }
            finally {
                if($NoAlert -and $a) {throw (expandErr  "Expected no alert but an alert of '$($a.Text)' was displayed") }
            }
            if(     $a -and -not $Value -and $PassThru) {return $a}
            elseif ($a -and -not $value) {return}
        }
        else   {
            foreach ($s in $Selection) {
                $gSEParams =  @{By=$By; Selection=$s}
                if ($Timeout) {$gSEParams['Timeout'] = $Timeout}
                try           {$e = Get-SeElement @gSEParams }
                catch         {throw (expandErr $_.Exception.Message)}

                #throw if we didn't get the element, and if were only asked to check it was there, return gracefully
                if (-not $e) {throw (expandErr "Didn't find '$s' by $by")}
                elseif ($PSCmdlet.ParameterSetName -eq "DefaultPS" ) {
                    Write-Verbose "Found $s"
                    if ($PassThru) { $e}
                }
                else {
                #if we got here we're not in the default parameter set, weren't asked for URL or title, and got an element
                #with an element property (text, displayed etc or an attribute name. Get that ... )
                    switch ($with) {
                        'Text'      {$newTestItem = $e.Text}
                        'Displayed' {$newTestItem = $e.Displayed}
                        'Enabled'   {$newTestItem = $e.Enabled}
                        'TagName'   {$newTestItem = $e.TagName}
                        'X'         {$newTestItem = $e.Location.X}
                        'Y'         {$newTestItem = $e.Location.Y}
                        'Width'     {$newTestItem = $e.Size.Width}
                        'Height'    {$newTestItem = $e.Size.Height}
                        default     {$newTestItem = $e.GetAttribute($with)}
                    }
                    if (-not $newTestItem -and ($value -ne '')) {
                        throw (expandErr "Didn't find '$with' on element")
                    }
                    else {
                        Write-Verbose  "Found $with = '$newTestItem'"
                        $testitems += $newTestItem
                    }
                }
            }
        }
        #if we either had URl, Title or Element with either a property name or attribute name and we got one of those ... Now compare with Value
        #if operator was not passed, allow it to be taken from an alias for the -value
        if (-not $PSBoundParameters.ContainsKey('operator') -and $lineText -match ' -(eq|ne|match|notmatch|like|notlike|gt|lt) ') {
            $Operator = $matches[1]
        }
        foreach ($t in $testitems) {
            Switch ($Operator) {
                'eq'       {$result = ($t -eq       $value)}
                'ne'       {$result = ($t -ne       $value)}
                'like'     {$result = ($t -like     $value)}
                'notlike'  {$result = ($t -notlike  $value)}
                'match'    {$result = ($t -match    $value)}
                'notmatch' {$result = ($t -notmatch $value)}
                'gt'       {$result = ($t -gt       $value)}
                'le'       {$result = ($t -lt       $value)}
            }
            if (-not $result) {throw (expandErr  "$with had value of '$testitems'. The comparison '-$operator $value' failed.") }
            elseif ($PassThru -and $Alert) {$a}
            elseif ($PassThru)             {$e}
        }
    }
}