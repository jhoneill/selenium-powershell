# [Unreleased]

## Added

- DLL loading is now from the PSD1 file instead of the PSM1 file.
- Smoother setting of web driver path on Mac or Linux platform.

- Send-SeKeys
  -Element and -Keys parameters are both mandatory.
  Selenium-keys are now cached in a script-scoped variable for performance. (Get-SEKeys can be dropped)

- Valdate-URL
   Function Renamed to validateUrl (private helper function does not need to use Verb-Noun naming and "validate" verb upsets script analyzer)
   Added the old name back as an alias (which _is_ allowed. )...
   Moved functionality into a validation class which so URLs are validated before enering functions

    - Checks for valid webdriver also moved to a validation class.

- Get-SeCookie & Invoke-SeScreenshot
   In line with other functions, -Driver prarameter is renamed -Target, with alias "Driver", and will come from $Global:SeDriver and functions throw an error if no driver is passed or found from Global var

- Start-SeFirefox & Start-SeChrome
    - Added aliases "SeChrome" and "SeFirefox" - user can now write in a sort of DSL "SeCHROME $URL - asDefault" etc.
    - Made -StartUrl the first parameter and use a parameter-validation class to check it.
    - Now call the .Navigate method directly instead of via Enter-SeURL if -StartURL is specified
    - Now use parameter-sets to avoid the conflict between Minimized/Maximized/FullScreen/Headless
    - Added parameter aliases so "Incognito" works with Firefox and "PrivateBrowsing" works with Chrome
    - Now trap a failure to return the driver, and simplified the subsequent if statements
    - Set ImplicitWait to 10 seconds
    - Added Parameter -AsDefaultDriver which sets $Global:SeDriver, instead of returning the driver
    - Added Parameter -Quiet to run tests without web driver prattle.
    - Added Parameter -DriverPath to Start-SeChrome for specifying a custom binary path
    - Allow AzureDevops environment vars (GeckoWebDriver / ChromeWebDriver)  or commandline parameter -WebDriverDirectory to specify source of Webdriver.
    - Removed a possible bug with piped input, by returning the driver in the process block instead of the end block

- Start-SeInternetExplorer
    - Added Aliases "SeInternetExplorer" & "SeIE"
    - Added parameters -StartUrl ,  -AsDefaultDriver which sets $Global:SeDriver, instead of returning the driver, and -Quiet to run tests without web driver prattle.
    - Allow AzureDevops environment var or commandline parameter -WebDriverDirectory to specify source of Webdriver
    - Set ImplicitWait to 10 seconds

- Start-SeEdge
    - Added Aliases "Start-SeLegacyEdge" & "LegacyEdge" . [Did have an alias of MSEdge but this conflicts with the Exe name ]
    - Added Parameters -StartUrl, -Maximized, -Minimized, FullScreen and -PrivateBrowsing (alias Incognito)
    - Added Parameters Plus -AsDefaultDriver which sets $Global:SeDriver, instead of returning the driver, and -Quiet to run tests without web driver prattle.
    - Added message if driver load errors - Webdriver is now added as a Windows Feature and will be found from windows\system32, which is on the path. Removed old driver and SHA file.
    - Set ImplicitWait to 10 seconds

- Stop-SeDriver
    Updated to handle the default driver being in $global:SeDriver. Added alias SeClose to mirror new function SeOpen

- Start-SeNewEdge  **New Function** with alias, credge to support Chromium based edge.
    Also added V79 webdriver. Specifying V80 / v81 from an ENV variable (EdgeWebDriver), or via command line requires it to be in the same directory as MSEdge.exe and -InPrivate, -Headless, and -BinaryLocation options don't work the web driver doesn't add them to the command line.

- Start-SeRemote **New Function**
  Connects to remote driver by URL and requested capablites tested with testingbot.com
```
                        Start-SeChrome  Start-SeEdge Start-SeFirefox  Start-SeInternetExplorer  Start-SeNewEdge  Start-SeRemote
    AsDefaultDriver	         X               X            X                X                         X                X
    Fullscreen               X               X            X                X                         X
    Headless                 X               X            X                X                         X
    ImplicitWait             X               X            X                X                         X                X
    Maximized                X               X            X                X                         X
    Minimized                X               X            X                X                         X
    Incognito                X             As alias	    As alias         As alias                  As alias
    PrivateBrowsing	       As alias          X            X                X                         X
    Quiet                    X               X            X                X                         X
    StartURL                 X               X            X                X                         X                 X
    WebDriverDirectory       X                            X                X                         X
    or Env var         ChromeWebDriver       -       GeckoWebDriver      IEWebDriver             EdgeWebDriver
    BinaryPath               X                                                                       X
    ProfileDirectoryPath     X                                                                       X
    DefaultDownloadPath      X                            X                                          X
    Arguments                X                            X
```

-SeOpen **New Function**
    - Takes parameter -In (name of browser) and opens that browser and optionally naviagates to destination in -URL - The browser is opened as with -AsDefault and other options passed in -Options.
    - If the browser name is omitted the function will look for $env:DefaultBrowser.
    - Mainly for Pester/DSL freindliness - by changing the environment variable a different browser can be used in different test runs

- Send-SeClick **New Function** (with alias SeClick) {replacement for Invoke-SeClick}
    - Element can be passed by position (first)
    - -JavaScriptClick can be abreviated -JS
    - Driver is ignored  (element.wrapedDriver used instead)
    - Added parameters -SleepSeconds and -PassThru

- New-SeScreenShot: **New Function** with alias SeScreenshot. Merges functionality of Invoke-ScreenShot and SaveSeScreenShot; can pass
```
        -Path,
        -Passthru (Alias PT)
        -Path & -Passthru
   or   -AsBase64
```
`       -Target` with alias "Driver", can come from $Global:SeDriver and throws an error if no driver passed or found from Global var

- Open-SeURL *New Function** with alias SeNavigate
    - Url is first on the commandline and mandatory
    - Driver is renamed target, with alias "Driver", and can come from $Global:SeDriver or from the pipeline. Throws an error if no driver passed or found from the Global variable
    - Made interoperable with Enter-SeURL via aliases
    - Commented out original Enter-SeURL and changed PSD1,

- Get-SeElement **New Function** with alias seElement
    - New -By parameter with possible values "CssSelector", "Name", "Id", "ClassName", "LinkText","PartialLinkText", "TagName", "XPath"
    combined with a -Selection parameter which holds the value, replaces the 8 parameters with those names and parametersets.
    - Wait made optional; -Wait without -Timeout sets timeout to 30 , and -Timeout 5 works without -Wait
    - Element and -Driver are merged as -Target which has aliases of 'Element' and 'Driver' and accepts pipeline input.
    - Made interoperable with Find-seElement by aliases
    - Commented out Find-seElement and updated PSD1.

- Switch-SeFrame **New Function** with alias SeFrame
    Selects a frame.

- Clear-SeAlert **New Function** with aliases SeAccept, SeDismiss
    - Accepts or dismisses an alert box so execution can continue
    - If called using the aliases, automatically selects Accept or dismiss.

- SeType **New Function**
     - More DSL friendly form of Send-SeKeys
     - Accepts element via the pipeline;  -Keys is the first parameter
     - Supports clearing the text box, submit and passthru, (Passthru allows the element to be fetched and keys typed, and then Pester to throw if element was not found.)

- Get-SeSelectionOption **New Function** alias SeSelection
    For setting and checking options in a dropdown lists.

- SeShouldHave **New Function**
    - Drives pester tests as a single command See examples for details, but allows things like
    `SeShouldHave -url -match 'packages\?q=selenium'`
    `SeShouldHave -Selection $linkpath -By XPath -With Text -Operator Like  -Value "*selenium*"`
    which can be shortened to
    `seShouldHave $linkpath text -Like "*selenium*"`

- Added some new examples - showing DSL style in Pester.

- Reorder PSD file for easy reading & add file list.

- Create CI directory with files for doing a build / test pass is Azure devops.