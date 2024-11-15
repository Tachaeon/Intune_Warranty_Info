#Requires -Version 7

#Adjust sleep timer for Selenium websites. If it goes too fast or you see errors, increase this number.
$SleepTimer = "3"
$CSVSavePath = "C:\Install\Intune_Warranty_Info.csv"

#region Selenium
#Adjust Selenium path here.
$SelDir = "C:\Projects\SAWAA\Selenium"
Add-Type -Path "$SelDir\WebDriver.dll"

$DriverFolder = Join-Path -Path $SelDir -ChildPath "BrowserDrivers"
$ManagerPath = Join-Path -Path $SelDir -ChildPath "selenium-manager.exe"

if (-not (Test-Path $DriverFolder)) {
    New-Item -ItemType Directory -Path $DriverFolder -Force | Out-Null
}

#Get the correct chrome driver
$SeleniumManagerResults = & $ManagerPath --browser chrome --cache-path $DriverFolder

#Find the location of the exe that was either downloaded or already cached
if (($SeleniumManagerResults | Where-Object { $_ -like "*Driver path*" }) -match "Driver path:\s(.+)$") {
    $DriverPath = $Matches[1]
}
else {
    throw "Unable to determine web driver path!"
}

#Create the Chrome options
$options = New-Object OpenQA.Selenium.Chrome.ChromeOptions
$options.ImplicitWaitTimeout = [System.TimeSpan]::FromSeconds(2)
#$options.AddArgument("--headless")
$options.AddArgument("--silent")
$options.AddArgument("--log-level=3")
#$options.AddArgument("start-minimized")
$options.AddArgument("start-maximized")
$options.AddArgument("--no-sandbox")
#$options.AddArgument("--disable-gpu")
$options.AddExtension("$SelDir\uBlock-Origin-Chrome-Web-Store.crx")
#$options.addArguments("force-device-scale-factor=0.75")
#$options.addArguments("high-dpi-support=0.75")

#Create the Chrome driver service, using the driver that was retrieved from the driver manager
$service = [OpenQA.Selenium.Chrome.ChromeDriverService]::CreateDefaultService($DriverPath)
$service.SuppressInitialDiagnosticInformation = $true
$service.HideCommandPromptWindow = $true

#Create the driver session
$Selenium = New-Object OpenQA.Selenium.Chrome.ChromeDriver($service, $options)
#endregion

#Collect Information from Graph
$IntuneWarrantyData = @()
Connect-MgGraph -Scopes "DeviceManagementManagedDevices.Read.All"
$devices = Get-MgDeviceManagementManagedDevice -All
$deviceInfo = $devices | Select-Object @{Name = 'SerialNumber'; Expression = { $_.SerialNumber } }, @{Name = 'Manufacturer'; Expression = { $_.Manufacturer } }, @{Name = 'UserDisplayName'; Expression = { $_.UserDisplayName } }, @{Name = 'UserPrincipalName'; Expression = { $_.UserPrincipalName } }

#region Lenovo
foreach ($lenovo in $deviceInfo) {
    $Manufacturer = $lenovo.Manufacturer
    if ($Manufacturer -eq "LENOVO") {
        try {
            $serialnumber = $lenovo.serialnumber
            $username = $lenovo.UserDisplayName
            $email = $lenovo.UserPrincipalName
            # Get device info from Lenovo API
            $Device_Info = Invoke-RestMethod "https://pcsupport.lenovo.com/us/en/api/v4/mse/getproducts?productId=$serialNumber"
            $Device_ID = $Device_Info.id
            $Warranty_url = "https://pcsupport.lenovo.com/us/en/products/$Device_ID/warranty"
    
            # Optional: Add custom headers if needed
            # $headers = @{
            # "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
            # }
        }
        Catch {
            Write-Warning "Cannot get information for the serial number: $serialNumber"
            Continue
        }

        try {
            $Web_Response = Invoke-WebRequest -Uri $Warranty_url -Method GET # -Headers $headers
        }
        Catch {
            Write-Warning "Cannot get warranty info for the serial number: $serialNumber"
            Continue
        }
    
        if ($Web_Response.StatusCode -eq 200) {
            $HTML_Content = $Web_Response.Content
    
            # Regular expressions to extract warranty information
            $Pattern_Status = '"warrantystatus":"(.*?)"'
            $Pattern_Status2 = '"StatusV2":"(.*?)"'
            $Pattern_StartDate = '"Start":"(.*?)"'
            $Pattern_EndDate = '"End":"(.*?)"'
            $Pattern_DeviceModel = '"Name":"(.*?)"'
            
            # Extract information using regex
            $Status_Result = ([regex]::Match($HTML_Content, $Pattern_Status)).Groups[1].Value.Trim()
            $Statusv2_Result = ([regex]::Match($HTML_Content, $Pattern_Status2)).Groups[1].Value.Trim()
            $StartDate_Result = ([regex]::Match($HTML_Content, $Pattern_StartDate)).Groups[1].Value.Trim()
            $EndDate_Result = ([regex]::Match($HTML_Content, $Pattern_EndDate)).Groups[1].Value.Trim()
            $Model_Result = ([regex]::Match($HTML_Content, $Pattern_DeviceModel)).Groups[1].Value.Trim()
    
            # Fallbacks in case data is missing
            $Status_Result = if ($Status_Result) { $Status_Result } else { "Cannot get status info" }
            $Statusv2_Result = if ($Statusv2_Result) { $Statusv2_Result } else { "Cannot get status info" }
        }
        else {
            Write-Output "Failed to retrieve warranty information. Status Code: $($Web_Response.StatusCode)"
            Continue
        }
        $Warranty_Object = [PSCustomObject]@{
            Manufacturer = $Manufacturer
            Username     = $username
            Email        = $email
            SerialNumber = $serialNumber
            Model        = $Model_Result
            Status       = $Status_Result
            IsActive     = $Statusv2_Result
            StartDate    = $StartDate_Result
            EndDate      = $EndDate_Result
        }
        $IntuneWarrantyData += $Warranty_Object
        Write-Host "Added $Warranty_Object" -ForegroundColor Yellow
    }
}
#endregion

#region Dell
foreach ($dell in $deviceInfo) {
    $Manufacturer = $dell.Manufacturer
    if ($Manufacturer -eq "Dell Inc.") {
        try {
            $DellSerial = $dell.SerialNumber
            $username = $dell.UserDisplayName
            $email = $dell.UserPrincipalName
            $Selenium.Navigate().GoToUrl("https://www.dell.com/support/home/en-us/product-support/servicetag/$DellSerial/overview")
            Start-Sleep $SleepTimer
            $Selenium.FindElement([OpenQA.Selenium.By]::cssSelector("#viewDetailsWarranty")).Click()
            Start-Sleep $SleepTimer
            $Model_Result = $Selenium.FindElement([OpenQA.Selenium.By]::cssSelector("#warrantyDetailsPopup > div > div > div.modal-body.pt-25 > div:nth-child(1) > div.d-none.d-sm-block.card.pl-5.pt-4.pr-6.pb-5.mb-5 > div > div > div.flex-column.w-100 > div:nth-child(1) > h1")).Text
            $Status_Result = $Selenium.FindElement([OpenQA.Selenium.By]::cssSelector("#supp-svc-status-txt > span")).Text
            $Statusv2_Result = $Selenium.FindElement([OpenQA.Selenium.By]::cssSelector("#supp-svc-status-txt > span")).Text
            $StartDate_Result = $Selenium.FindElement([OpenQA.Selenium.By]::cssSelector("#dsk-purchaseDt")).Text
            $EndDate_Result = $Selenium.FindElement([OpenQA.Selenium.By]::cssSelector("#dsk-expirationDt")).Text

            $Warranty_Object = [PSCustomObject]@{
                Manufacturer = $Manufacturer
                Username     = $username
                Email        = $email
                SerialNumber = $DellSerial
                Model        = $Model_Result
                Status       = $Status_Result
                IsActive     = $Statusv2_Result
                StartDate    = $StartDate_Result
                EndDate      = $EndDate_Result
            }
            $IntuneWarrantyData += $Warranty_Object
            Write-Host "Added $Warranty_Object" -ForegroundColor Yellow
        }
        catch {
            Write-Host "Either $DellSerial was not found or Selenium automation is too fast." -ForegroundColor Red
            Continue
        }
    }
}
#endregion

#region HP
foreach ($HP in $deviceInfo) {
    $Manufacturer = $HP.Manufacturer
    if ($Manufacturer -eq "HP") {
        try {
            $Selenium.Navigate().GoToUrl("https://support.hp.com/us-en/check-warranty")
            $Selenium.FindElement([OpenQA.Selenium.By]::cssSelector("#inputtextpfinder")).SendKeys($HP.serialnumber)
            $Selenium.FindElement([OpenQA.Selenium.By]::cssSelector("#FindMyProduct")).Click()
            Start-Sleep $SleepTimer
            $Selenium.ExecuteScript("document.body.style.zoom = '.30';")
        
            $HPSerial = $HP.serialnumber
            $username = $HP.UserDisplayName
            $email = $HP.UserPrincipalName
            $Model_Result = $Selenium.FindElement([OpenQA.Selenium.By]::cssSelector("#directionTracker > app-layout > app-check-warranty > div > div > div.check-warranty-container-intra > app-warranty-details > div > div.details-container.ng-tns-c1772239254-0 > main > div.product-info.ng-tns-c1772239254-0 > div.product-info-text.ng-tns-c1772239254-0 > h2")).Text
            $Status_Result = $Selenium.FindElement([OpenQA.Selenium.By]::cssSelector("#directionTracker > app-layout > app-check-warranty > div > div > div.check-warranty-container-intra > app-warranty-details > div > div.details-container.ng-tns-c1772239254-0 > main > div.additional-information.ng-tns-c1772239254-0.ng-star-inserted > div > div.ng-trigger.ng-trigger-slideInOut.ng-tns-c1772239254-0.ng-star-inserted > div > div > div:nth-child(1) > div:nth-child(3) > div.text.ng-tns-c1772239254-0")).Text
            $Statusv2_Result = $Selenium.FindElement([OpenQA.Selenium.By]::cssSelector("#warrantyStatus > div.warrantyInfo.ng-star-inserted > div.warrantyStatus")).Text
            $StartDate_Result = $Selenium.FindElement([OpenQA.Selenium.By]::cssSelector("#directionTracker > app-layout > app-check-warranty > div > div > div.check-warranty-container-intra > app-warranty-details > div > div.details-container.ng-tns-c1772239254-0 > main > div.additional-information.ng-tns-c1772239254-0.ng-star-inserted > div > div.ng-trigger.ng-trigger-slideInOut.ng-tns-c1772239254-0.ng-star-inserted > div > div > div:nth-child(1) > div:nth-child(4) > div.text.ng-tns-c1772239254-0")).Text
            $EndDate_Result = $Selenium.FindElement([OpenQA.Selenium.By]::cssSelector("#directionTracker > app-layout > app-check-warranty > div > div > div.check-warranty-container-intra > app-warranty-details > div > div.details-container.ng-tns-c1772239254-0 > main > div.additional-information.ng-tns-c1772239254-0.ng-star-inserted > div > div.ng-trigger.ng-trigger-slideInOut.ng-tns-c1772239254-0.ng-star-inserted > div > div > div:nth-child(1) > div:nth-child(5) > div.text.ng-tns-c1772239254-0")).Text

            $Warranty_Object = [PSCustomObject]@{
                Manufacturer = $Manufacturer
                Username     = $username
                Email        = $email
                SerialNumber = $HPSerial
                Model        = $Model_Result
                Status       = $Status_Result
                IsActive     = $Statusv2_Result
                StartDate    = $StartDate_Result
                EndDate      = $EndDate_Result
            }
            $IntuneWarrantyData += $Warranty_Object
            Write-Host "Added $Warranty_Object" -ForegroundColor Yellow
            $Selenium.ExecuteScript("document.body.style.zoom = '1';")
        }
        catch {
            Write-Host "Either $HPSerial was not found or Selenium automation is too fast." -ForegroundColor Red
            Continue
        }
    }
}
#endregion

#Save to CSV
$IntuneWarrantyData | Export-Csv -Path $CSVSavePath -NoTypeInformation -Encoding UTF8

#Ding fries are done.
[void] [System.Reflection.Assembly]::LoadWithPartialName("PresentationFramework")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Speech")
$Speak = New-Object System.Speech.Synthesis.SpeechSynthesizer
$Speak.SelectVoice("Microsoft Zira Desktop")
$Speak.Speak('Scripts done.')
[System.Windows.MessageBox]::Show("Scripts done.", "Complete", 'OK', 'Information')