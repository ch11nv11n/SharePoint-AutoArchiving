#loads sharepoint snapin if not loaded
if (-not(Get-PSSnapin -Name "Microsoft.Sharepoint.Powershell")) {
    Add-PSSnapin -Name "Microsoft.Sharepoint.Powershell"
    Write-Host "`nSharepoint Snapin loaded successfully...`n" -ForegroundColor Yellow 
}

#declaring var as name of site
$ordersArchiveUrl = 'http://portal/ordersarchive'

#returns portal site collection and stores in a var
$site = Get-SPSite http://portal

#returns all subsites in the site collection and stores in var
$web = Get-SpWeb http://portal/ordersarchive

#gets collection of all lists within the ordersarchive subsite
$SPListCollection = $web.Lists

#returns collection of custom list templates from subsite and stores in var
$custTemplate = $site.GetCustomListTemplates($web)

#vars to get current year and month name
$year = (Get-Date).year
$monthNum = (Get-Date).Month
$monthName = (Get-Culture).DateTimeFormat.GetMonthName($monthNum)

#creating folder names based on date info from previous vars
$archiveFolderName = "Archive_$($year)_$($monthName)"
$alreadyEnteredFolderName = "Archive_AlreadyEntered_$($year)_$($monthName)"

#creating blank description
$description = ""

#if folder name equals null (doesn't exist) in the collection list
if (($SPListCollection.TryGetList($archiveFolderName)) -eq $null) {
    
    #Adds folder with specified folder name, description, and template vars
    $SPListCollection.Add($archiveFolderName,$description,$custTemplate["Orders Archive New"]);

    #selecting folder that was created and setting value to var
    $archiveLibrary = $SPListCollection[$archiveFolderName]

    #disabling folder on quick launch
    $archiveLibrary.OnQuickLaunch = $false

    #updating folder
    $archiveLibrary.Update()
}

#if folder name equals null (doesn't exist) in the collection list
if (($SPListCollection.TryGetList($alreadyEnteredFolderName)) -eq $null) {
    
    #Adds folder with specified folder name, description, and template vars
    $SPListCollection.Add($alreadyEnteredFolderName,$description,$custTemplate["Archive Already Entered New"]);

    #selecting folder that was created and setting value to var
    $archiveAlreadyEnteredLibrary = $SPListCollection[$alreadyEnteredFolderName]

    #disabling folder on quick launch
    $archiveAlreadyEnteredLibrary.OnQuickLaunch = $false

    #updating folder
    $archiveAlreadyEnteredLibrary.Update()
    }
<## Changing Destination Document Library for the Archive Settings ##>

#Creates the number of the last month and stores it in a var
$lastMonthNum = (Get-Date).Month - 1

#Converts the number of the last month into the name
$lastMonthName = (Get-Culture).DateTimeFormat.GetMonthName($lastMonthNum)

#Selecting the orders url and storing in a var
$ordersUrl = "http://portal/orders"

#returns all subsites in the site collection and stores in var
$ordersWeb = Get-SPWeb $ordersUrl

#Selecting the specific list and setting to a var
$archiveSettingList = $ordersWeb.Lists["Archive Settings"]

#Selecting the column "destination document library, where the column is equal to the last month's archive folder name and setting to a var
$SPItem = $archiveSettingList.Items | Where {$_["Destination Document Library"] -eq "Archive_$($year)_$($lastMonthName)"}

#Setting the column value to the name of the current archive folder name
$SPItem["Destination Document Library"] = $archiveFolderName

#Updates the destination document library value
$SPItem.Update()

#Declaring var for date formatting
$timestamp = Get-Date -Format g | foreach {$_ -replace ":", "."} | foreach {$_ -replace "/","-"}

#Sending email for successful creation
if (($SPListCollection.TryGetList($archiveFolderName) -ne $null) -and ($SPListCollection.TryGetList($alreadyEnteredFolderName) -ne $null)) {
    Send-MailMessage -to "USER <EMAIL@domain.com>", "USER2 <EMAIL2@domain.com>" -Subject "Archive Folder Creation Log $($timestamp)"  -SmtpServer 0.0.0.0 -from "SERVER@domain.com" -BodyAsHtml `
    -Body "The following $($ordersArchiveUrl) Document Libraries have been created successfully for the month of: <b>$($monthName)</b>`n
<ul> <li><b>$($archiveFolderName)</b></li> <li><b>$($alreadyEnteredFolderName)</b></li> </ul>`n
The Archive Settings Destination Document Library has successfully been changed to: <b>$($archiveFolderName)</b>"     
}

#Sending email if unsuccessful creation
elif (($SPListCollection.TryGetList($archiveFolderName) -eq $null) -and ($SPListCollection.TryGetList($alreadyEnteredFolderName) -eq $null)) {
        Send-MailMessage -to "USER <EMAIL@domain.com>", "USER2 <EMAIL2@schoolhealth.com>" -Subject "Archive Folder Creation Log $($timestamp)"  -SmtpServer 0.0.0.0 -from "SERVER@domain.com" -BodyAsHtml `
    -Body "The following $($ordersArchiveUrl) Document Library creations <b>FAILED</b> for the month of: <b>$($monthName) </b>`n
<ul> <li><b>$($archiveFolderName)</b></li> <li><b>$($alreadyEnteredFolderName)</b></li> </ul>`n
Please check the portal for more information"
}

#removes connections to the portal that we established to prevent memory leaks
$site.Dispose()
$web.Dispose()
$ordersWeb.Dispose()