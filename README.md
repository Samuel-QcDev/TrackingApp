# TrackingApp
VBA web-scrapping automation application

Installation instructions:

1. Save the 4 main files (TrackingAppPart1.xlam, TRACKING_MACRO_FILTERS.xslx, Logins Transport-entrepot FOR Macro.xlsx, TrackingApp Excel MenuTab.exportedUI)on the desktop of your computer, in a file called "TrackingApp" (if these files are in a different location, the Path for these files will need to be adjusted in the Module "ModPublicVariables");
2. Download the software "Selenium Basics" (https://github.com/florentbr/SeleniumBasic/releases/tag/v2.0.9.0), and install it without any included drivers;
3. Download the driver for Google Chrome (https://chromedriver.chromium.org/downloads) and move it to the Selenium Basics folder (C:\Users\~username~\AppData\Local\SeleniumBasic)
4. Make sure the Microsoft .NET Framework version 3.5 or earlier is installed on your computer (https://www.microsoft.com/en-ca/download/details.aspx?id=21)
5. Open Excel and select "Options -> Add-Ins
6. At the bottom of the Options window, beside Manage: , make sure "Excel Add-ins"is selected and click "GO";
7. Click "BROWSE", find & Select the file TrackingAppPart1.xlam, that you downloaded to the folder on your desktop. Click on OK.
8. In an Excel window, right-click in the empty space at the right of the ribbon (Menu icons at the top of the window) and select "Customize the Ribbon"
9. On the new options window, click the button "Import/Exxport", just above the OK Button and select "Import the customization file"
10. Browse to the file "TrackingApp Excel MenuTab.exportedUI", in the folder "../Desktop/TrackingApp", select it and click OK;
11. You should now have a new Tab at the top of the screen called "Tracking" and the App is installed

How to use the App:
1. Open the Test file
2. Click on the "Tracking Tab"
3. Click on either "Filter 1" or "Filter 2"
4. Click "Ok" to the message box that displays once the macro is done running
5. If the Window title is "Manual Track", close it and the Excel window under it will be called "Automatic Track"
6. While on the "Automatic Track" window, select the tab "Tracking" again and click on one of the buttons "Track"
7. Chrome should open and the tracking process should start. If you get an "Automation error message", it means you do not have the required .NET Framework installed. If you get an error message about the version of Chrome, you might need to download the latest version of the Chromedriver and copy it to Selenium Basics folder (overwrite the previous version)
