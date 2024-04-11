Use the PS1 format if you can directly run script files, otherwise use the Text format and copy/paste into PowerShell

1. Default export directory and file is set as: "C:\PowerBIReports\Power BI Information Extract.xlsx" - change if needed (if you want to leverage in a refreshable Power BI dataset, sync a SharePoint folder to your computer and have the directory pointed there to ensure it's an online/web source.

2. Copy and paste text into PowerShell and run. You will have confirm your Power BI / Microsoft credentials from the script.



If you do not have all the required modules already installed, it will attempt to auto-install them (the 3 are listed below):

NuGet Provider

MicrosoftPowerBIMgmt

ImportExcel

This will also check the PowerShell Execution Policy and update if needed (to allow the above modules to work).

All of this is scoped to the Current User - this means it will typically not require an Admin account and is only done at the user level, not computer. 
