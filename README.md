The script will output an Excel doc with all the Power BI info (based on the permissions of the user running) related to:
 <br />
  <br />
- Workspaces
- Datasets
- Reports
- Pages within Reports
- Apps
- Reports within Apps
 <br />
  <br />
Use the PS1 format if you can directly run script files, otherwise use the Text format and copy/paste into PowerShell

1. The default export directory and file is set as: "C:\PowerBIReports\Power BI Information Extract.xlsx" - change if needed. (if you want to leverage in a refreshable Power BI dataset, sync a SharePoint folder to your computer and have the directory pointed there to ensure it's an online/web source.

2. Copy and paste text into PowerShell and run. You will have confirm your Power BI / Microsoft credentials from the script.

 <br />
 <br />
  <br />
   <br />
If you do not have all the required modules already installed, it will attempt to auto-install them (the 3 are listed below): <br />
 <br />

NuGet Provider <br />
MicrosoftPowerBIMgmt <br />
ImportExcel <br />




This will also check the PowerShell Execution Policy and update if needed (to allow the above modules to work).

All of this is scoped to the Current User - this means it will typically not require an Admin account and is only done at the user level, not computer. 
