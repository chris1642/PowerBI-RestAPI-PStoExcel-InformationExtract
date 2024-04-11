# Check and set the execution policy
$currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
if ($currentPolicy -eq 'Restricted') {
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
}

# Check and install the necessary modules
$requiredModules = @('MicrosoftPowerBIMgmt', 'ImportExcel')
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Install-Module -Name $module -Scope CurrentUser -Force
    }
}

# Connect to the Power BI Service
Connect-PowerBIServiceAccount

# Initialize collections
$workspacesInfo = @()
$datasetsInfo = @()
$reportsInfo = @()
$appsInfo = @()
$reportPagesInfo = @()

# Fetch list of Workspaces available to the user
$workspaces = Get-PowerBIWorkspace

foreach ($workspace in $workspaces) {
    # Store basic workspace info
    $workspaceInfo = [PSCustomObject]@{
        WorkspaceId = $workspace.Id
        WorkspaceName = $workspace.Name
    }
    $workspacesInfo += $workspaceInfo

    # Datasets
    $workspaceDatasets = Get-PowerBIDataset -WorkspaceId $workspace.Id
    foreach ($dataset in $workspaceDatasets) {
        $datasetInfo = [PSCustomObject]@{
            WorkspaceName = $workspace.Name
            WorkspaceId = $workspace.Id
            DatasetId = $dataset.Id
            DatasetName = $dataset.Name
        }
        $datasetsInfo += $datasetInfo
    }

    # Reports
    $workspaceReports = Get-PowerBIReport -WorkspaceId $workspace.Id
    foreach ($report in $workspaceReports) {
        # Fetch the dataset associated with the report
        $reportDataset = $workspaceDatasets | Where-Object { $_.Id -eq $report.DatasetId }

        $reportInfo = [PSCustomObject]@{
            WorkspaceName = $workspace.Name
            WorkspaceId = $workspace.Id
            DatasetId = $reportDataset.Id
            DatasetName = $reportDataset.Name
            ReportId = $report.Id
            ReportName = $report.Name
        }
        $reportsInfo += $reportInfo
    }
}

# Fetch list of Apps available to the user
$appsUrl = "https://api.powerbi.com/v1.0/myorg/apps"
$apps = Invoke-PowerBIRestMethod -Method GET -Url $appsUrl | ConvertFrom-Json
foreach ($app in $apps.value) {
    $appInfo = [PSCustomObject]@{
        AppId = $app.id
        AppName = $app.name
        AppWorkspaceId = $app.workspaceId
    }
    $appsInfo += $appInfo
}

# Fetch Report Pages within Workspaces
foreach ($workspace in $workspaces) {
    $workspaceReports = Get-PowerBIReport -WorkspaceId $workspace.Id
    foreach ($report in $workspaceReports) {
        $pagesUrl = "https://api.powerbi.com/v1.0/myorg/groups/$($workspace.Id)/reports/$($report.Id)/pages"
        $pages = Invoke-PowerBIRestMethod -Method GET -Url $pagesUrl | ConvertFrom-Json
        
        foreach ($page in $pages.value) {
            $pageInfo = [PSCustomObject]@{
                WorkspaceId = $workspace.Id
                WorkspaceName = $workspace.Name
                ReportId = $report.Id
                ReportName = $report.Name
                PageDisplayName = $page.displayName
                PageName = $page.Name
                PageOrder = $page.order
            }
            $reportPagesInfo += $pageInfo
        }
    }
}

# Define the Excel file path
$excelFile = "C:\PowerBIReports\Power BI Information Extract.xlsx"

# Check if the Excel file already exists and delete it if it does
if (Test-Path $excelFile) {
    Remove-Item $excelFile -Force
}

# Export information to Excel
$workspacesInfo | Export-Excel -Path $excelFile -WorksheetName "Workspaces" -AutoSize
$datasetsInfo | Export-Excel -Path $excelFile -WorksheetName "Datasets" -AutoSize -Append
$reportsInfo | Export-Excel -Path $excelFile -WorksheetName "Reports" -AutoSize -Append
$appsInfo | Export-Excel -Path $excelFile -WorksheetName "Apps" -AutoSize -Append
$reportPagesInfo | Export-Excel -Path $excelFile -WorksheetName "ReportPages" -AutoSize -Append

# Notify completion
Write-Host "Export completed. Data is saved to $excelFile"
