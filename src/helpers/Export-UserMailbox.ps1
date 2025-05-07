# filepath: Export-UserMailbox-GUI/src/helpers/Export-UserMailbox.ps1
# Parameters
param (
    [string]$caseName = "MailboxExportCase",
    [string]$searchName = "UserMailboxSearch",
    [string]$userEmail = "user@domain.com"  # Replace with the target user's email address
)

# Connect to Security & Compliance PowerShell
Write-Host "Connecting to Security & Compliance PowerShell..."
Connect-IPPSSession

# Create a new eDiscovery case
Write-Host "Creating eDiscovery case: $caseName"
New-ComplianceCase -Name $caseName -Description "Case for exporting user mailbox" -ErrorAction Stop

# Create a content search for the user's mailbox
Write-Host "Creating content search: $searchName"
New-ComplianceSearch -Case $caseName -Name $searchName -ExchangeLocation $userEmail -Description "Search for user mailbox export" -ErrorAction Stop

# Start the content search
Write-Host "Starting content search: $searchName"
Start-ComplianceSearch -Identity $searchName

# Wait for the search to complete
Write-Host "Waiting for search to complete..."
do {
    Start-Sleep -Seconds 30
    $searchStatus = Get-ComplianceSearch -Identity $searchName | Select-Object Status
    Write-Host "Search status: $($searchStatus.Status)"
} until ($searchStatus.Status -eq "Completed")

# Initiate the export action
Write-Host "Initiating export action for: $searchName"
New-ComplianceSearchAction -SearchName $searchName -Export -Format FxStream

# Retrieve export details
Write-Host "Retrieving export details..."
$exportAction = "$searchName_Export"
Get-ComplianceSearchAction -Identity $exportAction -IncludeCredential | Format-List ContainerUrl, ExportSasToken

Write-Host "Export action created. To download the PST file, go to the Microsoft Purview compliance portal, navigate to Content Search, select the search ($searchName), and use the Export tab to download results via the eDiscovery Export Tool."