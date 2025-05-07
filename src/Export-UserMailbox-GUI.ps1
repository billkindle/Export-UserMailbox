# PowerShell GUI for Exporting User Mailbox
# This script provides a graphical user interface (GUI) for exporting a user's mailbox using Microsoft 365 Compliance tools.

# Load required .NET assembly for Windows Forms
Add-Type -AssemblyName System.Windows.Forms

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Export User Mailbox" # Title of the form
$form.Size = New-Object System.Drawing.Size(400,300) # Initial size of the form
$form.StartPosition = "CenterScreen" # Center the form on the screen

# Create a label for the email input
$labelEmail = New-Object System.Windows.Forms.Label
$labelEmail.Text = "User Email:" # Label text
$labelEmail.Location = New-Object System.Drawing.Point(10,20) # Position of the label
$form.Controls.Add($labelEmail)

# Create a text box for the user to input their email
$textBoxEmail = New-Object System.Windows.Forms.TextBox
$textBoxEmail.Location = New-Object System.Drawing.Point(110,20) # Position of the text box
$textBoxEmail.Width = 250 # Width of the text box
$form.Controls.Add($textBoxEmail)

# Create a button to start the export process
$buttonExport = New-Object System.Windows.Forms.Button
$buttonExport.Text = "Export Mailbox" # Button text
$buttonExport.Location = New-Object System.Drawing.Point(10,60) # Position of the button
$form.Controls.Add($buttonExport)

# Create a label to display status messages
$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.Location = New-Object System.Drawing.Point(10,100) # Position of the label
$labelStatus.Size = New-Object System.Drawing.Size(350,50) # Size of the label
$form.Controls.Add($labelStatus)

# Create a "Close Application" button
$buttonClose = New-Object System.Windows.Forms.Button
$buttonClose.Text = "Close Application" # Button text
$buttonClose.Size = New-Object System.Drawing.Size(120, 30) # Button size
$buttonClose.Visible = $false # Initially hidden
$buttonClose.Add_Click({ $form.Close() }) # Close the form when clicked
$form.Controls.Add($buttonClose)

# Function to update status messages
function Update-Status {
    param ([string]$message)
    $labelStatus.Text = $message # Update the status label with the provided message
}

# Function to log actions
function Write-ActionLog {
    param ([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss" # Get the current timestamp
    "$timestamp - $message" | Out-File -FilePath "ActionLog.txt" -Append # Append the timestamped message to the log file
}

# Define the button click event
$buttonExport.Add_Click({
    try {
        # Get the email address from the text box
        $userEmail = $textBoxEmail.Text
        if (-not [string]::IsNullOrWhiteSpace($userEmail)) {
            # Validate the email format
            if ($userEmail -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
                Update-Status "Invalid email format. Please enter a valid email address."
                return
            }

            # Extract the username from the email address
            $username = $userEmail.Split('@')[0]

            # Define case and search names
            $caseName = "$username-MailboxExportCase"
            $searchName = "$username-UserMailboxSearch"

            # Connect to Security & Compliance PowerShell
            Update-Status "Connecting to Security & Compliance PowerShell..."
            Connect-IPPSSession -ErrorAction Stop

            # Create an eDiscovery case
            Update-Status "Creating eDiscovery case: $caseName"
            New-ComplianceCase -Name $caseName -Description "Case for exporting user mailbox" -ErrorAction Stop
            Write-ActionLog "eDiscovery case created: $caseName"

            # Create a content search
            Update-Status "Creating content search: $searchName"
            New-ComplianceSearch -Case $caseName -Name $searchName -ExchangeLocation $userEmail -Description "Search for user mailbox export" -ErrorAction Stop

            # Start the content search
            Update-Status "Starting content search: $searchName"
            Start-ComplianceSearch -Identity $searchName -ErrorAction Stop

            # Wait for the search to complete
            Update-Status "Waiting for search to complete..."
            do {
                Start-Sleep -Seconds 30
                $searchStatus = Get-ComplianceSearch -Identity $searchName | Select-Object Status
                Update-Status "Search status: $($searchStatus.Status)"
            } until ($searchStatus.Status -eq "Completed")

            # Initiate the export action
            Update-Status "Initiating export action for: $searchName"
            New-ComplianceSearchAction -SearchName $searchName -Export -Format FxStream -ErrorAction Stop -Confirm:$false

            # Notify the user of completion
            Update-Status "Export action created. Check the compliance portal for download instructions."

            # Make the "Close Application" button visible after successful completion
            $buttonClose.Visible = $true
        } else {
            Update-Status "Please enter a valid email address."
        }
    } catch {
        # Handle errors and log them
        $errorMessage = "An error occurred: $($_.Exception.Message)"
        Update-Status $errorMessage
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss" # Get the current timestamp
        "$timestamp - $errorMessage" | Out-File -FilePath "ErrorLog.txt" -Append # Append the timestamped error message to the log file
    }
})

# Adjust the "Close Application" button position when the form is resized
$form.Add_Resize({
    $buttonClose.Location = New-Object System.Drawing.Point(
        [int]($form.ClientSize.Width - $buttonClose.Width - 10), # 10px padding from the right
        [int]($form.ClientSize.Height - $buttonClose.Height - 10) # 10px padding from the bottom
    )
})

# Show the form
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()