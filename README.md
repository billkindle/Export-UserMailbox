# Export User Mailbox GUI

This project provides a PowerShell GUI application that allows users to export a user's mailbox from Microsoft 365. The application simplifies the process of creating an eDiscovery case, performing a content search, and exporting the mailbox data.

## Project Structure

```
Export-UserMailbox-GUI
├── src
│   ├── Export-UserMailbox-GUI.ps1  # Main GUI script
│   └── helpers
│       └── Export-UserMailbox.ps1   # Helper script with mailbox export commands
└── README.md                         # Documentation for the project
```

## Prerequisites

- PowerShell 5.1 or later
- Microsoft 365 Compliance Center permissions to create eDiscovery cases and perform content searches
- Required modules:
  - Exchange Online Management
  - Compliance Center PowerShell

## How to Run the Application

1. Open PowerShell as an administrator.
2. Navigate to the project directory:
   ```powershell
   cd path\to\Export-UserMailbox-GUI\src
   ```
3. Run the GUI application:
   ```powershell
   .\Export-UserMailbox-GUI.ps1
   ```
4. Fill in the required fields in the GUI:
   - **User Email**: Enter the email address of the user whose mailbox you want to export.
   - **Case Name**: Provide a name for the eDiscovery case.
   - **Search Name**: Provide a name for the content search.

5. Click the "Export Mailbox" button to initiate the export process.

## Notes

- Ensure that you have the necessary permissions to perform mailbox exports.
- The export results will be available in the Microsoft Purview compliance portal under the Content Search section.

For any issues or contributions, please refer to the project's repository or contact the maintainer.