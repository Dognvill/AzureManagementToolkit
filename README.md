# Azure User Management Toolkit

## Overview
This PowerShell script provides a comprehensive toolkit for managing Entra ID users, including:
- User Onboarding
- User Offboarding
- Email and Calendar Management
- Authentication Configuration

### Step-by-Step Installation

1. **Download the Toolkit**
   - Create a dedicated folder (e.g., `C:\AzureManagement`)
   - Place the following files in this folder:
     - `AzureManagement.ps1`
     - `Launch.bat`
     
2. **Launch Batch File**
  - Double-click Launch.bat
  - The script will automatically:
    - Create required folders

3. **Run Validate Environment**
  - Select option 4 to validate your enviroment and install the required modules

### System Requirements
- PowerShell 7.0 or later
- Windows 10/11 or Windows Server 2019/2022
- Appropiate Azure Permissions

### Required PowerShell Modules
- Microsoft.Graph (v2.0.0 or later)
- ExchangeOnlineManagement (v3.0.0 or later)

### Security Considerations
- Run the script with appropriate administrative permissions
- Keep authentication configurations secure and private
- Regularly review and rotate credentials
- Follow principle of least privilege

### Disclaimer
This script is provided as-is. Always test in a controlled environment before production use. 
The authors are not responsible for any unintended consequences of using this tool.
