# Azure User Management Toolkit

## Overview
This PowerShell script provides a comprehensive toolkit for managing Entra ID users, designed to streamline and secure user lifecycle management.

### Overview

### Step-by-Step Installation
1. **User Onboarding**
   - Create new user accounts in Microsoft Entra ID
   - Assign licenses and group memberships
   - Generate initial passwords
   - Create detailed onboarding reports


2. **User Offboarding**
   - Securely remove user access and resources
   - Remove group memberships
   - Revoke licenses
   - Disable user accounts
   - Convert mailboxes to shared resources
   - Generate comprehensive offboarding reports


3. **Email & Calendar Management**
   - Manage email delegation
   - Configure calendar permissions
   - Add/remove delegate access
   - Convert mailboxes to shared resources

4. **Validate Environment**
   - Check and install required PowerShell modules
   - Verify Microsoft Graph and Exchange Online modules
   - Ensure system meets minimum requirements
   - Install/update necessary components

5. **Configure Authentication**
   - Create and manage authentication profiles
   - List available configurations
   - Test authentication for saved configurations
   - Manage app registration details securely

6. **Documentation**
   - Interactive help and usage guidance

### Installation
1. **Download the Toolkit**
   - Create a dedicated folder (e.g., `C:\AzureManagement`)
   - Place the following files in this folder:
     - `AzureManagement.ps1`
     - `Launch.bat`
     
2. **Start Batch File**
     - Double-click `Launch.bat`

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

# MIT License

Copyright (c) 2024 [John Bignold]

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
