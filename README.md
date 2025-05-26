# SharePoint Shortcut Generator

This project simplifies the process of creating shortcuts to individual SharePoint sites and folders for specific users. It allows you to select users, sites, and specific folders within those sites, and then creates shortcuts in each user's OneDrive. The script checks if the user has permission to access the site and whether a shortcut already exists. Note that the Microsoft Graph API does not provide a list of existing shortcuts, and this project does not implement a search function for them. If a shortcut already exists, a warning is displayed and the script continues with the remaining items.

Additionally, new users with valid licenses do not have their OneDrive provisioned automatically. You must either log in with the account or provision OneDrive manually. For this purpose, we have included a script called `OneDrive-Provisioning.ps1`. With this script, you can select the users you want to provision, and it will provision OneDrive for all selected users.

## Requirements

- PowerShell 7 (will be installed if not available)
- WebView2 (will be installed if not available)
- A working tenant
- OneDrive/SharePoint license for users
- Entra application with correct redirect URI and API permissions
    - API permissions:
        - Microsoft Graph permissions:
            - files.readwrite.all
            - user.read.all
        - SharePoint permissions:
            - allsites.fullcontrol
            - allsites.manage
            - myfiles.read
            - myfiles.write
            - user.readwrite.all
    - Redirect URI: https://login.microsoftonline.com/common/oauth2/nativeclient

## Usage

You can run the `Sharepoint-Shortcut-Generator` script as follows:

```powershell
# Without Verbose
.\Sharepoint-Shortcut-Generator.ps1 -tenantId 'your-tenant-id' -clientId 'your-client-id'

# With Verbose
.\Sharepoint-Shortcut-Generator.ps1 -tenantId 'your-tenant-id' -clientId 'your-client-id' -Verbose
```

The same usage applies to the `OneDrive-Provisioning` script:

```powershell
# Without Verbose
.\OneDrive-Provisioning.ps1 -tenantId 'your-tenant-id' -clientId 'your-client-id'

# With Verbose
.\OneDrive-Provisioning.ps1 -tenantId 'your-tenant-id' -clientId 'your-client-id' -Verbose
```