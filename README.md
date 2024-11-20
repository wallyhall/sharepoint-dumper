# sharepoint-dumper
Dumps the contents of a Sharepoint site to a local disk or network share

This script requires the PnP Powershell module to run and an Entra ID Application.

Installing their module locally can be performed thusly:

```
Install-Module -Name `"PnP.PowerShell`" -AllowClobber
```

Creation of the Entra ID Application is [documented by the PnP folk](https://pnp.github.io/powershell/articles/registerapplication.html).  YMMV ... (partial) automation of this is shipped within the script:

```
./sharepoint.ps1 -CreateAppId -TenantId <YOUR_AZURE_TENANT_ID>
```

Provided your user has the relevant permissions (e.g. is a Site Admin), you can now attempt to dump a site:

```
.\sharepoint.ps1 -ExportPath D:\Sharepoint -CompanyUrl myorg.sharepoint.com -Site mysitetodump -AppId aaaaaaaa-bbbb-cccc-dddd-eeeeeeee
```

The following Excel formula may be useful for mass-creating one-liners to dump multiple sites from a Sharepoint Admin CSV site export ... (replace every `***` with your necessary values):

```
=CONCAT(".\sharepoint.ps1 -ExportPath D:\Sharepoint -CompanyUrl *****.sharepoint.com -Site ", SUBSTITUTE(B1, "https://*****.sharepoint.com/sites/", ""), " -AppId *****-*****-****-****-******")
```
