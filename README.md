# M365UserLicenseChange

Easily adjusts Microsoft 365 Licenses assigned to users (in bulk)  

- For each user listed in the CSV file, specify a comma separated list of licenses to add (and optionally to remove).  
- If licenses are already in place, the user is skipped.  So it's safe to run the script multiple times, or to interrupt it and run again.

Main Screen  
<img src=https://raw.githubusercontent.com/ITAutomator/Assets/main/M365UserLicenseChange/M365UserLicenseChange.png alt="screenshot" width="400">

User guide (pdf): Click [here](https://github.com/ITAutomator/M365UserLicenseChange/blob/main/M365UserLicenseChange%20Readme.pdf)  

Download from GitHub as [ZIP](https://github.com/ITAutomator/M365UserLicenseChange/archive/refs/heads/main.zip)  
Or Go to GitHub [here](https://github.com/ITAutomator/M365UserLicenseChange) and click `Code` (the green button) `> Download Zip`  

In this example  

- User1 licenses will be replaced with M365 E5
- User2 licenses will be replaced with M365 Business Premium and Information Governance
- User3 will have Information Governance added
- User4 will have Information Governance added

## Fields in the CSV

- User
- LicensesToAdd
- LicensesToRemove

## Valid License Names

Microsoft publishes a valid list of Sku’s here: [link](https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference) (CSV version is here: [link](https://download.microsoft.com/download/e/3/e/e3e9faf2-f28b-490a-9ada-c6089a1fc5b0/Product%20names%20and%20service%20plan%20identifiers%20for%20licensing.csv)).  
This program uses the values from the String ID field.

## Notes

- Use email address or sign in name for user
- Use can use a comma separated list of licenses
- You can use the keyword ‘<all>’ in LicensesToRemove 
- If there is a License in both ToAdd and ToRemove, adding wins.
- User Licenses will first be checked to see if they already comply with request.
- An Invalid SKU will pause the code.  In this case a valid list of SKUs will be displayed so the entry can be fixed.
- Licenses must be available, otherwise the entry will be skipped.
- The script is designed so that it does nothing if nothing needs to be done.  It can be run repeatedly safely.

## PowerShell Modules

These modules are required.  You can use the included `ModuleManager.cmd` script to install them  

- Microsoft.Graph.Authentication
- Microsoft.Graph.Users
- Microsoft.Graph.Identity.DirectoryManagement

