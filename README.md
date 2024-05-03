# M365UserLicenseChange
Changes Microsoft 365 Licenses assigned to users (in bulk)
![image](https://github.com/ITAutomator/M365UserLicenseChange/assets/135157036/36144eb9-75a7-4d5a-97a1-370de44dc25e)


Overview
Changes Microsoft 365 Licenses assigned to users (in bulk).
![image](https://github.com/ITAutomator/M365UserLicenseChange/assets/135157036/be9e59e0-f190-427c-a4dc-f32a04b358d6)

In this example 
User1 licenses will be replaced with M365 E5
User2 licenses will be replaced with M365 Business Premium and Information Governance
User3 will have Information Governance added
User4 will have Information Governance added

Fields in the CSV
User
LicensesToAdd
LicensesToRemove

Valid Licenses
Microsoft publishes a valid list of Sku’s here: link (pdf version is here: link).  
This program uses the values from the String ID field.

Notes
•	Use email address or sign in name for user
•	Use can use a comma separated list of licenses
•	You can use the keyword ‘<all>’ in LicensesToRemove 
•	If there is a License in both ToAdd and ToRemove, adding wins.
•	User Licenses will first be checked to see if they already comply with request.
•	An Invalid SKU will pause the code.  In this case a valid list of SKUs will be displayed so the entry can be fixed.
•	Licenses must be available, otherwise the entry will be skipped.
•	The script is designed so that it does nothing if nothing needs to be done.  It can be run repeatedly safely.
