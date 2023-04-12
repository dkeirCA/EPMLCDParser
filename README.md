# EPMLCDParser
Script to parse CyberArk EPM Admin report into Vault upload format

## Introduction

This PowerShell script was created to help with the onboarding of local administrators on “loosely” connected devices, for both Windows and MacOS platforms. The script requires the EPM Administrator report as an input, which then gets parsed to create a CSV upload file for importing into your CyberArk Vault.

It has been tested on data sets with both Windows and Mac hosts, with over 6000 endpoints. As an indication, a 6000 endpoint report will take ~ 1 minute to parse.

## Requirements

- PowerShell 5 or higher
- [Import-Excel](https://www.powershellgallery.com/packages/ImportExcel)
  - Used to import and parse xlsx files

## How to use
1. Login to EPM, and browse to the Reports section of your Set
2. Choose the Local Administrative Groups Report category
3. Download the "Users in Local Administrator Group" report, specifically the Full Excel report
4. Open PowerShell and run the script with the report as a parameter
  `$ .\EPMLCDParser.ps1 .\EPMReport.xlsx`
5. Upload the resulting CSV(accounts-upload.csv) into your Vault, either via PVWA bulk upload or API 

## Notes

Adjustments can be made in either the script, or resulting CSV.

- PlatformIDs (WinLooselyDevice and MacLooselyDevice) are default, adjust as per your needs
- Safe is a placeholder, adjust as per your needs
- The report does not contain the hosts FQDN. Verify the output with reality, especially if multiple domains are present
- No need to include passwords in the CSV (preferably don't, as a best practice. EPM will manage the account once onboarded)
