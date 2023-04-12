#Requires -Version 5.0
using namespace System.Management.Automation

# EPMLCDParser v1.0, April 2023
# https://github.com/dkeirCA/EPMLCDParser

<#
    .SYNOPSIS
        Writes messages to the information stream, optionally with
        color when written to the host.
    .DESCRIPTION
        An alternative to Write-Host which will write to the information stream
        and the host (optionally in colors specified) but will honor the
        $InformationPreference of the calling context.
        In PowerShell 5.0+ Write-Host calls through to Write-Information but
        will _always_ treat $InformationPreference as 'Continue', so the caller
        cannot use other options to the preference variable as intended.
#>
Function Write-InformationColored {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [Object]$MessageData,
        [ConsoleColor]$ForegroundColor = $Host.UI.RawUI.ForegroundColor, # Make sure we use the current colours by default
        [ConsoleColor]$BackgroundColor = $Host.UI.RawUI.BackgroundColor,
        [Switch]$NoNewline
    )

    $msg = [HostInformationMessage]@{
        Message         = $MessageData
        ForegroundColor = $ForegroundColor
        BackgroundColor = $BackgroundColor
        NoNewline       = $NoNewline.IsPresent
    }

    Write-Information $msg
}

$InformationPreference = 'Continue'
# Prompt the user for their domain and store it as a variable called DOMAIN
Write-InformationColored -MessageData "======================================" -ForegroundColor Green
Write-Information -MessageData "Make sure you are happy with the intended output Safe and platform ID's. If not, either edit in the output or in this script."
Write-Information -MessageData ""
Write-Information -MessageData "We will be building the FQDN of the machine using the domain provided. Be aware that in some scenarios, the FQDN you provide may not be correct. The EPM report does not have this information."
Write-InformationColored -MessageData "======================================" -ForegroundColor Green
$DOMAIN = Read-Host "Please enter your domain"

# Import the XLSX file specified as an argument
$inputFile = $args[0]
$data = Import-Excel -Path $inputFile -WorkSheetname "Sheet 1" -StartRow 2
# Create the output CSV file with the name "accounts-upload.csv"
$outputFile = "accounts-upload.csv"


# Create an array to hold the output data
$outputData = @()

# Create a hash table to count instances of domains found
$domains=@{}

# Loop through the input data and transform it into the output format
foreach ($row in $data) {
	$computerName, $userName = $row."User/Group".Split('\')
	#We only want users, not groups. So ignore groups. However there is an opportunity to report on found domains to alert the user that they may be uploading incorrect FQDNs
	if ($row."Type" -eq "DomainGroup"){
		if ($domains.$computerName) {
		
			$domains.$computerName = $domains.$computerName + 1
		}
		else {
			#Create a key value pair for the first instance
			
			$domains.Add($computerName,1)
		}
	} 
	else {
		#If type is user, then we want to parse for adding to our CSV
		if ($row."Group" -eq "Administrators"){
			$platformID = "WinLooselyDevice"
			$address = "$computerName.$DOMAIN"
		}
		else {
			$platformID = "MACLooselyDevice"
			$address = "$computerName.$DOMAIN"
			$computerName = ""
		}
		
		$safeName = "WorkstationLA"
		$secret = ""
		
		
		$outputRow = [ordered]@{
			"userName" = $userName
			"address" = $address
			"safeName" = $safeName
			"platformID" = $platformID
			"secret" = $secret
			"automaticManagementEnabled" = "FALSE"
			#Note: If automaticManagementEnabled is set to TRUE, then manualManagementReason must be blank
			"manualManagementReason" = "EPM LCD Import"
			"groupName" = ""
			"logonDomain" = $computerName
		
		}
		$outputData += New-Object psobject -Property $outputRow
    }
    
    
}
#Output some potentially interesting information
if ($domains.count -gt 0) {
	Write-InformationColored -MessageData "======================================" -ForegroundColor Green
	Write-Information -MessageData "The following domain-related info was captured, please analyze and take appropriate actions to correct the CSV or re-run the script with a sanitized input file."
	Write-Information -MessageData "This information was obtained based on the Domain\User NETBIOS format in column two of the report. The number is the number of instances the domain was found,"
	Write-Information -MessageData "and is not representative of the number of machines, but provided as a proportional measure of different domain-joined machines in this dataset."
	#Show message if orphaned SID is found 
	foreach ($key in $domains.Keys) {
		if ($key.StartsWith("S-1")) {
			Write-InformationColored -MessageData "Note: Orphaned SIDs should be investigated." -ForegroundColor Green
			break
		}
	}
	Write-Information -MessageData ""
	#Write-Information -MessageData $domains
	Write-Output $domains
}


# Export the output data to the CSV file
$outputData | Export-Csv $outputFile -NoTypeInformation -Encoding UTF8 -Delimiter ","
Write-Information ""
(Get-Content $outputFile).replace('"', '') | Set-Content $outputFile
Write-Information "Done!"

