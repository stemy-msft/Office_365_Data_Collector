#############################################################################
#                    Core_Assemble_Azure_Excel.ps1		 					#
#                                     			 							#
#                               4.0.2    		 							#
#                                     			 							#
#   This Sample Code is provided for the purpose of illustration only       #
#   and is not intended to be used in a production environment.  THIS       #
#   SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT    #
#   WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT    #
#   LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS     #
#   FOR A PARTICULAR PURPOSE.  We grant You a nonexclusive, royalty-free    #
#   right to use and modify the Sample Code and to reproduce and distribute #
#   the object code form of the Sample Code, provided that You agree:       #
#   (i) to not use Our name, logo, or trademarks to market Your software    #
#   product in which the Sample Code is embedded; (ii) to include a valid   #
#   copyright notice on Your software product in which the Sample Code is   #
#   embedded; and (iii) to indemnify, hold harmless, and defend Us and      #
#   Our suppliers from and against any claims or lawsuits, including        #
#   attorneys' fees, that arise or result from the use or distribution      #
#   of the Sample Code.                                                     #
#                                     			 							#
#############################################################################
Param($RunLocation)

$ErrorActionPreference = "Stop"
Trap {
$ErrorText = "Core_Assemble_Azure_Excel " + "`n" + $server + "`n"
$ErrorText += $_

$ErrorLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$ErrorLog.MachineName = "."
$ErrorLog.Source = "O365DC"
#$ErrorLog.WriteEntry($ErrorText,"Error", 100)
}

# Increase this value if adding new sheets
$SheetsInNewWorkbook = 19
function Convert-Datafile{
    param ([int]$NumberOfColumns, `
			[array]$DataFromFile, `
			$Wsheet, `
			[int]$ExcelVersion)
		$RowCount = $DataFromFile.Count
        $ArrayRow = 0
        $BadArrayValue = @()
        $DataArray = New-Object 'object[,]' -ArgumentList $RowCount,$NumberOfColumns
		Foreach ($DataRow in $DataFromFile)
        {
            $DataField = $DataRow.Split("`t")
            for ($ArrayColumn = 0 ; $ArrayColumn -lt $NumberOfColumns ; $ArrayColumn++)
            {
                # Excel chokes if field starts with = so we'll try to prepend the ' to the string if it does
                Try{If ($DataField[$ArrayColumn].substring(0,1) -eq "=") {$DataField[$ArrayColumn] = "'"+$DataField[$ArrayColumn]}}
				Catch{}
                # Excel 2003 limit of 1823 characters
                if ($DataField[$ArrayColumn].length -lt 1823)
                    {$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]}
                # Excel 2007 limit of 8203 characters
                elseif (($ExcelVersion -ge 12) -and ($DataField[$ArrayColumn].length -lt 8203))
                    {$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]}
                # No known Excel 2010 limit
                elseif ($ExcelVersion -ge 14)
                    {$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]}
                else
                {
                    Write-Host -Object "Number of characters in array member exceeds the version of limitations of this version of Excel" -ForegroundColor Yellow
                    Write-Host -Object "-- Writing value to temp variable" -ForegroundColor Yellow
                    $DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]
                    $BadArrayValue += "$ArrayRow,$ArrayColumn"
                }
            }
            $ArrayRow++
        }

        # Replace big values in $DataArray
        $BadArrayValue_count = $BadArrayValue.count
        $BadArrayValue_Temp = @()
        for ($i = 0 ; $i -lt $BadArrayValue_count ; $i++)
        {
            $BadArray_Split = $badarrayvalue[$i].Split(",")
            $BadArrayValue_Temp += $DataArray[$BadArray_Split[0],$BadArray_Split[1]]
            $DataArray[$BadArray_Split[0],$BadArray_Split[1]] = "**TEMP**"
            Write-Host -Object "-- Replacing long value with **TEMP**" -ForegroundColor Yellow
        }

        $EndCellRow = ($RowCount+1)
        $Data_range = $Wsheet.Range("a2","$EndCellColumn$EndCellRow")
        $Data_range.Value2 = $DataArray

        # Paste big values back into the spreadsheet
        for ($i = 0 ; $i -lt $BadArrayValue_count ; $i++)
        {
            $BadArray_Split = $badarrayvalue[$i].Split(",")
            # Adjust for header and $i=0
            $CellRow = [int]$BadArray_Split[0] + 2
            # Adjust for $i=0
            $CellColumn = [int]$BadArray_Split[1] + 1

            $Range = $Wsheet.cells.item($CellRow,$CellColumn)
            $Range.Value2 = $BadArrayValue_Temp[$i]
            Write-Host -Object "-- Pasting long value back in spreadsheet" -ForegroundColor Yellow
        }
    }

function Get-ColumnLetter{
	param([int]$HeaderCount)

	If ($headercount -ge 27)
	{
		$i = [int][math]::Floor($Headercount/26)
		$j = [int]($Headercount -($i*26))
		# This doesn't work on factors of 26
		# 52 become "b@" instead of "az"
		if ($j -eq 0)
		{
			$i--
			$j=26
		}
		$i_char = [char]($i+64)
		$j_char = [char]($j+64)
	}
	else
	{
		$j_char = [char]($headercount+64)
	}
	return [string]$i_char+[string]$j_char
}

set-location -LiteralPath $RunLocation

$EventLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$EventLog.MachineName = "."
$EventLog.Source = "O365DC"
#$EventLog.WriteEntry("Starting Core_Assemble_Azure_Excel","Information", 42)

Write-Host -Object "---- Starting to create com object for Excel"
$Excel_Azure = New-Object -ComObject excel.application
Write-Host -Object "---- Hiding Excel"
$Excel_Azure.visible = $false
Write-Host -Object "---- Setting ShowStartupDialog to false"
$Excel_Azure.ShowStartupDialog = $false
Write-Host -Object "---- Setting DefaultFilePath"
$Excel_Azure.DefaultFilePath = $RunLocation + "\output"
Write-Host -Object "---- Setting SheetsInNewWorkbook"
$Excel_Azure.SheetsInNewWorkbook = $SheetsInNewWorkbook
Write-Host -Object "---- Checking Excel version"
$Excel_Version = $Excel_Azure.version
if ($Excel_version -ge 12)
{
	$Excel_Azure.DefaultSaveFormat = 51
	$excel_Extension = ".xlsx"
}
else
{
	$Excel_Azure.DefaultSaveFormat = 56
	$excel_Extension = ".xls"
}
Write-Host -Object "---- Excel version $Excel_version and DefaultSaveFormat $Excel_extension"

# Create new Excel workbook
Write-Host -Object "---- Adding workbook"
$Excel_Azure_workbook = $Excel_Azure.workbooks.add()
Write-Host -Object "---- Setting output file"
$O365DC_Azure_XLS = $RunLocation + "\output\O365DC_Azure" + $excel_Extension

Write-Host -Object "---- Setting workbook properties"
$Excel_Azure_workbook.author = "Office 365 Data Collector v4 (O365DC v4)"
$Excel_Azure_workbook.title = "O365DC v4 - Exchange Organization"
$Excel_Azure_workbook.comments = "O365DC v4.0.2"

$intSheetCount = 1
$intColorIndex_AzureAd = 45

$intColorIndex = 0

# AzureAd
$intColorIndex = $intColorIndex_AzureAd

#Region Get-AzureAdApplication sheet
Write-Host -Object "---- Starting Get-AzureAdApplication"
	$Worksheet = $Excel_Azure_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Application"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "DisplayName"
	$header += "AllowGuestsSignIn"
	$header += "AllowPassthroughUsers"
	$header += "AppId"
	$header += "AppLogoUrl"
	#$header += "AppRoles.AllowedMemberTypes"
	#$header += "AppRoles.Description"
	#$header += "AppRoles.DisplayName"
	#$header += "AppRoles.Id"
	#$header += "AppRoles.IsEnabled"
	#$header += "AppRoles.Value"
	$header += "AvailableToOtherTenants"
	$header += "DeletionTimestamp"
	$header += "ErrorUrl"
	$header += "GroupMembershipClaims"
	$header += "Homepage"
	$header += "IdentifierUris"
	$header += "IsDeviceOnlyAuthSupported"
	$header += "IsDisabled"
	$header += "KeyCredentials"
	$header += "KnownClientApplications"
	$header += "LogoutUrl"
	$header += "Oauth2AllowImplicitFlow"
	$header += "Oauth2AllowUrlPathMatching"
	$header += "Oauth2Permissions.AdminConsentDescription"
	$header += "Oauth2Permissions.AdminConsentDisplayName"
	$header += "Oauth2Permissions.Id"
	$header += "Oauth2Permissions.IsEnabled"
	$header += "Oauth2Permissions.Type"
	$header += "Oauth2Permissions.UserConsentDescription"
	$header += "Oauth2Permissions.UserConsentDisplayName"
	$header += "Oauth2Permissions.Value"
	$header += "Oauth2RequirePostResponse"
	$header += "ObjectId"
	$header += "ObjectType"
	$header += "OptionalClaims"
	$header += "OrgRestrictions"
	$header += "ParentalControlSettings.CountriesBlockedForMinors"
	$header += "ParentalControlSettings.LegalAgeGroupRule"
	$header += "PreAuthorizedApplications"
	$header += "PublicClient"
	$header += "PublisherDomain"
	$header += "RecordConsentConditions"
	$header += "ReplyUrls"
	$header += "SamlMetadataUrl"
	$header += "SignInAudience"
	$header += "WwwHomepage"
	$HeaderCount = $header.count
	$EndCellColumn = Get-ColumnLetter $HeaderCount
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\Azure\Azure_AzureAdApplication.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Azure\Azure_AzureAdApplication.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

	#EndRegion Get-AzureAdApplication sheet

#Region Get-AzureAdContact sheet
Write-Host -Object "---- Starting Get-AzureAdContact"
$Worksheet = $Excel_Azure_workbook.worksheets.item($intSheetCount)
$Worksheet.name = "Contact"
$Worksheet.Tab.ColorIndex = $intColorIndex
$row = 1
$header = @()
$header += "DisplayName"
$header += "City"
$header += "CompanyName"
$header += "Country"
$header += "DeletionTimestamp"
$header += "Department"
$header += "DirSyncEnabled"
$header += "FacsimileTelephoneNumber"
$header += "GivenName"
$header += "JobTitle"
$header += "LastDirSyncTime"
$header += "Mail"
$header += "MailNickName"
$header += "Mobile"
$header += "ObjectId"
$header += "ObjectType"
$header += "PhysicalDeliveryOfficeName"
$header += "PostalCode"
$header += "ProvisioningErrors"
$header += "ProxyAddresses"
$header += "SipProxyAddress"
$header += "State"
$header += "StreetAddress"
$header += "Surname"
$header += "TelephoneNumber"
$HeaderCount = $header.count
$EndCellColumn = Get-ColumnLetter $HeaderCount
$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
$Header_range.value2 = $header
$Header_range.cells.interior.colorindex = 45
$Header_range.cells.font.colorindex = 0
$Header_range.cells.font.bold = $true
$row++
$intSheetCount++
$ColumnCount = $header.Count
$DataFile = @()
$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\Azure\Azure_AzureAdContact.txt") -eq $true)
{
$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Azure\Azure_AzureAdContact.txt")
# Send the data to the function to process and add to the Excel worksheet
Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-AzureAdContact sheet

#Region Get-AzureAdDevice sheet
Write-Host -Object "---- Starting Get-AzureAdDevice"
$Worksheet = $Excel_Azure_workbook.worksheets.item($intSheetCount)
$Worksheet.name = "Device"
$Worksheet.Tab.ColorIndex = $intColorIndex
$row = 1
$header = @()
$header += "DisplayName"
$header += "AccountEnabled"
$header += "ApproximateLastLogonTimeStamp"
$header += "ComplianceExpiryTime"
$header += "DeletionTimestamp"
$header += "DeviceId"
$header += "DeviceMetadata"
$header += "DeviceObjectVersion"
$header += "DeviceOSType"
$header += "DeviceOSVersion"
$header += "DeviceTrustType"
$header += "DirSyncEnabled"
$header += "IsCompliant"
$header += "IsManaged"
$header += "LastDirSyncTime"
$header += "ObjectId"
$header += "ObjectType"
$header += "ProfileType"
$HeaderCount = $header.count
$EndCellColumn = Get-ColumnLetter $HeaderCount
$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
$Header_range.value2 = $header
$Header_range.cells.interior.colorindex = 45
$Header_range.cells.font.colorindex = 0
$Header_range.cells.font.bold = $true
$row++
$intSheetCount++
$ColumnCount = $header.Count
$DataFile = @()
$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\Azure\Azure_AzureAdDevice.txt") -eq $true)
{
$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Azure\Azure_AzureAdDevice.txt")
# Send the data to the function to process and add to the Excel worksheet
Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-AzureAdDevice sheet

#Region Get-AzureAdDeviceRegisteredOwner sheet
Write-Host -Object "---- Starting Get-AzureAdDeviceRegisteredOwner"
$Worksheet = $Excel_Azure_workbook.worksheets.item($intSheetCount)
$Worksheet.name = "DeviceRegisteredOwner"
$Worksheet.Tab.ColorIndex = $intColorIndex
$row = 1
$header = @()
$header += "Device DisplayName"
$header += "DisplayName"
$header += "AccountEnabled"
$header += "AgeGroup"
$header += "AssignedLicenses"
$header += "AssignedPlans"
$header += "City"
$header += "CompanyName"
$header += "ConsentProvidedForMinor"
$header += "Country"
$header += "CreationType"
$header += "DeletionTimestamp"
$header += "Department"
$header += "DirSyncEnabled"
$header += "ExtensionProperty.createdDateTime"
$header += "ExtensionProperty.employeeId"
$header += "ExtensionProperty.odata.type"
$header += "ExtensionProperty.onPremisesDistinguishedName"
$header += "ExtensionProperty.userIdentities"
$header += "FacsimileTelephoneNumber"
$header += "GivenName"
$header += "ImmutableId"
$header += "IsCompromised"
$header += "JobTitle"
$header += "LastDirSyncTime"
$header += "LegalAgeGroupClassification"
$header += "Mail"
$header += "MailNickName"
$header += "Mobile"
$header += "ObjectId"
$header += "ObjectType"
$header += "OnPremisesSecurityIdentifier"
$header += "OtherMails"
$header += "PasswordPolicies"
$header += "PasswordProfile"
$header += "PhysicalDeliveryOfficeName"
$header += "PostalCode"
$header += "PreferredLanguage"
$header += "ProvisionedPlans"
$header += "ProvisioningErrors"
$header += "ProxyAddresses"
$header += "RefreshTokensValidFromDateTime"
$header += "ShowInAddressList"
$header += "SignInNames"
$header += "SipProxyAddress"
$header += "State"
$header += "StreetAddress"
$header += "Surname"
$header += "TelephoneNumber"
$header += "UsageLocation"
$header += "UserPrincipalName"
$header += "UserState"
$header += "UserStateChangedOn"
$header += "UserType"
$HeaderCount = $header.count
$EndCellColumn = Get-ColumnLetter $HeaderCount
$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
$Header_range.value2 = $header
$Header_range.cells.interior.colorindex = 45
$Header_range.cells.font.colorindex = 0
$Header_range.cells.font.bold = $true
$row++
$intSheetCount++
$ColumnCount = $header.Count
$DataFile = @()
$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\Azure\Azure_AzureAdDeviceRegisteredOwner.txt") -eq $true)
{
$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Azure\Azure_AzureAdDeviceRegisteredOwner.txt")
# Send the data to the function to process and add to the Excel worksheet
Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-AzureAdDeviceRegisteredOwner sheet

#Region Get-AzureAdDeviceRegisteredUser sheet
Write-Host -Object "---- Starting Get-AzureAdDeviceRegisteredUser"
$Worksheet = $Excel_Azure_workbook.worksheets.item($intSheetCount)
$Worksheet.name = "DeviceRegisteredUser"
$Worksheet.Tab.ColorIndex = $intColorIndex
$row = 1
$header = @()
$header += "Device DisplayName"
$header += "DisplayName"
$header += "AccountEnabled"
$header += "AgeGroup"
$header += "AssignedLicenses"
$header += "AssignedPlans"
$header += "City"
$header += "CompanyName"
$header += "ConsentProvidedForMinor"
$header += "Country"
$header += "CreationType"
$header += "DeletionTimestamp"
$header += "Department"
$header += "DirSyncEnabled"
$header += "ExtensionProperty.createdDateTime"
$header += "ExtensionProperty.employeeId"
$header += "ExtensionProperty.odata.type"
$header += "ExtensionProperty.onPremisesDistinguishedName"
$header += "ExtensionProperty.userIdentities"
$header += "FacsimileTelephoneNumber"
$header += "GivenName"
$header += "ImmutableId"
$header += "IsCompromised"
$header += "JobTitle"
$header += "LastDirSyncTime"
$header += "LegalAgeGroupClassification"
$header += "Mail"
$header += "MailNickName"
$header += "Mobile"
$header += "ObjectId"
$header += "ObjectType"
$header += "OnPremisesSecurityIdentifier"
$header += "OtherMails"
$header += "PasswordPolicies"
$header += "PasswordProfile"
$header += "PhysicalDeliveryOfficeName"
$header += "PostalCode"
$header += "PreferredLanguage"
$header += "ProvisionedPlans"
$header += "ProvisioningErrors"
$header += "ProxyAddresses"
$header += "RefreshTokensValidFromDateTime"
$header += "ShowInAddressList"
$header += "SignInNames"
$header += "SipProxyAddress"
$header += "State"
$header += "StreetAddress"
$header += "Surname"
$header += "TelephoneNumber"
$header += "UsageLocation"
$header += "UserPrincipalName"
$header += "UserState"
$header += "UserStateChangedOn"
$header += "UserType"
$HeaderCount = $header.count
$EndCellColumn = Get-ColumnLetter $HeaderCount
$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
$Header_range.value2 = $header
$Header_range.cells.interior.colorindex = 45
$Header_range.cells.font.colorindex = 0
$Header_range.cells.font.bold = $true
$row++
$intSheetCount++
$ColumnCount = $header.Count
$DataFile = @()
$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\Azure\Azure_AzureAdDeviceRegisteredUser.txt") -eq $true)
{
$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Azure\Azure_AzureAdDeviceRegisteredUser.txt")
# Send the data to the function to process and add to the Excel worksheet
Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-AzureAdDeviceRegisteredUser sheet

#Region Get-AzureAdDirectoryRole sheet
Write-Host -Object "---- Starting Get-AzureAdDirectoryRole"
$Worksheet = $Excel_Azure_workbook.worksheets.item($intSheetCount)
$Worksheet.name = "DirectoryRole"
$Worksheet.Tab.ColorIndex = $intColorIndex
$row = 1
$header = @()
$header += "DisplayName"
$header += "DeletionTimestamp"
$header += "Description"
$header += "IsSystem"
$header += "ObjectId"
$header += "ObjectType"
$header += "RoleDisabled"
$header += "RoleTemplateId"
$HeaderCount = $header.count
$EndCellColumn = Get-ColumnLetter $HeaderCount
$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
$Header_range.value2 = $header
$Header_range.cells.interior.colorindex = 45
$Header_range.cells.font.colorindex = 0
$Header_range.cells.font.bold = $true
$row++
$intSheetCount++
$ColumnCount = $header.Count
$DataFile = @()
$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\Azure\Azure_AzureAdDirectoryRole.txt") -eq $true)
{
$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Azure\Azure_AzureAdDirectoryRole.txt")
# Send the data to the function to process and add to the Excel worksheet
Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-AzureAdDirectoryRole sheet

#Region Get-AzureAdDomain sheet
Write-Host -Object "---- Starting Get-AzureAdDomain"
$Worksheet = $Excel_Azure_workbook.worksheets.item($intSheetCount)
$Worksheet.name = "Domain"
$Worksheet.Tab.ColorIndex = $intColorIndex
$row = 1
$header = @()
$header += "Name"
$header += "AuthenticationType"
$header += "AvailabilityStatus"
$header += "ForceDeleteState"
$header += "IsAdminManaged"
$header += "IsDefault"
$header += "IsInitial"
$header += "IsRoot"
$header += "IsVerified"
$header += "State"
$header += "SupportedServices"
$HeaderCount = $header.count
$EndCellColumn = Get-ColumnLetter $HeaderCount
$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
$Header_range.value2 = $header
$Header_range.cells.interior.colorindex = 45
$Header_range.cells.font.colorindex = 0
$Header_range.cells.font.bold = $true
$row++
$intSheetCount++
$ColumnCount = $header.Count
$DataFile = @()
$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\Azure\Azure_AzureAdDomain.txt") -eq $true)
{
$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Azure\Azure_AzureAdDomain.txt")
# Send the data to the function to process and add to the Excel worksheet
Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-AzureAdDomain sheet

#Region Get-AzureAdDomainServiceConfigurationRecord sheet
Write-Host -Object "---- Starting Get-AzureAdDomainServiceConfigurationRecord"
$Worksheet = $Excel_Azure_workbook.worksheets.item($intSheetCount)
$Worksheet.name = "DomainServiceConfigRecord"
$Worksheet.Tab.ColorIndex = $intColorIndex
$row = 1
$header = @()
$header += "Domain Name"
$header += "DnsRecordId"
$header += "IsOptional"
$header += "Label"
$header += "NameTarget"
$header += "Port"
$header += "Priority"
$header += "Protocol"
$header += "RecordType"
$header += "Service"
$header += "SupportedService"
$header += "Ttl"
$header += "Weight"
$HeaderCount = $header.count
$EndCellColumn = Get-ColumnLetter $HeaderCount
$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
$Header_range.value2 = $header
$Header_range.cells.interior.colorindex = 45
$Header_range.cells.font.colorindex = 0
$Header_range.cells.font.bold = $true
$row++
$intSheetCount++
$ColumnCount = $header.Count
$DataFile = @()
$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\Azure\Azure_AzureAdDomainServiceConfigurationRecord.txt") -eq $true)
{
$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Azure\Azure_AzureAdDomainServiceConfigurationRecord.txt")
# Send the data to the function to process and add to the Excel worksheet
Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-AzureAdDomainServiceConfigurationRecord sheet

#Region Get-AzureAdDomainVerificationDnsRecord sheet
Write-Host -Object "---- Starting Get-AzureAdDomainVerificationDnsRecord"
$Worksheet = $Excel_Azure_workbook.worksheets.item($intSheetCount)
$Worksheet.name = "DomainVerificationDnsRecord"
$Worksheet.Tab.ColorIndex = $intColorIndex
$row = 1
$header = @()
$header += "Domain Name"
$header += "DnsRecordId"
$header += "IsOptional"
$header += "Label"
$header += "RecordType"
$header += "SupportedService"
$header += "Text"
$header += "Ttl"
$HeaderCount = $header.count
$EndCellColumn = Get-ColumnLetter $HeaderCount
$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
$Header_range.value2 = $header
$Header_range.cells.interior.colorindex = 45
$Header_range.cells.font.colorindex = 0
$Header_range.cells.font.bold = $true
$row++
$intSheetCount++
$ColumnCount = $header.Count
$DataFile = @()
$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\Azure\Azure_AzureAdDomainVerificationDnsRecord.txt") -eq $true)
{
$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Azure\Azure_AzureAdDomainVerificationDnsRecord.txt")
# Send the data to the function to process and add to the Excel worksheet
Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-AzureAdDomainVerificationDnsRecord sheet

#Region Get-AzureAdGroup sheet
Write-Host -Object "---- Starting Get-AzureAdGroup"
$Worksheet = $Excel_Azure_workbook.worksheets.item($intSheetCount)
$Worksheet.name = "Group"
$Worksheet.Tab.ColorIndex = $intColorIndex
$row = 1
$header = @()
$header += "DisplayName"
$header += "DeletionTimestamp"
$header += "Description"
$header += "DirSyncEnabled"
$header += "LastDirSyncTime"
$header += "Mail"
$header += "MailEnabled"
$header += "MailNickName"
$header += "ObjectId"
$header += "ObjectType"
$header += "OnPremisesSecurityIdentifier"
$header += "ProvisioningErrors"
$header += "ProxyAddresses"
$header += "SecurityEnabled"
$HeaderCount = $header.count
$EndCellColumn = Get-ColumnLetter $HeaderCount
$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
$Header_range.value2 = $header
$Header_range.cells.interior.colorindex = 45
$Header_range.cells.font.colorindex = 0
$Header_range.cells.font.bold = $true
$row++
$intSheetCount++
$ColumnCount = $header.Count
$DataFile = @()
$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\Azure\Azure_AzureAdGroup.txt") -eq $true)
{
$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Azure\Azure_AzureAdGroup.txt")
# Send the data to the function to process and add to the Excel worksheet
Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-AzureAdGroup sheet

#Region Get-AzureAdGroupMember sheet
Write-Host -Object "---- Starting Get-AzureAdGroupMember"
$Worksheet = $Excel_Azure_workbook.worksheets.item($intSheetCount)
$Worksheet.name = "GroupMember"
$Worksheet.Tab.ColorIndex = $intColorIndex
$row = 1
$header = @()
$header += "Group DisplayName"
$header += "DisplayName"
$header += "AccountEnabled"
$header += "AgeGroup"
$header += "AssignedLicenses"
$header += "AssignedPlans"
$header += "Mail"
$header += "MailNickName"
$header += "ObjectType"
$header += "ProxyAddresses"
$header += "UserPrincipalName"
$header += "UserType"
$HeaderCount = $header.count
$EndCellColumn = Get-ColumnLetter $HeaderCount
$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
$Header_range.value2 = $header
$Header_range.cells.interior.colorindex = 45
$Header_range.cells.font.colorindex = 0
$Header_range.cells.font.bold = $true
$row++
$intSheetCount++
$ColumnCount = $header.Count
$DataFile = @()
$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\Azure\Azure_AzureAdGroupMember.txt") -eq $true)
{
$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Azure\Azure_AzureAdGroupMember.txt")
# Send the data to the function to process and add to the Excel worksheet
Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-AzureAdGroupMember sheet

#Region Get-AzureAdGroupOwner sheet
Write-Host -Object "---- Starting Get-AzureAdGroupOwner"
$Worksheet = $Excel_Azure_workbook.worksheets.item($intSheetCount)
$Worksheet.name = "GroupOwner"
$Worksheet.Tab.ColorIndex = $intColorIndex
$row = 1
$header = @()
$header += "Group DisplayName"
$header += "DisplayName"
$header += "AccountEnabled"
$header += "AgeGroup"
$header += "AssignedLicenses"
$header += "AssignedPlans"
$header += "Mail"
$header += "MailNickName"
$header += "ObjectType"
$header += "ProxyAddresses"
$header += "UserPrincipalName"
$header += "UserType"
$HeaderCount = $header.count
$EndCellColumn = Get-ColumnLetter $HeaderCount
$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
$Header_range.value2 = $header
$Header_range.cells.interior.colorindex = 45
$Header_range.cells.font.colorindex = 0
$Header_range.cells.font.bold = $true
$row++
$intSheetCount++
$ColumnCount = $header.Count
$DataFile = @()
$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\Azure\Azure_AzureAdGroupOwner.txt") -eq $true)
{
$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Azure\Azure_AzureAdGroupOwner.txt")
# Send the data to the function to process and add to the Excel worksheet
Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-AzureAdGroupOwner sheet

#Region Get-AzureAdSubscribedSku sheet
Write-Host -Object "---- Starting Get-AzureAdSubscribedSku"
$Worksheet = $Excel_Azure_workbook.worksheets.item($intSheetCount)
$Worksheet.name = "SubscribedSku"
$Worksheet.Tab.ColorIndex = $intColorIndex
$row = 1
$header = @()
$header += "SkuPartNumber"
$header += "AppliesTo"
$header += "CapabilityStatus"
$header += "ConsumedUnits"
$header += "PrepaidUnits.enabled"
$header += "PrepaidUnits.suspended"
$header += "PrepaidUnits.warning"
$header += "ServicePlans.ServicePlanNames"
$header += "SkuId"
$HeaderCount = $header.count
$EndCellColumn = Get-ColumnLetter $HeaderCount
$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
$Header_range.value2 = $header
$Header_range.cells.interior.colorindex = 45
$Header_range.cells.font.colorindex = 0
$Header_range.cells.font.bold = $true
$row++
$intSheetCount++
$ColumnCount = $header.Count
$DataFile = @()
$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\Azure\Azure_AzureAdSubscribedSku.txt") -eq $true)
{
$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Azure\Azure_AzureAdSubscribedSku.txt")
# Send the data to the function to process and add to the Excel worksheet
Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-AzureAdSubscribedSku sheet

#Region Get-AzureAdTenantDetail sheet
Write-Host -Object "---- Starting Get-AzureAdTenantDetail"
$Worksheet = $Excel_Azure_workbook.worksheets.item($intSheetCount)
$Worksheet.name = "TenantDetail"
$Worksheet.Tab.ColorIndex = $intColorIndex
$row = 1
$header = @()
$header += "DisplayName"
$header += "VerifiedDomains.Name"
$header += "AssignedPlans"
$header += "City"
$header += "CompanyLastDirSyncTime"
$header += "Country"
$header += "CountryLetterCode"
$header += "DirSyncEnabled"
$header += "MarketingNotificationEmails"
$header += "ObjectType"
$header += "PostalCode"
$header += "PreferredLanguage"
$header += "PrivacyProfile"
$header += "ProvisionedPlans"
$header += "ProvisioningErrors"
$header += "SecurityComplianceNotificationMails"
$header += "SecurityComplianceNotificationPhones"
$header += "State"
$header += "Street"
$header += "TechnicalNotificationMails"
$header += "TelephoneNumber"
$HeaderCount = $header.count
$EndCellColumn = Get-ColumnLetter $HeaderCount
$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
$Header_range.value2 = $header
$Header_range.cells.interior.colorindex = 45
$Header_range.cells.font.colorindex = 0
$Header_range.cells.font.bold = $true
$row++
$intSheetCount++
$ColumnCount = $header.Count
$DataFile = @()
$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\Azure\Azure_AzureAdTenantDetail.txt") -eq $true)
{
$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Azure\Azure_AzureAdTenantDetail.txt")
# Send the data to the function to process and add to the Excel worksheet
Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-AzureAdTenantDetail sheet

#Region Get-AzureAdUser sheet
Write-Host -Object "---- Starting Get-AzureAdUser"
$Worksheet = $Excel_Azure_workbook.worksheets.item($intSheetCount)
$Worksheet.name = "User"
$Worksheet.Tab.ColorIndex = $intColorIndex
$row = 1
$header = @()
$header += "UserPrincipalName"
$header += "AccountEnabled"
$header += "AgeGroup"
$header += "AssignedLicenses"
$header += "AssignedPlans"
$header += "City"
$header += "CompanyName"
$header += "ConsentProvidedForMinor"
$header += "Country"
$header += "CreationType"
$header += "CreationType"
$header += "DeletionTimestamp"
$header += "DirSyncEnabled"
$header += "DisplayName"
$header += "ExtensionProperty.createdDateTime"
$header += "ExtensionProperty.employeeId"
$header += "ExtensionProperty.odata.type"
$header += "ExtensionProperty.onPremisesDistinguishedName"
$header += "ExtensionProperty.userIdentities"
$header += "FacsimileTelephoneNumber"
$header += "GivenName"
$header += "ImmutableId"
$header += "IsCompromised"
$header += "JobTitle"
$header += "LastDirSyncTime"
$header += "LegalAgeGroupClassification"
$header += "Mail"
$header += "MailNickName"
$header += "Mobile"
$header += "ObjectType"
$header += "OnPremisesSecurityIdentifier"
$header += "OtherMails"
$header += "PasswordPolicies"
$header += "PasswordProfile.ForceChangePasswordNextLogin"
$header += "PasswordProfile.EnforceChangePasswordPolicy"
$header += "PhysicalDeliveryOfficeName"
$header += "PostalCode"
$header += "PreferredLanguage"
$header += "ProvisionedPlans"
$header += "ProvisioningErrors"
$header += "ProxyAddresses"
$header += "RefreshTokensValidFromDateTime"
$header += "ShowInAddressList"
$header += "SignInNames"
$header += "SipProxyAddress"
$header += "State"
$header += "StreetAddress"
$header += "Surname"
$header += "TelephoneNumber"
$header += "UsageLocation"
$header += "UserState"
$header += "UserStateChangedOn"
$header += "UserType"
$HeaderCount = $header.count
$EndCellColumn = Get-ColumnLetter $HeaderCount
$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
$Header_range.value2 = $header
$Header_range.cells.interior.colorindex = 45
$Header_range.cells.font.colorindex = 0
$Header_range.cells.font.bold = $true
$row++
$intSheetCount++
$ColumnCount = $header.Count
$DataFile = @()
$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\Azure\Azure_AzureAdUser.txt") -eq $true)
{
$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Azure\Azure_AzureAdUser.txt")
# Send the data to the function to process and add to the Excel worksheet
Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-AzureAdUser sheet


#Region Get-AzureAdUserLicenseDetail sheet
Write-Host -Object "---- Starting Get-AzureAdUserLicenseDetail"
$Worksheet = $Excel_Azure_workbook.worksheets.item($intSheetCount)
$Worksheet.name = "UserLicenseDetail"
$Worksheet.Tab.ColorIndex = $intColorIndex
$row = 1
$header = @()
$header += "User DisplayName"
$header += "SkuPartNumber"
$header += "SkuId"
$header += "ServicePlan.AppliesTo"
$header += "ServicePlan.ProvisioningStatus"
$header += "ServicePlan.ServicePlanId"
$header += "ServicePlan.ServicePlanName"
$header += "UserType"
$HeaderCount = $header.count
$EndCellColumn = Get-ColumnLetter $HeaderCount
$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
$Header_range.value2 = $header
$Header_range.cells.interior.colorindex = 45
$Header_range.cells.font.colorindex = 0
$Header_range.cells.font.bold = $true
$row++
$intSheetCount++
$ColumnCount = $header.Count
$DataFile = @()
$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\Azure\Azure_AzureAdUserLicenseDetail.txt") -eq $true)
{
$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Azure\Azure_AzureAdUserLicenseDetail.txt")
# Send the data to the function to process and add to the Excel worksheet
Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-AzureAdUserLicenseDetail sheet


#Region Get-AzureAdUserMembership sheet
Write-Host -Object "---- Starting Get-AzureAdUserMembership"
	$Worksheet = $Excel_Azure_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "UserMembership"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "User DisplayName"
	$header += "DisplayName"
	$header += "ObjectType"
	$header += "Description"
	$header += "DirSyncEnabled"
	$header += "LastDirSyncTime"
	$header += "Mail"
	$header += "MailEnabled"
	$header += "MailNickName"
	$header += "OnPremisesSecurityIdentifier"
	$header += "ProxyAddresses"
	$header += "SecurityEnabled"
	$HeaderCount = $header.count
	$EndCellColumn = Get-ColumnLetter $HeaderCount
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\Azure\Azure_AzureAdUserMembership.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Azure\Azure_AzureAdUserMembership.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

	#EndRegion Get-AzureAdUserMembership sheet


#Region Get-AzureAdUserOwnedDevice sheet
Write-Host -Object "---- Starting Get-AzureAdUserOwnedDevice"
	$Worksheet = $Excel_Azure_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "UserOwnedDevice"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "User DisplayName"
	$header += "DisplayName"
	$header += "AccountEnabled"
	$header += "ApproximateLastLogonTimeStamp"
	$header += "ComplianceExpiryTime"
	$header += "DeviceId"
	$header += "DeviceMetadata"
	$header += "DeviceObjectVersion"
	$header += "DeviceOSType"
	$header += "DeviceOSVersion"
	$header += "DeviceTrustType"
	$header += "DirSyncEnabled"
	$header += "IsCompliant"
	$header += "IsManaged"
	$header += "LastDirSyncTime"
	$header += "ObjectType"
	$header += "ProfileType"
	$header += "SystemLabels"
	$HeaderCount = $header.count
	$EndCellColumn = Get-ColumnLetter $HeaderCount
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\Azure\Azure_AzureAdUserOwnedDevice.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Azure\Azure_AzureAdUserOwnedDevice.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

	#EndRegion Get-AzureAdUserOwnedDevice sheet


#Region Get-AzureAdUserRegisteredDevice sheet
Write-Host -Object "---- Starting Get-AzureAdUserRegisteredDevice"
	$Worksheet = $Excel_Azure_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "UserRegisteredDevice"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "User DisplayName"
	$header += "DisplayName"
	$header += "AccountEnabled"
	$header += "ApproximateLastLogonTimeStamp"
	$header += "ComplianceExpiryTime"
	$header += "DeviceId"
	$header += "DeviceMetadata"
	$header += "DeviceObjectVersion"
	$header += "DeviceOSType"
	$header += "DeviceOSVersion"
	$header += "DeviceTrustType"
	$header += "DirSyncEnabled"
	$header += "IsCompliant"
	$header += "IsManaged"
	$header += "LastDirSyncTime"
	$header += "ObjectType"
	$header += "ProfileType"
	$header += "SystemLabels"
	$HeaderCount = $header.count
	$EndCellColumn = Get-ColumnLetter $HeaderCount
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\Azure\Azure_AzureAdUserRegisteredDevice.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Azure\Azure_AzureAdUserRegisteredDevice.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

	#EndRegion Get-AzureAdUserRegisteredDevice sheet




# Autofit columns
Write-Host -Object "---- Starting Autofit"
$Excel_AzureWorksheetCount = $Excel_Azure_workbook.worksheets.count
$AutofitSheetCount = 1
while ($AutofitSheetCount -le $Excel_AzureWorksheetCount)
{
	$ActiveWorksheet = $Excel_Azure_workbook.worksheets.item($AutofitSheetCount)
	$objRange = $ActiveWorksheet.usedrange
	[Void]	$objRange.entirecolumn.autofit()
	$AutofitSheetCount++
}
$Excel_Azure_workbook.saveas($O365DC_Azure_XLS)
Write-Host -Object "---- Spreadsheet saved"
$Excel_Azure.workbooks.close()
Write-Host -Object "---- Workbook closed"
$Excel_Azure.quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel_Azure)
Remove-Variable -Name Excel_Azure
# If the ReleaseComObject doesn't do it..
#spps -n excel

	$EventLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
	$EventLog.MachineName = "."
	$EventLog.Source = "O365DC"
	try{$EventLog.WriteEntry("Ending Core_Assemble_Azure_Excel","Information", 43)}catch{}

