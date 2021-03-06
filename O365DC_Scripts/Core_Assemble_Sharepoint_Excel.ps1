#############################################################################
#                    Core_Assemble_Spo_Excel.ps1		 					#
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
$ErrorText = "Core_Assemble_Spo_Excel " + "`n" + $server + "`n"
$ErrorText += $_

$ErrorLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$ErrorLog.MachineName = "."
$ErrorLog.Source = "O365DC"
#$ErrorLog.WriteEntry($ErrorText,"Error", 100)
}

# Increase this value if adding new sheets
$SheetsInNewWorkbook = 8
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
#$EventLog.WriteEntry("Starting Core_Assemble_Spo_Excel","Information", 42)

Write-Host -Object "---- Starting to create com object for Excel"
$Excel_Spo = New-Object -ComObject excel.application
Write-Host -Object "---- Hiding Excel"
$Excel_Spo.visible = $false
Write-Host -Object "---- Setting ShowStartupDialog to false"
$Excel_Spo.ShowStartupDialog = $false
Write-Host -Object "---- Setting DefaultFilePath"
$Excel_Spo.DefaultFilePath = $RunLocation + "\output"
Write-Host -Object "---- Setting SheetsInNewWorkbook"
$Excel_Spo.SheetsInNewWorkbook = $SheetsInNewWorkbook
Write-Host -Object "---- Checking Excel version"
$Excel_Version = $Excel_Spo.version
if ($Excel_version -ge 12)
{
	$Excel_Spo.DefaultSaveFormat = 51
	$excel_Extension = ".xlsx"
}
else
{
	$Excel_Spo.DefaultSaveFormat = 56
	$excel_Extension = ".xls"
}
Write-Host -Object "---- Excel version $Excel_version and DefaultSaveFormat $Excel_extension"

# Create new Excel workbook
Write-Host -Object "---- Adding workbook"
$Excel_Spo_workbook = $Excel_Spo.workbooks.add()
Write-Host -Object "---- Setting output file"
$O365DC_Spo_XLS = $RunLocation + "\output\O365DC_Sharepoint" + $excel_Extension

Write-Host -Object "---- Setting workbook properties"
$Excel_Spo_workbook.author = "Office 365 Data Collector v4 (O365DC v4)"
$Excel_Spo_workbook.title = "O365DC v4 - Exchange Organization"
$Excel_Spo_workbook.comments = "O365DC v4.0.2"

$intSheetCount = 1
$intColorIndex_SpoAd = 45

$intColorIndex = 0

# AzureAd
$intColorIndex = $intColorIndex_SpoAd

#Region Get-SpoDeletedSite sheet
Write-Host -Object "---- Starting Get-SpoDeletedSite"
	$Worksheet = $Excel_Spo_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "DeletedSite"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Url"
	$header += "DaysRemaining"
	$header += "DeletionTime"
	$header += "ResourceQuota"
	$header += "SiteId"
	$header += "Status"
	$header += "StorageQuota"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Sharepoint\Spo_SpoDeletedSite.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Sharepoint\Spo_SpoDeletedSite.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-SpoDeletedSite sheet

#Region Get-SpoExternalUser sheet
Write-Host -Object "---- Starting Get-SpoExternalUser"
	$Worksheet = $Excel_Spo_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "ExternalUser"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "NotImplemented"
	$header += "NotImplemented"
	$header += "NotImplemented"
	$header += "NotImplemented"
	$header += "NotImplemented"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Sharepoint\Spo_SpoExternalUser.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Sharepoint\Spo_SpoExternalUser.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-SpoExternalUser sheet

#Region Get-SpoGeoStorageQuota sheet
Write-Host -Object "---- Starting Get-SpoGeoStorageQuota"
	$Worksheet = $Excel_Spo_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "GeoStorageQuota"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "GeoLocation"
	$header += "GeoAllocatedStorageMB"
	$header += "GeoAvailableStorageMB"
	$header += "GeoUsedStorageMB"
	$header += "QuotaType"
	$header += "TenantStorageMB"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Sharepoint\Spo_SpoGeoStorageQuota.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Sharepoint\Spo_SpoGeoStorageQuota.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-SpoGeoStorageQuota sheet

#Region Get-SpoSite sheet
Write-Host -Object "---- Starting Get-SpoSite"
	$Worksheet = $Excel_Spo_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Site"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Title"
	$header += "Url"
	$header += "AllowDownloadingNonWebViewableFiles"
	$header += "AllowEditing"
	$header += "AllowSelfServiceUpgrade"
	$header += "CommentsOnSitePagesDisabled"
	$header += "CompatibilityLevel"
	$header += "ConditionalAccessPolicy"
	$header += "DefaultLinkPermission"
	$header += "DefaultSharingLinkType"
	$header += "DenyAddAndCustomizePages"
	$header += "DisableAppViews"
	$header += "DisableCompanyWideSharingLinks"
	$header += "DisableFlows"
	$header += "DisableSharingForNonOwnersStatus"
	$header += "HubSiteId"
	$header += "IsHubSite"
	$header += "LastContentModifiedDate"
	$header += "LimitedAccessFileType"
	$header += "LocaleId"
	$header += "LockIssue"
	$header += "LockState"
	$header += "Owner"
	$header += "PWAEnabled"
	$header += "ResourceQuota"
	$header += "ResourceQuotaWarningLevel"
	$header += "ResourceUsageAverage"
	$header += "ResourceUsageCurrent"
	$header += "RestrictedToGeo"
	$header += "SandboxedCodeActivationCapability"
	$header += "SensitivityLabel"
	$header += "SharingAllowedDomainList"
	$header += "SharingBlockedDomainList"
	$header += "SharingCapability"
	$header += "SharingDomainRestrictionMode"
	$header += "ShowPeoplePickerSuggestionsForGuestUsers"
	$header += "SiteDefinedSharingCapability"
	$header += "SocialBarOnSitePagesDisabled"
	$header += "Status"
	$header += "StorageQuota"
	$header += "StorageQuotaType"
	$header += "StorageQuotaWarningLevel"
	$header += "StorageUsageCurrent"
	$header += "Template"
	$header += "WebsCount"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Sharepoint\Spo_SpoSite.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Sharepoint\Spo_SpoSite.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-SpoSite sheet

#Region Get-SpoTenant sheet
Write-Host -Object "---- Starting Get-SpoTenant"
	$Worksheet = $Excel_Spo_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Tenant"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Tenant"
	$header += "AllowDownloadingNonWebViewableFiles"
	$header += "AllowEditing"
	$header += "ApplyAppEnforcedRestrictionsToAdHocRecipients"
	$header += "BccExternalSharingInvitations"
	$header += "BccExternalSharingInvitationsList"
	$header += "CommentsOnSitePagesDisabled"
	$header += "CompatibilityRange"
	$header += "ConditionalAccessPolicy"
	$header += "DefaultLinkPermission"
	$header += "DefaultSharingLinkType"
	$header += "DisabledWebPartIds"
	$header += "DisallowInfectedFileDownload"
	$header += "DisplayStartASiteOption"
	$header += "EmailAttestationReAuthDays"
	$header += "EmailAttestationRequired"
	$header += "EnableGuestSignInAcceleration"
	$header += "EnableMinimumVersionRequirement"
	$header += "ExternalServicesEnabled"
	$header += "FileAnonymousLinkType"
	$header += "FilePickerExternalImageSearchEnabled"
	$header += "FolderAnonymousLinkType"
	$header += "IPAddressAllowList"
	$header += "IPAddressEnforcement"
	$header += "IPAddressWACTokenLifetime"
	$header += "LegacyAuthProtocolsEnabled"
	$header += "LimitedAccessFileType"
	$header += "NoAccessRedirectUrl"
	$header += "NotificationsInOneDriveForBusinessEnabled"
	$header += "NotificationsInSharePointEnabled"
	$header += "NotifyOwnersWhenInvitationsAccepted"
	$header += "NotifyOwnersWhenItemsReshared"
	$header += "ODBAccessRequests"
	$header += "ODBMembersCanShare"
	$header += "OfficeClientADALDisabled"
	$header += "OneDriveForGuestsEnabled"
	$header += "OneDriveStorageQuota"
	$header += "OrgNewsSiteUrl"
	$header += "OrphanedPersonalSitesRetentionPeriod"
	$header += "OwnerAnonymousNotification"
	$header += "PermissiveBrowserFileHandlingOverride"
	$header += "PreventExternalUsersFromResharing"
	$header += "ProvisionSharedWithEveryoneFolder"
	$header += "PublicCdnAllowedFileTypes"
	$header += "PublicCdnEnabled"
	$header += "PublicCdnOrigins"
	$header += "RequireAcceptingAccountMatchInvitedAccount"
	$header += "RequireAnonymousLinksExpireInDays"
	$header += "ResourceQuota"
	$header += "ResourceQuotaAllocated"
	$header += "SearchResolveExactEmailOrUPN"
	$header += "SharingAllowedDomainList"
	$header += "SharingBlockedDomainList"
	$header += "SharingCapability"
	$header += "SharingDomainRestrictionMode"
	$header += "ShowAllUsersClaim"
	$header += "ShowEveryoneClaim"
	$header += "ShowEveryoneExceptExternalUsersClaim"
	$header += "ShowPeoplePickerSuggestionsForGuestUsers"
	$header += "SignInAccelerationDomain"
	$header += "SocialBarOnSitePagesDisabled"
	$header += "SpecialCharactersStateInFileFolderNames"
	$header += "StartASiteFormUrl"
	$header += "StorageQuota"
	$header += "StorageQuotaAllocated"
	$header += "SyncPrivacyProfileProperties"
	$header += "UseFindPeopleInPeoplePicker"
	$header += "UsePersistentCookiesForExplorerView"
	$header += "UserVoiceForFeedbackEnabled"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Sharepoint\Spo_SpoTenant.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Sharepoint\Spo_SpoTenant.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-SpoTenant sheet

#Region Get-SpoTenantSyncClientRestriction sheet
Write-Host -Object "---- Starting Get-SpoTenantSyncClientRestriction"
	$Worksheet = $Excel_Spo_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "TenantSyncClientRestriction"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "TenantSyncClientRestriction"
	$header += "TenantRestrictionEnabled"
	$header += "AllowedDomainList"
	$header += "BlockMacSync"
	$header += "DisableReportProblemDialog"
	$header += "ExcludedFileExtensions"
	$header += "OptOutOfGrooveBlock"
	$header += "OptOutOfGrooveSoftBlock"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Sharepoint\Spo_SpoTenantSyncClientRestriction.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Sharepoint\Spo_SpoTenantSyncClientRestriction.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-SpoTenantSyncClientRestriction sheet

#Region Get-SpoUser sheet
Write-Host -Object "---- Starting Get-SpoUser"
	$Worksheet = $Excel_Spo_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "User"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Site Title"
	$header += "Site Url"
	$header += "DisplayName"
	$header += "Groups"
	$header += "IsGroup"
	$header += "IsSiteAdmin"
	$header += "LoginName"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Sharepoint\Spo_SpoUser.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Sharepoint\Spo_SpoUser.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-SpoUser sheet

#Region Get-SpoWebTemplate sheet
Write-Host -Object "---- Starting Get-SpoWebTemplate"
	$Worksheet = $Excel_Spo_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "WebTemplate"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Name"
	$header += "Title"
	$header += "CompatibilityLevel"
	$header += "Description"
	$header += "DisplayCategory"
	$header += "LocaleId"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Sharepoint\Spo_SpoWebTemplate.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Sharepoint\Spo_SpoWebTemplate.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-SpoWebTemplate sheet



# Autofit columns
Write-Host -Object "---- Starting Autofit"
$Excel_SpoWorksheetCount = $Excel_Spo_workbook.worksheets.count
$AutofitSheetCount = 1
while ($AutofitSheetCount -le $Excel_SpoWorksheetCount)
{
	$ActiveWorksheet = $Excel_Spo_workbook.worksheets.item($AutofitSheetCount)
	$objRange = $ActiveWorksheet.usedrange
	[Void]	$objRange.entirecolumn.autofit()
	$AutofitSheetCount++
}
$Excel_Spo_workbook.saveas($O365DC_Spo_XLS)
Write-Host -Object "---- Spreadsheet saved"
$Excel_Spo.workbooks.close()
Write-Host -Object "---- Workbook closed"
$Excel_Spo.quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel_Spo)
Remove-Variable -Name Excel_Spo
# If the ReleaseComObject doesn't do it..
#spps -n excel

	$EventLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
	$EventLog.MachineName = "."
	$EventLog.Source = "O365DC"
	try{$EventLog.WriteEntry("Ending Core_Assemble_Spo_Excel","Information", 43)}catch{}

