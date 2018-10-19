#############################################################################
#                    Core_Assemble_Skype_Excel.ps1		 					#
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
$ErrorText = "Core_Assemble_Skype_Excel " + "`n" + $server + "`n"
$ErrorText += $_

$ErrorLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$ErrorLog.MachineName = "."
$ErrorLog.Source = "O365DC"
#$ErrorLog.WriteEntry($ErrorText,"Error", 100)
}

# Increase this value if adding new sheets
$SheetsInNewWorkbook = 67
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
#$EventLog.WriteEntry("Starting Core_Assemble_Skype_Excel","Information", 42)

Write-Host -Object "---- Starting to create com object for Excel"
$Excel_Skype = New-Object -ComObject excel.application
Write-Host -Object "---- Hiding Excel"
$Excel_Skype.visible = $false
Write-Host -Object "---- Setting ShowStartupDialog to false"
$Excel_Skype.ShowStartupDialog = $false
Write-Host -Object "---- Setting DefaultFilePath"
$Excel_Skype.DefaultFilePath = $RunLocation + "\output"
Write-Host -Object "---- Setting SheetsInNewWorkbook"
$Excel_Skype.SheetsInNewWorkbook = $SheetsInNewWorkbook
Write-Host -Object "---- Checking Excel version"
$Excel_Version = $Excel_Skype.version
if ($Excel_version -ge 12)
{
	$Excel_Skype.DefaultSaveFormat = 51
	$excel_Extension = ".xlsx"
}
else
{
	$Excel_Skype.DefaultSaveFormat = 56
	$excel_Extension = ".xls"
}
Write-Host -Object "---- Excel version $Excel_version and DefaultSaveFormat $Excel_extension"

# Create new Excel workbook
Write-Host -Object "---- Adding workbook"
$Excel_Skype_workbook = $Excel_Skype.workbooks.add()
Write-Host -Object "---- Setting output file"
$O365DC_Skype_XLS = $RunLocation + "\output\O365DC_Skype" + $excel_Extension

Write-Host -Object "---- Setting workbook properties"
$Excel_Skype_workbook.author = "Office 365 Data Collector v4 (O365DC v4)"
$Excel_Skype_workbook.title = "O365DC v4 - Exchange Organization"
$Excel_Skype_workbook.comments = "O365DC v4.0.2"

$intSheetCount = 1
$intColorIndex_CsGeneral = 45
$intColorIndex_CsOnline = 11
$intColorIndex_CsTeams = 45
$intColorIndex_CsTenant = 11

$intColorIndex = 0

#Region CsGeneral
# 26 Functions
$intColorIndex = $intColorIndex_CsGeneral

#Region Get-CsAudioConferencingProvider sheet
Write-Host -Object "---- Starting Get-CsAudioConferencingProvider"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsAudioConferencingProvider"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "Name"
	$header += "Domain"
	$header += "Port"
	$header += "Url"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsAudioConferencingProvider.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsAudioConferencingProvider.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsAudioConferencingProvider sheet

#Region Get-CsBroadcastMeetingConfiguration sheet
Write-Host -Object "---- Starting Get-CsBroadcastMeetingConfiguration"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsBroadcastMeetingConfig"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "BroadcastMeetingSupportUrl"
	$header += "EnableAnonymousBroadcastMeeting"
	$header += "EnableBroadcastMeeting"
	$header += "EnableBroadcastMeetingRecording"
	$header += "EnableOpenBroadcastMeeting"
	$header += "EnableSdnProviderForBroadcastMeeting"
	$header += "EnableTechPreviewFeatures"
	$header += "EnforceBroadcastMeetingRecording"
	$header += "SdnFallbackAttendeeThresholdCountForBroadcastMeeting"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsBroadcastMeetingConfiguration.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsBroadcastMeetingConfiguration.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsBroadcastMeetingConfiguration sheet

#Region Get-CsBroadcastMeetingPolicy sheet
Write-Host -Object "---- Starting Get-CsBroadcastMeetingPolicy"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsBroadcastMeetingPolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "AllowAnonymousBroadcastMeeting"
	$header += "AllowBroadcastMeeting"
	$header += "AllowBroadcastMeetingRecording"
	$header += "AllowOpenBroadcastMeeting"
	$header += "BroadcastMeetingRecordingEnforced"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsBroadcastMeetingPolicy.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsBroadcastMeetingPolicy.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsBroadcastMeetingPolicy sheet

#Region Get-CsCallerIdPolicy sheet
Write-Host -Object "---- Starting Get-CsCallerIdPolicy"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsCallerIdPolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "Name"
	$header += "CallerIDSubstitute"
	$header += "Description"
	$header += "EnableUserOverride"
	$header += "ServiceNumber"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsCallerIdPolicy.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsCallerIdPolicy.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsCallerIdPolicy sheet

#Region Get-CsCallingLineIdentity sheet
Write-Host -Object "---- Starting Get-CsCallingLineIdentity"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsCallingLineIdentity"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "BlockIncomingPstnCallerID"
	$header += "CallingIDSubstitute"
	$header += "Description"
	$header += "EnableUserOverride"
	$header += "ServiceNumber"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsCallingLineIdentity.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsCallingLineIdentity.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsCallingLineIdentity sheet

#Region Get-CsClientPolicy sheet
Write-Host -Object "---- Starting Get-CsClientPolicy"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsClientPolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "AddressBookAvailability"
	$header += "AttendantSafeTransfer"
	$header += "AutoDiscoveryRetryInterval"
	$header += "BlockConversationFromFederatedContacts"
	$header += "CalendarStatePublicationInterval"
	$header += "ConferenceIMIdleTimeout"
	$header += "CustomizedHelpUrl"
	$header += "CustomLinkInErrorMessages"
	$header += "CustomStateUrl"
	$header += "Description"
	$header += "DGRefreshInterval"
	$header += "DisableCalendarPresence"
	$header += "DisableContactCardOrganizationTab"
	$header += "DisableEmailComparisonCheck"
	$header += "DisableEmoticons"
	$header += "DisableFederatedPromptDisplayName"
	$header += "DisableFeedsTab"
	$header += "DisableFreeBusyInfo"
	$header += "DisableHandsetOnLockedMachine"
	$header += "DisableHtmlIm"
	$header += "DisableInkIM"
	$header += "DisableMeetingSubjectAndLocation"
	$header += "DisableOneNote12Integration"
	$header += "DisableOnlineContextualSearch"
	$header += "DisablePhonePresence"
	$header += "DisablePICPromptDisplayName"
	$header += "DisablePoorDeviceWarnings"
	$header += "DisablePoorNetworkWarnings"
	$header += "DisablePresenceNote"
	$header += "DisableRTFIM"
	$header += "DisableSavingIM"
	$header += "DisplayPhoto"
	$header += "EnableAppearOffline"
	$header += "EnableCallLogAutoArchiving"
	$header += "EnableClientAutoPopulateWithTeam"
	$header += "EnableClientMusicOnHold"
	$header += "EnableConversationWindowTabs"
	$header += "EnableEnterpriseCustomizedHelp"
	$header += "EnableEventLogging"
	$header += "EnableExchangeContactsFolder"
	$header += "EnableExchangeContactSync"
	$header += "EnableExchangeDelegateSync"
	$header += "EnableFullScreenVideo"
	$header += "EnableHighPerformanceConferencingAppSharing"
	$header += "EnableHighPerformanceP2PAppSharing"
	$header += "EnableHotdesking"
	$header += "EnableIMAutoArchiving"
	$header += "EnableMediaRedirection"
	$header += "EnableMeetingEngagement"
	$header += "EnableNotificationForNewSubscribers"
	$header += "EnableOnlineFeedback"
	$header += "EnableOnlineFeedbackScreenshots"
	$header += "EnableServerConversationHistory"
	$header += "EnableSkypeUI"
	$header += "EnableSQMData"
	$header += "EnableTracing"
	$header += "EnableUnencryptedFileTransfer"
	$header += "EnableURL"
	$header += "EnableViewBasedSubscriptionMode"
	$header += "EnableVOIPCallDefault"
	$header += "ExcludedContactFolders"
	$header += "HelpEnvironment"
	$header += "HotdeskingTimeout"
	$header += "IMLatencyErrorThreshold"
	$header += "IMLatencySpinnerDelay"
	$header += "IMWarning"
	$header += "MAPIPollInterval"
	$header += "MaximumDGsAllowedInContactList"
	$header += "MaximumNumberOfContacts"
	$header += "MaxPhotoSizeKB"
	$header += "MusicOnHoldAudioFile"
	$header += "P2PAppSharingEncryption"
	$header += "PlayAbbreviatedDialTone"
	$header += "PolicyEntry"
	$header += "PublicationBatchDelay"
	$header += "RateMyCallAllowCustomUserFeedback"
	$header += "RateMyCallDisplayPercentage"
	$header += "RequireContentPin"
	$header += "SearchPrefixFlags"
	$header += "ShowManagePrivacyRelationships"
	$header += "ShowRecentContacts"
	$header += "ShowSharepointPhotoEditLink"
	$header += "SPSearchCenterExternalURL"
	$header += "SPSearchCenterInternalURL"
	$header += "SPSearchExternalURL"
	$header += "SPSearchInternalURL"
	$header += "SupportModernFilePicker"
	$header += "TabURL"
	$header += "TelemetryTier"
	$header += "TracingLevel"
	$header += "WebServicePollInterval"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsClientPolicy.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsClientPolicy.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsClientPolicy sheet

#Region Get-CsCloudMeetingPolicy sheet
Write-Host -Object "---- Starting Get-CsCloudMeetingPolicy"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsCloudMeetingPolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "AllowAutoSchedule"
	$header += "IsModernSchedulingEnabled"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsCloudMeetingPolicy.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsCloudMeetingPolicy.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsCloudMeetingPolicy sheet

#Region Get-CsConferencingPolicy sheet
Write-Host -Object "---- Starting Get-CsConferencingPolicy"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsConferencingPolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "AllowAnnotations"
	$header += "AllowAnonymousParticipantsInMeetings"
	$header += "AllowAnonymousUsersToDialOut"
	$header += "AllowConferenceRecording"
	$header += "AllowExternalUserControl"
	$header += "AllowExternalUsersToRecordMeeting"
	$header += "AllowExternalUsersToSaveContent"
	$header += "AllowFederatedParticipantJoinAsSameEnterprise"
	$header += "AllowIPAudio"
	$header += "AllowIPVideo"
	$header += "AllowLargeMeetings"
	$header += "AllowMultiView"
	$header += "AllowNonEnterpriseVoiceUsersToDialOut"
	$header += "AllowOfficeContent"
	$header += "AllowParticipantControl"
	$header += "AllowPolls"
	$header += "AllowQandA"
	$header += "AllowSharedNotes"
	$header += "AllowUserToScheduleMeetingsWithAppSharing"
	$header += "ApplicationSharingMode"
	$header += "AppSharingBitRateKb"
	$header += "AudioBitRateKb"
	$header += "Description"
	$header += "DisablePowerPointAnnotations"
	$header += "EnableAppDesktopSharing"
	$header += "EnableDataCollaboration"
	$header += "EnableDialInConferencing"
	$header += "EnableFileTransfer"
	$header += "EnableMultiViewJoin"
	$header += "EnableOnlineMeetingPromptForLyncResources"
	$header += "EnableP2PFileTransfer"
	$header += "EnableP2PRecording"
	$header += "EnableP2PVideo"
	$header += "EnableReliableConferenceDeletion"
	$header += "FileTransferBitRateKb"
	$header += "MaxMeetingSize"
	$header += "MaxVideoConferenceResolution"
	$header += "TotalReceiveVideoBitRateKb"
	$header += "VideoBitRateKb"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsConferencingPolicy.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsConferencingPolicy.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsConferencingPolicy sheet

#Region Get-CsDialPlan sheet
Write-Host -Object "---- Starting Get-CsDialPlan"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsDialPlan"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "City"
	$header += "CountryCode"
	$header += "Description"
	$header += "DialinConferencingRegion"
	$header += "ExternalAccessPrefix"
	$header += "ITUCountryPrefix"
	$header += "NormalizationRules"
	$header += "OptimizeDeviceDialing"
	$header += "SimpleName"
	$header += "State"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsDialPlan.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsDialPlan.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsDialPlan sheet

#Region Get-CsExternalAccessPolicy sheet
Write-Host -Object "---- Starting Get-CsExternalAccessPolicy"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsExternalAccessPolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "Description"
	$header += "EnableFederationAccess"
	$header += "EnableOutsideAccess"
	$header += "EnablePublicCloudAccess"
	$header += "EnablePublicCloudAudioVideoAccess"
	$header += "EnableXmppAccess"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsExternalAccessPolicy.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsExternalAccessPolicy.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsExternalAccessPolicy sheet

#Region Get-CsExternalUserCommunicationPolicy sheet
Write-Host -Object "---- Starting Get-CsExternalUserCommunicationPolicy"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsExternalUserCommPolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "AllowPresenceVisibility"
	$header += "AllowTitleVisibility"
	$header += "EnableFileTransfer"
	$header += "EnableP2PFileTransfer"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsExternalUserCommunicationPolicy.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsExternalUserCommunicationPolicy.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsExternalUserCommunicationPolicy sheet

#Region Get-CsGraphPolicy sheet
Write-Host -Object "---- Starting Get-CsGraphPolicy"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsGraphPolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "Description"
	$header += "EnableMeetingsGraph"
	$header += "EnableSharedLinks"
	$header += "UseEWSDirectDownload"
	$header += "UseStorageService"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsGraphPolicy.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsGraphPolicy.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsGraphPolicy sheet

#Region Get-CsHostedVoicemailPolicy sheet
Write-Host -Object "---- Starting Get-CsHostedVoicemailPolicy"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsHostedVoicemailPolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "BusinessVoiceEnabled"
	$header += "Description"
	$header += "Destination"
	$header += "Organization"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsHostedVoicemailPolicy.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsHostedVoicemailPolicy.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsHostedVoicemailPolicy sheet

#Region Get-CsHostingProvider sheet
Write-Host -Object "---- Starting Get-CsHostingProvider"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsHostingProvider"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "AutodiscoverUrl"
	$header += "Enabled"
	$header += "EnabledSharedAddressSpace"
	$header += "HostsOCSUsers"
	$header += "IsLocal"
	$header += "Name"
	$header += "ProxyFqdn"
	$header += "VerificationLevel"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsHostingProvider.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsHostingProvider.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsHostingProvider sheet

#Region Get-CsImFilterConfiguration sheet
Write-Host -Object "---- Starting Get-CsImFilterConfiguration"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsImFilterConfig"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "Action"
	$header += "AllowMessage"
	$header += "BlockFileExtension"
	$header += "Enabled"
	$header += "IgnoreLocal"
	$header += "Prefixes"
	$header += "WarnMessage"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsImFilterConfiguration.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsImFilterConfiguration.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsImFilterConfiguration sheet

#Region Get-CsIPPhonePolicy sheet
Write-Host -Object "---- Starting Get-CsIPPhonePolicy"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsIPPhonePolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "BetterTogetherOverEthernetPairingMode"
	$header += "DateTimeFormat"
	$header += "EnableBetterTogetherOverEthernet"
	$header += "EnableDeviceUpdate"
	$header += "EnableExchangeCalendaring"
	$header += "EnableOneTouchVoicemail"
	$header += "EnablePowerSaveMode"
	$header += "KeyboardLockMaxPinRetry"
	$header += "LocalProvisioningServerAddress"
	$header += "LocalProvisioningServerPassword"
	$header += "LocalProvisioningServerType"
	$header += "LocalProvisioningServerUser"
	$header += "PowerSaveDuringOfficeHoursTimeoutMS"
	$header += "PowerSavePostOfficeHoursTimeoutMS"
	$header += "PrioritizedCodecsList"
	$header += "UserDialTimeoutMS"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsIPPhonePolicy.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsIPPhonePolicy.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsIPPhonePolicy sheet

#Region Get-CsMeetingConfiguration sheet
Write-Host -Object "---- Starting Get-CsMeetingConfiguration"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsMeetingConfig"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "AdmitAnonymousUsersByDefault"
	$header += "AllowCloudRecordingService"
	$header += "AllowConferenceRecording"
	$header += "AssignedConferenceTypeByDefault"
	$header += "CustomFooterText"
	$header += "DesignateAsPresenter"
	$header += "EnableAssignedConferenceType"
	$header += "EnableMeetingReport"
	$header += "HelpURL"
	$header += "LegalURL"
	$header += "LogoURL"
	$header += "PstnCallersBypassLobby"
	$header += "RequireRoomSystemsAuthorization"
	$header += "UserUriFormatForStUser"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsMeetingConfiguration.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsMeetingConfiguration.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsMeetingConfiguration sheet

#Region Get-CsMeetingMigrationStatus sheet
Write-Host -Object "---- Starting Get-CsMeetingMigrationStatus"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsMeetingMigrationStatus"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "CorrelationId"
	$header += "CreateDate"
	$header += "FailedMeetings"
	$header += "InvitesUpdated"
	$header += "LastMessage"
	$header += "ModifiedDate"
	$header += "RetryCount"
	$header += "State"
	$header += "SucceededMeetings"
	$header += "TotalMeetings"
	$header += "UserId"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsMeetingMigrationStatus.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsMeetingMigrationStatus.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsMeetingMigrationStatus sheet

#Region Get-CsMobilityPolicy sheet
Write-Host -Object "---- Starting Get-CsMobilityPolicy"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsMobilityPolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "AllowAutomaticPstnFallback"
	$header += "AllowCustomerExperienceImprovementProgram"
	$header += "AllowDeviceContactsSync"
	$header += "AllowExchangeConnectivity"
	$header += "AllowSaveCallLogs"
	$header += "AllowSaveCredentials"
	$header += "AllowSaveIMHistory"
	$header += "Description"
	$header += "EnableIPAudioVideo"
	$header += "EnableMobility"
	$header += "EnableOutsideVoice"
	$header += "EnablePushNotifications"
	$header += "EncryptAppData"
	$header += "RequireIntune"
	$header += "RequireWIFIForIPVideo"
	$header += "RequireWiFiForSharing"
	$header += "VoiceSettings"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsMobilityPolicy.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsMobilityPolicy.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsMobilityPolicy sheet

#Region Get-CsPresencePolicy sheet
Write-Host -Object "---- Starting Get-CsPresencePolicy"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsPresencePolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "Description"
	$header += "MaxCategorySubscription"
	$header += "MaxPromptedSubscriber"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsPresencePolicy.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsPresencePolicy.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsPresencePolicy sheet

#Region Get-CsPrivacyConfiguration sheet
Write-Host -Object "---- Starting Get-CsPrivacyConfiguration"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsPrivacyConfig"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "AutoInitiateContacts"
	$header += "DisplayPublishedPhotoDefault"
	$header += "EnablePrivacyMode"
	$header += "PublishLocationDataDefault"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsPrivacyConfiguration.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsPrivacyConfiguration.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsPrivacyConfiguration sheet

#Region Get-CsPushNotificationConfiguration sheet
Write-Host -Object "---- Starting Get-CsPushNotificationConfiguration"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsPushNotificationConfig"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "EnableApplePushNotificationService"
	$header += "EnableMicrosoftPushNotificationService"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsPushNotificationConfiguration.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsPushNotificationConfiguration.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsPushNotificationConfiguration sheet

#Region Get-CsUCPhoneConfiguration sheet
Write-Host -Object "---- Starting Get-CsUCPhoneConfiguration"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsUCPhoneConfig"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "CalendarPollInterval"
	$header += "EnforcePhoneLock"
	$header += "LoggingLevel"
	$header += "MinPhonePinLength"
	$header += "PhoneLockTimeout"
	$header += "SIPSecurityMode"
	$header += "Voice8021p"
	$header += "VoiceDiffServTag"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsUCPhoneConfiguration.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsUCPhoneConfiguration.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsUCPhoneConfiguration sheet

#Region Get-CsUserServicesPolicy sheet
Write-Host -Object "---- Starting Get-CsUserServicesPolicy"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsUserServicesPolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "EnableAwaySinceIndication"
	$header += "MigrationDelayInDays"
	$header += "UcsAllowed"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsUserServicesPolicy.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsUserServicesPolicy.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsUserServicesPolicy sheet

#Region Get-CsVoicePolicy sheet
Write-Host -Object "---- Starting Get-CsVoicePolicy"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsVoicePolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Name"
	$header += "Identity"
	$header += "AllowCallForwarding"
	$header += "AllowPSTNReRouting"
	$header += "AllowSimulRing"
	$header += "BusinessVoiceEnabled"
	$header += "CallForwardingSimulRingUsageType"
	$header += "CustomCallForwardingSimulRingUsages"
	$header += "Description"
	$header += "EnableBusyOptions"
	$header += "EnableBWPolicyOverride"
	$header += "EnableCallPark"
	$header += "EnableCallTransfer"
	$header += "EnableDelegation"
	$header += "EnableMaliciousCallTracing"
	$header += "EnableTeamCall"
	$header += "EnableVoicemailEscapeTimer"
	$header += "PreventPSTNTollBypass"
	$header += "PstnUsages"
	$header += "PSTNVoicemailEscapeTimer"
	$header += "TenantAdminEnabled"
	$header += "VoiceDeploymentMode"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsVoicePolicy.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsVoicePolicy.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsVoicePolicy sheet

#Region Get-CsVoiceRoutingPolicy sheet
Write-Host -Object "---- Starting Get-CsVoiceRoutingPolicy"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsVoiceRoutingPolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Name"
	$header += "Identity"
	$header += "AllowInternationalCalls"
	$header += "Description"
	$header += "HybridPSTNSiteIndex"
	$header += "PstnUsages"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsVoiceRoutingPolicy.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsVoiceRoutingPolicy.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsVoiceRoutingPolicy sheet

#EndRegion CsGeneral

#Region CsOnline
# 16 Functions
$intColorIndex = $intColorIndex_CsOnline

#Region Get-CsOnlineDialInConferencingBridge sheet
Write-Host -Object "---- Starting Get-CsOnlineDialInConferencingBridge"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsOnlineDialInConfBridge"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Name"
	$header += "DefaultServiceNumber"
	$header += "Identity"
	$header += "IsDefault"
	$header += "Region"
	$header += "ServiceNumbers.Number"
	$header += "ServiceNumbers.City"
	$header += "ServiceNumbers.PrimaryLanguage"
	$header += "ServiceNumbers.SecondaryLanguages"
	$header += "ServiceNumbers.BridgeId"
	$header += "ServiceNumbers.IsShared"
	$header += "ServiceNumbers.Type"
	$header += "TenantId"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsOnlineDialInConferencingBridge.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsOnlineDialInConferencingBridge.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsOnlineDialInConferencingBridge sheet

#Region Get-CsOnlineDialinConferencingPolicy sheet
Write-Host -Object "---- Starting Get-CsOnlineDialinConferencingPolicy"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsOnlineDialinConfPolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "AllowService"
	$header += "Description"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsOnlineDialinConferencingPolicy.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsOnlineDialinConferencingPolicy.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsOnlineDialinConferencingPolicy sheet

#Region Get-CsOnlineDialInConferencingServiceNumber sheet
Write-Host -Object "---- Starting Get-CsOnlineDialInConferencingServiceNumber"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsOnlineDialInConfServiceNumber"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "BridgeId"
	$header += "City"
	$header += "IsShared"
	$header += "Number"
	$header += "PrimaryLanguage"
	$header += "SecondaryLanguages"
	$header += "TenantId"
	$header += "Type"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsOnlineDialInConferencingServiceNumber.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsOnlineDialInConferencingServiceNumber.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsOnlineDialInConferencingServiceNumber sheet

#Region Get-CsOnlineDialinConferencingTenantConfiguration sheet
Write-Host -Object "---- Starting Get-CsOnlineDialinConferencingTenantConfiguration"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsOnlineDialinConfTenantConfig"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "Status"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsOnlineDialinConferencingTenantConfiguration.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsOnlineDialinConferencingTenantConfiguration.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsOnlineDialinConferencingTenantConfiguration sheet

#Region Get-CsOnlineDialInConferencingTenantSettings sheet
Write-Host -Object "---- Starting Get-CsOnlineDialInConferencingTenantSettings"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsOnlineDialInConfTenantSetting"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "AllowPSTNOnlyMeetingsByDefault"
	$header += "AutomaticallyMigrateUserMeetings"
	$header += "AutomaticallyReplaceAcpProvider"
	$header += "AutomaticallySendEmailsToUsers"
	$header += "EnableDialOutJoinConfirmation"
	$header += "EnableEntryExitNotifications"
	$header += "EnableNameRecording"
	$header += "EntryExitAnnouncementsType"
	$header += "IncludeTollFreeNumberInMeetingInvites"
	$header += "MigrateServiceNumbersOnCrossForestMove"
	$header += "PinLength"
	$header += "SendEmailFromAddress"
	$header += "SendEmailFromDisplayName"
	$header += "SendEmailFromOverride"
	$header += "UseUniqueConferenceIds"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsOnlineDialInConferencingTenantSettings.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsOnlineDialInConferencingTenantSettings.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsOnlineDialInConferencingTenantSettings sheet

#Region Get-CsOnlineDialInConferencingUser sheet
Write-Host -Object "---- Starting Get-CsOnlineDialInConferencingUser"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsOnlineDialInConfUser"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "AllowPstnOnlyMeetings"
	$header += "AllowTollFreeDialIn"
	$header += "BridgeId"
	$header += "BridgeName"
	$header += "ConferenceId"
	$header += "LeaderPin"
	$header += "ServiceNumber"
	$header += "SipAddress"
	$header += "Tenant"
	$header += "TollFreeServiceNumber"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsOnlineDialInConferencingUser.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsOnlineDialInConferencingUser.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsOnlineDialInConferencingUser sheet

#Region Get-CsOnlineDialInConferencingUserInfo sheet
Write-Host -Object "---- Starting Get-CsOnlineDialInConferencingUserInfo"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsOnlineDialInConfUserInfo"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "DisplayName"
	$header += "Identity"
	$header += "ConferenceId"
	$header += "DefaultTollFreeNumbers"
	$header += "DefaultTollNumber"
	$header += "ObjectId"
	$header += "Provider"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsOnlineDialInConferencingUserInfo.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsOnlineDialInConferencingUserInfo.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsOnlineDialInConferencingUserInfo sheet

#Region Get-CsOnlineDialInConferencingUserState sheet
Write-Host -Object "---- Starting Get-CsOnlineDialInConferencingUserState"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsOnlineDialInConfUserState"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "DisplayName"
	$header += "Identity"
	$header += "ConferenceId"
	$header += "Domain"
	$header += "Provider"
	$header += "PstnConferencingLicenseState"
	$header += "SipAddress"
	$header += "TollFreeNumbers"
	$header += "TollNumber"
	$header += "Url"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsOnlineDialInConferencingUserState.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsOnlineDialInConferencingUserState.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsOnlineDialInConferencingUserState sheet

#Region Get-CsOnlineDialOutPolicy sheet
Write-Host -Object "---- Starting Get-CsOnlineDialOutPolicy"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsOnlineDialOutPolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "AllowPSTNConferencingDialOutType"
	$header += "AllowPSTNOutboundCallingType"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsOnlineDialOutPolicy.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsOnlineDialOutPolicy.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsOnlineDialOutPolicy sheet

#Region Get-CsOnlineDirectoryTenant sheet
Write-Host -Object "---- Starting Get-CsOnlineDirectoryTenant"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsOnlineDirectoryTenant"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Id"
	$header += "AnnouncementsDisabled"
	$header += "Bridges"
	$header += "DefaultBridge"
	$header += "DefaultPoolFqdn"
	$header += "Domains"
	$header += "NameRecordingDisabled"
	$header += "Pools"
	$header += "ServiceNumberCount"
	$header += "SubscriberNumberCount"
	$header += "TnmAccountId"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsOnlineDirectoryTenant.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsOnlineDirectoryTenant.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsOnlineDirectoryTenant sheet

#Region Get-CsOnlineTelephoneNumberInventoryTypes sheet
Write-Host -Object "---- Starting Get-CsOnlineTelephoneNumberInventoryTypes"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsOnlineTeleNumberInvTypes"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Id"
	$header += "Description"
	$header += "Regions"
	$header += "Reservations"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsOnlineTelephoneNumberInventoryTypes.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsOnlineTelephoneNumberInventoryTypes.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsOnlineTelephoneNumberInventoryTypes sheet

#Region Get-CsOnlineTelephoneNumberReservationsInformation sheet
Write-Host -Object "---- Starting Get-CsOnlineTelephoneNumberReservationsInformation"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsOnlineTeleNumberReservInfo"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "ActiveReservationsCount"
	$header += "ActiveReservedNumbersCount"
	$header += "MaximumActiveReservationsCount"
	$header += "MaximumActiveReservedNumbersCount"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsOnlineTelephoneNumberReservationsInformation.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsOnlineTelephoneNumberReservationsInformation.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsOnlineTelephoneNumberReservationsInformation sheet

#Region Get-CsOnlineUser sheet
Write-Host -Object "---- Starting Get-CsOnlineUser"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsOnlineUser"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Alias"
	$header += "DisplayName"
#	$header += "AcpInfo"
	$header += "Acp.Default"
	$header += "Acp.TollNumber"
	$header += "Acp.ParticipantPassCode"
	$header += "Acp.domain"
	$header += "Acp.name"
	$header += "Acp.url"
	$header += "AddressBookPolicy"
	$header += "AdminDescription"
	$header += "ArchivingPolicy"
	$header += "AssignedPlan"
	$header += "AudioVideoDisabled"
	$header += "BaseSimpleUrl"
	$header += "BroadcastMeetingPolicy"
	$header += "CallerIdPolicy"
	$header += "CallingLineIdentity"
	$header += "CallViaWorkPolicy"
	$header += "City"
	$header += "ClientPolicy"
	$header += "ClientUpdateOverridePolicy"
	$header += "ClientUpdatePolicy"
	$header += "ClientVersionPolicy"
	$header += "CloudMeetingOpsPolicy"
	$header += "CloudMeetingPolicy"
	$header += "CloudVideoInteropPolicy"
	$header += "Company"
	$header += "ConferencingPolicy"
	$header += "ContactOptionFlags"
	$header += "CountryAbbreviation"
	$header += "CountryOrRegionDisplayName"
	$header += "Department"
	$header += "Description"
	$header += "DialPlan"
	$header += "DirSyncEnabled"
	$header += "DistinguishedName"
	$header += "Enabled"
	$header += "EnabledForRichPresence"
	$header += "EnterpriseVoiceEnabled"
	$header += "ExchangeArchivingPolicy"
	$header += "ExchUserHoldPolicies"
	$header += "ExperiencePolicy"
	$header += "ExternalAccessPolicy"
	$header += "ExternalUserCommunicationPolicy"
	$header += "ExUmEnabled"
	$header += "Fax"
	$header += "FirstName"
	$header += "GraphPolicy"
	$header += "Guid"
	$header += "HideFromAddressLists"
	$header += "HomePhone"
	$header += "HomeServer"
	$header += "HostedVoiceMail"
	$header += "HostedVoicemailPolicy"
	$header += "HostingProvider"
	$header += "Id"
	$header += "Identity"
	$header += "InterpretedUserType"
	$header += "IPPBXSoftPhoneRoutingEnabled"
	$header += "IPPhone"
	$header += "IPPhonePolicy"
	$header += "IsByPassValidation"
	$header += "IsValid"
	$header += "LastName"
	$header += "LastProvisionTimeStamp"
	$header += "LastPublishTimeStamp"
	$header += "LastSubProvisionTimeStamp"
	$header += "LastSyncTimeStamp"
	$header += "LegalInterceptPolicy"
	$header += "LicenseRemovalTimestamp"
	$header += "LineServerURI"
	$header += "LineURI"
	$header += "LocationPolicy"
	$header += "Manager"
	$header += "MCOValidationError"
	$header += "MNCReady"
	$header += "MobilePhone"
	$header += "MobilityPolicy"
	$header += "Name"
	$header += "NonPrimaryResource"
	$header += "ObjectCategory"
	$header += "ObjectClass"
	$header += "ObjectId"
	$header += "ObjectState"
	$header += "Office"
	$header += "OnlineDialinConferencingPolicy"
	$header += "OnlineDialOutPolicy"
	$header += "OnlineVoicemailPolicy"
	$header += "OnlineVoiceRoutingPolicy"
	$header += "OnPremEnterpriseVoiceEnabled"
	$header += "OnPremHideFromAddressLists"
	$header += "OnPremHostingProvider"
	$header += "OnPremLineURI"
	$header += "OnPremLineURIManuallySet"
	$header += "OnPremOptionFlags"
	$header += "OnPremSipAddress"
	$header += "OnPremSIPEnabled"
	$header += "OptionFlags"
	$header += "OriginalPreferredDataLocation"
	$header += "OriginatingServer"
	$header += "OriginatorSid"
	$header += "OtherTelephone"
	$header += "OverridePreferredDataLocation"
	$header += "OwnerUrn"
	$header += "PendingDeletion"
	$header += "Phone"
	$header += "PinPolicy"
	$header += "PostalCode"
	$header += "PreferredDataLocation"
	$header += "PreferredDataLocationOverwritePolicy"
	$header += "PreferredLanguage"
	$header += "PresencePolicy"
	$header += "PrivateLine"
	$header += "ProvisionedPlan"
	$header += "ProvisioningCounter"
	$header += "ProvisioningStamp"
	$header += "ProxyAddresses"
	$header += "PublishingCounter"
	$header += "PublishingStamp"
	$header += "Puid"
	$header += "RegistrarPool"
	$header += "RemoteCallControlTelephonyEnabled"
	$header += "SamAccountName"
	$header += "ServiceInfo"
	$header += "ServiceInstance"
	$header += "ShadowProxyAddresses"
	$header += "Sid"
	$header += "SipAddress"
	$header += "SipProxyAddress"
	$header += "SmsServicePolicy"
	$header += "SoftDeletionTimestamp"
	$header += "StateOrProvince"
	$header += "Street"
	$header += "StreetAddress"
	$header += "StsRefreshTokensValidFrom"
	$header += "SubProvisioningCounter"
	$header += "SubProvisioningStamp"
	$header += "SubProvisionLineType"
	$header += "SyncingCounter"
	$header += "TargetRegistrarPool"
	$header += "TargetServerIfMoving"
	$header += "TeamsAppPermissionPolicy"
	$header += "TeamsAppSetupPolicy"
	$header += "TeamsCallingPolicy"
	$header += "TeamsCortanaPolicy"
	$header += "TeamsInteropPolicy"
	$header += "TeamsMeetingBroadcastPolicy"
	$header += "TeamsMeetingPolicy"
	$header += "TeamsMessagingPolicy"
	$header += "TeamsOwnersPolicy"
	$header += "TeamsUpgradeEffectiveMode"
	$header += "TeamsUpgradeNotificationsEnabled"
	$header += "TeamsUpgradeOverridePolicy"
	$header += "TeamsUpgradePolicy"
	$header += "TeamsUpgradePolicyIsReadOnly"
	$header += "TeamsVideoInteropServicePolicy"
	$header += "TeamsWorkLoadPolicy"
	$header += "TenantDialPlan"
	$header += "TenantId"
	$header += "ThirdPartyVideoSystemPolicy"
	$header += "ThumbnailPhoto"
	$header += "Title"
	$header += "UpgradeRetryCounter"
	$header += "UsageLocation"
	$header += "UserAccountControl"
	$header += "UserPrincipalName"
	$header += "UserProvisionType"
	$header += "UserRoutingGroupId"
	$header += "UserServicesPolicy"
	$header += "VoicePolicy"
	$header += "VoiceRoutingPolicy"
	$header += "WebPage"
	$header += "WhenChanged"
	$header += "WhenCreated"
	$header += "WindowsEmailAddress"
	$header += "XForestMovePolicy"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsOnlineUser.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsOnlineUser.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsOnlineUser sheet

#Region Get-CsOnlineVoicemailPolicy sheet
Write-Host -Object "---- Starting Get-CsOnlineVoicemailPolicy"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsOnlineVoicemailPolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "EnableTranscription"
	$header += "EnableTranscriptionProfanityMasking"
	$header += "ShareData"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsOnlineVoicemailPolicy.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsOnlineVoicemailPolicy.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsOnlineVoicemailPolicy sheet

#Region Get-CsOnlineVoiceRoute sheet
Write-Host -Object "---- Starting Get-CsOnlineVoiceRoute"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsOnlineVoiceRoute"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "Description"
	$header += "Name"
	$header += "NumberPattern"
	$header += "OnlinePstnGatewayList"
	$header += "OnlinePstnUsages"
	$header += "Priority"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsOnlineVoiceRoute.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsOnlineVoiceRoute.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsOnlineVoiceRoute sheet

#Region Get-CsOnlineVoiceRoutingPolicy sheet
Write-Host -Object "---- Starting Get-CsOnlineVoiceRoutingPolicy"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsOnlineVoiceRoutingPolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "Description"
	$header += "OnlinePstnUsages"
	$header += "RouteType"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsOnlineVoiceRoutingPolicy.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsOnlineVoiceRoutingPolicy.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsOnlineVoiceRoutingPolicy sheet

#EndRegion CsOnline

#Region CsTeams
# 16 Functions
$intColorIndex = $intColorIndex_CsTeams

#Region Get-CsTeamsCallingPolicy sheet
Write-Host -Object "---- Starting Get-CsTeamsCallingPolicy"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsTeamsCallingPolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "AllowCalling"
	$header += "AllowPrivateCalling"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsTeamsCallingPolicy.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsTeamsCallingPolicy.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsTeamsCallingPolicy sheet

#Region Get-CsTeamsClientConfiguration sheet
Write-Host -Object "---- Starting Get-CsTeamsClientConfiguration"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsTeamsClientConfig"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "AllowBox"
	$header += "AllowDropBox"
	$header += "AllowEmailIntoChannel"
	$header += "AllowGoogleDrive"
	$header += "AllowGuestUser"
	$header += "AllowOrganizationTab"
	$header += "AllowResourceAccountSendMessage"
	$header += "AllowScopedPeopleSearchandAccess"
	$header += "AllowShareFile"
	$header += "AllowSkypeBusinessInterop"
	$header += "AllowTBotProactiveMessaging"
	$header += "ContentPin"
	$header += "ResourceAccountContentAccess"
	$header += "RestrictedSenderList"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsTeamsClientConfiguration.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsTeamsClientConfiguration.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsTeamsClientConfiguration sheet

#Region Get-CsTeamsGuestCallingConfiguration sheet
Write-Host -Object "---- Starting Get-CsTeamsGuestCallingConfiguration"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsTeamsGuestCallingConfig"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "AllowPrivateCalling"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsTeamsGuestCallingConfiguration.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsTeamsGuestCallingConfiguration.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsTeamsGuestCallingConfiguration sheet

#Region Get-CsTeamsGuestMeetingConfiguration sheet
Write-Host -Object "---- Starting Get-CsTeamsGuestMeetingConfiguration"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsTeamsGuestMeetingConfig"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "AllowIpVideo"
	$header += "AllowMeetNow"
	$header += "ScreenSharingMode"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsTeamsGuestMeetingConfiguration.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsTeamsGuestMeetingConfiguration.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsTeamsGuestMeetingConfiguration sheet

#Region Get-CsTeamsGuestMessagingConfiguration sheet
Write-Host -Object "---- Starting Get-CsTeamsGuestMessagingConfiguration"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsTeamsGuestMessagingConfig"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "AllowGiphy"
	$header += "AllowMemes"
	$header += "AllowStickers"
	$header += "AllowUserChat"
	$header += "AllowUserDeleteMessage"
	$header += "AllowUserEditMessage"
	$header += "GiphyRatingType"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsTeamsGuestMessagingConfiguration.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsTeamsGuestMessagingConfiguration.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsTeamsGuestMessagingConfiguration sheet

#Region Get-CsTeamsInterop sheet
Write-Host -Object "---- Starting Get-CsTeamsInteropPolicy"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsTeamsInteropPolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "AllowEndUserClientOverride"
	$header += "CallingDefaultClient"
	$header += "ChatDefaultClient"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsTeamsInteropPolicy.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsTeamsInteropPolicy.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsTeamsInterop sheet

#Region Get-CsTeamsMeetingBroadcastConfiguration sheet
Write-Host -Object "---- Starting Get-CsTeamsMeetingBroadcastConfiguration"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsTeamsMeetingBroadcastConfig"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "AllowSdnProviderForBroadcastMeeting"
	$header += "SupportURL"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsTeamsMeetingBroadcastConfiguration.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsTeamsMeetingBroadcastConfiguration.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsTeamsMeetingBroadcastConfiguration sheet

#Region Get-CsTeamsMeetingBroadcastPolicy sheet
Write-Host -Object "---- Starting Get-CsTeamsMeetingBroadcastPolicy"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsTeamsMeetingBroadcastPolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "AllowBroadcastScheduling"
	$header += "AllowBroadcastTranscription"
	$header += "BroadcastAttendeeVisibilityMode"
	$header += "BroadcastRecordingMode"
	$header += "Description"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsTeamsMeetingBroadcastPolicy.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsTeamsMeetingBroadcastPolicy.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsTeamsMeetingBroadcastPolicy sheet

#Region Get-CsTeamsMeetingConfiguration sheet
Write-Host -Object "---- Starting Get-CsTeamsMeetingConfiguration"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsTeamsMeetingConfig"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "ClientAppSharingPort"
	$header += "ClientAppSharingPortRange"
	$header += "ClientAudioPort"
	$header += "ClientAudioPortRange"
	$header += "ClientMediaPortRangeEnabled"
	$header += "ClientVideoPort"
	$header += "ClientVideoPortRange"
	$header += "CustomFooterText"
	$header += "DisableAnonymousJoin"
	$header += "EnableQoS"
	$header += "HelpURL"
	$header += "LegalURL"
	$header += "LogoURL"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsTeamsMeetingConfiguration.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsTeamsMeetingConfiguration.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsTeamsMeetingConfiguration sheet

#Region Get-CsTeamsMeetingPolicy sheet
Write-Host -Object "---- Starting Get-CsTeamsMeetingPolicy"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsTeamsMeetingPolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "AllowAnonymousUsersToDialOut"
	$header += "AllowAnonymousUsersToStartMeeting"
	$header += "AllowChannelMeetingScheduling"
	$header += "AllowCloudRecording"
	$header += "AllowExternalParticipantGiveRequestControl"
	$header += "AllowIPVideo"
	$header += "AllowMeetNow"
	$header += "AllowOutlookAddIn"
	$header += "AllowParticipantGiveRequestControl"
	$header += "AllowPowerPointSharing"
	$header += "AllowPrivateMeetingScheduling"
	$header += "AllowSharedNotes"
	$header += "AllowTranscription"
	$header += "AllowWhiteboard"
	$header += "AutoAdmittedUsers"
	$header += "Description"
	$header += "MediaBitRateKb"
	$header += "ScreenSharingMode"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsTeamsMeetingPolicy.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsTeamsMeetingPolicy.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsTeamsMeetingPolicy sheet

#Region Get-CsTeamsMessagingPolicy sheet
Write-Host -Object "---- Starting Get-CsTeamsMessagingPolicy"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsTeamsMessagingPolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "AllowGiphy"
	$header += "AllowMemes"
	$header += "AllowOwnerDeleteMessage"
	$header += "AllowStickers"
	$header += "AllowUrlPreviews"
	$header += "AllowUserChat"
	$header += "AllowUserDeleteMessage"
	$header += "AllowUserEditMessage"
	$header += "AllowUserTranslation"
	$header += "Description"
	$header += "GiphyRatingType"
	$header += "ReadReceiptsEnabledType"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsTeamsMessagingPolicy.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsTeamsMessagingPolicy.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsTeamsMessagingPolicy sheet

#Region Get-CsTeamsMigrationConfiguration sheet
Write-Host -Object "---- Starting Get-CsTeamsMigrationConfiguration"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsTeamsMigrationConfig"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "EnableLegacyClientInterOp"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsTeamsMigrationConfiguration.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsTeamsMigrationConfiguration.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsTeamsMigrationConfiguration sheet

#Region Get-CsTeamsUpgradeConfiguration sheet
Write-Host -Object "---- Starting Get-CsTeamsUpgradeConfiguration"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsTeamsUpgradeConfig"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "DownloadTeams"
	$header += "SfBMeetingJoinUx"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsTeamsUpgradeConfiguration.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsTeamsUpgradeConfiguration.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsTeamsUpgradeConfiguration sheet

#Region Get-CsTeamsUpgradePolicy sheet
Write-Host -Object "---- Starting Get-CsTeamsUpgradePolicy"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsTeamsUpgradePolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "Action"
	$header += "Description"
	$header += "Mode"
	$header += "NotifySfbUsers"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsTeamsUpgradePolicy.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsTeamsUpgradePolicy.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsTeamsUpgradePolicy sheet

#Region Get-CsTeamsUpgradeStatus sheet
Write-Host -Object "---- Starting Get-CsTeamsUpgradeStatus"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsTeamsUpgradeStatus"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "TenantId"
	$header += "LastStateChangeDate"
	$header += "OptInEligibleDate"
	$header += "State"
	$header += "UpgradeDate"
	$header += "UpgradeScheduledDate"
	$header += "UserNotificationDate"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsTeamsUpgradeStatus.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsTeamsUpgradeStatus.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsTeamsUpgradeStatus sheet

#Region Get-CsTeamsVideoInteropServicePolicy sheet
Write-Host -Object "---- Starting Get-CsTeamsVideoInteropServicePolicy"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsTeamsVideoInteropServPolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "Description"
	$header += "Enabled"
	$header += "ProviderName"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsTeamsVideoInteropServicePolicy.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsTeamsVideoInteropServicePolicy.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsTeamsVideoInteropServicePolicy sheet

#EndRegion CsTeams

#Region CsTenant
# 9 Functions
$intColorIndex = $intColorIndex_CsTenant

#Region Get-CsTenant sheet
Write-Host -Object "---- Starting Get-CsTenant"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsTenant"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Name"
	$header += "AdminDescription"
	$header += "AllowedDataLocation"
	$header += "AssignedPlan"
	$header += "City"
	#$header += "CompanyPartnership"
	#$header += "CompanyTags"
	$header += "CountryAbbreviation"
	$header += "CountryOrRegionDisplayName"
	$header += "Description"
	$header += "DirSyncEnabled"
	$header += "DisableExoPlanProvisioning"
	$header += "DisableTeamsProvisioning"
	$header += "DisplayName"
	$header += "DistinguishedName"
	$header += "Domains"
	$header += "DomainUrlMap"
	$header += "ExperiencePolicy"
	$header += "Guid"
	$header += "Id"
	$header += "Identity"
	$header += "IsByPassValidation"
	$header += "IsMNC"
	$header += "IsReadinessUploaded"
	$header += "IsUpgradeReady"
	$header += "IsValid"
	$header += "LastProvisionTimeStamp"
	$header += "LastPublishTimeStamp"
	$header += "LastSubProvisionTimeStamp"
	$header += "LastSyncTimeStamp"
	$header += "MNCEnableTimeStamp"
	$header += "MNCReady"
	$header += "NonPrimaryResource"
	$header += "ObjectCategory"
	$header += "ObjectClass"
	$header += "ObjectId"
	$header += "ObjectState"
	$header += "OcoDomainsTracked"
	$header += "OriginalRegistrarPool"
	$header += "OriginatingServer"
	$header += "PendingDeletion"
	$header += "Phone"
	$header += "PostalCode"
	$header += "PreferredLanguage"
	$header += "ProvisionedPlan"
	$header += "ProvisioningCounter"
	$header += "ProvisioningStamp"
	$header += "PublicProvider"
	$header += "PublishingCounter"
	$header += "PublishingStamp"
	$header += "RegistrarPool"
	$header += "ServiceInfo"
	$header += "ServiceInstance"
	$header += "StateOrProvince"
	$header += "Street"
	$header += "SubProvisioningCounter"
	$header += "SubProvisioningStamp"
	$header += "SyncingCounter"
	$header += "TeamsUpgradeEffectiveMode"
	$header += "TeamsUpgradeEligible"
	$header += "TeamsUpgradeNotificationsEnabled"
	$header += "TeamsUpgradeOverridePolicy"
	$header += "TeamsUpgradePolicyIsReadOnly"
	$header += "TenantId"
	$header += "TenantNotified"
	$header += "TenantPoolExtension"
	$header += "UpgradeRetryCounter"
	$header += "UserRoutingGroupIds"
	$header += "WhenChanged"
	$header += "WhenCreated"
	$header += "XForestMovePolicy"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsTenant.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsTenant.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsTenant sheet

#Region Get-CsTenantBlockedCallingNumbers sheet
Write-Host -Object "---- Starting Get-CsTenantBlockedCallingNumbers"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsTenantBlockedCallingNumbers"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "Enabled"
	$header += "InboundBlockedNumberPatterns"
	$header += "Name"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsTenantBlockedCallingNumbers.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsTenantBlockedCallingNumbers.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsTenantBlockedCallingNumbers sheet

#Region Get-CsTenantDialPlan sheet
Write-Host -Object "---- Starting Get-CsTenantDialPlan"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsTenantDialPlan"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "Description"
	$header += "ExternalAccessPrefix"
	$header += "NormalizationRules"
	$header += "OptimizeDeviceDialing"
	$header += "SimpleName"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsTenantDialPlan.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsTenantDialPlan.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsTenantDialPlan sheet

#Region Get-CsTenantFederationConfiguration sheet
Write-Host -Object "---- Starting Get-CsTenantFederationConfiguration"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsTenantFederationConfig"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "AllowedDomains"
	$header += "AllowFederatedUsers"
	$header += "AllowPublicUsers"
	$header += "BlockedDomains"
	$header += "SharedSipAddressSpace"
	$header += "TreatDiscoveredPartnersAsUnverified"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsTenantFederationConfiguration.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsTenantFederationConfiguration.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsTenantFederationConfiguration sheet

#Region Get-CsTenantHybridConfiguration sheet
Write-Host -Object "---- Starting Get-CsTenantHybridConfiguration"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsTenantHybridConfig"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "HybridConfigServiceExternalUrl"
	$header += "HybridConfigServiceInternalUrl"
	$header += "HybridPSTNAppliances"
	$header += "HybridPSTNSites"
	$header += "PeerDestination"
	$header += "TenantUpdateTimeWindows"
	$header += "UseOnPremDialPlan"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsTenantHybridConfiguration.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsTenantHybridConfiguration.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsTenantHybridConfiguration sheet

#Region Get-CsTenantLicensingConfiguration sheet
Write-Host -Object "---- Starting Get-CsTenantLicensingConfiguration"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsTenantLicensingConfig"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "Status"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsTenantLicensingConfiguration.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsTenantLicensingConfiguration.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsTenantLicensingConfiguration sheet

#Region Get-CsTenantMigrationConfiguration sheet
Write-Host -Object "---- Starting Get-CsTenantMigrationConfiguration"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsTenantMigrationConfig"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "ACPMeetingMigrationTriggerEnabled"
	$header += "MeetingMigrationEnabled"
	$header += "MeetingMigrationSourceMeetingTypes"
	$header += "MeetingMigrationTargetMeetingTypes"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsTenantMigrationConfiguration.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsTenantMigrationConfiguration.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsTenantMigrationConfiguration sheet

#Region Get-CsTenantPublicProvider sheet
Write-Host -Object "---- Starting Get-CsTenantPublicProvider"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsTenantPublicProvider"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Domain"
	$header += "Provider"
	$header += "Status"
	$header += "TimeStamp"
	$header += "Detail"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsTenantPublicProvider.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsTenantPublicProvider.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsTenantPublicProvider sheet

#Region Get-CsTenantUpdateTimeWindow sheet
Write-Host -Object "---- Starting Get-CsTenantUpdateTimeWindow"
	$Worksheet = $Excel_Skype_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CsTenantUpdateTimeWindow"
	$Worksheet.Tab.ColorIndex = $intColorIndex
	$row = 1
	$header = @()
	$header += "Identity"
	$header += "DayOfMonth"
	$header += "DaysOfWeek"
	$header += "Duration"
	$header += "StartTime"
	$header += "Type"
	$header += "WeeksOfMonth"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Skype\Skype_CsTenantUpdateTimeWindow.txt") -eq $true)
{
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Skype\Skype_CsTenantUpdateTimeWindow.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Convert-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

#EndRegion Get-CsTenantUpdateTimeWindow sheet

#EndRegion CsTenant

# Autofit columns
Write-Host -Object "---- Starting Autofit"
$Excel_SkypeWorksheetCount = $Excel_Skype_workbook.worksheets.count
$AutofitSheetCount = 1
while ($AutofitSheetCount -le $Excel_SkypeWorksheetCount)
{
	$ActiveWorksheet = $Excel_Skype_workbook.worksheets.item($AutofitSheetCount)
	$objRange = $ActiveWorksheet.usedrange
	[Void]	$objRange.entirecolumn.autofit()
	$AutofitSheetCount++
}
$Excel_Skype_workbook.saveas($O365DC_Skype_XLS)
Write-Host -Object "---- Spreadsheet saved"
$Excel_Skype.workbooks.close()
Write-Host -Object "---- Workbook closed"
$Excel_Skype.quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel_Skype)
Remove-Variable -Name Excel_Skype
# If the ReleaseComObject doesn't do it..
#spps -n excel

	$EventLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
	$EventLog.MachineName = "."
	$EventLog.Source = "O365DC"
	try{$EventLog.WriteEntry("Ending Core_Assemble_Skype_Excel","Information", 43)}catch{}

