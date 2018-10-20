#############################################################################
#                     Exchange_TransportConfig.ps1							#
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
Param($location,$server,$i,$PSSession)

$a = get-date

$ErrorActionPreference = "Stop"
Trap {
$ErrorText = "Exchange_TransportConfig " + "`n" + $server + "`n"
$ErrorText += $_

$ErrorLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$ErrorLog.MachineName = "."
$ErrorLog.Source = "O365DC"
Try{$ErrorLog.WriteEntry($ErrorText,"Error", 100)}catch{}
}

set-location -LiteralPath $location
$output_location = $location + "\output\Exchange"

if ((Test-Path -LiteralPath $output_location) -eq $false)
    {New-Item -Path $output_location -ItemType directory -Force}

$Exchange_TransportConfig_outputfile = $output_location + "\Exchange_TransportConfig.txt"

@(Get-TransportConfig) | ForEach-Object `
{
	$output_Exchange_TransportConfig = "Transport Config" + "`t" + `
	$_.AddressBookPolicyRoutingEnabled + "`t" + `
	$_.AnonymousSenderToRecipientRatePerHour + "`t" + `
	$_.ClearCategories + "`t" + `
	$_.ConvertDisclaimerWrapperToEml + "`t" + `
	$_.DSNConversionMode + "`t" + `
	$_.ExternalDelayDsnEnabled + "`t" + `
	$_.ExternalDsnDefaultLanguage + "`t" + `
	$_.ExternalDsnLanguageDetectionEnabled + "`t" + `
	$_.ExternalDsnMaxMessageAttachSize + "`t" + `
	$_.ExternalDsnReportingAuthority + "`t" + `
	$_.ExternalDsnSendHtml + "`t" + `
	$_.ExternalPostmasterAddress + "`t" + `
	$_.GenerateCopyOfDSNFor + "`t" + `
	$_.HeaderPromotionModeSetting + "`t" + `
	$_.HygieneSuite + "`t" + `
	$_.InternalDelayDsnEnabled + "`t" + `
	$_.InternalDsnDefaultLanguage + "`t" + `
	$_.InternalDsnLanguageDetectionEnabled + "`t" + `
	$_.InternalDsnMaxMessageAttachSize + "`t" + `
	$_.InternalDsnReportingAuthority + "`t" + `
	$_.InternalDsnSendHtml + "`t" + `
	$_.InternalSMTPServers + "`t" + `
	$_.JournalArchivingEnabled + "`t" + `
	$_.JournalingReportNdrTo + "`t" + `
	$_.LegacyArchiveJournalingEnabled + "`t" + `
	$_.LegacyArchiveLiveJournalingEnabled + "`t" + `
	$_.LegacyJournalingMigrationEnabled + "`t" + `
	$_.MaxDumpsterSizePerDatabase + "`t" + `
	$_.MaxDumpsterTime + "`t" + `
	$_.MaxReceiveSize + "`t" + `
	$_.MaxRecipientEnvelopeLimit + "`t" + `
	$_.MaxRetriesForLocalSiteShadow + "`t" + `
	$_.MaxRetriesForRemoteSiteShadow + "`t" + `
	$_.MaxSendSize + "`t" + `
	$_.MigrationEnabled + "`t" + `
	$_.OpenDomainRoutingEnabled + "`t" + `
	$_.RedirectDLMessagesForLegacyArchiveJournaling + "`t" + `
	$_.RedirectUnprovisionedUserMessagesForLegacyArchiveJournaling + "`t" + `
	$_.RejectMessageOnShadowFailure + "`t" + `
	$_.Rfc2231EncodingEnabled + "`t" + `
	$_.SafetyNetHoldTime + "`t" + `
	$_.ShadowHeartbeatFrequency + "`t" + `
	$_.ShadowMessageAutoDiscardInterval + "`t" + `
	$_.ShadowMessagePreferenceSetting + "`t" + `
	$_.ShadowRedundancyEnabled + "`t" + `
	$_.ShadowResubmitTimeSpan + "`t" + `
	$_.SmtpClientAuthenticationDisabled + "`t" + `
	$_.SupervisionTags + "`t" + `
	$_.TLSReceiveDomainSecureList + "`t" + `
	$_.TLSSendDomainSecureList + "`t" + `
	$_.VerifySecureSubmitEnabled + "`t" + `
	$_.VoicemailJournalingEnabled + "`t" + `
	$_.Xexch50Enabled
	$output_Exchange_TransportConfig | Out-File -FilePath $Exchange_TransportConfig_outputfile -append 
}

$EventText = "Exchange_TransportConfig " + "`n" + $server + "`n"
$RunTimeInSec = [int](((get-date) - $a).totalseconds)
$EventText += "Process run time`t`t" + $RunTimeInSec + " sec `n" 

$EventLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$EventLog.MachineName = "."
$EventLog.Source = "O365DC"
try{$EventLog.WriteEntry($EventText,"Information", 35)}catch{}
