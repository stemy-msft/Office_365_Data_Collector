#############################################################################
#                    Exchange_UmMailboxPolicy.ps1		 					#
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
$ErrorText = "Exchange_UmMailboxPolicy " + "`n" + $server + "`n"
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

$Exchange_UmMailboxPolicy_outputfile = $output_location + "\Exchange_UmMailboxPolicy.txt"

@(Get-UmMailboxPolicy) | ForEach-Object `
{
	$output_Exchange_UmMailboxPolicy = $_.name + "`t" + `
		$_.Identity + "`t" + `
		$_.MaxGreetingDuration + "`t" + `
		$_.MaxLogonAttempts + "`t" + `
		$_.AllowCommonPatterns + "`t" + `
		$_.PINLifetime + "`t" + `
		$_.PINHistoryCount + "`t" + `
		$_.AllowSMSNotification + "`t" + `
		$_.ProtectUnauthenticatedVoiceMail + "`t" + `
		$_.ProtectAuthenticatedVoiceMail + "`t" + `
		$_.ProtectedVoiceMailText + "`t" + `
		$_.RequireProtectedPlayOnPhone + "`t" + `
		$_.MinPINLength + "`t" + `
		$_.FaxMessageText + "`t" + `
		$_.UMEnabledText + "`t" + `
		$_.ResetPINText + "`t" + `
		$_.SourceForestPolicyNames + "`t" + `
		$_.VoiceMailText + "`t" + `
		$_.UMDialPlan + "`t" + `
		$_.FaxServerURI + "`t" + `
		$_.AllowedInCountryOrRegionGroups + "`t" + `
		$_.AllowedInternationalGroups + "`t" + `
		$_.AllowDialPlanSubscribers + "`t" + `
		$_.AllowExtensions + "`t" + `
		$_.LogonFailuresBeforePINReset + "`t" + `
		$_.AllowMissedCallNotifications + "`t" + `
		$_.AllowFax + "`t" + `
		$_.AllowTUIAccessToCalendar + "`t" + `
		$_.AllowTUIAccessToEmail + "`t" + `
		$_.AllowSubscriberAccess + "`t" + `
		$_.AllowTUIAccessToDirectory + "`t" + `
		$_.AllowTUIAccessToPersonalContacts + "`t" + `
		$_.AllowAutomaticSpeechRecognition + "`t" + `
		$_.AllowPlayOnPhone + "`t" + `
		$_.AllowVoiceMailPreview + "`t" + `
		$_.AllowCallAnsweringRules + "`t" + `
		$_.AllowMessageWaitingIndicator + "`t" + `
		$_.AllowPinlessVoiceMailAccess + "`t" + `
		$_.AllowVoiceResponseToOtherMessageTypes + "`t" + `
		$_.AllowVoiceMailAnalysis + "`t" + `
		$_.AllowVoiceNotification + "`t" + `
		$_.InformCallerOfVoiceMailAnalysis + "`t" + `
		$_.VoiceMailPreviewPartnerAddress + "`t" + `
		$_.VoiceMailPreviewPartnerAssignedID + "`t" + `
		$_.VoiceMailPreviewPartnerMaxMessageDuration + "`t" + `
		$_.VoiceMailPreviewPartnerMaxDeliveryDelay + "`t" + `
		$_.IsDefault + "`t" + `
		$_.IsValid
	$output_Exchange_UmMailboxPolicy | Out-File -FilePath $Exchange_UmMailboxPolicy_outputfile -append 
}

$EventText = "Exchange_UmMailboxPolicy " + "`n" + $server + "`n"
$RunTimeInSec = [int](((get-date) - $a).totalseconds)
$EventText += "Process run time`t`t" + $RunTimeInSec + " sec `n" 

$EventLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$EventLog.MachineName = "."
$EventLog.Source = "O365DC"
try{$EventLog.WriteEntry($EventText,"Information", 35)}catch{}
