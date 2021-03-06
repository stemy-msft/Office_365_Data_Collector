#############################################################################
#                        Exchange_UmMailbox.ps1		 						#
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
$ErrorText = "Exchange_UmMailbox " + "`n" + $server + "`n"
$ErrorText += $_

$ErrorLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$ErrorLog.MachineName = "."
$ErrorLog.Source = "O365DC"
Try{$ErrorLog.WriteEntry($ErrorText,"Error", 100)}catch{}
}

set-location -LiteralPath $location
$output_location = $location + "\output\Exchange\GetUmMailbox"

if ((Test-Path -LiteralPath $output_location) -eq $false)
    {New-Item -Path $output_location -ItemType directory -Force}

@(Get-Content -Path ".\CheckedMailbox.Set$i.txt") | ForEach-Object `
{
	$Exchange_UmMailbox_outputfile = $output_location + "\\Set$i~~GetUmMailbox.txt"
    $mailbox = $_
	if (((get-mailbox $mailbox).UMEnabled) -eq $true)
	{
		@(Get-UmMailbox $mailbox) | ForEach-Object `
		{
			$output_Exchange_UmMailbox = $_.name + "`t" + `
				$_.Identity + "`t" + `
				$_.EmailAddresses + "`t" + `
				$_.UMAddresses + "`t" + `
				$_.LegacyExchangeDN + "`t" + `
				$_.LinkedMasterAccount + "`t" + `
				$_.PrimarySmtpAddress + "`t" + `
				$_.SamAccountName + "`t" + `
				$_.ServerLegacyDN + "`t" + `
				$_.ServerName + "`t" + `
				$_.UMDtmfMap + "`t" + `
				$_.UMEnabled + "`t" + `
				$_.TUIAccessToCalendarEnabled + "`t" + `
				$_.FaxEnabled + "`t" + `
				$_.TUIAccessToEmailEnabled + "`t" + `
				$_.SubscriberAccessEnabled + "`t" + `
				$_.MissedCallNotificationEnabled + "`t" + `
				$_.UMSMSNotificationOption + "`t" + `
				$_.PinlessAccessToVoiceMailEnabled + "`t" + `
				$_.AnonymousCallersCanLeaveMessages + "`t" + `
				$_.AutomaticSpeechRecognitionEnabled + "`t" + `
				$_.PlayOnPhoneEnabled + "`t" + `
				$_.CallAnsweringRulesEnabled + "`t" + `
				$_.AllowUMCallsFromNonUsers + "`t" + `
				$_.OperatorNumber + "`t" + `
				$_.PhoneProviderId + "`t" + `
				$_.UMDialPlan + "`t" + `
				$_.UMMailboxPolicy + "`t" + `
				$_.Extensions + "`t" + `
				$_.CallAnsweringAudioCodec + "`t" + `
				$_.SIPResourceIdentifier + "`t" + `
				$_.PhoneNumber + "`t" + `
				$_.AirSyncNumbers + "`t" + `
				$_.IsValid
			$output_Exchange_UmMailbox | Out-File -FilePath $Exchange_UmMailbox_outputfile -append 
		}
	}
}

$EventText = "Exchange_UmMailbox " + "`n" + $server + "`n"
$RunTimeInSec = [int](((get-date) - $a).totalseconds)
$EventText += "Process run time`t`t" + $RunTimeInSec + " sec `n" 

$EventLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$EventLog.MachineName = "."
$EventLog.Source = "O365DC"
try{$EventLog.WriteEntry($EventText,"Information", 35)}catch{}
