#############################################################################
#                      Exchange_RemoteDomain.ps1		 					#
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
$ErrorText = "Exchange_ReceiveConnector " + "`n" + $server + "`n"
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

$Exchange_RemoteDomain_outputfile = $output_location + "\Exchange_RemoteDomain.txt"

@(Get-RemoteDomain) | ForEach-Object `
{
	$output_Exchange_RemoteDomain = $_.Identity + "`t" + `
		$_.DomainName + "`t" + `
		$_.IsInternal + "`t" + `
		$_.TargetDeliveryDomain + "`t" + `
		$_.CharacterSet + "`t" + `
		$_.NonMimeCharacterSet + "`t" + `
		$_.AllowedOOFType + "`t" + `
		$_.AutoReplyEnabled + "`t" + `
		$_.AutoForwardEnabled + "`t" + `
		$_.DeliveryReportEnabled + "`t" + `
		$_.NDREnabled + "`t" + `
		$_.MeetingForwardNotificationEnabled + "`t" + `
		$_.ContentType + "`t" + `
		$_.DisplaySenderName + "`t" + `
		$_.TNEFEnabled + "`t" + `
		$_.LineWrapSize + "`t" + `
		$_.TrustedMailOutboundEnabled + "`t" + `
		$_.TrustedMailInboundEnabled + "`t" + `
		$_.UseSimpleDisplayName
	$output_Exchange_RemoteDomain | Out-File -FilePath $Exchange_RemoteDomain_outputfile -append 
}

$EventText = "Exchange_RemoteConnector " + "`n" + $server + "`n"
$RunTimeInSec = [int](((get-date) - $a).totalseconds)
$EventText += "Process run time`t`t" + $RunTimeInSec + " sec `n" 

$EventLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$EventLog.MachineName = "."
$EventLog.Source = "O365DC"
Try{$EventLog.WriteEntry($EventText,"Information", 35)}catch{}
