#############################################################################
#                  Exchange_DynamicDistributionGroup.ps1					#
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

$a=get-date

$ErrorActionPreference = "Stop"
Trap {
$ErrorText = "Exchange_DynamicDistributionGroup " + "`n" + $server + "`n"
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

$Exchange_DDG_outputfile = $output_location + "\Exchange_DynamicDistributionGroup.txt"

@(Get-DynamicDistributionGroup -resultsize unlimited) | ForEach-Object `
{	
	$output_Exchange_DDG = $_.Name + "`t" + `
		$_.RecipientContainer + "`t" + `
		$_.RecipientFilter + "`t" + `
		$_.LdapRecipientFilter + "`t" + `
		$_.IncludedRecipients + "`t" + `
		$_.ManagedBy + "`t" + `
		$_.ExpansionServer + "`t" + `
		$_.ReportToManagerEnabled + "`t" + `
		$_.ReportToOriginatorEnabled + "`t" + `
		$_.SendOofMessageToOriginatorEnabled + "`t" + `
		$_.AcceptMessagesOnlyFrom + "`t" + `
		$_.AcceptMessagesOnlyFromDLMembers + "`t" + `
		$_.AcceptMessagesOnlyFromSendersOrMembers + "`t" + `
		$_.Alias + "`t" + `
		$_.OrganizationalUnit + "`t" + `
		$_.CustomAttribute1 + "`t" + `
		$_.CustomAttribute10 + "`t" + `
		$_.CustomAttribute11 + "`t" + `
		$_.CustomAttribute12 + "`t" + `
		$_.CustomAttribute13 + "`t" + `
		$_.CustomAttribute14 + "`t" + `
		$_.CustomAttribute15 + "`t" + `
		$_.CustomAttribute2 + "`t" + `
		$_.CustomAttribute3 + "`t" + `
		$_.CustomAttribute4 + "`t" + `
		$_.CustomAttribute5 + "`t" + `
		$_.CustomAttribute6 + "`t" + `
		$_.CustomAttribute7 + "`t" + `
		$_.CustomAttribute8 + "`t" + `
		$_.CustomAttribute9 + "`t" + `
		$_.ExtensionCustomAttribute1 + "`t" + `
		$_.ExtensionCustomAttribute2 + "`t" + `
		$_.ExtensionCustomAttribute3 + "`t" + `
		$_.ExtensionCustomAttribute4 + "`t" + `
		$_.ExtensionCustomAttribute5 + "`t" + `
		$_.DisplayName + "`t" + `
		$_.GrantSendOnBehalfTo + "`t" + `
		$_.HiddenFromAddressListsEnabled + "`t" + `
		$_.MaxSendSize + "`t" + `
		$_.MaxReceiveSize + "`t" + `
		$_.ModeratedBy + "`t" + `
		$_.ModerationEnabled + "`t" + `
		$_.PrimarySmtpAddress + "`t" + `
		$_.RecipientType + "`t" + `
		$_.RecipientTypeDetails + "`t" + `
		$_.RejectMessagesFrom + "`t" + `
		$_.RejectMessagesFromDLMembers + "`t" + `
		$_.RejectMessagesFromSendersOrMembers + "`t" + `
		$_.RequireSenderAuthenticationEnabled + "`t" + `
		$_.WhenCreatedUTC + "`t" + `
		$_.WhenChangedUTC + "`t" + `
		$_.IsValid
	$output_Exchange_DDG | Out-File -FilePath $Exchange_DDG_outputfile -append 
}

$EventText = "Exchange_DynamicDistributionGroup " + "`n" + $server + "`n"
$RunTimeInSec = [int](((get-date) - $a).totalseconds)
$EventText += "Process run time`t`t" + $RunTimeInSec + " sec `n" 

$EventLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$EventLog.MachineName = "."
$EventLog.Source = "O365DC"
Try{$EventLog.WriteEntry($EventText,"Information", 35)}catch{}
