#############################################################################
#                   Exchange_EmailAddressPolicy.ps1		 					#
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
$ErrorText = "Exchange_EmailAddressPolicy " + "`n" + $server + "`n"
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

$Exchange_EAP_outputfile = $output_location + "\Exchange_EmailAddressPolicy.txt"

@(Get-EmailAddressPolicy) | ForEach-Object `
{
	$output_Exchange_EAP = $_.Name + "`t" + `
		$_.IsValid + "`t" + `
		$_.RecipientFilter + "`t" + `
		$_.LdapRecipientFilter + "`t" + `
		$_.LastUpdatedRecipientFilter + "`t" + `
		$_.RecipientFilterApplied + "`t" + `
		$_.IncludedRecipients + "`t" + `
		$_.ConditionalDepartment + "`t" + `
		$_.ConditionalCompany + "`t" + `
		$_.ConditionalStateOrProvince + "`t" + `
		$_.ConditionalCustomAttribute1 + "`t" + `
		$_.ConditionalCustomAttribute2 + "`t" + `
		$_.ConditionalCustomAttribute3 + "`t" + `
		$_.ConditionalCustomAttribute4 + "`t" + `
		$_.ConditionalCustomAttribute5 + "`t" + `
		$_.ConditionalCustomAttribute6 + "`t" + `
		$_.ConditionalCustomAttribute7 + "`t" + `
		$_.ConditionalCustomAttribute8 + "`t" + `
		$_.ConditionalCustomAttribute9 + "`t" + `
		$_.ConditionalCustomAttribute10 + "`t" + `
		$_.ConditionalCustomAttribute11 + "`t" + `
		$_.ConditionalCustomAttribute12 + "`t" + `
		$_.ConditionalCustomAttribute13 + "`t" + `
		$_.ConditionalCustomAttribute14 + "`t" + `
		$_.ConditionalCustomAttribute15 + "`t" + `
		$_.RecipientContainer + "`t" + `
		$_.RecipientFilterType + "`t" + `
		$_.Priority + "`t" + `
		$_.EnabledPrimarySMTPAddressTemplate + "`t" + `
		$_.EnabledEmailAddressTemplates + "`t" + `
		$_.DisabledEmailAddressTemplates + "`t" + `
		$_.HasEmailAddressSetting + "`t" + `
		$_.HasMailboxManagerSetting + "`t" + `
		$_.NonAuthoritativeDomains + "`t" + `
		$_.ExchangeVersion
	$output_Exchange_EAP | Out-File -FilePath $Exchange_EAP_outputfile -append 
}

$EventText = "Exchange_EmailAddressPolicy " + "`n" + $server + "`n"
$RunTimeInSec = [int](((get-date) - $a).totalseconds)
$EventText += "Process run time`t`t" + $RunTimeInSec + " sec `n" 

$EventLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$EventLog.MachineName = "."
$EventLog.Source = "O365DC"
Try{$EventLog.WriteEntry($EventText,"Information", 35)}catch{}
