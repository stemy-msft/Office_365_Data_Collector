#############################################################################
#                        Exchange_Contact.ps1	 							#
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
$ErrorText = "Exchange_Contact " + "`n" + $server + "`n"
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

$Exchange_Contact_outputfile = $output_location + "\Exchange_Contact.txt"

@(Get-Contact) | ForEach-Object `
{
	$output_Exchange_Contact = $_.DisplayName + "`t" + `
	$_.AdministrativeUnits + "`t" + `
	$_.AllowUMCallsFromNonUsers + "`t" + `
	$_.AssistantName + "`t" + `
	$_.City + "`t" + `
	$_.Company + "`t" + `
	$_.CountryOrRegion + "`t" + `
	$_.Department + "`t" + `
	$_.DirectReports + "`t" + `
	$_.Fax + "`t" + `
	$_.FirstName + "`t" + `
	$_.GeoCoordinates + "`t" + `
	$_.Guid + "`t" + `
	$_.HomePhone + "`t" + `
	$_.Identity + "`t" + `
	$_.Initials + "`t" + `
	$_.IsDirSynced + "`t" + `
	$_.IsValid + "`t" + `
	$_.LastName + "`t" + `
	$_.Manager + "`t" + `
	$_.MobilePhone + "`t" + `
	$_.Name + "`t" + `
	$_.Notes + "`t" + `
	$_.Office + "`t" + `
	$_.OrganizationId + "`t" + `
	$_.OtherFax + "`t" + `
	$_.OtherHomePhone + "`t" + `
	$_.OtherTelephone + "`t" + `
	$_.Pager + "`t" + `
	$_.Phone + "`t" + `
	$_.PhoneticDisplayName + "`t" + `
	$_.PostalCode + "`t" + `
	$_.PostOfficeBox + "`t" + `
	$_.RecipientType + "`t" + `
	$_.RecipientTypeDetails + "`t" + `
	$_.SeniorityIndex + "`t" + `
	$_.SimpleDisplayName + "`t" + `
	$_.StateOrProvince + "`t" + `
	$_.StreetAddress + "`t" + `
	$_.TelephoneAssistant + "`t" + `
	$_.Title + "`t" + `
	$_.UMCallingLineIds + "`t" + `
	$_.UMDialPlan + "`t" + `
	$_.UMDtmfMap + "`t" + `
	$_.VoiceMailSettings + "`t" + `
	$_.WebPage + "`t" + `
	$_.WhenChangedUTC + "`t" + `
	$_.WhenCreatedUTC + "`t" + `
	$_.WindowsEmailAddress
	$output_Exchange_Contact | Out-File -FilePath $Exchange_Contact_outputfile -append 
}

$EventText = "Exchange_Contact " + "`n" + $server + "`n"
$RunTimeInSec = [int](((get-date) - $a).totalseconds)
$EventText += "Process run time`t`t" + $RunTimeInSec + " sec `n" 

$EventLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$EventLog.MachineName = "."
$EventLog.Source = "O365DC"
Try{$EventLog.WriteEntry($EventText,"Information", 35)}catch{}