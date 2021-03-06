#############################################################################
#                      Exchange_UnifiedGroup.ps1							#
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
$ErrorText = "Exchange_UnifiedGroup " + "`n" + $server + "`n"
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

$Exchange_DG_outputfile = $output_location + "\Exchange_UnifiedGroup.txt"

@(Get-UnifiedGroup -resultsize unlimited) | ForEach-Object `
{	
	$Notes = $_.Notes -replace "`n"," " -replace "`r"," "
	$output_Exchange_DG = $_.DisplayName + "`t" + `
	$_.AcceptMessagesOnlyFrom + "`t" + `
	$_.AcceptMessagesOnlyFromDLMembers + "`t" + `
	$_.AcceptMessagesOnlyFromSendersOrMembers + "`t" + `
	$_.AccessType + "`t" + `
	$_.AddressListMembership + "`t" + `
	$_.AdministrativeUnits + "`t" + `
	$_.Alias + "`t" + `
	$_.AllowAddGuests + "`t" + `
	$_.AlwaysSubscribeMembersToCalendarEvents + "`t" + `
	$_.AutoSubscribeNewMembers + "`t" + `
	$_.BypassModerationFromSendersOrMembers + "`t" + `
	$_.CalendarMemberReadOnly + "`t" + `
	$_.CalendarUrl + "`t" + `
	$_.Classification + "`t" + `
	$_.ConnectorsEnabled + "`t" + `
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
	$_.DataEncryptionPolicy + "`t" + `
	$_.EmailAddresses + "`t" + `
	$_.EmailAddressPolicyEnabled + "`t" + `
	$_.ExchangeGuid + "`t" + `
	$_.ExchangeVersion + "`t" + `
	$_.ExpansionServer + "`t" + `
	$_.ExpirationTime + "`t" + `
	$_.ExtensionCustomAttribute1 + "`t" + `
	$_.ExtensionCustomAttribute2 + "`t" + `
	$_.ExtensionCustomAttribute3 + "`t" + `
	$_.ExtensionCustomAttribute4 + "`t" + `
	$_.ExtensionCustomAttribute5 + "`t" + `
	$_.ExternalDirectoryObjectId + "`t" + `
	$_.FileNotificationsSettings + "`t" + `
	$_.GrantSendOnBehalfTo + "`t" + `
	$_.GroupExternalMemberCount + "`t" + `
	$_.GroupMemberCount + "`t" + `
	$_.GroupPersonification + "`t" + `
	$_.GroupSKU + "`t" + `
	$_.GroupType + "`t" + `
	$_.Guid + "`t" + `
	$_.HiddenFromAddressListsEnabled + "`t" + `
	$_.HiddenFromExchangeClientsEnabled + "`t" + `
	$_.HiddenGroupMembershipEnabled + "`t" + `
	$_.Identity + "`t" + `
	$_.InboxUrl + "`t" + `
	$_.IsDirSynced + "`t" + `
	$_.IsExternalResourcesPublished + "`t" + `
	$_.IsMailboxConfigured + "`t" + `
	$_.IsMembershipDynamic + "`t" + `
	$_.IsValid + "`t" + `
	$_.Language + "`t" + `
	$_.LastExchangeChangedTime + "`t" + `
	$_.MailboxProvisioningConstraint + "`t" + `
	$_.MailboxRegion + "`t" + `
	$_.MailTip + "`t" + `
	$_.MailTipTranslations + "`t" + `
	$_.ManagedBy + "`t" + `
	$_.ManagedByDetails + "`t" + `
	$_.MaxReceiveSize + "`t" + `
	$_.MaxSendSize + "`t" + `
	$_.MigrationToUnifiedGroupInProgress + "`t" + `
	$_.ModeratedBy + "`t" + `
	$_.ModerationEnabled + "`t" + `
	$_.Name + "`t" + `
	$Notes + "`t" + `
	$_.OrganizationId + "`t" + `
	$_.PeopleUrl + "`t" + `
	$_.PhotoUrl + "`t" + `
	$_.PoliciesExcluded + "`t" + `
	$_.PoliciesIncluded + "`t" + `
	$_.PrimarySmtpAddress + "`t" + `
	$_.ProvisioningOption + "`t" + `
	$_.RecipientType + "`t" + `
	$_.RecipientTypeDetails + "`t" + `
	$_.RejectMessagesFrom + "`t" + `
	$_.RejectMessagesFromDLMembers + "`t" + `
	$_.RejectMessagesFromSendersOrMembers + "`t" + `
	$_.ReportToManagerEnabled + "`t" + `
	$_.ReportToOriginatorEnabled + "`t" + `
	$_.RequireSenderAuthenticationEnabled + "`t" + `
	$_.SendModerationNotifications + "`t" + `
	$_.SendOofMessageToOriginatorEnabled + "`t" + `
	$_.SharePointDocumentsUrl + "`t" + `
	$_.SharePointNotebookUrl + "`t" + `
	$_.SharePointSiteUrl + "`t" + `
	$_.SubscriptionEnabled + "`t" + `
	$_.WelcomeMessageEnabled + "`t" + `
	$_.WhenChangedUTC + "`t" + `
	$_.WhenCreatedUTC + "`t" + `
	$_.WhenSoftDeleted + "`t" + `
	$_.YammerEmailAddress
	$output_Exchange_DG | Out-File -FilePath $Exchange_DG_outputfile -append 
}

$EventText = "Exchange_UnifiedGroup " + "`n" + $server + "`n"
$RunTimeInSec = [int](((get-date) - $a).totalseconds)
$EventText += "Process run time`t`t" + $RunTimeInSec + " sec `n" 

$EventLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$EventLog.MachineName = "."
$EventLog.Source = "O365DC"
Try{$EventLog.WriteEntry($EventText,"Information", 35)}catch{}
