#############################################################################
#                     	Exchange_OrgConfig.ps1								#
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
$ErrorText = "Exchange_OrgConfig " + "`n" + $server + "`n"
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

$Exchange_OrgConfig_outputfile = $output_location + "\Exchange_OrgConfig.txt"

@(Get-OrganizationConfig) | ForEach-Object `
{
	$output_Exchange_OrgConfig = $_.Name + "`t" + `
		$_.DefaultPublicFolderAgeLimit + "`t" + `
		$_.DefaultPublicFolderIssueWarningQuota + "`t" + `
		$_.DefaultPublicFolderProhibitPostQuota + "`t" + `
		$_.DefaultPublicFolderMaxItemSize + "`t" + `
		$_.DefaultPublicFolderDeletedItemRetention + "`t" + `
		$_.DefaultPublicFolderMovedItemRetention + "`t" + `
		$_.PublicFoldersLockedForMigration + "`t" + `
		$_.PublicFolderMigrationComplete + "`t" + `
		$_.PublicFolderMailboxesLockedForNewConnections + "`t" + `
		$_.PublicFolderMailboxesMigrationComplete + "`t" + `
		$_.PublicFolderShowClientControl + "`t" + `
		$_.PublicFoldersEnabled + "`t" + `
		$_.ActivityBasedAuthenticationTimeoutInterval + "`t" + `
		$_.ActivityBasedAuthenticationTimeoutEnabled + "`t" + `
		$_.ActivityBasedAuthenticationTimeoutWithSingleSignOnEnabled + "`t" + `


		$_.AppsForOfficeEnabled + "`t" + `
		$_.AppsForOfficeCorpCatalogAppsCount + "`t" + `
		$_.PrivateCatalogAppsCount + "`t" + `
		$_.AVAuthenticationService + "`t" + `
		$_.CustomerFeedbackEnabled + "`t" + `
		$_.DistributionGroupDefaultOU + "`t" + `
		$_.DistributionGroupNameBlockedWordsList + "`t" + `
		$_.DistributionGroupNamingPolicy + "`t" + `
		$_.EwsAllowEntourage + "`t" + `
		$_.EwsAllowList + "`t" + `
		$_.EwsAllowMacOutlook + "`t" + `
		$_.EwsAllowOutlook + "`t" + `
		$_.EwsApplicationAccessPolicy + "`t" + `
		$_.EwsBlockList + "`t" + `
		$_.EwsEnabled + "`t" + `
		$_.IPListBlocked + "`t" + `
		$_.ElcProcessingDisabled + "`t" + `
		$_.AutoExpandingArchiveEnabled + "`t" + `
		$_.ExchangeNotificationEnabled + "`t" + `
		$_.ExchangeNotificationRecipients + "`t" + `
		$_.HierarchicalAddressBookRoot + "`t" + `
		$_.Industry + "`t" + `
		$_.MailTipsAllTipsEnabled + "`t" + `
		$_.MailTipsExternalRecipientsTipsEnabled + "`t" + `
		$_.MailTipsGroupMetricsEnabled + "`t" + `
		$_.MailTipsLargeAudienceThreshold + "`t" + `
		$_.MailTipsMailboxSourcedTipsEnabled + "`t" + `
		$_.ReadTrackingEnabled + "`t" + `
		$_.SCLJunkThreshold + "`t" + `
		$_.MaxConcurrentMigrations + "`t" + `
		$_.IntuneManagedStatus + "`t" + `
		$_.AzurePremiumSubscriptionStatus + "`t" + `
		$_.HybridConfigurationStatus + "`t" + `
		$_.ReleaseTrack + "`t" + `
		$_.CompassEnabled + "`t" + `
		$_.SharePointUrl + "`t" + `
		$_.MapiHttpEnabled + "`t" + `
		$_.RealTimeLogServiceEnabled + "`t" + `
		$_.CustomerLockboxEnabled + "`t" + `
		$_.UnblockUnsafeSenderPromptEnabled + "`t" + `
		$_.IsMixedMode + "`t" + `
		$_.ServicePlan + "`t" + `
		$_.DefaultDataEncryptionPolicy + "`t" + `
		$_.MailboxDataEncryptionEnabled + "`t" + `
		$_.GuestsEnabled + "`t" + `
		$_.GroupsCreationEnabled + "`t" + `
		$_.GroupsNamingPolicy + "`t" + `
		$_.OrganizationSummary
	$output_Exchange_OrgConfig | Out-File -FilePath $Exchange_OrgConfig_outputfile -append 
}

$EventText = "Exchange_OrgConfig " + "`n" + $server + "`n"
$RunTimeInSec = [int](((get-date) - $a).totalseconds)
$EventText += "Process run time`t`t" + $RunTimeInSec + " sec `n" 

$EventLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$EventLog.MachineName = "."
$EventLog.Source = "O365DC"
Try{$EventLog.WriteEntry($EventText,"Information", 35)}catch{}
