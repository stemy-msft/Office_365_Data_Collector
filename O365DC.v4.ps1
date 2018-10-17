<#
.SYNOPSIS
	Collects data in a Office 365 environment and assembles the output into Word and Excel files.
.DESCRIPTION
	O365DC is used to collect a large amount of information about an Office 365 environment with minimal effort.
	The data is initially collected into a series of text files that can then be assembled into reports on the
	data collection server or another workstation.  This script was originally written for use by
	Microsoft Premier Services engineers during onsite engagements.

	Script guidelines:
	* Complete data collection requires elevated credentials for both Office 365 components
	* Data collection does not require that Office is installed as output files are txt and xml
	* The O365DC folder can be forklifted to a workstation with Office to generate output reports
.PARAMETER JobCount_Exchange
	Max number of jobs for Exchange cmdlet functions (Default = 10)
	Caution: The OOB throttling policy sets PowershellMaxConcurrency at 18 sessions per user per server
	Modifying this value without increasing the throttling policy can cause Exchange jobs to immediately fail.
.PARAMETER JobPolling_Exchange
	Polling interval for job completion for Exchange cmdlet functions (Default = 5 sec)
.PARAMETER Timeout_Exchange_Job
	Job timeout for Exchange functions  (Default = 3600 sec)
	The default value is 3600 seconds but should be adjusted for organizations with a large number of mailboxes or servers over slow connections.
.PARAMETER ServerForPSSession
	Exchange server used for Powershell sessions
.PARAMETER INI_Exchange
	Specify INI file for Exchange Tests configuration
.PARAMETER NoEMS
	Use this switch to launch the tool in Powershell (No Exchange cmdlets)
.PARAMETER MFA
	Use this switch if Multi-Factor Authentication is required for the environment.
	If is recommended that the Trusted IPs be set in Azure AD Conditional Access to allow the admin account to use traditional user name
	and password when run from a trusted IP.  If this switch is set, multi-threading will be disabled.
.PARAMETER ForceNewConnection
	Use this switch to force Powershell to make a new connection to Office 365 instead of trying to re-use an existing session.
.EXAMPLE
	.\O365DC.v4.ps1 -JobCount_Exchange 12
	This results in O365DC using 12 active Exchange jobs instead of the default of 10.
.EXAMPLE
	.\O365DC.v4.ps1 -JobPolling_Exchange 30
	This results in O365DC polling for completed Exchange jobs every 30 seconds.
.EXAMPLE
	.\O365DC.v4.ps1 -Timeout_Exchange_Job 7200
	This results in O365DC killing Exchange jobs that have exceeded 7200 seconds at the next polling interval.
.EXAMPLE
	.\O365DC.v4.ps1 -INI_Server ".\Templates\Template1_INI_Server.ini"
	This results in O365DC loading the specified template on start up.
.INPUTS
	None.
.OUTPUTS
	This script has no output objects.  O365DC creates txt, xml, docx, and xlsx output.
.NOTES
	NAME        :   O365DC.v4.ps1
	AUTHOR      :   Stemy Mynhier [MSFT]
	VERSION     :   4.0.2 build a1
	LAST EDIT   :   Sept-2018
.LINK
	https://gallery.technet.microsoft.com/office/
#>

Param(	[int]$JobCount_Exchange = 2,`
		[int]$JobPolling_Exchange = 5,`
		[int]$Timeout_Exchange_Job = 3600,`
		[string]$ServerForPSSession = $null,`
		[string]$INI_Exchange,`
		[switch]$NoEMS,`
		[switch]$MFA,`
		[switch]$ForceNewConnection)

function New-O365DCForm {
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null

#region *** Initialize Form ***

#region Main Form
$form1 = New-Object System.Windows.Forms.Form
$tab_Master = New-Object System.Windows.Forms.TabControl
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
$Menu_Main = new-object System.Windows.Forms.MenuStrip
$Menu_File = new-object System.Windows.Forms.ToolStripMenuItem('&File')
$Menu_Toggle = new-object System.Windows.Forms.ToolStripMenuItem('&Toggle')
$Menu_Help  = new-object System.Windows.Forms.ToolStripMenuItem('&Help')
$Submenu_LoadTargets = new-object System.Windows.Forms.ToolStripMenuItem('&Load all Targets from files')
$Submenu_PackageLogs = new-object System.Windows.Forms.ToolStripMenuItem('&Package application log')
$Submenu_Targets_CheckAll = new-object System.Windows.Forms.ToolStripMenuItem('&Check All Targets')
$Submenu_Targets_UnCheckAll = new-object System.Windows.Forms.ToolStripMenuItem('&Uncheck All Targets')
$Submenu_Tests_CheckAll = new-object System.Windows.Forms.ToolStripMenuItem('&Check All Tests')
$Submenu_Tests_UnCheckAll = new-object System.Windows.Forms.ToolStripMenuItem('&Uncheck All Tests')
$Submenu_Help = new-object System.Windows.Forms.ToolStripMenuItem('&Help')
$Submenu_About = new-object System.Windows.Forms.ToolStripMenuItem('&About')
#endregion Main Form

#region Step1 - Targets

#region Step1 Main
$tab_Step1 = New-Object System.Windows.Forms.TabPage
$btn_Step1_Discover = New-Object System.Windows.Forms.Button
$btn_Step1_Populate = New-Object System.Windows.Forms.Button
$tab_Step1_Master = New-Object System.Windows.Forms.TabControl
$status_Step1 = New-Object System.Windows.Forms.StatusBar
#endregion Step1 Main

#region Step1 Mailboxes Tab
$tab_Step1_Mailboxes = New-Object System.Windows.Forms.TabPage
$bx_Mailboxes_List = New-Object System.Windows.Forms.GroupBox
$btn_Step1_Mailboxes_Discover = New-Object System.Windows.Forms.Button
$btn_Step1_Mailboxes_Populate = New-Object System.Windows.Forms.Button
$clb_Step1_Mailboxes_List = New-Object system.Windows.Forms.CheckedListBox
$txt_MailboxesTotal = New-Object System.Windows.Forms.TextBox
$btn_Step1_Mailboxes_CheckAll = New-Object System.Windows.Forms.Button
$btn_Step1_Mailboxes_UncheckAll = New-Object System.Windows.Forms.Button
#endregion Step1 Mailboxes Tab

#region Step1 AzureAdUser Tab
$tab_Step1_AzureAdUser = New-Object System.Windows.Forms.TabPage
$bx_AzureAdUser_List = New-Object System.Windows.Forms.GroupBox
$btn_Step1_AzureAdUser_Discover = New-Object System.Windows.Forms.Button
$btn_Step1_AzureAdUser_Populate = New-Object System.Windows.Forms.Button
$clb_Step1_AzureAdUser_List = New-Object system.Windows.Forms.CheckedListBox
$txt_AzureAdUserTotal = New-Object System.Windows.Forms.TextBox
$btn_Step1_AzureAdUser_CheckAll = New-Object System.Windows.Forms.Button
$btn_Step1_AzureAdUser_UncheckAll = New-Object System.Windows.Forms.Button
#endregion Step1 AzureAdUser Tab

#endregion Step1 - Targets

#region Step2 - Templates
$tab_Step2 = New-Object System.Windows.Forms.TabPage
$bx_Step2_Templates = New-Object System.Windows.Forms.GroupBox
$rb_Step2_Template_1 = New-Object System.Windows.Forms.RadioButton
$rb_Step2_Template_2 = New-Object System.Windows.Forms.RadioButton
$rb_Step2_Template_3 = New-Object System.Windows.Forms.RadioButton
$rb_Step2_Template_4 = New-Object System.Windows.Forms.RadioButton
$Status_Step2 = New-Object System.Windows.Forms.StatusBar
#endregion Step2 - Templates

#Region Step3 - Tests

#region Step3 Main Tier1
$tab_Step3 = New-Object System.Windows.Forms.TabPage
$tab_Step3_Master = New-Object System.Windows.Forms.TabControl
$status_Step3 = New-Object System.Windows.Forms.StatusBar
$lbl_Step3_Execute = New-Object System.Windows.Forms.Label
$btn_Step3_Execute = New-Object System.Windows.Forms.Button
#endregion Step3 Main Tier1

#region Step3 Exchange Tier2
$tab_Step3_Exchange = New-Object System.Windows.Forms.TabPage
$tab_Step3_Exchange_Tier2 = New-Object System.Windows.Forms.TabControl
#endregion Step3 Exchange Tier2

#region Step3 Azure Tier2
$tab_Step3_Azure = New-Object System.Windows.Forms.TabPage
$tab_Step3_Azure_Tier2 = New-Object System.Windows.Forms.TabControl
#endregion Step3 Azure Tier2

#region Step3 Sharepoint Tier2
$tab_Step3_Sharepoint = New-Object System.Windows.Forms.TabPage
$tab_Step3_Sharepoint_Tier2 = New-Object System.Windows.Forms.TabControl
#endregion Step3 Sharepoint Tier2

#region Step3 Skype Tier2
$tab_Step3_Skype = New-Object System.Windows.Forms.TabPage
$tab_Step3_Skype_Tier2 = New-Object System.Windows.Forms.TabControl
#endregion Step3 Skype Tier2

#region Step3 Exchange tabs

#region Step3 Exchange - Client Access tab
$tab_Step3_ClientAccess = New-Object System.Windows.Forms.TabPage
$bx_ClientAccess_Functions = New-Object System.Windows.Forms.GroupBox
$btn_Step3_ClientAccess_CheckAll = New-Object System.Windows.Forms.Button
$btn_Step3_ClientAccess_UncheckAll = New-Object System.Windows.Forms.Button
$chk_Org_Get_ActiveSyncOrgSettings = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_MobileDevice = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_MobileDevicePolicy = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_AvailabilityAddressSpace = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_OWAMailboxPolicy = New-Object System.Windows.Forms.CheckBox

#endregion Step3 Client Access tab

#region Step3 Exchange - Global tab
$tab_Step3_Global = New-Object System.Windows.Forms.TabPage
$bx_Global_Functions = New-Object System.Windows.Forms.GroupBox
$btn_Step3_Global_CheckAll = New-Object System.Windows.Forms.Button
$btn_Step3_Global_UncheckAll = New-Object System.Windows.Forms.Button
$chk_Org_Get_AddressBookPolicy  = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_AddressList  = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_AntiPhishPolicy = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_AntiSpoofingPolicy = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_AtpPolicyForO365 = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_EmailAddressPolicy = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_GlobalAddressList = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_OfflineAddressBook = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_OnPremisesOrganization = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_OrgConfig = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_Rbac = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_RetentionPolicy = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_RetentionPolicyTag = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_SmimeConfig = New-Object System.Windows.Forms.CheckBox
#endregion Step3 Global tab

#region Step3 Exchange - Recipient Tab
$tab_Step3_Recipient = New-Object System.Windows.Forms.TabPage
$bx_Recipient_Functions = New-Object System.Windows.Forms.GroupBox
$btn_Step3_Recipient_CheckAll = New-Object System.Windows.Forms.Button
$btn_Step3_Recipient_UncheckAll = New-Object System.Windows.Forms.Button
$chk_Org_Get_CalendarProcessing = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_CASMailbox = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_CASMailboxPlan = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_Contact = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_DistributionGroup = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_DynamicDistributionGroup = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_Mailbox = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_MailboxFolderStatistics = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_MailboxPermission = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_MailboxPlan = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_MailboxStatistics = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_MailUser = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_PublicFolder = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_PublicFolderStatistics = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_UnifiedGroup = New-Object System.Windows.Forms.CheckBox
$chk_Org_Quota = New-Object System.Windows.Forms.CheckBox
#endregion Step3 Recipient Tab

#region Step3 Exchange - Transport Tab
$tab_Step3_Transport = New-Object System.Windows.Forms.TabPage
$bx_Transport_Functions = New-Object System.Windows.Forms.GroupBox
$btn_Step3_Transport_CheckAll = New-Object System.Windows.Forms.Button
$btn_Step3_Transport_UncheckAll = New-Object System.Windows.Forms.Button
$chk_Org_Get_AcceptedDomain = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_DkimSigningConfig = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_InboundConnector = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_RemoteDomain = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_OutboundConnector = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_TransportConfig = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_TransportRule = New-Object System.Windows.Forms.CheckBox
#endregion Step3 Transport Tab

#region Step3 Exchange - Unified Messaging tab
$tab_Step3_UM = New-Object System.Windows.Forms.TabPage
$bx_UM_Functions = New-Object System.Windows.Forms.GroupBox
$btn_Step3_UM_CheckAll = New-Object System.Windows.Forms.Button
$btn_Step3_UM_UncheckAll = New-Object System.Windows.Forms.Button
$chk_Org_Get_UmAutoAttendant = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_UmDialPlan = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_UmIpGateway = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_UmMailbox = New-Object System.Windows.Forms.CheckBox
#$chk_Org_Get_UmMailboxConfiguration = New-Object System.Windows.Forms.CheckBox
#$chk_Org_Get_UmMailboxPin = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_UmMailboxPolicy = New-Object System.Windows.Forms.CheckBox
#endregion Step3 Unified Messaging tab

#region Step3 Exchange - Misc Tab
$tab_Step3_Misc = New-Object System.Windows.Forms.TabPage
$bx_Misc_Functions = New-Object System.Windows.Forms.GroupBox
$btn_Step3_Misc_CheckAll = New-Object System.Windows.Forms.Button
$btn_Step3_Misc_UncheckAll = New-Object System.Windows.Forms.Button
$chk_Org_Get_AdminGroups = New-Object System.Windows.Forms.CheckBox
#endregion Step3 Misc Tab
#endregion Step3 Exchange tabs

#region Step3 Azure tabs

#region Step3 Azure - AzureAD Tab
$tab_Step3_AzureAD = New-Object System.Windows.Forms.TabPage
$bx_AzureAD_Functions = New-Object System.Windows.Forms.GroupBox
$btn_Step3_AzureAD_CheckAll = New-Object System.Windows.Forms.Button
$btn_Step3_AzureAD_UncheckAll = New-Object System.Windows.Forms.Button
$chk_Azure_Get_AzureADApplication = New-Object System.Windows.Forms.CheckBox
$chk_Azure_Get_AzureADContact = New-Object System.Windows.Forms.CheckBox
$chk_Azure_Get_AzureADDevice = New-Object System.Windows.Forms.CheckBox
$chk_Azure_Get_AzureADDeviceRegisteredOwner = New-Object System.Windows.Forms.CheckBox
$chk_Azure_Get_AzureADDeviceRegisteredUser = New-Object System.Windows.Forms.CheckBox
$chk_Azure_Get_AzureADDirectoryRole = New-Object System.Windows.Forms.CheckBox
$chk_Azure_Get_AzureADDomain = New-Object System.Windows.Forms.CheckBox
$chk_Azure_Get_AzureAdDomainServiceConfigurationRecord = New-Object System.Windows.Forms.CheckBox
$chk_Azure_Get_AzureAdDomainVerificationDnsRecord = New-Object System.Windows.Forms.CheckBox
$chk_Azure_Get_AzureADGroup = New-Object System.Windows.Forms.CheckBox
$chk_Azure_Get_AzureADGroupMember = New-Object System.Windows.Forms.CheckBox
$chk_Azure_Get_AzureADGroupOwner = New-Object System.Windows.Forms.CheckBox
$chk_Azure_Get_AzureADSubscribedSku = New-Object System.Windows.Forms.CheckBox
$chk_Azure_Get_AzureADTenantDetail = New-Object System.Windows.Forms.CheckBox
$chk_Azure_Get_AzureADUser = New-Object System.Windows.Forms.CheckBox
$chk_Azure_Get_AzureADUserLicenseDetail = New-Object System.Windows.Forms.CheckBox
$chk_Azure_Get_AzureADUserMembership = New-Object System.Windows.Forms.CheckBox
$chk_Azure_Get_AzureADUserOwnedDevice = New-Object System.Windows.Forms.CheckBox
$chk_Azure_Get_AzureADUserRegisteredDevice = New-Object System.Windows.Forms.CheckBox
#endregion Step3 Azure - AzureAD Tab

#endregion Step3 Azure tabs

#region Step3 Sharepoint tabs

#region Step3 Sharepoint - SPO Tab
$tab_Step3_SPO = New-Object System.Windows.Forms.TabPage
$bx_SPO_Functions = New-Object System.Windows.Forms.GroupBox
$btn_Step3_SPO_CheckAll = New-Object System.Windows.Forms.Button
$btn_Step3_SPO_UncheckAll = New-Object System.Windows.Forms.Button
$chk_Org_Get_SPODeletedSite = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_SPOExternalUser = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_SPOGeoStorageQuota = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_SPOSite = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_SPOTenant = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_SPOTenantSyncClientRestriction = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_SPOUser = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_SPOWebTemplate = New-Object System.Windows.Forms.CheckBox
#endregion Step3 Sharepoint - SPO Tab

#endregion Step3 Sharepoint tabs

#region Step3 Skype tabs

#region Step3 Skype - Misc Tab
#$tab_Step3_Misc = New-Object System.Windows.Forms.TabPage
#$bx_Misc_Functions = New-Object System.Windows.Forms.GroupBox
#$btn_Step3_Misc_CheckAll = New-Object System.Windows.Forms.Button
#$btn_Step3_Misc_UncheckAll = New-Object System.Windows.Forms.Button
#$chk_Org_Get_AdminGroups = New-Object System.Windows.Forms.CheckBox
#endregion Step3 Skype - Misc Tab

#endregion Step3 Skype tabs

#EndRegion Step3 - Tests

#region Step4 - Reporting
$tab_Step4 = New-Object System.Windows.Forms.TabPage
$btn_Step4_Assemble = New-Object System.Windows.Forms.Button
$lbl_Step4_Assemble = New-Object System.Windows.Forms.Label
$bx_Step4_Functions = New-Object System.Windows.Forms.GroupBox
#$chk_Step4_DC_Report = New-Object System.Windows.Forms.CheckBox
#$chk_Step4_Ex_Report = New-Object System.Windows.Forms.CheckBox
$chk_Step4_Exchange_Report = New-Object System.Windows.Forms.CheckBox
$chk_Step4_Azure_Report = New-Object System.Windows.Forms.CheckBox
#$chk_Step4_Exchange_Environment_Doc = New-Object System.Windows.Forms.CheckBox
$Status_Step4 = New-Object System.Windows.Forms.StatusBar
#endregion Step4 - Reporting

#region Step5 - Having Trouble?
#$tab_Step5 = New-Object System.Windows.Forms.TabPage
#$bx_Step5_Functions = New-Object System.Windows.Forms.GroupBox
#$Status_Step5 = New-Object System.Windows.Forms.StatusBar
#endregion Step5 - Having Trouble?

#endregion *** Initialize Form ***

#region *** Events ***

#region "Main Menu" Events
$handler_Submenu_LoadTargets=
{
	Import-TargetsMailboxes
}

$handler_Submenu_PackageLogs=
{
	.\O365DC_Scripts\Core_Package_Logs.ps1 -RunLocation $location
}

$handler_Submenu_Targets_CheckAll=
{
	Enable-TargetsMailbox
}

$handler_Submenu_Targets_UnCheckAll=
{
	Disable-TargetsMailbox
}

$handler_Submenu_Tests_CheckAll=
{
	# Exchange Functions - All
	Set-AllFunctionsClientAccess -Check $true
	Set-AllFunctionsGlobal -Check $true
	Set-AllFunctionsRecipient -Check $true
	Set-AllFunctionsTransport -Check $true
	Set-AllFunctionsMisc -Check $true
	Set-AllFunctionsUm -Check $true
	Set-AllFunctionsAzureAd -Check $true
	Set-AllFunctionsSpo -Check $true
}

$handler_Submenu_Tests_UnCheckAll=
{
	# Exchange Functions - All
	Set-AllFunctionsClientAccess -Check $False
	Set-AllFunctionsGlobal -Check $False
	Set-AllFunctionsRecipient -Check $False
	Set-AllFunctionsTransport -Check $False
	Set-AllFunctionsMisc -Check $False
	Set-AllFunctionsUm -Check $False
	Set-AllFunctionsAzureAd -Check $False
	Set-AllFunctionsSpo -Check $False
}

$handler_Submenu_Help=
{
	$Message_Help = "Would you like to open the Help document?"
	$Title_Help = "O365DC Help"
	$MessageBox_Help = [Windows.Forms.MessageBox]::Show($Message_Help, $Title_Help, [Windows.Forms.MessageBoxButtons]::YesNo, [Windows.Forms.MessageBoxIcon]::Information)
	if ($MessageBox_Help -eq [Windows.Forms.DialogResult]::Yes)
	{
		$ie = New-Object -ComObject "InternetExplorer.Application"
		$ie.visible = $true
		$ie.navigate((get-location).path + "\Help\Documentation_O365DC.v.4.mht")
	}
}

$handler_Submenu_About=
{
	$Message_About = ""
	$Message_About = "Office 365 Data Collector `n`n"
	$Message_About = $Message_About += "Version: 4.0.2 Build a1 `n`n"
	$Message_About = $Message_About += "Release Date: October 2018 `n`n"
	$Message_About = $Message_About += "Written by: Stemy Mynhier`nstemy@microsoft.com `n`n"
	$Message_About = $Message_About += "This script is provided AS IS with no warranties, and confers no rights.  "
	$Message_About = $Message_About += "Use of any portion or all of this script are subject to the terms specified at https://www.microsoft.com/en-us/legal/intellectualproperty/copyright/default.aspx."
	$Title_About = "About Office 365 Data Collector (O365DC)"
	[Windows.Forms.MessageBox]::Show($Message_About, $Title_About, [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Information)
}
#endregion "Main Menu" Events

#region "Step1 - Targets" Events
$handler_btn_Step1_Mailboxes_Discover=
{
	Disable-AllTargetsButtons
    $status_Step1.Text = "Step 1 Status: Running"
	$EventLog = New-Object System.Diagnostics.EventLog('Application')
	$EventLog.MachineName = "."
	$EventLog.Source = "O365DC"
	try{$EventLog.WriteEntry("Starting O365DC Step 1 - Discover mailboxes","Information", 10)} catch{}
	$Mailbox_outputfile = ".\Mailbox.txt"
	if ((Test-Path ".\mailbox.txt") -eq $true)
	{
	    $status_Step1.Text = "Step 1 Status: Failed - mailbox.txt already present.  Please remove and rerun or select Populate."
		write-host "Mailbox.txt is already present in this folder." -ForegroundColor Red
		write-host "Loading values from text file that is present." -ForegroundColor Red
	}
	else
	{
		New-Item $Mailbox_outputfile -type file -Force
	    $MailboxList = @()
		get-mailbox -resultsize unlimited | where-object {$_.RecipientTypeDetails -ne "DiscoveryMailbox"} | ForEach-Object `
		{
			$MailboxList += $_.alias
		}

	    $MailboxListSorted = $MailboxList | sort-object
		$MailboxListSorted | out-file $Mailbox_outputfile -append
		$status_Step1.Text = "Step 1 Status: Idle"
	}
    $File_Location = $location + "\mailbox.txt"
	if ((Test-Path $File_Location) -eq $true)
	{
	    $array_Mailboxes = @(([System.IO.File]::ReadAllLines($File_Location)) | sort-object -Unique)
		$intMailboxTotal = 0
		$clb_Step1_Mailboxes_List.items.clear()
	    foreach ($member_Mailbox in $array_Mailboxes | where-object {$_ -ne ""})
	    {
	        $clb_Step1_Mailboxes_List.items.add($member_Mailbox)
			$intMailboxTotal++
	    }
		For ($i=0;$i -le ($intMailboxTotal - 1);$i++)
		{
			$clb_Step1_Mailboxes_List.SetItemChecked($i,$true)
		}
		$txt_MailboxesTotal.Text = "Mailbox count = " + $intMailboxTotal
		$txt_MailboxesTotal.visible = $true
	}
	else
	{
		write-host	"The file mailbox.txt is not present.  Run Discover to create the file."
		$status_Step1.Text = "Step 1 Status: Failed - mailbox.txt file not found.  Run Discover to create the file."
	}
	$EventLog = New-Object System.Diagnostics.EventLog('Application')
	$EventLog.MachineName = "."
	$EventLog.Source = "O365DC"
	try{$EventLog.WriteEntry("Ending O365DC Step 1 - Discover mailboxes","Information", 11)} catch{}
	Enable-AllTargetsButtons
}

$handler_btn_Step1_Mailboxes_Populate=
{
	 Import-TargetsMailboxes
}

$handler_btn_Step1_Mailboxes_CheckAll=
{
	Enable-TargetsMailbox
}

$handler_btn_Step1_Mailboxes_UncheckAll=
{
	Disable-TargetsMailbox
}

$handler_btn_Step1_AzureAdUser_Discover=
{
	Disable-AllTargetsButtons
    $status_Step1.Text = "Step 1 Status: Running"
	$EventLog = New-Object System.Diagnostics.EventLog('Application')
	$EventLog.MachineName = "."
	$EventLog.Source = "O365DC"
	try{$EventLog.WriteEntry("Starting O365DC Step 1 - Discover AzureAdUser","Information", 10)} catch{}
	$AzureAdUser_outputfile = ".\AzureAdUser.txt"
	if ((Test-Path ".\AzureAdUser.txt") -eq $true)
	{
	    $status_Step1.Text = "Step 1 Status: Failed - AzureAdUser.txt already present.  Please remove and rerun or select Populate."
		write-host "AzureAdUser.txt is already present in this folder." -ForegroundColor Red
		write-host "Loading values from text file that is present." -ForegroundColor Red
	}
	else
	{
		New-Item $AzureAdUser_outputfile -type file -Force
	    $AzureAdUserList = @()
		get-AzureAdUser | ForEach-Object `
		{
			$AzureAdUserList += $_.UserPrincipalName
		}

	    $AzureAdUserListSorted = $AzureAdUserList | sort-object
		$AzureAdUserListSorted | out-file $AzureAdUser_outputfile -append
		$status_Step1.Text = "Step 1 Status: Idle"
	}
    $File_Location = $location + "\AzureAdUser.txt"
	if ((Test-Path $File_Location) -eq $true)
	{
	    $array_AzureAdUser = @(([System.IO.File]::ReadAllLines($File_Location)) | sort-object -Unique)
		$intAzureAdUserTotal = 0
		$clb_Step1_AzureAdUser_List.items.clear()
	    foreach ($member_AzureAdUser in $array_AzureAdUser | where-object {$_ -ne ""})
	    {
	        $clb_Step1_AzureAdUser_List.items.add($member_AzureAdUser)
			$intAzureAdUserTotal++
	    }
		For ($i=0;$i -le ($intAzureAdUserTotal - 1);$i++)
		{
			$clb_Step1_AzureAdUser_List.SetItemChecked($i,$true)
		}
		$txt_AzureAdUserTotal.Text = "AzureAdUser count = " + $intAzureAdUserTotal
		$txt_AzureAdUserTotal.visible = $true
	}
	else
	{
		write-host	"The file AzureAdUser.txt is not present.  Run Discover to create the file."
		$status_Step1.Text = "Step 1 Status: Failed - AzureAdUser.txt file not found.  Run Discover to create the file."
	}
	$EventLog = New-Object System.Diagnostics.EventLog('Application')
	$EventLog.MachineName = "."
	$EventLog.Source = "O365DC"
	try{$EventLog.WriteEntry("Ending O365DC Step 1 - Discover AzureAdUser","Information", 11)} catch{}
	Enable-AllTargetsButtons
}

$handler_btn_Step1_AzureAdUser_Populate=
{
	 Import-TargetsAzureAdUser
}

$handler_btn_Step1_AzureAdUser_CheckAll=
{
	Enable-TargetsAzureAdUser
}

$handler_btn_Step1_AzureAdUser_UncheckAll=
{
	Disable-TargetsAzureAdUser
}

#endregion "Step1 - Targets" Events

#Region "Step2" Events
$handler_rb_Step2_Template_1=
{
	# Uncheck all other radio buttons
	$rb_Step2_Template_2.Checked = $false
	$rb_Step2_Template_3.Checked = $false
	$rb_Step2_Template_4.Checked = $false
	#Load the templates
	if ($NoEMS -eq $false)
	{
		try{& ".\O365DC_Scripts\Core_Parse_Ini_File.ps1" -IniFile ".\Templates\Template1_INI_Exchange.ini"} catch{}
	}
}

$handler_rb_Step2_Template_2=
{
	# Uncheck all other radio buttons
	$rb_Step2_Template_1.Checked = $false
	$rb_Step2_Template_3.Checked = $false
	$rb_Step2_Template_4.Checked = $false
	#Load the templates
	if ($NoEMS -eq $false)
	{
		try{& ".\O365DC_Scripts\Core_Parse_Ini_File.ps1" -IniFile ".\Templates\Template2_INI_Exchange.ini"} catch{}
	}
}

$handler_rb_Step2_Template_3=
{
	# Uncheck all other radio buttons
	$rb_Step2_Template_1.Checked = $false
	$rb_Step2_Template_2.Checked = $false
	$rb_Step2_Template_4.Checked = $false
	# Since this is the Environmental Doc template, warn if no EMS
	if ($NoEMS -eq $true)
	{
		write-host "This template is designed to run with the Exchange cmdlets.  NoEMS switch detected." -foregroundcolor yellow
		write-host "Data collection will be incomplete." -foregroundcolor yellow
	}
	#Load the templates
	if ($NoEMS -eq $false)
	{
		try{& ".\O365DC_Scripts\Core_Parse_Ini_File.ps1" -IniFile ".\Templates\Template3_INI_Exchange.ini"} catch {}
	}
}

$handler_rb_Step2_Template_4=
{
	# Uncheck all other radio buttons
	$rb_Step2_Template_1.Checked = $false
	$rb_Step2_Template_2.Checked = $false
	$rb_Step2_Template_3.Checked = $false
	#Load the templates
	if ($NoEMS -eq $false)
	{
		try{& ".\O365DC_Scripts\Core_Parse_Ini_File.ps1" -IniFile ".\Templates\Template4_INI_Exchange.ini"} catch{}
	}
}

#Endregion "Step2" Events

#region "Step3 - Tests" Events
$handler_btn_Step3_Execute_Click=
{
	try
	{
		Start-Transcript -Path (".\O365DC_Step3_Transcript_" + $append + ".txt")
	}
	catch [System.Management.Automation.CmdletInvocationException]
	{
		write-host "Transcription already started" -ForegroundColor red
		write-host "Restarting transcription" -ForegroundColor red
		Stop-Transcript
		Start-Transcript -Path (".\O365DC_Step3_Transcript_" + $append + ".txt")
	}
	$btn_Step3_Execute.enabled = $false
	$status_Step3.Text = "Step 3 Status: Running"
	$EventLog = New-Object System.Diagnostics.EventLog('Application')
	$EventLog.MachineName = "."
	$EventLog.Source = "O365DC"
	$EventLogText = "Starting O365DC Step 3`nMailboxes: $intMailboxTotal`nAzureAD Users: $intAzureADUserTotal"
	try{$EventLog.WriteEntry($EventLogText,"Information", 30)} catch{}
	#send the form to the back to expose the Powershell window when starting Step 3
	$form1.WindowState = "minimized"
	write-host "O365DC Form minimized." -ForegroundColor Green

	#Region Executing Exchange Organization Tests
	write-host "Starting Exchange Organization..." -ForegroundColor Green
	If (Get-ExchangeBoxStatus = $true)
	{
		# Save checked mailboxes to file for use by jobs
		$Mailbox_Checked_outputfile = ".\CheckedMailbox.txt"
		if ((Test-Path $Mailbox_Checked_outputfile) -eq $true)
		{
			Remove-Item $Mailbox_Checked_outputfile -Force
		}
		write-host "-- Building the checked mailbox list..."
		foreach ($item in $clb_Step1_Mailboxes_List.checkeditems)
		{
			$item.tostring() | out-file $Mailbox_Checked_outputfile -append -Force
		}

		If (Get-ExchangeMbxBoxStatus = $true)
		{
			# Avoid this path if we're not running mailbox tests
			# Splitting CheckedMailboxes file 10 times
			Split-List10 -InputFile "CheckedMailbox" -OutputFile "CheckedMailbox" -Text "Mailbox"

			# Old Code
			<#
			write-host "-- Splitting the list of checked mailboxes... "
			$File_Location = $location + "\CheckedMailbox.txt"
			If ((Test-Path $File_Location) -eq $false)
			{
				# Create empty Mailbox.txt file if not present
				write-host "No mailboxes appear to be selected.  Mailbox tests will produce no output." -ForegroundColor Red
				"" | Out-File $File_Location
			}
			$CheckedMailbox = [System.IO.File]::ReadAllLines($File_Location)
			$CheckedMailboxCount = $CheckedMailbox.count
			$CheckedMailboxCountSplit = [int]$CheckedMailboxCount/10
			if ((Test-Path ".\CheckedMailbox.Set1.txt") -eq $true) {Remove-Item ".\CheckedMailbox.Set1.txt" -Force}
			if ((Test-Path ".\CheckedMailbox.Set2.txt") -eq $true) {Remove-Item ".\CheckedMailbox.Set2.txt" -Force}
			if ((Test-Path ".\CheckedMailbox.Set3.txt") -eq $true) {Remove-Item ".\CheckedMailbox.Set3.txt" -Force}
			if ((Test-Path ".\CheckedMailbox.Set4.txt") -eq $true) {Remove-Item ".\CheckedMailbox.Set4.txt" -Force}
			if ((Test-Path ".\CheckedMailbox.Set5.txt") -eq $true) {Remove-Item ".\CheckedMailbox.Set5.txt" -Force}
			if ((Test-Path ".\CheckedMailbox.Set6.txt") -eq $true) {Remove-Item ".\CheckedMailbox.Set6.txt" -Force}
			if ((Test-Path ".\CheckedMailbox.Set7.txt") -eq $true) {Remove-Item ".\CheckedMailbox.Set7.txt" -Force}
			if ((Test-Path ".\CheckedMailbox.Set8.txt") -eq $true) {Remove-Item ".\CheckedMailbox.Set8.txt" -Force}
			if ((Test-Path ".\CheckedMailbox.Set9.txt") -eq $true) {Remove-Item ".\CheckedMailbox.Set9.txt" -Force}
			if ((Test-Path ".\CheckedMailbox.Set10.txt") -eq $true) {Remove-Item ".\CheckedMailbox.Set10.txt" -Force}
			For ($Count = 0;$Count -lt ($CheckedMailboxCountSplit);$Count++)
				{$CheckedMailbox[$Count] | Out-File ".\CheckedMailbox.Set1.txt" -Append -Force}
			For (;$Count -lt (2*$CheckedMailboxCountSplit);$Count++)
				{$CheckedMailbox[$Count] | Out-File ".\CheckedMailbox.Set2.txt" -Append -Force}
			For (;$Count -lt (3*$CheckedMailboxCountSplit);$Count++)
				{$CheckedMailbox[$Count] | Out-File ".\CheckedMailbox.Set3.txt" -Append -Force}
			For (;$Count -lt (4*$CheckedMailboxCountSplit);$Count++)
				{$CheckedMailbox[$Count] | Out-File ".\CheckedMailbox.Set4.txt" -Append -Force}
			For (;$Count -lt (5*$CheckedMailboxCountSplit);$Count++)
				{$CheckedMailbox[$Count] | Out-File ".\CheckedMailbox.Set5.txt" -Append -Force}
			For (;$Count -lt (6*$CheckedMailboxCountSplit);$Count++)
				{$CheckedMailbox[$Count] | Out-File ".\CheckedMailbox.Set6.txt" -Append -Force}
			For (;$Count -lt (7*$CheckedMailboxCountSplit);$Count++)
				{$CheckedMailbox[$Count] | Out-File ".\CheckedMailbox.Set7.txt" -Append -Force}
			For (;$Count -lt (8*$CheckedMailboxCountSplit);$Count++)
				{$CheckedMailbox[$Count] | Out-File ".\CheckedMailbox.Set8.txt" -Append -Force}
			For (;$Count -lt (9*$CheckedMailboxCountSplit);$Count++)
				{$CheckedMailbox[$Count] | Out-File ".\CheckedMailbox.Set9.txt" -Append -Force}
			For (;$Count -lt (10*$CheckedMailboxCountSplit);$Count++)
				{$CheckedMailbox[$Count] | Out-File ".\CheckedMailbox.Set10.txt" -Append -Force}

			#>
		}

		# First we start the jobs that query the organization instead of the Exchange server
		#Region Exchange Non-server Functions
		If ($chk_Org_Get_AcceptedDomain.checked -eq $true)
		{
			write-host "Starting Get-AcceptedDomain" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetAcceptedDomain.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_ActiveSyncOrgSettings.checked -eq $true)
		{
			write-host "Starting Get-ActiveSyncOrgSettings" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetActiveSyncOrgSettings.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_AddressBookPolicy.checked -eq $true)
		{
			write-host "Starting Get-AddressBookPolicy" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetAddressBookPolicy.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_AddressList.checked -eq $true)
		{
			write-host "Starting Get-AddressList" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetAddressList.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_AdminGroups.checked -eq $true)
		{
			write-host "Starting Get-AdminGroups" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_Misc_AdminGroups.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_AntiPhishPolicy.checked -eq $true)
		{
			write-host "Starting Get-AntiPhishPolicy" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetAntiPhishPolicy.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_AntiSpoofingPolicy.checked -eq $true)
		{
			write-host "Starting Get-AntiSpoofingPolicy" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetAntiSpoofingPolicy.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_AtpPolicyForO365.checked -eq $true)
		{
			write-host "Starting Get-AtpPolicyForO365" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetAtpPolicyForO365.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_AvailabilityAddressSpace.checked -eq $true)
		{
			write-host "Starting Get-AvailabilityAddressSpace" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetAvailabilityAddressSpace.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_CalendarProcessing.checked -eq $true)
		{
			For ($i = 1;$i -lt 11;$i++)
			{
				write-host "Starting Get-CalendarProcessing job $i" -foregroundcolor green
				try
					{.\O365DC_Scripts\Exchange_GetCalendarProcessing.ps1 -location $location -i $i}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		}
		If ($chk_Org_Get_CASMailbox.checked -eq $true)
		{
			For ($i = 1;$i -lt 11;$i++)
			{
				write-host "Starting Get-CASMailbox job $i" -foregroundcolor green
				try
					{.\O365DC_Scripts\Exchange_GetCASMailbox.ps1 -location $location -i $i}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		}
		If ($chk_Org_Get_CasMailboxPlan.checked -eq $true)
		{
			write-host "Starting Get-CasMailboxPlan" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetCasMailboxPlan.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_Contact.checked -eq $true)
		{
			write-host "Starting Get-Contact" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetContact.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_DistributionGroup.checked -eq $true)
		{
			write-host "Starting Get-DistributionGroup" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetDistributionGroup.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
	If ($chk_Org_Get_DkimSigningConfig.checked -eq $true)
		{
			write-host "Starting Get-DkimSigningConfig" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetDkimSigningConfig.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_DynamicDistributionGroup.checked -eq $true)
		{
			write-host "Starting Get-DynamicDistributionGroup" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetDynamicDistributionGroup.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_EmailAddressPolicy.checked -eq $true)
		{
			write-host "Starting Get-EmailAddressPolicy" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetEmailAddressPolicy.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_GlobalAddressList.checked -eq $true)
		{
			write-host "Starting Get-GlobalAddressList job" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetGlobalAddressList.ps1 -location $location -i $i}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_Mailbox.checked -eq $true)
		{
			For ($i = 1;$i -lt 11;$i++)
			{
				write-host "Starting Get-Mailbox job $i" -foregroundcolor green
				try
					{.\O365DC_Scripts\Exchange_GetMbx.ps1 -location $location -i $i}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		}
		If ($chk_Org_Get_MailboxPlan.checked -eq $true)
		{
			write-host "Starting Get-MailboxPlan" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetMailboxPlan.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_MailUser.checked -eq $true)
		{
			write-host "Starting Get-MailUser" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetMailUser.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_InboundConnector.checked -eq $true)
		{
			write-host "Starting Get-InboundConnector" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetInboundConnector.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_MailboxFolderStatistics.checked -eq $true)
		{
			For ($i = 1;$i -lt 11;$i++)
			{
				write-host "Starting Get-MailboxFolderStatistics job $i" -foregroundcolor green
				try
					{.\O365DC_Scripts\Exchange_GetMbxFolderStatistics.ps1 -location $location -i $i}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		}
		If ($chk_Org_Get_MailboxPermission.checked -eq $true)
		{
			For ($i = 1;$i -lt 11;$i++)
			{
				write-host "Starting Get-MailboxPermission job$i" -foregroundcolor green
				try
					{.\O365DC_Scripts\Exchange_GetMbxPermission.ps1 -location $location -i $i}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		}
		If ($chk_Org_Get_MailboxStatistics.checked -eq $true)
		{
			For ($i = 1;$i -lt 11;$i++)
			{
				write-host "Starting Get-MailboxStatistics job $i" -foregroundcolor green
				try
					{.\O365DC_Scripts\Exchange_GetMbxStatistics.ps1 -location $location -i $i}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		}
		If ($chk_Org_Get_MobileDevice.checked -eq $true)
		{
			write-host "Starting Get-MobileDevice" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetMobileDevice.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_MobileDevicePolicy.checked -eq $true)
		{
			write-host "Starting Get-MobileDevicePolicy" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetMobileDeviceMbxPolicy.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_OfflineAddressBook.checked -eq $true)
		{
			write-host "Starting Get-OfflineAddressBook" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetOfflineAddressBook.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_OnPremisesOrganization.checked -eq $true)
		{
			write-host "Starting Get-OnPremisesOrganization" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetOnPremisesOrganization.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_OrgConfig.checked -eq $true)
		{
			write-host "Starting Get-OrgConfig" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetOrgConfig.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_OutboundConnector.checked -eq $true)
		{
			write-host "Starting Get-OutboundConnector" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetOutboundConnector.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_OwaMailboxPolicy.checked -eq $true)
		{
			write-host "Starting Get-OwaMailboxPolicy" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetOwaMailboxPolicy.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_PublicFolder.checked -eq $true)
		{
			write-host "Starting Get-PublicFolder" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetPublicFolder.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_PublicFolderStatistics.checked -eq $true)
		{
			write-host "Starting Get-PublicFolderStatistics" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetPublicFolderStats.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Quota.checked -eq $true)
		{
			For ($i = 1;$i -lt 11;$i++)
			{
				write-host "Starting Get-Quota job$i" -foregroundcolor green
				try
					{.\O365DC_Scripts\Exchange_Quota.ps1 -location $location -i $i}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		}
		If ($chk_Org_Get_Rbac.checked -eq $true)
		{
			write-host "Starting Get-Rbac" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetRbac.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_RemoteDomain.checked -eq $true)
		{
			write-host "Starting Get-RemoteDomain" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetRemoteDomain.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_RetentionPolicy.checked -eq $true)
		{
			write-host "Starting Get-RetentionPolicy" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetRetentionPolicy.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_RetentionPolicyTag.checked -eq $true)
		{
			write-host "Starting Get-RetentionPolicyTag" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetRetentionPolicyTag.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_SmimeConfig.checked -eq $true)
		{
			write-host "Starting Get-SmimeConfig" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetSmimeConfig.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_TransportConfig.checked -eq $true)
		{
			write-host "Starting Get-TransportConfig" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetTransportConfig.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_TransportRule.checked -eq $true)
		{
			write-host "Starting Get-TransportRule" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetTransportRule.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_UmAutoAttendant.checked -eq $true)
		{
			write-host "Starting Get-UmAutoAttendant" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetUmAutoAttendant.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_UmDialPlan.checked -eq $true)
		{
			write-host "Starting Get-UmDialPlan" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetUmDialPlan.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_UmIpGateway.checked -eq $true)
		{
			write-host "Starting Get-UmIpGateway" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetUmIpGateway.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_UmMailbox.checked -eq $true)
		{
			For ($i = 1;$i -lt 11;$i++)
			{
				write-host "Starting Get-UmMailbox job $i" -foregroundcolor green
				try
					{.\O365DC_Scripts\Exchange_GetUmMailbox.ps1 -location $location -i $i}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		}
#			If ($chk_Org_Get_UmMailboxConfiguration.checked -eq $true)
#			{
#				For ($i = 1;$i -lt 11;$i++)
#				{Start-O365DCJob -server $server -job "Get-UmMailboxConfiguration - Set $i" -JobType 0 -Location $location -JobScriptName "Exchange_GetUmMailboxConfiguration.ps1" -i $i -PSSession $session_0}
#			}
#			If ($chk_Org_Get_UmMailboxPin.checked -eq $true)
#			{
#				For ($i = 1;$i -lt 11;$i++)
#				{Start-O365DCJob -server $server -job "Get-UmMailboxPin - Set $i" -JobType 0 -Location $location -JobScriptName "Exchange_GetUmMailboxPin.ps1" -i $i -PSSession $session_0}
#			}
		If ($chk_Org_Get_UmMailboxPolicy.checked -eq $true)
		{
			write-host "Starting Get-UmMailboxPolicy" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetUmMailboxPolicy.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		If ($chk_Org_Get_UnifiedGroup.checked -eq $true)
		{
			write-host "Starting Get-UnifiedGroup" -foregroundcolor green
			try
				{.\O365DC_Scripts\Exchange_GetUnifiedGroup.ps1 -location $location}
			catch [System.Management.Automation.CommandNotFoundException]
				{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
		}
		#EndRegion Exchange Non-Server Functions
	}
	else
	{
		write-host "---- No Exchange Organization Functions selected"
	}
	#EndRegion Executing Exchange Organization Tests

	#Region Executing Azure Tests

	write-host "Starting Azure..." -ForegroundColor Green
	If (Get-AzureBoxStatus = $true)
	{
		# Save checked mailboxes to file for use by jobs
		$AzureAdUser_Checked_outputfile = ".\CheckedAzureAdUser.txt"
		if ((Test-Path $AzureAdUser_Checked_outputfile) -eq $true)
		{
			Remove-Item $AzureAdUser_Checked_outputfile -Force
		}
		write-host "-- Building the checked AzureAdUser list..."
		foreach ($item in $clb_Step1_AzureAdUser_List.checkeditems)
		{
			$item.tostring() | out-file $AzureAdUser_Checked_outputfile -append -Force
		}

		If (Get-AzureAdUserBoxStatus = $true)
		{
			# Avoid this path if we're not running mailbox tests
			# Splitting CheckedMailboxes file 10 times
			Split-List10 -InputFile "CheckedAzureAdUser" -OutputFile "CheckedAzureAdUser" -Text "AzureAdUser"

			#Region Azure Functions

			Test-CheckBoxAndRun -chkBox $chk_Azure_Get_AzureADApplication.Checked -Text "Get-AzureADApplication" -Script "Azure_GetAzureADApplication"
			Test-CheckBoxAndRun -chkBox $chk_Azure_Get_AzureADContact.Checked -Text "Get-AzureADContact" -Script "Azure_GetAzureADContact"
			Test-CheckBoxAndRun -chkBox $chk_Azure_Get_AzureADDevice.Checked -Text "Get-AzureADDevice" -Script "Azure_GetAzureADDevice"
			Test-CheckBoxAndRun -chkBox $chk_Azure_Get_AzureADDeviceRegisteredOwner.Checked -Text "Get-AzureADDeviceRegisteredOwner" -Script "Azure_GetAzureADDeviceRegisteredOwner"
			Test-CheckBoxAndRun -chkBox $chk_Azure_Get_AzureADDeviceRegisteredUser.Checked -Text "Get-AzureADDeviceRegisteredUser" -Script "Azure_GetAzureADDeviceRegisteredUser"
			Test-CheckBoxAndRun -chkBox $chk_Azure_Get_AzureADDirectoryRole.Checked -Text "Get-AzureADDirectoryRole" -Script "Azure_GetAzureADDirectoryRole"
			Test-CheckBoxAndRun -chkBox $chk_Azure_Get_AzureADDomain.Checked -Text "Get-AzureADDomain" -Script "Azure_GetAzureADDomain"
			Test-CheckBoxAndRun -chkBox $chk_Azure_Get_AzureADDomainServiceConfigurationRecord.Checked -Text "Get-AzureADDomainServiceConfigurationRecord" -Script "Azure_GetAzureADDomainServiceConfigurationRecord"
			Test-CheckBoxAndRun -chkBox $chk_Azure_Get_AzureADDomainVerificationDnsRecord.Checked -Text "Get-AzureADDomainVerificationDnsRecord" -Script "Azure_GetAzureADDomainVerificationDnsRecord"
			Test-CheckBoxAndRun -chkBox $chk_Azure_Get_AzureADGroup.Checked -Text "Get-AzureADGroup" -Script "Azure_GetAzureADGroup"
			Test-CheckBoxAndRun -chkBox $chk_Azure_Get_AzureADGroupMember.Checked -Text "Get-AzureADGroupMember" -Script "Azure_GetAzureADGroupMember"
			Test-CheckBoxAndRun -chkBox $chk_Azure_Get_AzureADGroupOwner.Checked -Text "Get-AzureADGroupOwner" -Script "Azure_GetAzureADGroupOwner"
			Test-CheckBoxAndRun -chkBox $chk_Azure_Get_AzureADSubscribedSku.Checked -Text "Get-AzureADSubscribedSku" -Script "Azure_GetAzureADSubscribedSku"
			Test-CheckBoxAndRun -chkBox $chk_Azure_Get_AzureADTenantDetail.Checked -Text "Get-AzureADTenantDetail" -Script "Azure_GetAzureADTenantDetail"
			Test-CheckBoxAndRun -chkBox $chk_Azure_Get_AzureAdUser.Checked -Text "Get-AzureADUser" -Script "Azure_GetAzureAdUser"
			Test-CheckBoxAndRun -chkBox $chk_Azure_Get_AzureADUserLicenseDetail.Checked -Text "Get-AzureADUserLicenseDetail" -Script "Azure_GetAzureADUserLicenseDetail"
			Test-CheckBoxAndRun -chkBox $chk_Azure_Get_AzureADUserMembership.Checked -Text "Get-AzureADUserMembership" -Script "Azure_GetAzureADUserMembership"
			Test-CheckBoxAndRun -chkBox $chk_Azure_Get_AzureADUserOwnedDevice.Checked -Text "Get-AzureADUserOwnedDevice" -Script "Azure_GetAzureADUserOwnedDevice"
			Test-CheckBoxAndRun -chkBox $chk_Azure_Get_AzureADUserRegisteredDevice.Checked -Text "Get-AzureADUserRegisteredDevice" -Script "Azure_GetAzureADUserRegisteredDevice"


		}
	}
	else
	{
		write-host "---- No Azure Functions selected"
	}
	#EndRegion Executing Azure Tests


	# Delay changing status to Idle until all jobs have finished
	Update-O365DCJobCount 1 15
	Remove-Item	".\RunningJobs.txt"
	# Remove Failed Jobs
	$colJobsFailed = @(Get-Job -State Failed)
	foreach ($objJobsFailed in $colJobsFailed)
	{
		if ($objJobsFailed.module -like "__DynamicModule*")
		{
			Remove-Job -Id $objJobsFailed.id
		}
		else
		{
            write-host "---- Failed job " $objJobsFailed.name -ForegroundColor Red
			$FailedJobOutput = ".\FailedJobs_" + $append + ".txt"
            if ((Test-Path $FailedJobOutput) -eq $false)
	        {
		      new-item $FailedJobOutput -type file -Force
	        }
	        "* * * * * * * * * * * * * * * * * * * *"  | Out-File $FailedJobOutput -Force -Append
            "Job Name: " + $objJobsFailed.name | Out-File $FailedJobOutput -Force -Append
	        "Job State: " + $objJobsFailed.state | Out-File $FailedJobOutput -Force	-Append
            if ($null -ne ($objJobsFailed.childjobs[0]))
            {
	           $objJobsFailed.childjobs[0].output | format-list | Out-File $FailedJobOutput -Force -Append
	           $objJobsFailed.childjobs[0].warning | format-list | Out-File $FailedJobOutput -Force -Append
	           $objJobsFailed.childjobs[0].error | format-list | Out-File $FailedJobOutput -Force -Append
			}
            $ErrorText = $objJobsFailed.name + "`n"
			$ErrorText += "Job failed"
			$ErrorLog = New-Object System.Diagnostics.EventLog('Application')
			$ErrorLog.MachineName = "."
			$ErrorLog.Source = "O365DC"
			Try{$ErrorLog.WriteEntry($ErrorText,"Error", 500)} catch{}
			Remove-Job -Id $objJobsFailed.id
		}
	}
	write-host "Restoring O365DC Form to normal." -ForegroundColor Green
	$form1.WindowState = "normal"
	$btn_Step3_Execute.enabled = $true
	$status_Step3.Text = "Step 3 Status: Idle"
	write-host "Step 3 jobs finished"
    Get-Job | Remove-Job -Force
	$EventLog = New-Object System.Diagnostics.EventLog('Application')
	$EventLog.MachineName = "."
	$EventLog.Source = "O365DC"
	try{$EventLog.WriteEntry("Ending O365DC Step 3","Information", 31)} catch{}
    Stop-Transcript
}

$handler_btn_Step3_ClientAccess_CheckAll_Click=
{
	Set-AllFunctionsClientAccess -Check $true
}

$handler_btn_Step3_ClientAccess_UncheckAll_Click=
{
	Set-AllFunctionsClientAccess -Check $False
}

$handler_btn_Step3_Global_CheckAll_Click=
{
	Set-AllFunctionsGlobal -Check $true
}

$handler_btn_Step3_Global_UncheckAll_Click=
{
	Set-AllFunctionsGlobal -Check $false
}

$handler_btn_Step3_Recipient_CheckAll_Click=
{
	Set-AllFunctionsRecipient -Check $true
}

$handler_btn_Step3_Recipient_UncheckAll_Click=
{
	Set-AllFunctionsRecipient -Check $False
}

$handler_btn_Step3_Transport_CheckAll_Click=
{
	Set-AllFunctionsTransport -Check $true
}

$handler_btn_Step3_Transport_UncheckAll_Click=
{
	Set-AllFunctionsTransport -Check $False
}

$handler_btn_Step3_Um_CheckAll_Click=
{
	Set-AllFunctionsUm -Check $true
}

$handler_btn_Step3_Um_UncheckAll_Click=
{
	Set-AllFunctionsUm -Check $False
}

$handler_btn_Step3_Misc_CheckAll_Click=
{
	Set-AllFunctionsMisc -Check $true
}

$handler_btn_Step3_Misc_UncheckAll_Click=
{
	Set-AllFunctionsMisc -Check $False
}

$handler_btn_Step3_AzureAd_CheckAll_Click=
{
	Set-AllFunctionsAzureAd -Check $true
}

$handler_btn_Step3_AzureAd_UncheckAll_Click=
{
	Set-AllFunctionsAzureAd -Check $False
}

$handler_btn_Step3_Spo_CheckAll_Click=
{
	Set-AllFunctionsSpo -Check $true
}

$handler_btn_Step3_Spo_UncheckAll_Click=
{
	Set-AllFunctionsSpo -Check $False
}

#endregion "Step3 - Tests" Events

#region "Step4 - Reporting" Events
$handler_btn_Step4_Assemble_Click=
{
	$btn_Step4_Assemble.enabled = $false
    $status_Step4.Text = "Step 4 Status: Running"
	$EventLog = New-Object System.Diagnostics.EventLog('Application')
	$EventLog.MachineName = "."
	$EventLog.Source = "O365DC"
	try{$EventLog.WriteEntry("Starting O365DC Step 4","Information", 40)} catch{}
	#Minimize form to the back to expose the Powershell window when starting Step 4
	$form1.WindowState = "minimized"
	write-host "O365DC Form minimized." -ForegroundColor Green
	if ((Test-Path registry::HKey_Classes_Root\Excel.Application\CurVer) -eq $true)
	{
<# 		if ($chk_Step4_DC_Report.checked -eq $true)
		{
			write-host "-- Starting to assemble the DC Spreadsheet"
				.\O365DC_Scripts\Core_assemble_dc_Excel.ps1 -RunLocation $location
				write-host "---- Completed the DC Spreadsheet" -ForegroundColor Green
		}
		if ($chk_Step4_Ex_Report.checked -eq $true)
		{
			write-host "-- Starting to assemble the Exchange Server Spreadsheet"
				.\O365DC_Scripts\Core_assemble_exch_Excel.ps1 -RunLocation $location
				write-host "---- Completed the Exchange Spreadsheet" -ForegroundColor Green
		}
 #>
 		if ($chk_Step4_Exchange_Report.checked -eq $true)
		{
			write-host "-- Starting to assemble the Exchange Organization Spreadsheet"
				.\O365DC_Scripts\Core_assemble_Exchange_Excel.ps1 -RunLocation $location
				write-host "---- Completed the Exchange Organization Spreadsheet" -ForegroundColor Green
		}
		if ($chk_Step4_Azure_Report.checked -eq $true)
		{
			write-host "-- Starting to assemble the Azure Spreadsheet"
				.\O365DC_Scripts\Core_assemble_Azure_Excel.ps1 -RunLocation $location
				write-host "---- Completed the Azure Spreadsheet" -ForegroundColor Green
		}
	}
	else
	{
		write-host "Excel does not appear to be installed on this server."
		$EventLog = New-Object System.Diagnostics.EventLog('Application')
		$EventLog.MachineName = "."
		$EventLog.Source = "O365DC"
		try{$EventLog.WriteEntry("Excel does not appear to be installed on this server.","Warning", 49)} catch{}
	}
<# 	if ((Test-Path registry::HKey_Classes_Root\Word.Application\CurVer) -eq $true)
	{
		if ($chk_Step4_Exchange_Environment_Doc.checked -eq $true)
		{
			write-host "-- Starting to assemble the Exchange Documentation using Word"
				.\O365DC_Scripts\Core_Assemble_ExDoc_Word.ps1 -RunLocation $location
				write-host "---- Completed the Exchange Documentation using Word" -ForegroundColor Green
		}
	}
	else
	{
		write-host "Word does not appear to be installed on this server."
		$EventLog = New-Object System.Diagnostics.EventLog('Application')
		$EventLog.MachineName = "."
		$EventLog.Source = "O365DC"
		try{$EventLog.WriteEntry("Word does not appear to be installed on this server.","Warning", 49)} catch{}
	}
 #>	write-host "Restoring O365DC Form to normal." -ForegroundColor Green
	$form1.WindowState = "normal"
	$EventLog = New-Object System.Diagnostics.EventLog('Application')
	$EventLog.MachineName = "."
	$EventLog.Source = "O365DC"
	try{$EventLog.WriteEntry("Ending O365DC Step 4","Information", 41)} catch{}
	$status_Step4.Text = "Step 4 Status: Idle"
    $btn_Step4_Assemble.enabled = $true
}
#endregion "Step4 - Reporting" Events

#region *** Events ***

#endregion *** Events ***

$OnLoadForm_StateCorrection=
{$form1.WindowState = $InitialFormWindowState}

#region *** Build Form ***

#Region Form Main
#Region Reusable
# Reusable fonts
	$font_Calibri_8pt_normal = 	New-Object System.Drawing.Font("Calibri",7.8,0,3,0)
	$font_Calibri_10pt_normal = New-Object System.Drawing.Font("Calibri",9.75,0,3,1)
	$font_Calibri_12pt_normal = New-Object System.Drawing.Font("Calibri",12,0,3,1)
	$font_Calibri_14pt_normal = New-Object System.Drawing.Font("Calibri",14.25,0,3,1)
	$font_Calibri_10pt_bold = 	New-Object System.Drawing.Font("Calibri",9.75,1,3,1)
# Reusable padding
	$System_Windows_Forms_Padding_Reusable = New-Object System.Windows.Forms.Padding
	$System_Windows_Forms_Padding_Reusable.All = 3
	$System_Windows_Forms_Padding_Reusable.Bottom = 3
	$System_Windows_Forms_Padding_Reusable.Left = 3
	$System_Windows_Forms_Padding_Reusable.Right = 3
	$System_Windows_Forms_Padding_Reusable.Top = 3
# Reusable button
	$Size_Buttons 			= New-Object System.Drawing.Size(110,38)
# Reusable status
	$Size_Status 			= New-Object System.Drawing.Size(651,22)
	$Loc_Status 			= New-Object System.Drawing.Point(3,653)
# Reusable tabs
	$Size_Tab_1 			= New-Object System.Drawing.Size(700,678)
# Reusable checkboxes
	$Size_Chk 				= New-Object System.Drawing.Size(225,20)
	$Size_Chk_long 			= New-Object System.Drawing.Size(400,20)
# Reusable Size
	$Size_Form 				= New-Object System.Drawing.Size(665,718)
	$Size_Tab_Small 		= New-Object System.Drawing.Size(542,488)
	$Size_Tab_Control		= New-Object System.Drawing.Size(100,32)
# Reusable Location
	$Loc_Tab_Control 		= New-Object System.Drawing.Point(0,0)
	$Loc_Tab_Tier1 			= New-Object System.Drawing.Point(4,36)
	$Loc_Tab_Tier3 			= New-Object System.Drawing.Point(4,33)
	$Loc_Btn_1 				= New-Object System.Drawing.Point(20,15)
	$Loc_Lbl_1 				= New-Object System.Drawing.Point(138,15)
	$Loc_Box_1				= New-Object System.Drawing.Point(3,3)

# Reusable text box in Step1
	$Size_TextBox 			= New-Object System.Drawing.Size(400,27)
# Reusable boxes in Step1 Tabs
	$Size_Box_1 			= New-Object System.Drawing.Size(536,482)
# Reusable check buttons in Step1 tabs
	$Size_Btn_1 			= New-Object System.Drawing.Size(150,25)
# Reusable check list boxes in Step1 tabs
	$Size_Clb_1 			= New-Object System.Drawing.Size(400,350)
	$Loc_Clb_1 				= New-Object System.Drawing.Point(50,50)
# Reusable Discover/populate buttons in Step1 tabs
	$Loc_Discover 			= New-Object System.Drawing.Point(50,15)
	$Loc_Populate 			= New-Object System.Drawing.Point(300,15)
# Reusable check/uncheck buttons in Step1 tabs
	$Loc_CheckAll_1 		= New-Object System.Drawing.Point(50,450)
	$Loc_UncheckAll_1 		= New-Object System.Drawing.Point(300,450)
# Reusable boxes in Step3 Tabs
	$Size_Box_3 			= New-Object System.Drawing.Size(536,400)
# Reusable check buttons in Step3 tabs
	$Size_Btn_3 			= New-Object System.Drawing.Size(150,25)
# Reusable check/uncheck buttons in Step3 tabs
	$Loc_CheckAll_3 		= New-Object System.Drawing.Point(50,400)
	$Loc_UncheckAll_3 		= New-Object System.Drawing.Point(300,400)
#EndRegion Reusable

#Region Main Form
$form1.BackColor = [System.Drawing.Color]::FromArgb(255,169,169,169)
	$form1.ClientSize = $Size_Form
	$form1.MaximumSize = $Size_Form
	$form1.Font = $font_Calibri_10pt_normal
	$form1.FormBorderStyle = 2
	$form1.MaximizeBox = $False
	$form1.Name = "form1"
	$form1.ShowIcon = $False
	$form1.StartPosition = 1
	$form1.Text = "Office 365 Data Collector v4.0.2"
#EndRegion Main Form

#Region Main Tabs
$TabIndex = 0
$tab_Master.Appearance = 2
	$tab_Master.Dock = 5
	$tab_Master.Font = $font_Calibri_14pt_normal
	$tab_Master.ItemSize = $Size_Tab_Control
	$tab_Master.Location = $Loc_Tab_Control
	$tab_Master.Name = "tab_Master"
	$tab_Master.SelectedIndex = 0
	$tab_Master.Size = $Size_Form
	$tab_Master.SizeMode = "filltoright"
	$tab_Master.TabIndex = $TabIndex++
	$form1.Controls.Add($tab_Master)
#EndRegion Main Tabs

#Region Menu Strip
$Menu_Main.Location = $Loc_Tab_Control
	$Menu_Main.Name = "MainMenu"
	$Menu_Main.Size = new-object System.Drawing.Size(1151, 24)
	$Menu_Main.TabIndex = $TabIndex++
	$Menu_Main.Text = "Main Menu"
	$form1.Controls.add($Menu_Main)
[Void]$Menu_File.DropDownItems.Add($Submenu_LoadTargets)
[Void]$Menu_File.DropDownItems.Add($Submenu_PackageLogs)
[Void]$Menu_Main.items.add($Menu_File)
[Void]$Menu_Toggle.DropDownItems.Add($Submenu_Targets_CheckAll)
[Void]$Menu_Toggle.DropDownItems.Add($Submenu_Targets_UnCheckAll)
[Void]$Menu_Toggle.DropDownItems.Add($Submenu_Tests_CheckAll)
[Void]$Menu_Toggle.DropDownItems.Add($Submenu_Tests_UnCheckAll)
[Void]$Menu_Main.items.add($Menu_Toggle)
[Void]$Menu_Help.DropDownItems.Add($Submenu_Help)
[Void]$Menu_Help.DropDownItems.Add($Submenu_About)
[Void]$Menu_Main.items.add($Menu_Help)
$Submenu_LoadTargets.add_click($handler_Submenu_LoadTargets)
$Submenu_PackageLogs.add_click($handler_Submenu_PackageLogs)
$Submenu_Targets_CheckAll.add_click($handler_Submenu_Targets_CheckAll)
$Submenu_Targets_UnCheckAll.add_click($handler_Submenu_Targets_UnCheckAll)
$Submenu_Tests_CheckAll.add_click($handler_Submenu_Tests_CheckAll)
$Submenu_Tests_UnCheckAll.add_click($handler_Submenu_Tests_UnCheckAll)
$Submenu_Help.add_click($handler_Submenu_Help)
$Submenu_About.add_click($handler_Submenu_About)
#EndRegion Menu Strip

#EndRegion Form Main

#Region "Step1 - Targets"

#Region Step1 Main
$TabIndex = 0
$tab_Step1.BackColor = [System.Drawing.Color]::FromArgb(0,255,255,255)
	$tab_Step1.Location = $Loc_Tab_Tier1
	$tab_Step1.Name = "tab_Step1"
	$tab_Step1.Padding = $System_Windows_Forms_Padding_Reusable
	$tab_Step1.TabIndex = $TabIndex++
	$tab_Step1.Text = "  Targets  "
	$tab_Step1.Size = $Size_Tab_1
	$tab_Master.Controls.Add($tab_Step1)
$btn_Step1_Discover.Font = $font_Calibri_14pt_normal
	$btn_Step1_Discover.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
	$btn_Step1_Discover.Location = $Loc_Btn_1
	$btn_Step1_Discover.Name = "btn_Step1_Discover"
	$btn_Step1_Discover.Size = $Size_Buttons
	$btn_Step1_Discover.TabIndex = $TabIndex++
	$btn_Step1_Discover.Text = "Discover"
	$btn_Step1_Discover.Visible = $false
	$btn_Step1_Discover.UseVisualStyleBackColor = $True
	$btn_Step1_Discover.add_Click($handler_btn_Step1_Discover_Click)
	$tab_Step1.Controls.Add($btn_Step1_Discover)
$btn_Step1_Populate.Font = $font_Calibri_14pt_normal
	$btn_Step1_Populate.Location = New-Object System.Drawing.Point(200,15)
	$btn_Step1_Populate.Name = "btn_Step1_Populate"
	$btn_Step1_Populate.Size = $Size_Buttons
	$btn_Step1_Populate.TabIndex = $TabIndex++
	$btn_Step1_Populate.Text = "Load from File"
	$btn_Step1_Populate.Visible = $false
	$btn_Step1_Populate.UseVisualStyleBackColor = $True
	$btn_Step1_Populate.add_Click($handler_btn_Step1_Populate_Click)
	$tab_Step1.Controls.Add($btn_Step1_Populate)
$tab_Step1_Master.Font = $font_Calibri_12pt_normal
	$tab_Step1_Master.Location = New-Object System.Drawing.Point(20,60)
	$tab_Step1_Master.Name = "tab_Step1_Master"
	$tab_Step1_Master.SelectedIndex = 0
	$tab_Step1_Master.Size = New-Object System.Drawing.Size(550,525)
	$tab_Step1_Master.TabIndex = $TabIndex++
	$tab_Step1.Controls.Add($tab_Step1_Master)
$status_Step1.Font = $font_Calibri_10pt_normal
	$status_Step1.Location = $Loc_Status
	$status_Step1.Name = "status_Step1"
	$status_Step1.Size = $Size_Status
	$status_Step1.TabIndex = $TabIndex++
	$status_Step1.Text = "Step 1 Status"
	$tab_Step1.Controls.Add($status_Step1)
#EndRegion Step1 Main

#Region Step1 Mailboxes tab
$TabIndex = 0
$tab_Step1_Mailboxes.Location = $Loc_Tab_Tier3
	$tab_Step1_Mailboxes.Name = "tab_Step1_Mailboxes"
	$tab_Step1_Mailboxes.Padding = $System_Windows_Forms_Padding_Reusable
	$tab_Step1_Mailboxes.Size = $Size_Tab_Small
	$tab_Step1_Mailboxes.TabIndex = $TabIndex++
	$tab_Step1_Mailboxes.Text = "Mailboxes"
	$tab_Step1_Mailboxes.UseVisualStyleBackColor = $True
	$tab_Step1_Master.Controls.Add($tab_Step1_Mailboxes)
$btn_Step1_Mailboxes_Discover.Font = $font_Calibri_10pt_normal
	$btn_Step1_Mailboxes_Discover.Location = $Loc_Discover
	$btn_Step1_Mailboxes_Discover.Name = "btn_Step1_Mailboxes_Discover"
	$btn_Step1_Mailboxes_Discover.Size = $Size_Btn_1
	$btn_Step1_Mailboxes_Discover.TabIndex = $TabIndex++
	$btn_Step1_Mailboxes_Discover.Text = "Discover"
	$btn_Step1_Mailboxes_Discover.UseVisualStyleBackColor = $True
	$btn_Step1_Mailboxes_Discover.add_Click($handler_btn_Step1_Mailboxes_Discover)
	$bx_Mailboxes_List.Controls.Add($btn_Step1_Mailboxes_Discover)
$btn_Step1_Mailboxes_Populate.Font = $font_Calibri_10pt_normal
	$btn_Step1_Mailboxes_Populate.Location = $Loc_Populate
	$btn_Step1_Mailboxes_Populate.Name = "btn_Step1_Mailboxes_Populate"
	$btn_Step1_Mailboxes_Populate.Size = $Size_Btn_1
	$btn_Step1_Mailboxes_Populate.TabIndex = $TabIndex++
	$btn_Step1_Mailboxes_Populate.Text = "Load from File"
	$btn_Step1_Mailboxes_Populate.UseVisualStyleBackColor = $True
	$btn_Step1_Mailboxes_Populate.add_Click($handler_btn_Step1_Mailboxes_Populate)
$bx_Mailboxes_List.Controls.Add($btn_Step1_Mailboxes_Populate)
	$bx_Mailboxes_List.Dock = 5
	$bx_Mailboxes_List.Font = $font_Calibri_10pt_bold
	$bx_Mailboxes_List.Location = $Loc_Box_1
	$bx_Mailboxes_List.Name = "bx_Mailboxes_List"
	$bx_Mailboxes_List.Size = $Size_Box_1
	$bx_Mailboxes_List.TabIndex = $TabIndex++
	$bx_Mailboxes_List.TabStop = $False
	$tab_Step1_Mailboxes.Controls.Add($bx_Mailboxes_List)
$clb_Step1_Mailboxes_List.Font = $font_Calibri_10pt_normal
	$clb_Step1_Mailboxes_List.Location = $Loc_Clb_1
	$clb_Step1_Mailboxes_List.Name = "clb_Step1_Mailboxes_List"
	$clb_Step1_Mailboxes_List.Size = $Size_Clb_1
	$clb_Step1_Mailboxes_List.TabIndex = $TabIndex++
	$clb_Step1_Mailboxes_List.horizontalscrollbar = $true
	$clb_Step1_Mailboxes_List.CheckOnClick = $true
	$bx_Mailboxes_List.Controls.Add($clb_Step1_Mailboxes_List)
$txt_MailboxesTotal.Font = $font_Calibri_10pt_normal
	$txt_MailboxesTotal.Location = New-Object System.Drawing.Point(50,410)
	$txt_MailboxesTotal.Name = "txt_MailboxesTotal"
	$txt_MailboxesTotal.Size = $Size_TextBox
	$txt_MailboxesTotal.TabIndex = $TabIndex++
	$txt_MailboxesTotal.Visible = $False
	$bx_Mailboxes_List.Controls.Add($txt_MailboxesTotal)
$btn_Step1_Mailboxes_CheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step1_Mailboxes_CheckAll.Location = $Loc_CheckAll_1
	$btn_Step1_Mailboxes_CheckAll.Name = "btn_Step1_Mailboxes_CheckAll"
	$btn_Step1_Mailboxes_CheckAll.Size = $Size_Btn_1
	$btn_Step1_Mailboxes_CheckAll.TabIndex = $TabIndex++
	$btn_Step1_Mailboxes_CheckAll.Text = "Check all on this tab"
	$btn_Step1_Mailboxes_CheckAll.UseVisualStyleBackColor = $True
	$btn_Step1_Mailboxes_CheckAll.add_Click($handler_btn_Step1_Mailboxes_CheckAll)
	$bx_Mailboxes_List.Controls.Add($btn_Step1_Mailboxes_CheckAll)
$btn_Step1_Mailboxes_UncheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step1_Mailboxes_UncheckAll.Location = $Loc_UncheckAll_1
	$btn_Step1_Mailboxes_UncheckAll.Name = "btn_Step1_Mailboxes_UncheckAll"
	$btn_Step1_Mailboxes_UncheckAll.Size = $Size_Btn_1
	$btn_Step1_Mailboxes_UncheckAll.TabIndex = $TabIndex++
	$btn_Step1_Mailboxes_UncheckAll.Text = "Uncheck all on this tab"
	$btn_Step1_Mailboxes_UncheckAll.UseVisualStyleBackColor = $True
	$btn_Step1_Mailboxes_UncheckAll.add_Click($handler_btn_Step1_Mailboxes_UncheckAll)
	$bx_Mailboxes_List.Controls.Add($btn_Step1_Mailboxes_UncheckAll)
#EndRegion Step1 Mailboxes tab

#Region Step1 AzureAdUser tab
$TabIndex = 0
$tab_Step1_AzureAdUser.Location = $Loc_Tab_Tier3
	$tab_Step1_AzureAdUser.Name = "tab_Step1_AzureAdUser"
	$tab_Step1_AzureAdUser.Padding = $System_Windows_Forms_Padding_Reusable
	$tab_Step1_AzureAdUser.Size = $Size_Tab_Small
	$tab_Step1_AzureAdUser.TabIndex = $TabIndex++
	$tab_Step1_AzureAdUser.Text = "AzureAD Users"
	$tab_Step1_AzureAdUser.UseVisualStyleBackColor = $True
	$tab_Step1_Master.Controls.Add($tab_Step1_AzureAdUser)
$btn_Step1_AzureAdUser_Discover.Font = $font_Calibri_10pt_normal
	$btn_Step1_AzureAdUser_Discover.Location = $Loc_Discover
	$btn_Step1_AzureAdUser_Discover.Name = "btn_Step1_AzureAdUser_Discover"
	$btn_Step1_AzureAdUser_Discover.Size = $Size_Btn_1
	$btn_Step1_AzureAdUser_Discover.TabIndex = $TabIndex++
	$btn_Step1_AzureAdUser_Discover.Text = "Discover"
	$btn_Step1_AzureAdUser_Discover.UseVisualStyleBackColor = $True
	$btn_Step1_AzureAdUser_Discover.add_Click($handler_btn_Step1_AzureAdUser_Discover)
	$bx_AzureAdUser_List.Controls.Add($btn_Step1_AzureAdUser_Discover)
$btn_Step1_AzureAdUser_Populate.Font = $font_Calibri_10pt_normal
	$btn_Step1_AzureAdUser_Populate.Location = $Loc_Populate
	$btn_Step1_AzureAdUser_Populate.Name = "btn_Step1_AzureAdUser_Populate"
	$btn_Step1_AzureAdUser_Populate.Size = $Size_Btn_1
	$btn_Step1_AzureAdUser_Populate.TabIndex = $TabIndex++
	$btn_Step1_AzureAdUser_Populate.Text = "Load from File"
	$btn_Step1_AzureAdUser_Populate.UseVisualStyleBackColor = $True
	$btn_Step1_AzureAdUser_Populate.add_Click($handler_btn_Step1_AzureAdUser_Populate)
$bx_AzureAdUser_List.Controls.Add($btn_Step1_AzureAdUser_Populate)
	$bx_AzureAdUser_List.Dock = 5
	$bx_AzureAdUser_List.Font = $font_Calibri_10pt_bold
	$bx_AzureAdUser_List.Location = $Loc_Box_1
	$bx_AzureAdUser_List.Name = "bx_AzureAdUser_List"
	$bx_AzureAdUser_List.Size = $Size_Box_1
	$bx_AzureAdUser_List.TabIndex = $TabIndex++
	$bx_AzureAdUser_List.TabStop = $False
	$tab_Step1_AzureAdUser.Controls.Add($bx_AzureAdUser_List)
$clb_Step1_AzureAdUser_List.Font = $font_Calibri_10pt_normal
	$clb_Step1_AzureAdUser_List.Location = $Loc_Clb_1
	$clb_Step1_AzureAdUser_List.Name = "clb_Step1_AzureAdUser_List"
	$clb_Step1_AzureAdUser_List.Size = $Size_Clb_1
	$clb_Step1_AzureAdUser_List.TabIndex = $TabIndex++
	$clb_Step1_AzureAdUser_List.horizontalscrollbar = $true
	$clb_Step1_AzureAdUser_List.CheckOnClick = $true
	$bx_AzureAdUser_List.Controls.Add($clb_Step1_AzureAdUser_List)
$txt_AzureAdUserTotal.Font = $font_Calibri_10pt_normal
	$txt_AzureAdUserTotal.Location = New-Object System.Drawing.Point(50,410)
	$txt_AzureAdUserTotal.Name = "txt_AzureAdUserTotal"
	$txt_AzureAdUserTotal.Size = $Size_TextBox
	$txt_AzureAdUserTotal.TabIndex = $TabIndex++
	$txt_AzureAdUserTotal.Visible = $False
	$bx_AzureAdUser_List.Controls.Add($txt_AzureAdUserTotal)
$btn_Step1_AzureAdUser_CheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step1_AzureAdUser_CheckAll.Location = $Loc_CheckAll_1
	$btn_Step1_AzureAdUser_CheckAll.Name = "btn_Step1_AzureAdUser_CheckAll"
	$btn_Step1_AzureAdUser_CheckAll.Size = $Size_Btn_1
	$btn_Step1_AzureAdUser_CheckAll.TabIndex = $TabIndex++
	$btn_Step1_AzureAdUser_CheckAll.Text = "Check all on this tab"
	$btn_Step1_AzureAdUser_CheckAll.UseVisualStyleBackColor = $True
	$btn_Step1_AzureAdUser_CheckAll.add_Click($handler_btn_Step1_AzureAdUser_CheckAll)
	$bx_AzureAdUser_List.Controls.Add($btn_Step1_AzureAdUser_CheckAll)
$btn_Step1_AzureAdUser_UncheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step1_AzureAdUser_UncheckAll.Location = $Loc_UncheckAll_1
	$btn_Step1_AzureAdUser_UncheckAll.Name = "btn_Step1_AzureAdUser_UncheckAll"
	$btn_Step1_AzureAdUser_UncheckAll.Size = $Size_Btn_1
	$btn_Step1_AzureAdUser_UncheckAll.TabIndex = $TabIndex++
	$btn_Step1_AzureAdUser_UncheckAll.Text = "Uncheck all on this tab"
	$btn_Step1_AzureAdUser_UncheckAll.UseVisualStyleBackColor = $True
	$btn_Step1_AzureAdUser_UncheckAll.add_Click($handler_btn_Step1_AzureAdUser_UncheckAll)
	$bx_AzureAdUser_List.Controls.Add($btn_Step1_AzureAdUser_UncheckAll)
#EndRegion Step1 AzureAdUser tab

#Endregion "Step1 - Targets"

#Region "Step2"
$TabIndex = 0
$tab_Step2.BackColor = [System.Drawing.Color]::FromArgb(0,255,255,255)
	$tab_Step2.Font = $font_Calibri_8pt_normal
	$tab_Step2.Location = $Loc_Tab_Tier1
	$tab_Step2.Name = "tab_Step2"
	$tab_Step2.Padding = $System_Windows_Forms_Padding_Reusable
	$tab_Step2.TabIndex = $TabIndex++
	$tab_Step2.Text = "  Templates  "
	$tab_Step2.Size = $Size_Tab_1
	$tab_Master.Controls.Add($tab_Step2)
$bx_Step2_Templates.Font = $font_Calibri_10pt_bold
	$bx_Step2_Templates.Location = New-Object System.Drawing.Point(27,91)
	$bx_Step2_Templates.Name = "bx_Step2_Templates"
	$bx_Step2_Templates.Size = $Size_Box_1  #New-Object System.Drawing.Size(536,487)
	$bx_Step2_Templates.TabIndex = $TabIndex++
	$bx_Step2_Templates.TabStop = $False
	$bx_Step2_Templates.Text = "Select a data collection template"
	$tab_Step2.Controls.Add($bx_Step2_Templates)
$rb_Step2_Template_1.Checked = $False
	$rb_Step2_Template_1.Font = $font_Calibri_10pt_normal
	$rb_Step2_Template_1.Location = New-Object System.Drawing.Point(50,25)
	$rb_Step2_Template_1.Name = "rb_Step2_Template_1"
	$rb_Step2_Template_1.Size = $Size_Chk_long
	$rb_Step2_Template_1.TabIndex = $TabIndex++
	$rb_Step2_Template_1.Text = "Recommended tests"
	$rb_Step2_Template_1.UseVisualStyleBackColor = $True
	$rb_Step2_Template_1.add_Click($handler_rb_Step2_Template_1)
	$bx_Step2_Templates.Controls.Add($rb_Step2_Template_1)
$rb_Step2_Template_2.Checked = $False
	$rb_Step2_Template_2.Font = $font_Calibri_10pt_normal
	$rb_Step2_Template_2.Location = New-Object System.Drawing.Point(50,50)
	$rb_Step2_Template_2.Name = "rb_Step2_Template_2"
	$rb_Step2_Template_2.Size = $Size_Chk_long
	$rb_Step2_Template_2.TabIndex = $TabIndex++
	$rb_Step2_Template_2.Text = "All tests"
	$rb_Step2_Template_2.UseVisualStyleBackColor = $True
	$rb_Step2_Template_2.add_Click($handler_rb_Step2_Template_2)
	$bx_Step2_Templates.Controls.Add($rb_Step2_Template_2)
$rb_Step2_Template_3.Checked = $False
	$rb_Step2_Template_3.Font = $font_Calibri_10pt_normal
	$rb_Step2_Template_3.Location = New-Object System.Drawing.Point(50,75)
	$rb_Step2_Template_3.Name = "rb_Step2_Template_3"
	$rb_Step2_Template_3.Size = $Size_Chk_long
	$rb_Step2_Template_3.TabIndex = $TabIndex++
	$rb_Step2_Template_3.Text = "Minimum tests for Environmental Document"
	$rb_Step2_Template_3.UseVisualStyleBackColor = $True
	$rb_Step2_Template_3.add_Click($handler_rb_Step2_Template_3)
	$bx_Step2_Templates.Controls.Add($rb_Step2_Template_3)
$rb_Step2_Template_4.Checked = $False
	$rb_Step2_Template_4.Font = $font_Calibri_10pt_normal
	$rb_Step2_Template_4.Location = New-Object System.Drawing.Point(50,100)
	$rb_Step2_Template_4.Name = "rb_Step2_Template_4"
	$rb_Step2_Template_4.Size = $Size_Chk_long
	$rb_Step2_Template_4.TabIndex = $TabIndex++
	$rb_Step2_Template_4.Text = "Custom Template 1"
	$rb_Step2_Template_4.UseVisualStyleBackColor = $True
	$rb_Step2_Template_4.add_Click($handler_rb_Step2_Template_4)
	$bx_Step2_Templates.Controls.Add($rb_Step2_Template_4)
$Status_Step2.Font = $font_Calibri_10pt_normal
	$Status_Step2.Location = $Loc_Status
	$Status_Step2.Name = "Status_Step2"
	$Status_Step2.Size = $Size_Status
	$Status_Step2.TabIndex = $TabIndex++
	$Status_Step2.Text = "Step 2 Status"
	$tab_Step2.Controls.Add($Status_Step2)
#Endregion "Step2"

#Region "Step3 - Tests"
#Region Step3 Main
$TabIndex = 0
$tab_Step3.Location = $Loc_Tab_Tier1
	$tab_Step3.Name = "tab_Step3"
	$tab_Step3.Padding = $System_Windows_Forms_Padding_Reusable
	$tab_Step3.TabIndex = $TabIndex++
	$tab_Step3.Text = "   Tests   "
	$tab_Step3.Size = $Size_Tab_1
	$tab_Master.Controls.Add($tab_Step3)
$tab_Step3_Master.Font = $font_Calibri_12pt_normal
	$tab_Step3_Master.Location = New-Object System.Drawing.Point(20,60)
	$tab_Step3_Master.Name = "tab_Step3_Master"
	$tab_Step3_Master.SelectedIndex = 0
	$tab_Step3_Master.Size = New-Object System.Drawing.Size(550,525)
	$tab_Step3_Master.TabIndex = $TabIndex++
	$tab_Step3.Controls.Add($tab_Step3_Master)
$btn_Step3_Execute.Font = $font_Calibri_14pt_normal
	$btn_Step3_Execute.Location = $Loc_Btn_1
	$btn_Step3_Execute.Name = "btn_Step3_Execute"
	$btn_Step3_Execute.Size = $Size_Buttons
	$btn_Step3_Execute.TabIndex = $TabIndex++
	$btn_Step3_Execute.Text = "Execute"
	$btn_Step3_Execute.UseVisualStyleBackColor = $True
	$btn_Step3_Execute.add_Click($handler_btn_Step3_Execute_Click)
	$tab_Step3.Controls.Add($btn_Step3_Execute)
$lbl_Step3_Execute.BackColor = [System.Drawing.Color]::FromArgb(0,255,255,255)
	$lbl_Step3_Execute.Font = $font_Calibri_10pt_normal
	$lbl_Step3_Execute.Location = $Loc_Lbl_1
	$lbl_Step3_Execute.Name = "lbl_Step3"
	$lbl_Step3_Execute.Size = New-Object System.Drawing.Size(510,38)
	$lbl_Step3_Execute.TabIndex = $TabIndex++
	$lbl_Step3_Execute.Text = "Select the functions below and click on the Execute button."
	$lbl_Step3_Execute.TextAlign = 16
	$tab_Step3.Controls.Add($lbl_Step3_Execute)
$status_Step3.Font = $font_Calibri_10pt_normal
	$status_Step3.Location = $Loc_Status
	$status_Step3.Name = "status_Step3"
	$status_Step3.Size = $Size_Status
	$status_Step3.TabIndex = $TabIndex++
	$status_Step3.Text = "Step 3 Status"
	$tab_Step3.Controls.Add($status_Step3)
#EndRegion Step3 Main

#Region Step3 Exchange - Tier 2
$TabIndex = 0
$tab_Step3_Exchange.Location = $Loc_Tab_Tier3
	$tab_Step3_Exchange.Name = "tab_Step3_Exchange"
	$tab_Step3_Exchange.Padding = $System_Windows_Forms_Padding_Reusable
	$tab_Step3_Exchange.Size = $Size_Tab_Small
	$tab_Step3_Exchange.TabIndex = $TabIndex++
	$tab_Step3_Exchange.Text = "Exchange Functions"
	$tab_Step3_Exchange.UseVisualStyleBackColor = $True
	$tab_Step3_Master.Controls.Add($tab_Step3_Exchange)
#EndRegion Step3 Exchange - Tier 2

#Region Step3 Azure - Tier 2
$TabIndex = 0
$tab_Step3_Azure.Location = $Loc_Tab_Tier3
	$tab_Step3_Azure.Name = "tab_Step3_Azure"
	$tab_Step3_Azure.Padding = $System_Windows_Forms_Padding_Reusable
	$tab_Step3_Azure.Size = $Size_Tab_Small
	$tab_Step3_Azure.TabIndex = $TabIndex++
	$tab_Step3_Azure.Text = "Azure Functions"
	$tab_Step3_Azure.UseVisualStyleBackColor = $True
	$tab_Step3_Master.Controls.Add($tab_Step3_Azure)
#EndRegion Step3 Azure - Tier 2

#Region Step3 Sharepoint - Tier 2
$TabIndex = 0
$tab_Step3_Sharepoint.Location = $Loc_Tab_Tier3
	$tab_Step3_Sharepoint.Name = "tab_Step3_Sharepoint"
	$tab_Step3_Sharepoint.Padding = $System_Windows_Forms_Padding_Reusable
	$tab_Step3_Sharepoint.Size = $Size_Tab_Small
	$tab_Step3_Sharepoint.TabIndex = $TabIndex++
	$tab_Step3_Sharepoint.Text = "Sharepoint Functions (Not yet Enabled)"
	$tab_Step3_Sharepoint.UseVisualStyleBackColor = $True
	$tab_Step3_Master.Controls.Add($tab_Step3_Sharepoint)
#EndRegion Step3 Sharepoint - Tier 2

#Region Exchange Tab Control
$TabIndex = 0
$tab_Step3_Exchange_Tier2.Appearance = 2
	$tab_Step3_Exchange_Tier2.Dock = 5
	$tab_Step3_Exchange_Tier2.Font = $font_Calibri_10pt_normal
	$tab_Step3_Exchange_Tier2.ItemSize = $Size_Tab_Control
	$tab_Step3_Exchange_Tier2.Location = $Loc_Tab_Control
	$tab_Step3_Exchange_Tier2.Name = "tab_Step3_Exchange_Tier2"
	$tab_Step3_Exchange_Tier2.SelectedIndex = 0
	$tab_Step3_Exchange_Tier2.Size = $Size_Form
	$tab_Step3_Exchange_Tier2.TabIndex = $TabIndex++
	$tab_Step3_Exchange.Controls.Add($tab_Step3_Exchange_Tier2)
#EndRegion Exchange Tab Control

#Region Azure Tab Control
$TabIndex = 0
$tab_Step3_Azure_Tier2.Appearance = 2
	$tab_Step3_Azure_Tier2.Dock = 5
	$tab_Step3_Azure_Tier2.Font = $font_Calibri_10pt_normal
	$tab_Step3_Azure_Tier2.ItemSize = $Size_Tab_Control
	$tab_Step3_Azure_Tier2.Location = $Loc_Tab_Control
	$tab_Step3_Azure_Tier2.Name = "tab_Step3_Azure_Tier2"
	$tab_Step3_Azure_Tier2.SelectedIndex = 0
	$tab_Step3_Azure_Tier2.Size = $Size_Form
	$tab_Step3_Azure_Tier2.TabIndex = $TabIndex++
	$tab_Step3_Azure.Controls.Add($tab_Step3_Azure_Tier2)
#EndRegion Azure Tab Control

#Region Sharepoint Tab Control
$TabIndex = 0
$tab_Step3_Sharepoint_Tier2.Appearance = 2
	$tab_Step3_Sharepoint_Tier2.Dock = 5
	$tab_Step3_Sharepoint_Tier2.Font = $font_Calibri_10pt_normal
	$tab_Step3_Sharepoint_Tier2.ItemSize = $Size_Tab_Control
	$tab_Step3_Sharepoint_Tier2.Location = $Loc_Tab_Control
	$tab_Step3_Sharepoint_Tier2.Name = "tab_Step3_Sharepoint_Tier2"
	$tab_Step3_Sharepoint_Tier2.SelectedIndex = 0
	$tab_Step3_Sharepoint_Tier2.Size = $Size_Form
	$tab_Step3_Sharepoint_Tier2.TabIndex = $TabIndex++
	$tab_Step3_Sharepoint.Controls.Add($tab_Step3_Sharepoint_Tier2)
#EndRegion Sharepoint Tab Control

#Region Step3 Exchange - Client Access tab
$TabIndex = 0
$tab_Step3_ClientAccess.Location = $Loc_Tab_Tier3
	$tab_Step3_ClientAccess.Name = "tab_Step3_ClientAccess"
	$tab_Step3_ClientAccess.Padding = $System_Windows_Forms_Padding_Reusable
	$tab_Step3_ClientAccess.Size = $Size_Tab_Small
	$tab_Step3_ClientAccess.TabIndex = $TabIndex++
	$tab_Step3_ClientAccess.Text = "Client Access"
	$tab_Step3_ClientAccess.UseVisualStyleBackColor = $True
	$tab_Step3_Exchange_Tier2.Controls.Add($tab_Step3_ClientAccess)
$bx_ClientAccess_Functions.Dock = 5
	$bx_ClientAccess_Functions.Font = $font_Calibri_10pt_bold
	$bx_ClientAccess_Functions.Location = $Loc_Box_1
	$bx_ClientAccess_Functions.Name = "bx_ClientAccess_Functions"
	$bx_ClientAccess_Functions.Size = $Size_Box_3
	$bx_ClientAccess_Functions.TabIndex = $TabIndex++
	$bx_ClientAccess_Functions.TabStop = $False
	$tab_Step3_ClientAccess.Controls.Add($bx_ClientAccess_Functions)
$btn_Step3_ClientAccess_CheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_ClientAccess_CheckAll.Location = $Loc_CheckAll_3
	$btn_Step3_ClientAccess_CheckAll.Name = "btn_Step3_ClientAccess_CheckAll"
	$btn_Step3_ClientAccess_CheckAll.Size = $Size_Btn_3
	$btn_Step3_ClientAccess_CheckAll.TabIndex = $TabIndex++
	$btn_Step3_ClientAccess_CheckAll.Text = "Check all on this tab"
	$btn_Step3_ClientAccess_CheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_ClientAccess_CheckAll.add_Click($handler_btn_Step3_ClientAccess_CheckAll_Click)
	$bx_ClientAccess_Functions.Controls.Add($btn_Step3_ClientAccess_CheckAll)
$btn_Step3_ClientAccess_UncheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_ClientAccess_UncheckAll.Location = $Loc_UncheckAll_3
	$btn_Step3_ClientAccess_UncheckAll.Name = "btn_Step3_ClientAccess_UncheckAll"
	$btn_Step3_ClientAccess_UncheckAll.Size = $Size_Btn_3
	$btn_Step3_ClientAccess_UncheckAll.TabIndex = $TabIndex++
	$btn_Step3_ClientAccess_UncheckAll.Text = "Uncheck all on this tab"
	$btn_Step3_ClientAccess_UncheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_ClientAccess_UncheckAll.add_Click($handler_btn_Step3_ClientAccess_UncheckAll_Click)
	$bx_ClientAccess_Functions.Controls.Add($btn_Step3_ClientAccess_UncheckAll)
$Col_1_loc = 35
$Col_2_loc = 290
$Row_1_loc = 0
$Row_2_loc = 0
$chk_Org_Get_ActiveSyncOrgSettings.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_ActiveSyncOrgSettings.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_ActiveSyncOrgSettings.Name = "chk_Org_Get_ActiveSyncOrgSettings"
	$chk_Org_Get_ActiveSyncOrgSettings.Size = $Size_Chk
	$chk_Org_Get_ActiveSyncOrgSettings.TabIndex = $TabIndex++
	$chk_Org_Get_ActiveSyncOrgSettings.Text = "Get-ActiveSyncOrgSettings"
	$chk_Org_Get_ActiveSyncOrgSettings.UseVisualStyleBackColor = $True
	$bx_ClientAccess_Functions.Controls.Add($chk_Org_Get_ActiveSyncOrgSettings)
$chk_Org_Get_AvailabilityAddressSpace.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_AvailabilityAddressSpace.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_AvailabilityAddressSpace.Name = "chk_Org_Get_AvailabilityAddressSpace"
	$chk_Org_Get_AvailabilityAddressSpace.Size = $Size_Chk
	$chk_Org_Get_AvailabilityAddressSpace.TabIndex = $TabIndex++
	$chk_Org_Get_AvailabilityAddressSpace.Text = "Get-AvailabilityAddressSpace"
	$chk_Org_Get_AvailabilityAddressSpace.UseVisualStyleBackColor = $True
	$bx_ClientAccess_Functions.Controls.Add($chk_Org_Get_AvailabilityAddressSpace)
$chk_Org_Get_MobileDevice.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_MobileDevice.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_MobileDevice.Name = "chk_Org_Get_MobileDevice"
	$chk_Org_Get_MobileDevice.Size = $Size_Chk
	$chk_Org_Get_MobileDevice.TabIndex = $TabIndex++
	$chk_Org_Get_MobileDevice.Text = "Get-MobileDevice"
	$chk_Org_Get_MobileDevice.UseVisualStyleBackColor = $True
	$bx_ClientAccess_Functions.Controls.Add($chk_Org_Get_MobileDevice)
$chk_Org_Get_MobileDevicePolicy.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_MobileDevicePolicy.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_MobileDevicePolicy.Name = "chk_Org_Get_MobileDevicePolicy"
	$chk_Org_Get_MobileDevicePolicy.Size = $Size_Chk
	$chk_Org_Get_MobileDevicePolicy.TabIndex = $TabIndex++
	$chk_Org_Get_MobileDevicePolicy.Text = "Get-MobileDeviceMailboxPolicy"
	$chk_Org_Get_MobileDevicePolicy.UseVisualStyleBackColor = $True
	$bx_ClientAccess_Functions.Controls.Add($chk_Org_Get_MobileDevicePolicy)
$chk_Org_Get_OwaMailboxPolicy.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_OwaMailboxPolicy.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_OwaMailboxPolicy.Name = "chk_Org_Get_OwaMailboxPolicy"
	$chk_Org_Get_OwaMailboxPolicy.Size = $Size_Chk
	$chk_Org_Get_OwaMailboxPolicy.TabIndex = $TabIndex++
	$chk_Org_Get_OwaMailboxPolicy.Text = "Get-OwaMailboxPolicy"
	$chk_Org_Get_OwaMailboxPolicy.UseVisualStyleBackColor = $True
	$bx_ClientAccess_Functions.Controls.Add($chk_Org_Get_OwaMailboxPolicy)
#EndRegion Step3 Exchange - Client Access tab

#Region Step3 Exchange - Global tab
$TabIndex = 0
$tab_Step3_Global.Location = $Loc_Tab_Tier3
	$tab_Step3_Global.Name = "tab_Step3_Global"
	$tab_Step3_Global.Padding = $System_Windows_Forms_Padding_Reusable
$tab_Step3_Global.Size = $Size_Tab_Small
	$tab_Step3_Global.TabIndex = $TabIndex++
	$tab_Step3_Global.Text = "Global and Database"
	$tab_Step3_Global.UseVisualStyleBackColor = $True
	$tab_Step3_Exchange_Tier2.Controls.Add($tab_Step3_Global)
$bx_Global_Functions.Dock = 5
	$bx_Global_Functions.Font = $font_Calibri_10pt_bold
	$bx_Global_Functions.Location = $Loc_Box_1
	$bx_Global_Functions.Name = "bx_Global_Functions"
	$bx_Global_Functions.Size = $Size_Box_3
	$bx_Global_Functions.TabIndex = $TabIndex++
	$bx_Global_Functions.TabStop = $False
	$tab_Step3_Global.Controls.Add($bx_Global_Functions)
$btn_Step3_Global_CheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_Global_CheckAll.Location = $Loc_CheckAll_3
	$btn_Step3_Global_CheckAll.Name = "btn_Step3_Global_CheckAll"
	$btn_Step3_Global_CheckAll.Size = $Size_Btn_3
	$btn_Step3_Global_CheckAll.TabIndex = $TabIndex++
	$btn_Step3_Global_CheckAll.Text = "Check all on this tab"
	$btn_Step3_Global_CheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_Global_CheckAll.add_Click($handler_btn_Step3_Global_CheckAll_Click)
	$bx_Global_Functions.Controls.Add($btn_Step3_Global_CheckAll)
$btn_Step3_Global_UncheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_Global_UncheckAll.Location = $Loc_UncheckAll_3
	$btn_Step3_Global_UncheckAll.Name = "btn_Step3_Global_UncheckAll"
	$btn_Step3_Global_UncheckAll.Size = $Size_Btn_3
	$btn_Step3_Global_UncheckAll.TabIndex = $TabIndex++
	$btn_Step3_Global_UncheckAll.Text = "Uncheck all on this tab"
	$btn_Step3_Global_UncheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_Global_UncheckAll.add_Click($handler_btn_Step3_Global_UncheckAll_Click)
	$bx_Global_Functions.Controls.Add($btn_Step3_Global_UncheckAll)
$Col_1_loc = 35
$Col_2_loc = 290
$Row_1_loc = 0
$Row_2_loc = 0
$chk_Org_Get_AddressBookPolicy.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_AddressBookPolicy.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_AddressBookPolicy.Name = "chk_Org_Get_AddressBookPolicy"
	$chk_Org_Get_AddressBookPolicy.Size = $Size_Chk
	$chk_Org_Get_AddressBookPolicy.TabIndex = $TabIndex++
	$chk_Org_Get_AddressBookPolicy.Text = "Get-AddressBookPolicy"
	$chk_Org_Get_AddressBookPolicy.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_AddressBookPolicy)
$chk_Org_Get_AddressList.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_AddressList.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_AddressList.Name = "chk_Org_Get_AddressList"
	$chk_Org_Get_AddressList.Size = $Size_Chk
	$chk_Org_Get_AddressList.TabIndex = $TabIndex++
	$chk_Org_Get_AddressList.Text = "Get-AddressList"
	$chk_Org_Get_AddressList.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_AddressList)
$chk_Org_Get_AntiPhishPolicy.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_AntiPhishPolicy.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_AntiPhishPolicy.Name = "chk_Org_Get_AntiPhishPolicy"
	$chk_Org_Get_AntiPhishPolicy.Size = $Size_Chk
	$chk_Org_Get_AntiPhishPolicy.TabIndex = $TabIndex++
	$chk_Org_Get_AntiPhishPolicy.Text = "Get-AntiPhishPolicy"
	$chk_Org_Get_AntiPhishPolicy.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_AntiPhishPolicy)
$chk_Org_Get_AntiSpoofingPolicy.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_AntiSpoofingPolicy.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_AntiSpoofingPolicy.Name = "chk_Org_Get_AntiSpoofingPolicy"
	$chk_Org_Get_AntiSpoofingPolicy.Size = $Size_Chk
	$chk_Org_Get_AntiSpoofingPolicy.TabIndex = $TabIndex++
	$chk_Org_Get_AntiSpoofingPolicy.Text = "Get-AntiSpoofingPolicy"
	$chk_Org_Get_AntiSpoofingPolicy.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_AntiSpoofingPolicy)
$chk_Org_Get_AtpPolicyForO365.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_AtpPolicyForO365.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_AtpPolicyForO365.Name = "chk_Org_Get_AtpPolicyForO365"
	$chk_Org_Get_AtpPolicyForO365.Size = $Size_Chk
	$chk_Org_Get_AtpPolicyForO365.TabIndex = $TabIndex++
	$chk_Org_Get_AtpPolicyForO365.Text = "Get-AtpPolicyForO365"
	$chk_Org_Get_AtpPolicyForO365.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_AtpPolicyForO365)
$chk_Org_Get_EmailAddressPolicy.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_EmailAddressPolicy.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_EmailAddressPolicy.Name = "chk_Org_Get_EmailAddressPolicy"
	$chk_Org_Get_EmailAddressPolicy.Size = $Size_Chk
	$chk_Org_Get_EmailAddressPolicy.TabIndex = $TabIndex++
	$chk_Org_Get_EmailAddressPolicy.Text = "Get-EmailAddressPolicy"
	$chk_Org_Get_EmailAddressPolicy.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_EmailAddressPolicy)
$chk_Org_Get_GlobalAddressList.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_GlobalAddressList.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_GlobalAddressList.Name = "chk_Org_Get_GlobalAddressList"
	$chk_Org_Get_GlobalAddressList.Size = $Size_Chk
	$chk_Org_Get_GlobalAddressList.TabIndex = $TabIndex++
	$chk_Org_Get_GlobalAddressList.Text = "Get-GlobalAddressList"
	$chk_Org_Get_GlobalAddressList.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_GlobalAddressList)
$chk_Org_Get_OfflineAddressBook.Font = $font_Calibri_10pt_normal
	$Row_2_loc += 25
	$chk_Org_Get_OfflineAddressBook.Location = New-Object System.Drawing.Point($Col_2_loc,$Row_2_loc)
	$chk_Org_Get_OfflineAddressBook.Name = "chk_Org_Get_OfflineAddressBook"
	$chk_Org_Get_OfflineAddressBook.Size = $Size_Chk
	$chk_Org_Get_OfflineAddressBook.TabIndex = $TabIndex++
	$chk_Org_Get_OfflineAddressBook.Text = "Get-OfflineAddressBook"
	$chk_Org_Get_OfflineAddressBook.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_OfflineAddressBook)
$chk_Org_Get_OnPremisesOrganization.Font = $font_Calibri_10pt_normal
	$Row_2_loc += 25
	$chk_Org_Get_OnPremisesOrganization.Location = New-Object System.Drawing.Point($Col_2_loc,$Row_2_loc)
	$chk_Org_Get_OnPremisesOrganization.Name = "chk_Org_Get_OnPremisesOrganization"
	$chk_Org_Get_OnPremisesOrganization.Size = $Size_Chk
	$chk_Org_Get_OnPremisesOrganization.TabIndex = $TabIndex++
	$chk_Org_Get_OnPremisesOrganization.Text = "Get-OnPremisesOrganization"
	$chk_Org_Get_OnPremisesOrganization.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_OnPremisesOrganization)
$chk_Org_Get_OrgConfig.Font = $font_Calibri_10pt_normal
	$Row_2_loc += 25
	$chk_Org_Get_OrgConfig.Location = New-Object System.Drawing.Point($Col_2_loc,$Row_2_loc)
	$chk_Org_Get_OrgConfig.Name = "chk_Org_Get_OrgConfig"
	$chk_Org_Get_OrgConfig.Size = $Size_Chk
	$chk_Org_Get_OrgConfig.TabIndex = $TabIndex++
	$chk_Org_Get_OrgConfig.Text = "Get-OrganizationConfig"
	$chk_Org_Get_OrgConfig.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_OrgConfig)
$chk_Org_Get_Rbac.Font = $font_Calibri_10pt_normal
	$Row_2_loc += 25
	$chk_Org_Get_Rbac.Location = New-Object System.Drawing.Point($Col_2_loc,$Row_2_loc)
	$chk_Org_Get_Rbac.Name = "chk_Org_Get_Rbac"
	$chk_Org_Get_Rbac.Size = $Size_Chk
	$chk_Org_Get_Rbac.TabIndex = $TabIndex++
	$chk_Org_Get_Rbac.Text = "Get-Rbac"
	$chk_Org_Get_Rbac.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_Rbac)
$chk_Org_Get_RetentionPolicy.Font = $font_Calibri_10pt_normal
	$Row_2_loc += 25
	$chk_Org_Get_RetentionPolicy.Location = New-Object System.Drawing.Point($Col_2_loc,$Row_2_loc)
	$chk_Org_Get_RetentionPolicy.Name = "chk_Org_Get_RetentionPolicy"
	$chk_Org_Get_RetentionPolicy.Size = $Size_Chk
	$chk_Org_Get_RetentionPolicy.TabIndex = $TabIndex++
	$chk_Org_Get_RetentionPolicy.Text = "Get-RetentionPolicy"
	$chk_Org_Get_RetentionPolicy.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_RetentionPolicy)
$chk_Org_Get_RetentionPolicyTag.Font = $font_Calibri_10pt_normal
	$Row_2_loc += 25
	$chk_Org_Get_RetentionPolicyTag.Location = New-Object System.Drawing.Point($Col_2_loc,$Row_2_loc)
	$chk_Org_Get_RetentionPolicyTag.Name = "chk_Org_Get_RetentionPolicyTag"
	$chk_Org_Get_RetentionPolicyTag.Size = $Size_Chk
	$chk_Org_Get_RetentionPolicyTag.TabIndex = $TabIndex++
	$chk_Org_Get_RetentionPolicyTag.Text = "Get-RetentionPolicyTag"
	$chk_Org_Get_RetentionPolicyTag.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_RetentionPolicyTag)
$chk_Org_Get_SmimeConfig.Font = $font_Calibri_10pt_normal
	$Row_2_loc += 25
	$chk_Org_Get_SmimeConfig.Location = New-Object System.Drawing.Point($Col_2_loc,$Row_2_loc)
	$chk_Org_Get_SmimeConfig.Name = "chk_Org_Get_SmimeConfig"
	$chk_Org_Get_SmimeConfig.Size = $Size_Chk
	$chk_Org_Get_SmimeConfig.TabIndex = $TabIndex++
	$chk_Org_Get_SmimeConfig.Text = "Get-SmimeConfig"
	$chk_Org_Get_SmimeConfig.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_SmimeConfig)
#EndRegion Step3 Exchange - Global tab

#Region Step3 Exchange - Recipient tab
$TabIndex = 0
$tab_Step3_Recipient.Location = $Loc_Tab_Tier3
	$tab_Step3_Recipient.Name = "tab_Step3_Recipient"
	$tab_Step3_Recipient.Padding = $System_Windows_Forms_Padding_Reusable
	$tab_Step3_Recipient.Size = $Size_Tab_Small
	$tab_Step3_Recipient.TabIndex = $TabIndex++
	$tab_Step3_Recipient.Text = "Recipient"
	$tab_Step3_Recipient.UseVisualStyleBackColor = $True
	$tab_Step3_Exchange_Tier2.Controls.Add($tab_Step3_Recipient)
$bx_Recipient_Functions.Dock = 5
	$bx_Recipient_Functions.Font = $font_Calibri_10pt_bold
	$bx_Recipient_Functions.Location = $Loc_Box_1
	$bx_Recipient_Functions.Name = "bx_Recipient_Functions"
	$bx_Recipient_Functions.Size = $Size_Box_3
	$bx_Recipient_Functions.TabIndex = $TabIndex++
	$bx_Recipient_Functions.TabStop = $False
	$tab_Step3_Recipient.Controls.Add($bx_Recipient_Functions)
$btn_Step3_Recipient_CheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_Recipient_CheckAll.Location = $Loc_CheckAll_3
	$btn_Step3_Recipient_CheckAll.Name = "btn_Step3_Recipient_CheckAll"
	$btn_Step3_Recipient_CheckAll.Size = $Size_Btn_3
	$btn_Step3_Recipient_CheckAll.TabIndex = $TabIndex++
	$btn_Step3_Recipient_CheckAll.Text = "Check all on this tab"
	$btn_Step3_Recipient_CheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_Recipient_CheckAll.add_Click($handler_btn_Step3_Recipient_CheckAll_Click)
	$bx_Recipient_Functions.Controls.Add($btn_Step3_Recipient_CheckAll)
$btn_Step3_Recipient_UncheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_Recipient_UncheckAll.Location = $Loc_UncheckAll_3
	$btn_Step3_Recipient_UncheckAll.Name = "btn_Step3_Recipient_UncheckAll"
	$btn_Step3_Recipient_UncheckAll.Size = $Size_Btn_3
	$btn_Step3_Recipient_UncheckAll.TabIndex = $TabIndex++
	$btn_Step3_Recipient_UncheckAll.Text = "Uncheck all on this tab"
	$btn_Step3_Recipient_UncheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_Recipient_UncheckAll.add_Click($handler_btn_Step3_Recipient_UncheckAll_Click)
	$bx_Recipient_Functions.Controls.Add($btn_Step3_Recipient_UncheckAll)
$Col_1_loc = 35
$Col_2_loc = 290
$Row_1_loc = 0
$Row_2_loc = 0
$chk_Org_Get_CalendarProcessing.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_CalendarProcessing.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_CalendarProcessing.Name = "chk_Org_Get_CalendarProcessing"
	$chk_Org_Get_CalendarProcessing.Size = $Size_Chk
	$chk_Org_Get_CalendarProcessing.TabIndex = $TabIndex++
	$chk_Org_Get_CalendarProcessing.Text = "Get-CalendarProcessing"
	$chk_Org_Get_CalendarProcessing.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_CalendarProcessing)
$chk_Org_Get_CASMailbox.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_CASMailbox.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_CASMailbox.Name = "chk_Org_Get_CASMailbox"
	$chk_Org_Get_CASMailbox.Size = $Size_Chk
	$chk_Org_Get_CASMailbox.TabIndex = $TabIndex++
	$chk_Org_Get_CASMailbox.Text = "Get-CASMailbox"
	$chk_Org_Get_CASMailbox.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_CASMailbox)
$chk_Org_Get_CasMailboxPlan.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_CasMailboxPlan.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_CasMailboxPlan.Name = "chk_Org_Get_CasMailboxPlan"
	$chk_Org_Get_CasMailboxPlan.Size = $Size_Chk
	$chk_Org_Get_CasMailboxPlan.TabIndex = $TabIndex++
	$chk_Org_Get_CasMailboxPlan.Text = "Get-CasMailboxPlan"
	$chk_Org_Get_CasMailboxPlan.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_CasMailboxPlan)
$chk_Org_Get_Contact.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_Contact.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_Contact.Name = "chk_Org_Get_Contact"
	$chk_Org_Get_Contact.Size = $Size_Chk
	$chk_Org_Get_Contact.TabIndex = $TabIndex++
	$chk_Org_Get_Contact.Text = "Get-Contact"
	$chk_Org_Get_Contact.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_Contact)
$chk_Org_Get_DistributionGroup.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_DistributionGroup.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_DistributionGroup.Name = "chk_Org_Get_DistributionGroup"
	$chk_Org_Get_DistributionGroup.Size = $Size_Chk
	$chk_Org_Get_DistributionGroup.TabIndex = $TabIndex++
	$chk_Org_Get_DistributionGroup.Text = "Get-DistributionGroup"
	$chk_Org_Get_DistributionGroup.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_DistributionGroup)
$chk_Org_Get_DynamicDistributionGroup.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_DynamicDistributionGroup.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_DynamicDistributionGroup.Name = "chk_Org_Get_DynamicDistributionGroup"
	$chk_Org_Get_DynamicDistributionGroup.Size = $Size_Chk
	$chk_Org_Get_DynamicDistributionGroup.TabIndex = $TabIndex++
	$chk_Org_Get_DynamicDistributionGroup.Text = "Get-DynamicDistributionGroup"
	$chk_Org_Get_DynamicDistributionGroup.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_DynamicDistributionGroup)
$chk_Org_Get_Mailbox.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_Mailbox.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_Mailbox.Name = "chk_Org_Get_Mailbox"
	$chk_Org_Get_Mailbox.Size = $Size_Chk
	$chk_Org_Get_Mailbox.TabIndex = $TabIndex++
	$chk_Org_Get_Mailbox.Text = "Get-Mailbox"
	$chk_Org_Get_Mailbox.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_Mailbox)
$chk_Org_Get_MailboxFolderStatistics.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_MailboxFolderStatistics.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_MailboxFolderStatistics.Name = "chk_Org_Get_MailboxFolderStatistics"
	$chk_Org_Get_MailboxFolderStatistics.Size = $Size_Chk
	$chk_Org_Get_MailboxFolderStatistics.TabIndex = $TabIndex++
	$chk_Org_Get_MailboxFolderStatistics.Text = "Get-MailboxFolderStatistics"
	$chk_Org_Get_MailboxFolderStatistics.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_MailboxFolderStatistics)
$chk_Org_Get_MailboxPermission.Font = $font_Calibri_10pt_normal
	$Row_2_loc += 25
	$chk_Org_Get_MailboxPermission.Location = New-Object System.Drawing.Point($Col_2_loc,$Row_2_loc)
	$chk_Org_Get_MailboxPermission.Name = "chk_Org_Get_MailboxPermission"
	$chk_Org_Get_MailboxPermission.Size = $Size_Chk
	$chk_Org_Get_MailboxPermission.TabIndex = $TabIndex++
	$chk_Org_Get_MailboxPermission.Text = "Get-MailboxPermission"
	$chk_Org_Get_MailboxPermission.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_MailboxPermission)
$chk_Org_Get_MailboxPlan.Font = $font_Calibri_10pt_normal
	$Row_2_loc += 25
	$chk_Org_Get_MailboxPlan.Location = New-Object System.Drawing.Point($Col_2_loc,$Row_2_loc)
	$chk_Org_Get_MailboxPlan.Name = "chk_Org_Get_MailboxPlan"
	$chk_Org_Get_MailboxPlan.Size = $Size_Chk
	$chk_Org_Get_MailboxPlan.TabIndex = $TabIndex++
	$chk_Org_Get_MailboxPlan.Text = "Get-MailboxPlan"
	$chk_Org_Get_MailboxPlan.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_MailboxPlan)
$chk_Org_Get_MailboxStatistics.Font = $font_Calibri_10pt_normal
	$Row_2_loc += 25
	$chk_Org_Get_MailboxStatistics.Location = New-Object System.Drawing.Point($Col_2_loc,$Row_2_loc)
	$chk_Org_Get_MailboxStatistics.Name = "chk_Org_Get_MailboxStatistics"
	$chk_Org_Get_MailboxStatistics.Size = $Size_Chk
	$chk_Org_Get_MailboxStatistics.TabIndex = $TabIndex++
	$chk_Org_Get_MailboxStatistics.Text = "Get-MailboxStatistics"
	$chk_Org_Get_MailboxStatistics.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_MailboxStatistics)
$chk_Org_Get_MailUser.Font = $font_Calibri_10pt_normal
	$Row_2_loc += 25
	$chk_Org_Get_MailUser.Location = New-Object System.Drawing.Point($Col_2_loc,$Row_2_loc)
	$chk_Org_Get_MailUser.Name = "chk_Org_Get_MailUser"
	$chk_Org_Get_MailUser.Size = $Size_Chk
	$chk_Org_Get_MailUser.TabIndex = $TabIndex++
	$chk_Org_Get_MailUser.Text = "Get-MailUser"
	$chk_Org_Get_MailUser.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_MailUser)
$chk_Org_Get_PublicFolder.Font = $font_Calibri_10pt_normal
	$Row_2_loc += 25
	$chk_Org_Get_PublicFolder.Location = New-Object System.Drawing.Point($Col_2_loc,$Row_2_loc)
	$chk_Org_Get_PublicFolder.Name = "chk_Org_Get_PublicFolder"
	$chk_Org_Get_PublicFolder.Size = $Size_Chk
	$chk_Org_Get_PublicFolder.TabIndex = $TabIndex++
	$chk_Org_Get_PublicFolder.Text = "Get-PublicFolder"
	$chk_Org_Get_PublicFolder.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_PublicFolder)
$chk_Org_Get_PublicFolderStatistics.Font = $font_Calibri_10pt_normal
	$Row_2_loc += 25
	$chk_Org_Get_PublicFolderStatistics.Location = New-Object System.Drawing.Point($Col_2_loc,$Row_2_loc)
	$chk_Org_Get_PublicFolderStatistics.Name = "chk_Org_Get_PublicFolderStatistics"
	$chk_Org_Get_PublicFolderStatistics.Size = $Size_Chk
	$chk_Org_Get_PublicFolderStatistics.TabIndex = $TabIndex++
	$chk_Org_Get_PublicFolderStatistics.Text = "Get-PublicFolderStatistics"
	$chk_Org_Get_PublicFolderStatistics.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_PublicFolderStatistics)
$chk_Org_Get_UnifiedGroup.Font = $font_Calibri_10pt_normal
	$Row_2_loc += 25
	$chk_Org_Get_UnifiedGroup.Location = New-Object System.Drawing.Point($Col_2_loc,$Row_2_loc)
	$chk_Org_Get_UnifiedGroup.Name = "chk_Org_Get_UnifiedGroup"
	$chk_Org_Get_UnifiedGroup.Size = $Size_Chk
	$chk_Org_Get_UnifiedGroup.TabIndex = $TabIndex++
	$chk_Org_Get_UnifiedGroup.Text = "Get-UnifiedGroup"
	$chk_Org_Get_UnifiedGroup.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_UnifiedGroup)
$chk_Org_Quota.Font = $font_Calibri_10pt_normal
	$Row_2_loc += 25
	$chk_Org_Quota.Location = New-Object System.Drawing.Point($Col_2_loc,$Row_2_loc)
	$chk_Org_Quota.Name = "chk_Org_Quota"
	$chk_Org_Quota.Size = $Size_Chk
	$chk_Org_Quota.TabIndex = $TabIndex++
	$chk_Org_Quota.Text = "Quota"
	$chk_Org_Quota.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Quota)
#EndRegion Step3 Exchange - Recipient tab

#Region Step3 Exchange - Transport tab
$TabIndex = 0
$tab_Step3_Transport.Location = $Loc_Tab_Tier3
	$tab_Step3_Transport.Name = "tab_Step3_Transport"
	$tab_Step3_Transport.Padding = $System_Windows_Forms_Padding_Reusable
	$tab_Step3_Transport.Size = $Size_Tab_Small
	$tab_Step3_Transport.TabIndex = $TabIndex++
	$tab_Step3_Transport.Text = "Transport"
	$tab_Step3_Transport.UseVisualStyleBackColor = $True
	$tab_Step3_Exchange_Tier2.Controls.Add($tab_Step3_Transport)
$bx_Transport_Functions.Dock = 5
	$bx_Transport_Functions.Font = $font_Calibri_10pt_bold
	$bx_Transport_Functions.Location = $Loc_Box_1
	$bx_Transport_Functions.Name = "bx_Transport_Functions"
	$bx_Transport_Functions.Size = $Size_Box_3
	$bx_Transport_Functions.TabIndex = $TabIndex++
	$bx_Transport_Functions.TabStop = $False
	$tab_Step3_Transport.Controls.Add($bx_Transport_Functions)
$btn_Step3_Transport_CheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_Transport_CheckAll.Location = $Loc_CheckAll_3
	$btn_Step3_Transport_CheckAll.Name = "btn_Step3_Transport_CheckAll"
	$btn_Step3_Transport_CheckAll.Size = $Size_Btn_3
	$btn_Step3_Transport_CheckAll.TabIndex = $TabIndex++
	$btn_Step3_Transport_CheckAll.Text = "Check all on this tab"
	$btn_Step3_Transport_CheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_Transport_CheckAll.add_Click($handler_btn_Step3_Transport_CheckAll_Click)
	$bx_Transport_Functions.Controls.Add($btn_Step3_Transport_CheckAll)
$btn_Step3_Transport_UncheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_Transport_UncheckAll.Location = $Loc_UncheckAll_3
	$btn_Step3_Transport_UncheckAll.Name = "btn_Step3_Transport_UncheckAll"
	$btn_Step3_Transport_UncheckAll.Size = $Size_Btn_3
	$btn_Step3_Transport_UncheckAll.TabIndex = $TabIndex++
	$btn_Step3_Transport_UncheckAll.Text = "Uncheck all on this tab"
	$btn_Step3_Transport_UncheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_Transport_UncheckAll.add_Click($handler_btn_Step3_Transport_UncheckAll_Click)
	$bx_Transport_Functions.Controls.Add($btn_Step3_Transport_UncheckAll)
$Col_1_loc = 35
$Col_2_loc = 290
$Row_1_loc = 0
$Row_2_loc = 0
$chk_Org_Get_AcceptedDomain.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_AcceptedDomain.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_AcceptedDomain.Name = "chk_Org_Get_AcceptedDomain"
	$chk_Org_Get_AcceptedDomain.Size = $Size_Chk
	$chk_Org_Get_AcceptedDomain.TabIndex = $TabIndex++
	$chk_Org_Get_AcceptedDomain.Text = "Get-AcceptedDomain"
	$chk_Org_Get_AcceptedDomain.UseVisualStyleBackColor = $True
	$bx_Transport_Functions.Controls.Add($chk_Org_Get_AcceptedDomain)
$chk_Org_Get_DkimSigningConfig.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_DkimSigningConfig.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_DkimSigningConfig.Name = "chk_Org_Get_DkimSigningConfig"
	$chk_Org_Get_DkimSigningConfig.Size = $Size_Chk
	$chk_Org_Get_DkimSigningConfig.TabIndex = $TabIndex++
	$chk_Org_Get_DkimSigningConfig.Text = "Get-DkimSigningConfig"
	$chk_Org_Get_DkimSigningConfig.UseVisualStyleBackColor = $True
	$bx_Transport_Functions.Controls.Add($chk_Org_Get_DkimSigningConfig)
$chk_Org_Get_InboundConnector.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_InboundConnector.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_InboundConnector.Name = "chk_Org_Get_InboundConnector"
	$chk_Org_Get_InboundConnector.Size = $Size_Chk
	$chk_Org_Get_InboundConnector.TabIndex = $TabIndex++
	$chk_Org_Get_InboundConnector.Text = "Get-InboundConnector"
	$chk_Org_Get_InboundConnector.UseVisualStyleBackColor = $True
	$bx_Transport_Functions.Controls.Add($chk_Org_Get_InboundConnector)
$chk_Org_Get_OutboundConnector.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_OutboundConnector.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_OutboundConnector.Name = "chk_Org_Get_OutboundConnector"
	$chk_Org_Get_OutboundConnector.Size = $Size_Chk
	$chk_Org_Get_OutboundConnector.TabIndex = $TabIndex++
	$chk_Org_Get_OutboundConnector.Text = "Get-OutboundConnector"
	$chk_Org_Get_OutboundConnector.UseVisualStyleBackColor = $True
	$bx_Transport_Functions.Controls.Add($chk_Org_Get_OutboundConnector)
$chk_Org_Get_RemoteDomain.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_RemoteDomain.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_RemoteDomain.Name = "chk_Org_Get_RemoteDomain"
	$chk_Org_Get_RemoteDomain.Size = $Size_Chk
	$chk_Org_Get_RemoteDomain.TabIndex = $TabIndex++
	$chk_Org_Get_RemoteDomain.Text = "Get-RemoteDomain"
	$chk_Org_Get_RemoteDomain.UseVisualStyleBackColor = $True
	$bx_Transport_Functions.Controls.Add($chk_Org_Get_RemoteDomain)
$chk_Org_Get_TransportConfig.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_TransportConfig.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_TransportConfig.Name = "chk_Org_Get_TransportConfig"
	$chk_Org_Get_TransportConfig.Size = $Size_Chk
	$chk_Org_Get_TransportConfig.TabIndex = $TabIndex++
	$chk_Org_Get_TransportConfig.Text = "Get-TransportConfig"
	$chk_Org_Get_TransportConfig.UseVisualStyleBackColor = $True
	$bx_Transport_Functions.Controls.Add($chk_Org_Get_TransportConfig)
$chk_Org_Get_TransportRule.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_TransportRule.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_TransportRule.Name = "chk_Org_Get_TransportRule"
	$chk_Org_Get_TransportRule.Size = $Size_Chk
	$chk_Org_Get_TransportRule.TabIndex = $TabIndex++
	$chk_Org_Get_TransportRule.Text = "Get-TransportRule"
	$chk_Org_Get_TransportRule.UseVisualStyleBackColor = $True
	$bx_Transport_Functions.Controls.Add($chk_Org_Get_TransportRule)
#EndRegion Step3 Exchange - Transport tab

#Region Step3 Exchange - UM tab
$TabIndex = 0
$tab_Step3_UM.Location = $Loc_Tab_Tier3
	$tab_Step3_UM.Name = "tab_Step3_Misc"
	$tab_Step3_UM.Padding = $System_Windows_Forms_Padding_Reusable
	$tab_Step3_UM.Size = $Size_Tab_Small #New-Object System.Drawing.Size(300,488)
	$tab_Step3_UM.TabIndex = $TabIndex++
	$tab_Step3_UM.Text = "    UM"
	$tab_Step3_UM.UseVisualStyleBackColor = $True
	$tab_Step3_Exchange_Tier2.Controls.Add($tab_Step3_UM)
$bx_UM_Functions.Dock = 5
	$bx_UM_Functions.Font = $font_Calibri_10pt_bold
	$bx_UM_Functions.Location = $Loc_Box_1
	$bx_UM_Functions.Name = "bx_Misc_Functions"
	$bx_UM_Functions.Size = $Size_Box_3
	$bx_UM_Functions.TabIndex = $TabIndex++
	$bx_UM_Functions.TabStop = $False
	$tab_Step3_UM.Controls.Add($bx_UM_Functions)
$btn_Step3_UM_CheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_UM_CheckAll.Location = $Loc_CheckAll_3
	$btn_Step3_UM_CheckAll.Name = "btn_Step3_Misc_CheckAll"
	$btn_Step3_UM_CheckAll.Size = $Size_Btn_3
	$btn_Step3_UM_CheckAll.TabIndex = $TabIndex++
	$btn_Step3_UM_CheckAll.Text = "Check all on this tab"
	$btn_Step3_UM_CheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_UM_CheckAll.add_Click($handler_btn_Step3_UM_CheckAll_Click)
	$bx_UM_Functions.Controls.Add($btn_Step3_UM_CheckAll)
$btn_Step3_UM_UncheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_UM_UncheckAll.Location = $Loc_UncheckAll_3
	$btn_Step3_UM_UncheckAll.Name = "btn_Step3_Misc_UncheckAll"
	$btn_Step3_UM_UncheckAll.Size = $Size_Btn_3
	$btn_Step3_UM_UncheckAll.TabIndex = $TabIndex++
	$btn_Step3_UM_UncheckAll.Text = "Uncheck all on this tab"
	$btn_Step3_UM_UncheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_UM_UncheckAll.add_Click($handler_btn_Step3_UM_UncheckAll_Click)
	$bx_UM_Functions.Controls.Add($btn_Step3_UM_UncheckAll)
$Col_1_loc = 35
$Col_2_loc = 290
$Row_1_loc = 0
$Row_2_loc = 0
$chk_Org_Get_UmAutoAttendant.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_UmAutoAttendant.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_UmAutoAttendant.Name = "chk_Org_Get_UmAutoAttendant"
	$chk_Org_Get_UmAutoAttendant.Size = $Size_Chk
	$chk_Org_Get_UmAutoAttendant.TabIndex = $TabIndex++
	$chk_Org_Get_UmAutoAttendant.Text = "Get-UmAutoAttendant"
	$chk_Org_Get_UmAutoAttendant.UseVisualStyleBackColor = $True
	$bx_UM_Functions.Controls.Add($chk_Org_Get_UmAutoAttendant)
$chk_Org_Get_UmDialPlan.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_UmDialPlan.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_UmDialPlan.Name = "chk_Org_Get_UmDialPlan"
	$chk_Org_Get_UmDialPlan.Size = $Size_Chk
	$chk_Org_Get_UmDialPlan.TabIndex = $TabIndex++
	$chk_Org_Get_UmDialPlan.Text = "Get-UmDialPlan"
	$chk_Org_Get_UmDialPlan.UseVisualStyleBackColor = $True
	$bx_UM_Functions.Controls.Add($chk_Org_Get_UmDialPlan)
$chk_Org_Get_UmIpGateway.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_UmIpGateway.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_UmIpGateway.Name = "chk_Org_Get_UmIpGateway"
	$chk_Org_Get_UmIpGateway.Size = $Size_Chk
	$chk_Org_Get_UmIpGateway.TabIndex = $TabIndex++
	$chk_Org_Get_UmIpGateway.Text = "Get-UmIpGateway"
	$chk_Org_Get_UmIpGateway.UseVisualStyleBackColor = $True
	$bx_UM_Functions.Controls.Add($chk_Org_Get_UmIpGateway)
$chk_Org_Get_UmMailbox.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_UmMailbox.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_UmMailbox.Name = "chk_Org_Get_UmMailbox"
	$chk_Org_Get_UmMailbox.Size = $Size_Chk
	$chk_Org_Get_UmMailbox.TabIndex = $TabIndex++
	$chk_Org_Get_UmMailbox.Text = "Get-UmMailbox"
	$chk_Org_Get_UmMailbox.UseVisualStyleBackColor = $True
	$bx_UM_Functions.Controls.Add($chk_Org_Get_UmMailbox)
#	$chk_Org_Get_UmMailboxConfiguration.Font = $font_Calibri_10pt_normal
#		$System_Drawing_Point = New-Object System.Drawing.Point
#		$System_Drawing_Point.X = $Col_1_loc
#		$System_Drawing_Point.Y = $Row_1_loc
#		$Row_1_loc += 25
#	$chk_Org_Get_UmMailboxConfiguration.Location = $System_Drawing_Point
#	$chk_Org_Get_UmMailboxConfiguration.Name = "chk_Org_Get_UmMailboxConfiguration"
#	$chk_Org_Get_UmMailboxConfiguration.Size = $Size_Chk
#	$chk_Org_Get_UmMailboxConfiguration.TabIndex = $TabIndex++
#	$chk_Org_Get_UmMailboxConfiguration.Text = "Get-UmMailboxConfiguration"
#	$chk_Org_Get_UmMailboxConfiguration.UseVisualStyleBackColor = $True
#	$bx_UM_Functions.Controls.Add($chk_Org_Get_UmMailboxConfiguration)
#	$chk_Org_Get_UmMailboxPin.Font = $font_Calibri_10pt_normal
#		$System_Drawing_Point = New-Object System.Drawing.Point
#		$System_Drawing_Point.X = $Col_1_loc
#		$System_Drawing_Point.Y = $Row_1_loc
#		$Row_1_loc += 25
#	$chk_Org_Get_UmMailboxPin.Location = $System_Drawing_Point
#	$chk_Org_Get_UmMailboxPin.Name = "chk_Org_Get_UmMailboxPin"
#	$chk_Org_Get_UmMailboxPin.Size = $Size_Chk
#	$chk_Org_Get_UmMailboxPin.TabIndex = $TabIndex++
#	$chk_Org_Get_UmMailboxPin.Text = "Get-UmMailboxPin"
#	$chk_Org_Get_UmMailboxPin.UseVisualStyleBackColor = $True
#	$bx_UM_Functions.Controls.Add($chk_Org_Get_UmMailboxPin)
$chk_Org_Get_UmMailboxPolicy.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_UmMailboxPolicy.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_UmMailboxPolicy.Name = "chk_Org_Get_UmMailboxPolicy"
	$chk_Org_Get_UmMailboxPolicy.Size = $Size_Chk
	$chk_Org_Get_UmMailboxPolicy.TabIndex = $TabIndex++
	$chk_Org_Get_UmMailboxPolicy.Text = "Get-UmMailboxPolicy"
	$chk_Org_Get_UmMailboxPolicy.UseVisualStyleBackColor = $True
	$bx_UM_Functions.Controls.Add($chk_Org_Get_UmMailboxPolicy)
#EndRegion Step3 Exchange - UM tab

#Region Step3 Exchange - Misc tab
$TabIndex = 0
$tab_Step3_Misc.Location = $Loc_Tab_Tier3
	$tab_Step3_Misc.Name = "tab_Step3_Misc"
	$tab_Step3_Misc.Padding = $System_Windows_Forms_Padding_Reusable
	$tab_Step3_Misc.Size = $Size_Tab_Small
	$tab_Step3_Misc.TabIndex = $TabIndex++
	$tab_Step3_Misc.Text = "Misc"
	$tab_Step3_Misc.UseVisualStyleBackColor = $True
	$tab_Step3_Exchange_Tier2.Controls.Add($tab_Step3_Misc)
$bx_Misc_Functions.Dock = 5
	$bx_Misc_Functions.Font = $font_Calibri_10pt_bold
	$bx_Misc_Functions.Location = $Loc_Box_1
	$bx_Misc_Functions.Name = "bx_Misc_Functions"
	$bx_Misc_Functions.Size = $Size_Box_3
	$bx_Misc_Functions.TabIndex = $TabIndex++
	$bx_Misc_Functions.TabStop = $False
	$tab_Step3_Misc.Controls.Add($bx_Misc_Functions)
$btn_Step3_Misc_CheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_Misc_CheckAll.Location = $Loc_CheckAll_3
	$btn_Step3_Misc_CheckAll.Name = "btn_Step3_Misc_CheckAll"
	$btn_Step3_Misc_CheckAll.Size = $Size_Btn_3
	$btn_Step3_Misc_CheckAll.TabIndex = $TabIndex++
	$btn_Step3_Misc_CheckAll.Text = "Check all on this tab"
	$btn_Step3_Misc_CheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_Misc_CheckAll.add_Click($handler_btn_Step3_Misc_CheckAll_Click)
	$bx_Misc_Functions.Controls.Add($btn_Step3_Misc_CheckAll)
$btn_Step3_Misc_UncheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_Misc_UncheckAll.Location = $Loc_UncheckAll_3
	$btn_Step3_Misc_UncheckAll.Name = "btn_Step3_Misc_UncheckAll"
	$btn_Step3_Misc_UncheckAll.Size = $Size_Btn_3
	$btn_Step3_Misc_UncheckAll.TabIndex = $TabIndex++
	$btn_Step3_Misc_UncheckAll.Text = "Uncheck all on this tab"
	$btn_Step3_Misc_UncheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_Misc_UncheckAll.add_Click($handler_btn_Step3_Misc_UncheckAll_Click)
	$bx_Misc_Functions.Controls.Add($btn_Step3_Misc_UncheckAll)
$Col_1_loc = 35
$Col_2_loc = 290
$Row_1_loc = 0
$Row_2_loc = 0
$chk_Org_Get_AdminGroups.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_AdminGroups.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_AdminGroups.Name = "chk_Org_Get_AdminGroups"
	$chk_Org_Get_AdminGroups.Size = $Size_Chk
	$chk_Org_Get_AdminGroups.TabIndex = $TabIndex++
	$chk_Org_Get_AdminGroups.Text = "Get memberships of admin groups"
	$chk_Org_Get_AdminGroups.UseVisualStyleBackColor = $True
	$bx_Misc_Functions.Controls.Add($chk_Org_Get_AdminGroups)
#EndRegion Step3 Exchange - Misc tab

#Region Step3 Azure - AzureAD tab
$TabIndex = 0
$tab_Step3_AzureAd.Location = $Loc_Tab_Tier3
	$tab_Step3_AzureAd.Name = "tab_Step3_AzureAd"
	$tab_Step3_AzureAd.Padding = $System_Windows_Forms_Padding_Reusable
	$tab_Step3_AzureAd.Size = $Size_Tab_Small
	$tab_Step3_AzureAd.TabIndex = $TabIndex++
	$tab_Step3_AzureAd.Text = "Get-AzureAD"
	$tab_Step3_AzureAd.UseVisualStyleBackColor = $True
	$tab_Step3_Azure_Tier2.Controls.Add($tab_Step3_AzureAd)
$bx_AzureAd_Functions.Dock = 5
	$bx_AzureAd_Functions.Font = $font_Calibri_10pt_bold
	$bx_AzureAd_Functions.Location = $Loc_Box_1
	$bx_AzureAd_Functions.Name = "bx_AzureAd_Functions"
	$bx_AzureAd_Functions.Size = $Size_Box_3
	$bx_AzureAd_Functions.TabIndex = $TabIndex++
	$bx_AzureAd_Functions.TabStop = $False
	$tab_Step3_AzureAd.Controls.Add($bx_AzureAd_Functions)
$btn_Step3_AzureAd_CheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_AzureAd_CheckAll.Location = $Loc_CheckAll_3
	$btn_Step3_AzureAd_CheckAll.Name = "btn_Step3_AzureAd_CheckAll"
	$btn_Step3_AzureAd_CheckAll.Size = $Size_Btn_3
	$btn_Step3_AzureAd_CheckAll.TabIndex = $TabIndex++
	$btn_Step3_AzureAd_CheckAll.Text = "Check all on this tab"
	$btn_Step3_AzureAd_CheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_AzureAd_CheckAll.add_Click($handler_btn_Step3_AzureAd_CheckAll_Click)
	$bx_AzureAd_Functions.Controls.Add($btn_Step3_AzureAd_CheckAll)
	$btn_Step3_AzureAd_UncheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_AzureAd_UncheckAll.Location = $Loc_UncheckAll_3
	$btn_Step3_AzureAd_UncheckAll.Name = "btn_Step3_AzureAd_UncheckAll"
	$btn_Step3_AzureAd_UncheckAll.Size = $Size_Btn_3
	$btn_Step3_AzureAd_UncheckAll.TabIndex = $TabIndex++
	$btn_Step3_AzureAd_UncheckAll.Text = "Uncheck all on this tab"
	$btn_Step3_AzureAd_UncheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_AzureAd_UncheckAll.add_Click($handler_btn_Step3_AzureAd_UncheckAll_Click)
	$bx_AzureAd_Functions.Controls.Add($btn_Step3_AzureAd_UncheckAll)
$Col_1_loc = 35
$Col_2_loc = 290
$Row_1_loc = 0
$Row_2_loc = 0
$chk_Azure_Get_AzureADApplication.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Azure_Get_AzureADApplication.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Azure_Get_AzureADApplication.Name = "chk_Azure_Get_AzureADApplication"
	$chk_Azure_Get_AzureADApplication.Size = $Size_Chk
	$chk_Azure_Get_AzureADApplication.TabIndex = $TabIndex++
	$chk_Azure_Get_AzureADApplication.Text = "Application"
	$chk_Azure_Get_AzureADApplication.UseVisualStyleBackColor = $True
	$bx_AzureAd_Functions.Controls.Add($chk_Azure_Get_AzureADApplication)
$chk_Azure_Get_AzureADContact.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Azure_Get_AzureADContact.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Azure_Get_AzureADContact.Name = "chk_Azure_Get_AzureADContact"
	$chk_Azure_Get_AzureADContact.Size = $Size_Chk
	$chk_Azure_Get_AzureADContact.TabIndex = $TabIndex++
	$chk_Azure_Get_AzureADContact.Text = "Contact"
	$chk_Azure_Get_AzureADContact.UseVisualStyleBackColor = $True
	$bx_AzureAd_Functions.Controls.Add($chk_Azure_Get_AzureADContact)
$chk_Azure_Get_AzureADDevice.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Azure_Get_AzureADDevice.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Azure_Get_AzureADDevice.Name = "chk_Azure_Get_AzureADDevice"
	$chk_Azure_Get_AzureADDevice.Size = $Size_Chk
	$chk_Azure_Get_AzureADDevice.TabIndex = $TabIndex++
	$chk_Azure_Get_AzureADDevice.Text = "Device"
	$chk_Azure_Get_AzureADDevice.UseVisualStyleBackColor = $True
	$bx_AzureAd_Functions.Controls.Add($chk_Azure_Get_AzureADDevice)
$chk_Azure_Get_AzureADDeviceRegisteredOwner.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Azure_Get_AzureADDeviceRegisteredOwner.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Azure_Get_AzureADDeviceRegisteredOwner.Name = "chk_Azure_Get_AzureADDeviceRegisteredOwner"
	$chk_Azure_Get_AzureADDeviceRegisteredOwner.Size = $Size_Chk
	$chk_Azure_Get_AzureADDeviceRegisteredOwner.TabIndex = $TabIndex++
	$chk_Azure_Get_AzureADDeviceRegisteredOwner.Text = "DeviceRegisteredOwner"
	$chk_Azure_Get_AzureADDeviceRegisteredOwner.UseVisualStyleBackColor = $True
	$bx_AzureAd_Functions.Controls.Add($chk_Azure_Get_AzureADDeviceRegisteredOwner)
$chk_Azure_Get_AzureADDeviceRegisteredUser.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Azure_Get_AzureADDeviceRegisteredUser.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Azure_Get_AzureADDeviceRegisteredUser.Name = "chk_Azure_Get_AzureADDeviceRegisteredUser"
	$chk_Azure_Get_AzureADDeviceRegisteredUser.Size = $Size_Chk
	$chk_Azure_Get_AzureADDeviceRegisteredUser.TabIndex = $TabIndex++
	$chk_Azure_Get_AzureADDeviceRegisteredUser.Text = "DeviceRegisteredUser"
	$chk_Azure_Get_AzureADDeviceRegisteredUser.UseVisualStyleBackColor = $True
	$bx_AzureAd_Functions.Controls.Add($chk_Azure_Get_AzureADDeviceRegisteredUser)
$chk_Azure_Get_AzureADDirectoryRole.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Azure_Get_AzureADDirectoryRole.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Azure_Get_AzureADDirectoryRole.Name = "chk_Azure_Get_AzureADDirectoryRole"
	$chk_Azure_Get_AzureADDirectoryRole.Size = $Size_Chk
	$chk_Azure_Get_AzureADDirectoryRole.TabIndex = $TabIndex++
	$chk_Azure_Get_AzureADDirectoryRole.Text = "DirectoryRole"
	$chk_Azure_Get_AzureADDirectoryRole.UseVisualStyleBackColor = $True
	$bx_AzureAd_Functions.Controls.Add($chk_Azure_Get_AzureADDirectoryRole)
$chk_Azure_Get_AzureADDomain.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Azure_Get_AzureADDomain.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Azure_Get_AzureADDomain.Name = "chk_Azure_Get_AzureADDomain"
	$chk_Azure_Get_AzureADDomain.Size = $Size_Chk
	$chk_Azure_Get_AzureADDomain.TabIndex = $TabIndex++
	$chk_Azure_Get_AzureADDomain.Text = "Domain"
	$chk_Azure_Get_AzureADDomain.UseVisualStyleBackColor = $True
	$bx_AzureAd_Functions.Controls.Add($chk_Azure_Get_AzureADDomain)
$chk_Azure_Get_AzureAdDomainServiceConfigurationRecord.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Azure_Get_AzureAdDomainServiceConfigurationRecord.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Azure_Get_AzureAdDomainServiceConfigurationRecord.Name = "chk_Azure_Get_AzureAdDomainServiceConfigurationRecord"
	$chk_Azure_Get_AzureAdDomainServiceConfigurationRecord.Size = $Size_Chk
	$chk_Azure_Get_AzureAdDomainServiceConfigurationRecord.TabIndex = $TabIndex++
	$chk_Azure_Get_AzureAdDomainServiceConfigurationRecord.Text = "DomainServiceConfigurationRecord"
	$chk_Azure_Get_AzureAdDomainServiceConfigurationRecord.UseVisualStyleBackColor = $True
	$bx_AzureAd_Functions.Controls.Add($chk_Azure_Get_AzureAdDomainServiceConfigurationRecord)
$chk_Azure_Get_AzureAdDomainVerificationDnsRecord.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Azure_Get_AzureAdDomainVerificationDnsRecord.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Azure_Get_AzureAdDomainVerificationDnsRecord.Name = "chk_Azure_Get_AzureAdDomainVerificationDnsRecord"
	$chk_Azure_Get_AzureAdDomainVerificationDnsRecord.Size = $Size_Chk
	$chk_Azure_Get_AzureAdDomainVerificationDnsRecord.TabIndex = $TabIndex++
	$chk_Azure_Get_AzureAdDomainVerificationDnsRecord.Text = "DomainVerificationDnsRecord"
	$chk_Azure_Get_AzureAdDomainVerificationDnsRecord.UseVisualStyleBackColor = $True
	$bx_AzureAd_Functions.Controls.Add($chk_Azure_Get_AzureAdDomainVerificationDnsRecord)
$chk_Azure_Get_AzureADGroup.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Azure_Get_AzureADGroup.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Azure_Get_AzureADGroup.Name = "chk_Azure_Get_AzureADGroup"
	$chk_Azure_Get_AzureADGroup.Size = $Size_Chk
	$chk_Azure_Get_AzureADGroup.TabIndex = $TabIndex++
	$chk_Azure_Get_AzureADGroup.Text = "Group"
	$chk_Azure_Get_AzureADGroup.UseVisualStyleBackColor = $True
	$bx_AzureAd_Functions.Controls.Add($chk_Azure_Get_AzureADGroup)
$chk_Azure_Get_AzureADGroupMember.Font = $font_Calibri_10pt_normal
	$Row_2_loc += 25
	$chk_Azure_Get_AzureADGroupMember.Location = New-Object System.Drawing.Point($Col_2_loc,$Row_2_loc)
	$chk_Azure_Get_AzureADGroupMember.Name = "chk_Azure_Get_AzureADGroupMember"
	$chk_Azure_Get_AzureADGroupMember.Size = $Size_Chk
	$chk_Azure_Get_AzureADGroupMember.TabIndex = $TabIndex++
	$chk_Azure_Get_AzureADGroupMember.Text = "GroupMember"
	$chk_Azure_Get_AzureADGroupMember.UseVisualStyleBackColor = $True
	$bx_AzureAd_Functions.Controls.Add($chk_Azure_Get_AzureADGroupMember)
$chk_Azure_Get_AzureADGroupOwner.Font = $font_Calibri_10pt_normal
	$Row_2_loc += 25
	$chk_Azure_Get_AzureADGroupOwner.Location = New-Object System.Drawing.Point($Col_2_loc,$Row_2_loc)
	$chk_Azure_Get_AzureADGroupOwner.Name = "chk_Azure_Get_AzureADGroupOwner"
	$chk_Azure_Get_AzureADGroupOwner.Size = $Size_Chk
	$chk_Azure_Get_AzureADGroupOwner.TabIndex = $TabIndex++
	$chk_Azure_Get_AzureADGroupOwner.Text = "GroupOwner"
	$chk_Azure_Get_AzureADGroupOwner.UseVisualStyleBackColor = $True
	$bx_AzureAd_Functions.Controls.Add($chk_Azure_Get_AzureADGroupOwner)
$chk_Azure_Get_AzureADSubscribedSku.Font = $font_Calibri_10pt_normal
	$Row_2_loc += 25
	$chk_Azure_Get_AzureADSubscribedSku.Location = New-Object System.Drawing.Point($Col_2_loc,$Row_2_loc)
	$chk_Azure_Get_AzureADSubscribedSku.Name = "chk_Azure_Get_AzureADSubscribedSku"
	$chk_Azure_Get_AzureADSubscribedSku.Size = $Size_Chk
	$chk_Azure_Get_AzureADSubscribedSku.TabIndex = $TabIndex++
	$chk_Azure_Get_AzureADSubscribedSku.Text = "SubscribedSku"
	$chk_Azure_Get_AzureADSubscribedSku.UseVisualStyleBackColor = $True
	$bx_AzureAd_Functions.Controls.Add($chk_Azure_Get_AzureADSubscribedSku)
$chk_Azure_Get_AzureADTenantDetail.Font = $font_Calibri_10pt_normal
	$Row_2_loc += 25
	$chk_Azure_Get_AzureADTenantDetail.Location = New-Object System.Drawing.Point($Col_2_loc,$Row_2_loc)
	$chk_Azure_Get_AzureADTenantDetail.Name = "chk_Azure_Get_AzureADTenantDetail"
	$chk_Azure_Get_AzureADTenantDetail.Size = $Size_Chk
	$chk_Azure_Get_AzureADTenantDetail.TabIndex = $TabIndex++
	$chk_Azure_Get_AzureADTenantDetail.Text = "TenantDetail"
	$chk_Azure_Get_AzureADTenantDetail.UseVisualStyleBackColor = $True
	$bx_AzureAd_Functions.Controls.Add($chk_Azure_Get_AzureADTenantDetail)
$chk_Azure_Get_AzureADUser.Font = $font_Calibri_10pt_normal
	$Row_2_loc += 25
	$chk_Azure_Get_AzureADUser.Location = New-Object System.Drawing.Point($Col_2_loc,$Row_2_loc)
	$chk_Azure_Get_AzureADUser.Name = "chk_Azure_Get_AzureADUser"
	$chk_Azure_Get_AzureADUser.Size = $Size_Chk
	$chk_Azure_Get_AzureADUser.TabIndex = $TabIndex++
	$chk_Azure_Get_AzureADUser.Text = "User"
	$chk_Azure_Get_AzureADUser.UseVisualStyleBackColor = $True
	$bx_AzureAd_Functions.Controls.Add($chk_Azure_Get_AzureADUser)
$chk_Azure_Get_AzureADUserLicenseDetail.Font = $font_Calibri_10pt_normal
	$Row_2_loc += 25
	$chk_Azure_Get_AzureADUserLicenseDetail.Location = New-Object System.Drawing.Point($Col_2_loc,$Row_2_loc)
	$chk_Azure_Get_AzureADUserLicenseDetail.Name = "chk_Azure_Get_AzureADUserLicenseDetail"
	$chk_Azure_Get_AzureADUserLicenseDetail.Size = $Size_Chk
	$chk_Azure_Get_AzureADUserLicenseDetail.TabIndex = $TabIndex++
	$chk_Azure_Get_AzureADUserLicenseDetail.Text = "UserLicenseDetail"
	$chk_Azure_Get_AzureADUserLicenseDetail.UseVisualStyleBackColor = $True
	$bx_AzureAd_Functions.Controls.Add($chk_Azure_Get_AzureADUserLicenseDetail)
$chk_Azure_Get_AzureADUserMembership.Font = $font_Calibri_10pt_normal
	$Row_2_loc += 25
	$chk_Azure_Get_AzureADUserMembership.Location = New-Object System.Drawing.Point($Col_2_loc,$Row_2_loc)
	$chk_Azure_Get_AzureADUserMembership.Name = "chk_Azure_Get_AzureADUserMembership"
	$chk_Azure_Get_AzureADUserMembership.Size = $Size_Chk
	$chk_Azure_Get_AzureADUserMembership.TabIndex = $TabIndex++
	$chk_Azure_Get_AzureADUserMembership.Text = "UserMembership"
	$chk_Azure_Get_AzureADUserMembership.UseVisualStyleBackColor = $True
	$bx_AzureAd_Functions.Controls.Add($chk_Azure_Get_AzureADUserMembership)
$chk_Azure_Get_AzureADUserOwnedDevice.Font = $font_Calibri_10pt_normal
	$Row_2_loc += 25
	$chk_Azure_Get_AzureADUserOwnedDevice.Location = New-Object System.Drawing.Point($Col_2_loc,$Row_2_loc)
	$chk_Azure_Get_AzureADUserOwnedDevice.Name = "chk_Azure_Get_AzureADUserOwnedDevice"
	$chk_Azure_Get_AzureADUserOwnedDevice.Size = $Size_Chk
	$chk_Azure_Get_AzureADUserOwnedDevice.TabIndex = $TabIndex++
	$chk_Azure_Get_AzureADUserOwnedDevice.Text = "UserOwnedDevice"
	$chk_Azure_Get_AzureADUserOwnedDevice.UseVisualStyleBackColor = $True
	$bx_AzureAd_Functions.Controls.Add($chk_Azure_Get_AzureADUserOwnedDevice)
$chk_Azure_Get_AzureADUserRegisteredDevice.Font = $font_Calibri_10pt_normal
	$Row_2_loc += 25
	$chk_Azure_Get_AzureADUserRegisteredDevice.Location = New-Object System.Drawing.Point($Col_2_loc,$Row_2_loc)
	$chk_Azure_Get_AzureADUserRegisteredDevice.Name = "chk_Azure_Get_AzureADUserRegisteredDevice"
	$chk_Azure_Get_AzureADUserRegisteredDevice.Size = $Size_Chk
	$chk_Azure_Get_AzureADUserRegisteredDevice.TabIndex = $TabIndex++
	$chk_Azure_Get_AzureADUserRegisteredDevice.Text = "UserRegisteredDevice"
	$chk_Azure_Get_AzureADUserRegisteredDevice.UseVisualStyleBackColor = $True
	$bx_AzureAd_Functions.Controls.Add($chk_Azure_Get_AzureADUserRegisteredDevice)

	#EndRegion Step3 Azure - AzureAd tab

#Region Step3 Sharepoint - Spo tab
$TabIndex = 0
$tab_Step3_Spo.Location = $Loc_Tab_Tier3
	$tab_Step3_Spo.Name = "tab_Step3_Spo"
	$tab_Step3_Spo.Padding = $System_Windows_Forms_Padding_Reusable
	$tab_Step3_Spo.Size = $Size_Tab_Small
	$tab_Step3_Spo.TabIndex = $TabIndex++
	$tab_Step3_Spo.Text = "SPO"
	$tab_Step3_Spo.UseVisualStyleBackColor = $True
	$tab_Step3_Sharepoint_Tier2.Controls.Add($tab_Step3_Spo)
$bx_Spo_Functions.Dock = 5
	$bx_Spo_Functions.Font = $font_Calibri_10pt_bold
	$bx_Spo_Functions.Location = $Loc_Box_1
	$bx_Spo_Functions.Name = "bx_Spo_Functions"
	$bx_Spo_Functions.Size = $Size_Box_3
	$bx_Spo_Functions.TabIndex = $TabIndex++
	$bx_Spo_Functions.TabStop = $False
	$tab_Step3_Spo.Controls.Add($bx_Spo_Functions)
$btn_Step3_Spo_CheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_Spo_CheckAll.Location = $Loc_CheckAll_3
	$btn_Step3_Spo_CheckAll.Name = "btn_Step3_Spo_CheckAll"
	$btn_Step3_Spo_CheckAll.Size = $Size_Btn_3
	$btn_Step3_Spo_CheckAll.TabIndex = $TabIndex++
	$btn_Step3_Spo_CheckAll.Text = "Check all on this tab"
	$btn_Step3_Spo_CheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_Spo_CheckAll.add_Click($handler_btn_Step3_Spo_CheckAll_Click)
	$bx_Spo_Functions.Controls.Add($btn_Step3_Spo_CheckAll)
$btn_Step3_Spo_UncheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_Spo_UncheckAll.Location = $Loc_UncheckAll_3
	$btn_Step3_Spo_UncheckAll.Name = "btn_Step3_Spo_UncheckAll"
	$btn_Step3_Spo_UncheckAll.Size = $Size_Btn_3
	$btn_Step3_Spo_UncheckAll.TabIndex = $TabIndex++
	$btn_Step3_Spo_UncheckAll.Text = "Uncheck all on this tab"
	$btn_Step3_Spo_UncheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_Spo_UncheckAll.add_Click($handler_btn_Step3_Spo_UncheckAll_Click)
	$bx_Spo_Functions.Controls.Add($btn_Step3_Spo_UncheckAll)
$Col_1_loc = 35
$Col_2_loc = 290
$Row_1_loc = 0
$Row_2_loc = 0
$chk_Org_Get_SPODeletedSite.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_SPODeletedSite.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_SPODeletedSite.Name = "chk_Org_Get_SPODeletedSite"
	$chk_Org_Get_SPODeletedSite.Size = $Size_Chk
	$chk_Org_Get_SPODeletedSite.TabIndex = $TabIndex++
	$chk_Org_Get_SPODeletedSite.Text = "GetDeletedSite"
	$chk_Org_Get_SPODeletedSite.UseVisualStyleBackColor = $True
	$bx_Spo_Functions.Controls.Add($chk_Org_Get_SPODeletedSite)
$chk_Org_Get_SPOExternalUser.Font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_SPOExternalUser.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_SPOExternalUser.Name = "chk_Org_Get_SPOExternalUser"
	$chk_Org_Get_SPOExternalUser.Size = $Size_Chk
	$chk_Org_Get_SPOExternalUser.TabIndex = $TabIndex++
	$chk_Org_Get_SPOExternalUser.Text = "Get-ExternalUser"
	$chk_Org_Get_SPOExternalUser.UseVisualStyleBackColor = $True
	$bx_Spo_Functions.Controls.Add($chk_Org_Get_SPOExternalUser)
$chk_Org_Get_SPOGeoStorageQuota.font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_SPOGeoStorageQuota.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_SPOGeoStorageQuota.Name = "chk_Org_Get_SPOGeoStorageQuota"
	$chk_Org_Get_SPOGeoStorageQuota.Size = $Size_Chk
	$chk_Org_Get_SPOGeoStorageQuota.TabIndex = $TabIndex++
	$chk_Org_Get_SPOGeoStorageQuota.Text = "Get-SPOGeoStorageQuota"
	$chk_Org_Get_SPOGeoStorageQuota.UseVisualStyleBackColor = $True
	$bx_Spo_Functions.Controls.Add($chk_Org_Get_SPOGeoStorageQuota)
$chk_Org_Get_SPOSite.font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_SPOSite.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_SPOSite.Name = "chk_Org_Get_SPOSite"
	$chk_Org_Get_SPOSite.Size = $Size_Chk
	$chk_Org_Get_SPOSite.TabIndex = $TabIndex++
	$chk_Org_Get_SPOSite.Text = "Get-SPOSite"
	$chk_Org_Get_SPOSite.UseVisualStyleBackColor = $True
	$bx_Spo_Functions.Controls.Add($chk_Org_Get_SPOSite)
$chk_Org_Get_SPOTenant.font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_SPOTenant.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_SPOTenant.Name = "chk_Org_Get_SPOTenant"
	$chk_Org_Get_SPOTenant.Size = $Size_Chk
	$chk_Org_Get_SPOTenant.TabIndex = $TabIndex++
	$chk_Org_Get_SPOTenant.Text = "Get-SPOTenant"
	$chk_Org_Get_SPOTenant.UseVisualStyleBackColor = $True
	$bx_Spo_Functions.Controls.Add($chk_Org_Get_SPOTenant)
$chk_Org_Get_SPOTenantSYncClientRestriction.font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_SPOTenantSYncClientRestriction.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_SPOTenantSYncClientRestriction.Name = "chk_Org_Get_SPOTenantSYncClientRestriction"
	$chk_Org_Get_SPOTenantSYncClientRestriction.Size = $Size_Chk
	$chk_Org_Get_SPOTenantSYncClientRestriction.TabIndex = $TabIndex++
	$chk_Org_Get_SPOTenantSYncClientRestriction.Text = "Get-SPOTenantSYncClientRestriction"
	$chk_Org_Get_SPOTenantSYncClientRestriction.UseVisualStyleBackColor = $True
	$bx_Spo_Functions.Controls.Add($chk_Org_Get_SPOTenantSYncClientRestriction)
$chk_Org_Get_SPOUser.font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_SPOUser.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_SPOUser.Name = "chk_Org_Get_SPOUser"
	$chk_Org_Get_SPOUser.Size = $Size_Chk
	$chk_Org_Get_SPOUser.TabIndex = $TabIndex++
	$chk_Org_Get_SPOUser.Text = "Get-SPOUser"
	$chk_Org_Get_SPOUser.UseVisualStyleBackColor = $True
	$bx_Spo_Functions.Controls.Add($chk_Org_Get_SPOUser)
$chk_Org_Get_SPOWebTemplate.font = $font_Calibri_10pt_normal
	$Row_1_loc += 25
	$chk_Org_Get_SPOWebTemplate.Location = New-Object System.Drawing.Point($Col_1_loc,$Row_1_loc)
	$chk_Org_Get_SPOWebTemplate.Name = "chk_Org_Get_SPOWebTemplate"
	$chk_Org_Get_SPOWebTemplate.Size = $Size_Chk
	$chk_Org_Get_SPOWebTemplate.TabIndex = $TabIndex++
	$chk_Org_Get_SPOWebTemplate.Text = "Get-SPOWebTemplate"
	$chk_Org_Get_SPOWebTemplate.UseVisualStyleBackColor = $True
	$bx_Spo_Functions.Controls.Add($chk_Org_Get_SPOWebTemplate)

	#EndRegion Step3 Sharepoint - SPO tab



#EndRegion "Step3 - Tests"

#Region "Step4 - Reporting"
$TabIndex = 0
$tab_Step4.BackColor = [System.Drawing.Color]::FromArgb(0,255,255,255)
	$tab_Step4.Font = $font_Calibri_8pt_normal
	$tab_Step4.Location = $Loc_Tab_Tier1
	$tab_Step4.Name = "tab_Step4"
	$tab_Step4.Padding = $System_Windows_Forms_Padding_Reusable
	$tab_Step4.TabIndex = $TabIndex++
	$tab_Step4.Text = "  Reporting  "
	$tab_Step4.Size = $Size_Tab_1
	$tab_Master.Controls.Add($tab_Step4)
$btn_Step4_Assemble.Font = $font_Calibri_14pt_normal
	$btn_Step4_Assemble.Location = $Loc_Btn_1
	$btn_Step4_Assemble.Name = "btn_Step4_Assemble"
	$btn_Step4_Assemble.Size = $Size_Buttons
	$btn_Step4_Assemble.TabIndex = $TabIndex++
	$btn_Step4_Assemble.Text = "Execute"
	$btn_Step4_Assemble.UseVisualStyleBackColor = $True
	$btn_Step4_Assemble.add_Click($handler_btn_Step4_Assemble_Click)
	$tab_Step4.Controls.Add($btn_Step4_Assemble)
$lbl_Step4_Assemble.Font = $font_Calibri_10pt_normal
	$lbl_Step4_Assemble.Location = $Loc_Lbl_1
	$lbl_Step4_Assemble.Name = "lbl_Step4"
	$lbl_Step4_Assemble.Size = New-Object System.Drawing.Size(510,38)
	$lbl_Step4_Assemble.TabIndex = $TabIndex++
	$lbl_Step4_Assemble.Text = "If Office 2010 or later is installed, the Execute button can be used to assemble `nthe output from Tests into reports."
	$lbl_Step4_Assemble.TextAlign = 16
	$tab_Step4.Controls.Add($lbl_Step4_Assemble)
$bx_Step4_Functions.Font = $font_Calibri_10pt_bold
	$bx_Step4_Functions.Location = New-Object System.Drawing.Point(27,91)
	$bx_Step4_Functions.Name = "bx_Step4_Functions"
	$bx_Step4_Functions.Size = $Size_Box_1  #New-Object System.Drawing.Size(536,487)
	$bx_Step4_Functions.TabIndex = $TabIndex++
	$bx_Step4_Functions.TabStop = $False
	$bx_Step4_Functions.Text = "Report Generation Functions"
	$tab_Step4.Controls.Add($bx_Step4_Functions)
<# $chk_Step4_DC_Report.Checked = $True
	$chk_Step4_DC_Report.CheckState = 1
	$chk_Step4_DC_Report.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 50
	$System_Drawing_Point.Y = 25
	$chk_Step4_DC_Report.Location = $System_Drawing_Point
	$chk_Step4_DC_Report.Name = "chk_Step4_DC_Report"
	$chk_Step4_DC_Report.Size = $Size_Chk_long
	$chk_Step4_DC_Report.TabIndex = $TabIndex++
	$chk_Step4_DC_Report.Text = "Generate Excel for Domain Controllers"
	$chk_Step4_DC_Report.UseVisualStyleBackColor = $True
	$bx_Step4_Functions.Controls.Add($chk_Step4_DC_Report)
$chk_Step4_Ex_Report.Checked = $True
	$chk_Step4_Ex_Report.CheckState = 1
	$chk_Step4_Ex_Report.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 50
	$System_Drawing_Point.Y = 50
	$chk_Step4_Ex_Report.Location = $System_Drawing_Point
	$chk_Step4_Ex_Report.Name = "chk_Step4_Ex_Report"
	$chk_Step4_Ex_Report.Size = $Size_Chk_long
	$chk_Step4_Ex_Report.TabIndex = $TabIndex++
	$chk_Step4_Ex_Report.Text = "Generate Excel for Exchange servers"
	$chk_Step4_Ex_Report.UseVisualStyleBackColor = $True
	$bx_Step4_Functions.Controls.Add($chk_Step4_Ex_Report)
 #>
 $chk_Step4_Exchange_Report.Checked = $True
	$chk_Step4_Exchange_Report.CheckState = 1
	$chk_Step4_Exchange_Report.Font = $font_Calibri_10pt_normal
	$chk_Step4_Exchange_Report.Location = New-Object System.Drawing.Point(50,25)
	$chk_Step4_Exchange_Report.Name = "chk_Step4_Exchange_Report"
	$chk_Step4_Exchange_Report.Size = $Size_Chk_long
	$chk_Step4_Exchange_Report.TabIndex = $TabIndex++
	$chk_Step4_Exchange_Report.Text = "Generate Excel for Exchange Organization"
	$chk_Step4_Exchange_Report.UseVisualStyleBackColor = $True
	$bx_Step4_Functions.Controls.Add($chk_Step4_Exchange_Report)
$chk_Step4_Azure_Report.Checked = $True
	$chk_Step4_Azure_Report.CheckState = 1
	$chk_Step4_Azure_Report.Font = $font_Calibri_10pt_normal
	$chk_Step4_Azure_Report.Location = New-Object System.Drawing.Point(50,50)
	$chk_Step4_Azure_Report.Name = "chk_Step4_Azure_Report"
	$chk_Step4_Azure_Report.Size = $Size_Chk_long
	$chk_Step4_Azure_Report.TabIndex = $TabIndex++
	$chk_Step4_Azure_Report.Text = "Generate Excel for Azure"
	$chk_Step4_Azure_Report.UseVisualStyleBackColor = $True
	$bx_Step4_Functions.Controls.Add($chk_Step4_Azure_Report)
<# $chk_Step4_Exchange_Environment_Doc.Checked = $True
	$chk_Step4_Exchange_Environment_Doc.CheckState = 1
	$chk_Step4_Exchange_Environment_Doc.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 50
	$System_Drawing_Point.Y = 100
	$chk_Step4_Exchange_Environment_Doc.Location = $System_Drawing_Point
	$chk_Step4_Exchange_Environment_Doc.Name = "chk_Step4_Exchange_Environment_Doc"
	$chk_Step4_Exchange_Environment_Doc.Size = $Size_Chk_long
	$chk_Step4_Exchange_Environment_Doc.TabIndex = $TabIndex++
	$chk_Step4_Exchange_Environment_Doc.Text = "Generate Word for Exchange Documention"
	$chk_Step4_Exchange_Environment_Doc.UseVisualStyleBackColor = $True
	$bx_Step4_Functions.Controls.Add($chk_Step4_Exchange_Environment_Doc)
 #>$Status_Step4.Font = $font_Calibri_10pt_normal
	$Status_Step4.Location = $Loc_Status
	$Status_Step4.Name = "Status_Step4"
	$Status_Step4.Size = $Size_Status
	$Status_Step4.TabIndex = $TabIndex++
	$Status_Step4.Text = "Step 4 Status"
	$tab_Step4.Controls.Add($Status_Step4)
#EndRegion "Step4 - Reporting"

<#
#Region "Step5 - Having Trouble?"
$tab_Step5.BackColor = [System.Drawing.Color]::FromArgb(0,255,255,255)
$tab_Step5.Font = $font_Calibri_8pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 4
	$System_Drawing_Point.Y = 36
$tab_Step5.Location = $System_Drawing_Point
$tab_Step5.Name = "tab_Step5"
$tab_Step5.Padding = $System_Windows_Forms_Padding_Reusable
$tab_Step5.TabIndex = $TabIndex++
$tab_Step5.Text = "  Having Trouble?  "
$tab_Step5.Size = $System_Drawing_Size_tab_2
$tab_Step5.visible = $False
$tab_Master.Controls.Add($tab_Step5)

$bx_Step5_Functions.Font = $font_Calibri_10pt_bold
	$System_Drawing_Point_bx_Step5 = New-Object System.Drawing.Point
	$System_Drawing_Point_bx_Step5.X = 27	# 96-69
	$System_Drawing_Point_bx_Step5.Y = 91
$bx_Step5_Functions.Location = $System_Drawing_Point_bx_Step5
$bx_Step5_Functions.Name = "bx_Step5_Functions"
	$System_Drawing_Size_bx_Step5 = New-Object System.Drawing.Size
	$System_Drawing_Size_bx_Step5.Height = 487 #482 to short
	$System_Drawing_Size_bx_Step5.Width = 536
$bx_Step5_Functions.Size = $System_Drawing_Size_bx_Step5
$bx_Step5_Functions.TabIndex = $TabIndex++
$bx_Step5_Functions.TabStop = $False
$bx_Step5_Functions.Text = "If you're having trouble collecting data..."
$tab_Step5.Controls.Add($bx_Step5_Functions)

$Status_Step5.Font = $font_Calibri_10pt_normal
$Status_Step5.Location = $Loc_Status
$Status_Step5.Name = "Status_Step5"
$Status_Step5.Size = $Size_Status
$Status_Step5.TabIndex = $TabIndex++
$Status_Step5.Text = "Step 5 Status"
$tab_Step5.Controls.Add($Status_Step5)

#EndRegion "Step5 - Having Trouble?"
#>

#Region Set Tests Checkbox States
if (($INI_Exchange -ne ""))
{
	# Code to parse INI
	write-host "Importing INI settings"
	write-host "Exchange INI settings: " $ini_Exchange
	# Exchange INI
	write-host $ini_Exchange
	if (($ini_Exchange -ne "") -and ((Test-Path $ini_Exchange) -eq $true))
	{
		write-host "File specified using the -INI_Exchange switch" -ForegroundColor Green
		& ".\O365DC_Scripts\Core_Parse_Ini_File.ps1" -IniFile $INI_Exchange
	}
	elseif (($ini_Exchange -ne "") -and ((Test-Path $ini_Exchange) -eq $false))
	{
		write-host "File specified using the -INI_Exchange switch was not found" -ForegroundColor Red
	}
}
else
{
	# Exchange Functions
		Set-AllFunctionsClientAccess -Check $true
		Set-AllFunctionsGlobal -Check $true
		Set-AllFunctionsRecipient -Check $true
		Set-AllFunctionsTransport -Check $true
		Set-AllFunctionsMisc -Check $true
		Set-AllFunctionsUm -Check $true
}

#EndRegion Set Checkbox States

#endregion *** Build Form ***

#Save the initial state of the form
$InitialFormWindowState = $form1.WindowState
#Init the OnLoad event to correct the initial state of the form
$form1.add_Load($OnLoadForm_StateCorrection)
#Show the Form
$form1.ShowDialog()| Out-Null

}
#End Function

##############################################
# New-O365DCForm should not be above this line #
##############################################

#region *** Custom Functions ***

Trap {
$ErrorText = "O365DC " + "`n" + $server + "`n"
$ErrorText += $_

$ErrorLog = New-Object System.Diagnostics.EventLog('Application')
$ErrorLog.MachineName = "."
$ErrorLog.Source = "O365DC"
Try{$ErrorLog.WriteEntry($ErrorText,"Error", 100)} catch{}
}

Function Watch-O365DCKnownErrors()
{
	Trap [System.InvalidOperationException]{
		#write-host -Fore Red -back White $_.Exception.Message
		#write-host -Fore Red -back White $_.Exception.FullyQualifiedErrorId
		Continue
	}
	Trap [system.Management.Automation.ErrorRecord] {
		#write-host -Fore Red -back White $_.Exception.Message
		#write-host -Fore Red -back White $_.Exception.FullyQualifiedErrorId
		Continue
	}
	Trap [System.Management.AutomationRuntimeException] {
		write-host -Fore Red -back White $_.Exception.Message
		write-host -Fore Red -back White $_.Exception.FullyQualifiedErrorId
		Silently Continue
	}
	Trap [System.Management.Automation.MethodInvocationException] {
		write-host -Fore Red -back White $_.Exception.Message
		write-host -Fore Red -back White $_.Exception.FullyQualifiedErrorId
		Continue
	}
}

Function Disable-AllTargetsButtons()
{
	$btn_Step1_Mailboxes_Discover.enabled = $false
	$btn_Step1_Mailboxes_Populate.enabled = $false
}

Function Enable-AllTargetsButtons()
{
	$btn_Step1_Mailboxes_Discover.enabled = $true
	$btn_Step1_Mailboxes_Populate.enabled = $true
}

Function Limit-O365DCJob
{
	Param([int]$JobThrottleMaxJobs,`
		[int]$JobThrottlePolling)

	# Remove Failed Jobs
	$colJobsFailed = @(Get-Job -State Failed)
	foreach ($objJobsFailed in $colJobsFailed)
	{
		if ($objJobsFailed.module -like "__DynamicModule*")
		{
			Remove-Job -Id $objJobsFailed.id
		}
		else
		{
            write-host "---- Failed job " $objJobsFailed.name -ForegroundColor Red
			$FailedJobOutput = ".\FailedJobs_" + $append + ".txt"
            if ((Test-Path $FailedJobOutput) -eq $false)
	        {
		      new-item $FailedJobOutput -type file -Force
	        }
	        "* * * * * * * * * * * * * * * * * * * *"  | Out-File $FailedJobOutput -Force -Append
            "Job Name: " + $objJobsFailed.name | Out-File $FailedJobOutput -Force -Append
	        "Job State: " + $objJobsFailed.state | Out-File $FailedJobOutput -Force	-Append
            if ($null -ne ($objJobsFailed.childjobs[0]))
            {
	           $objJobsFailed.childjobs[0].output | format-list | Out-File $FailedJobOutput -Force -Append
	           $objJobsFailed.childjobs[0].warning | format-list | Out-File $FailedJobOutput -Force -Append
	           $objJobsFailed.childjobs[0].error | format-list | Out-File $FailedJobOutput -Force -Append
			}
            $ErrorText = $objJobsFailed.name + "`n"
			$ErrorText += "Job failed"
			$ErrorLog = New-Object System.Diagnostics.EventLog('Application')
			$ErrorLog.MachineName = "."
			$ErrorLog.Source = "O365DC"
			Try{$ErrorLog.WriteEntry($ErrorText,"Error", 500)} catch{}
			Remove-Job -Id $objJobsFailed.id
		}
	}

    $colJobsRunning = @((Get-Job -State Running) | where-object {$_.Module -ne "__DynamicModule*"})
	if ((Test-Path ".\RunningJobs.txt") -eq $false)
	{
		new-item ".\RunningJobs.txt" -type file -Force
	}
	$RunningJobsOutput = ""
	$Now = Get-Date
	foreach ($objJobsRunning in $colJobsRunning)
	{
		$JobPID = $objJobsRunning.childjobs[0].output[0]
		if ($null -ne $JobPID)
		{
			# Pass the variable assignment as a condition to reduce timing issues
			if(($JobStartTime = ((Get-Process | where-object {$_.id -eq $JobPID}).starttime)) -ne $null)
			{
				$JobRunningTime = [int](($Now - $JobStartTime).TotalMinutes)
				if ((($objJobsRunning.childjobs[0].output[1] -eq "WMI") -and ($JobRunningTime -gt ($intWMIJobTimeout/60))) `
					-or (($objJobsRunning.childjobs[0].output[1] -eq "Exchange") -and ($JobRunningTime -gt ($intExchJobTimeout/60))))
				{
					try
					{
						(Get-Process | where-object {$_.id -eq $JobPID}).kill()
						write-host "Timer expired.  Killing job process $JobPID - " + $objJobsRunning.name -ForegroundColor Red
						$ErrorText = $objJobsRunning.name + "`n"
						$ErrorText += "Process $JobPID killed`n"
						if ($objJobsRunning.childjobs[0].output[1] -eq "WMI") {$ErrorText += "Timeout $intWMIJobTimeout seconds exceeded"}
						if ($objJobsRunning.childjobs[0].output[1] -eq "Exchange") {$ErrorText += "Timeout $intExchJobTimeout seconds exceeded"}
						$ErrorLog = New-Object System.Diagnostics.EventLog('Application')
						$ErrorLog.MachineName = "."
						$ErrorLog.Source = "O365DC"
						Try{$ErrorLog.WriteEntry($ErrorText,"Error", 600)} catch{}
					}
					catch [System.Management.Automation.MethodInvocationException]
					{
						write-host "`tMethodInvocationException occured during Kill request for process $JobPID" -ForegroundColor Red
					}
					catch [System.Management.Automation.RuntimeException]
					{
						write-host "`tRuntimeException occured during Kill request for process $JobPID" -ForegroundColor Red
					}
				}
				$RunningJobsOutput += "Job Name: " + $objJobsRunning.name + "`n"
				$RunningJobsOutput += "Job State: " + $objJobsRunning.State + "`n"
				$RunningJobsOutput += "Job process PID: " + $JobPID + "`n"
				$RunningJobsOutput += "Job process time running: " +  $JobRunningTime + " min"
				$RunningJobsOutput += "`n`n"
			}
		}
	}
	$RunningJobsOutput | Out-File ".\RunningJobs.txt" -Force

	$intRunningJobs = $colJobsRunning.count

	if ($intRunningJobs -eq $null)
	{
		$intRunningJobs = "0"
	}

	$colJobsCompleted = @((Get-Job -State completed) | where-object {$null -ne $_.childjobs})
	foreach ($objJobsCompleted in $colJobsCompleted)
	{
		Remove-Job -Id $objJobsCompleted.id
		write-host "---- Finished job " $objJobsCompleted.name -ForegroundColor Green
	}

	do
	{
        ## Repeat bulk of function code to prevent recursive loop
        ##      and the dreaded System.Management.Automation.ScriptCallDepthException:
        ##      The script failed due to call depth overflow.
        ##      The call depth reached 1001 and the maximum is 1000.

            # Remove Failed Jobs
	        $colJobsFailed = @(Get-Job -State Failed)
	        foreach ($objJobsFailed in $colJobsFailed)
            {
		        if ($objJobsFailed.module -like "__DynamicModule*")
		        {
			        Remove-Job -Id $objJobsFailed.id
		        }
		        else
		        {
                    write-host "---- Failed job " $objJobsFailed.name -ForegroundColor Red
			        $FailedJobOutput = ".\FailedJobs_" + $append + ".txt"
                    if ((Test-Path $FailedJobOutput) -eq $false)
	                {
		              new-item $FailedJobOutput -type file -Force
	                }
	                "* * * * * * * * * * * * * * * * * * * *"  | Out-File $FailedJobOutput -Force -Append
                    "Job Name: " + $objJobsFailed.name | Out-File $FailedJobOutput -Force -Append
	                "Job State: " + $objJobsFailed.state | Out-File $FailedJobOutput -Force	-Append
                    if ($null -ne ($objJobsFailed.childjobs[0]))
                    {
	                   $objJobsFailed.childjobs[0].output | format-list | Out-File $FailedJobOutput -Force -Append
	                   $objJobsFailed.childjobs[0].warning | format-list | Out-File $FailedJobOutput -Force -Append
	                   $objJobsFailed.childjobs[0].error | format-list | Out-File $FailedJobOutput -Force -Append
			        }
                    $ErrorText = $objJobsFailed.name + "`n"
			        $ErrorText += "Job failed"
			        $ErrorLog = New-Object System.Diagnostics.EventLog('Application')
			        $ErrorLog.MachineName = "."
			        $ErrorLog.Source = "O365DC"
			        Try{$ErrorLog.WriteEntry($ErrorText,"Error", 500)} catch{}
			        Remove-Job -Id $objJobsFailed.id
		        }
	        }
            $colJobsRunning = @((Get-Job -State Running) | where-object {$_.Module -ne "__DynamicModule*"})
	        if ((Test-Path ".\RunningJobs.txt") -eq $false)
	        {
		        new-item ".\RunningJobs.txt" -type file -Force
	        }
	        $RunningJobsOutput = ""
	        $Now = Get-Date
	        foreach ($objJobsRunning in $colJobsRunning)
	        {
		        $JobPID = $objJobsRunning.childjobs[0].output[0]
		        if ($null -ne $JobPID)
		        {
			        # Pass the variable assignment as a condition to reduce timing issues
			        if (($JobStartTime = ((Get-Process | where-object {$_.id -eq $JobPID}).starttime)) -ne $null)
			        {
				        $JobRunningTime = [int](($Now - $JobStartTime).TotalMinutes)
				        if ((($objJobsRunning.childjobs[0].output[1] -eq "WMI") -and ($JobRunningTime -gt ($intWMIJobTimeout/60))) `
					        -or (($objJobsRunning.childjobs[0].output[1] -eq "Exchange") -and ($JobRunningTime -gt ($intExchJobTimeout/60))))
				        {
					        try
					        {
						        (Get-Process | where-object {$_.id -eq $JobPID}).kill()
						        write-host "Timer expired.  Killing job process $JobPID - " + $objJobsRunning.name -ForegroundColor Red
						        $ErrorText = $objJobsRunning.name + "`n"
						        $ErrorText += "Process $JobPID killed`n"
						        if ($objJobsRunning.childjobs[0].output[1] -eq "WMI") {$ErrorText += "Timeout $intWMIJobTimeout seconds exceeded"}
						        if ($objJobsRunning.childjobs[0].output[1] -eq "Exchange") {$ErrorText += "Timeout $intExchJobTimeout seconds exceeded"}
						        $ErrorLog = New-Object System.Diagnostics.EventLog('Application')
						        $ErrorLog.MachineName = "."
						        $ErrorLog.Source = "O365DC"
						        Try{$ErrorLog.WriteEntry($ErrorText,"Error", 600)} catch{}
					        }
					        catch [System.Management.Automation.MethodInvocationException]
					        {
						        write-host "`tMethodInvocationException occured during Kill request for process $JobPID" -ForegroundColor Red
					        }
					        catch [System.Management.Automation.RuntimeException]
					        {
						        write-host "`tRuntimeException occured during Kill request for process $JobPID" -ForegroundColor Red
					        }
				        }
				        $RunningJobsOutput += "Job Name: " + $objJobsRunning.name + "`n"
				        $RunningJobsOutput += "Job State: " + $objJobsRunning.State + "`n"
				        $RunningJobsOutput += "Job process PID: " + $JobPID + "`n"
				        $RunningJobsOutput += "Job process time running: " +  $JobRunningTime + " min"
				        $RunningJobsOutput += "`n`n"
			        }
		        }
	        }
	        $RunningJobsOutput | Out-File ".\RunningJobs.txt" -Force
	        $intRunningJobs = $colJobsRunning.count
	        if ($intRunningJobs -eq $null)
	        {
		        $intRunningJobs = "0"
	        }
	        $colJobsCompleted = @((Get-Job -State completed) | where-object {$null -ne $_.childjobs})
	        foreach ($objJobsCompleted in $colJobsCompleted)
	        {
		        Remove-Job -Id $objJobsCompleted.id
		        write-host "---- Finished job " $objJobsCompleted.name -ForegroundColor Green
	        }
        if ($intRunningJobs -ge $JobThrottleMaxJobs)
        {
            write-host "** Throttling at $intRunningJobs jobs." -ForegroundColor DarkYellow
            Start-Sleep -Seconds $JobThrottlePolling
        }
	} while ($intRunningJobs -ge $JobThrottleMaxJobs)


	write-host "** $intRunningJobs jobs running." -ForegroundColor DarkYellow

}

Function Update-O365DCJobCount
{
	Param([int]$JobCountMaxJobs,`
		[int]$JobCountPolling)

	# Remove Failed Jobs
	$colJobsFailed = @(Get-Job -State Failed)
	foreach ($objJobsFailed in $colJobsFailed)
	{
		if ($objJobsFailed.module -like "__DynamicModule*")
		{
			Remove-Job -Id $objJobsFailed.id
		}
		else
		{
            write-host "---- Failed job " $objJobsFailed.name -ForegroundColor Red
			$FailedJobOutput = ".\FailedJobs_" + $append + ".txt"
			if ((Test-Path $FailedJobOutput) -eq $false)
	        {
		      new-item $FailedJobOutput -type file -Force
	        }
	        "* * * * * * * * * * * * * * * * * * * *"  | Out-File $FailedJobOutput -Force -Append
	        "Job Name: " + $objJobsFailed.name | Out-File $FailedJobOutput -Force -Append
	        "Job State: " + $objJobsFailed.state | Out-File $FailedJobOutput -Force	-Append
            if ($null -ne ($objJobsFailed.childjobs[0]))
            {
	           $objJobsFailed.childjobs[0].output | format-list | Out-File $FailedJobOutput -Force -Append
	           $objJobsFailed.childjobs[0].warning | format-list | Out-File $FailedJobOutput -Force -Append
	           $objJobsFailed.childjobs[0].error | format-list | Out-File $FailedJobOutput -Force -Append
			}
            $ErrorText = $objJobsFailed.name + "`n"
			$ErrorText += "Job failed"
			$ErrorLog = New-Object System.Diagnostics.EventLog('Application')
			$ErrorLog.MachineName = "."
			$ErrorLog.Source = "O365DC"
			Try{$ErrorLog.WriteEntry($ErrorText,"Error", 500)} catch{}
			Remove-Job -Id $objJobsFailed.id
		}
	}

	$colJobsRunning = @((Get-Job -State Running) | where-object {$_.Module -ne "__DynamicModule*"})
	if ((Test-Path ".\RunningJobs.txt") -eq $false)
	{
		new-item ".\RunningJobs.txt" -type file -Force
	}

	$RunningJobsOutput = ""
	$Now = Get-Date
	foreach ($objJobsRunning in $colJobsRunning)
	{
		$JobPID = $objJobsRunning.childjobs[0].output[0]
		if ($null -ne $JobPID)
		{
			# Pass the variable assignment as a condition to reduce timing issues
			if (($JobStartTime = ((Get-Process | where-object {$_.id -eq $JobPID}).starttime)) -ne $null)
			{
				$JobRunningTime = [int](($Now - $JobStartTime).TotalMinutes)
				if ((($objJobsRunning.childjobs[0].output[1] -eq "WMI") -and ($JobRunningTime -gt ($intWMIJobTimeout/60))) `
					-or (($objJobsRunning.childjobs[0].output[1] -eq "Exchange") -and ($JobRunningTime -gt ($intExchJobTimeout/60))))
				{
					try
                    {
                        (Get-Process | where-object {$_.id -eq $JobPID}).kill()
                    }
                    catch {}
					write-host "Timer expired.  Killing job process $JobPID - " $objJobsRunning.name -ForegroundColor Red
					$ErrorText = $objJobsRunning.name + "`n"
					$ErrorText += "Process $JobPID killed`n"
					if ($objJobsRunning.childjobs[0].output[1] -eq "WMI") {$ErrorText += "Timeout $intWMIJobTimeout seconds exceeded"}
					if ($objJobsRunning.childjobs[0].output[1] -eq "Exchange") {$ErrorText += "Timeout $intExchJobTimeout seconds exceeded"}
					$ErrorLog = New-Object System.Diagnostics.EventLog('Application')
					$ErrorLog.MachineName = "."
					$ErrorLog.Source = "O365DC"
					Try{$ErrorLog.WriteEntry($ErrorText,"Error", 600)} catch{}
				}
				$RunningJobsOutput += "Job Name: " + $objJobsRunning.name + "`n"
				$RunningJobsOutput += "Job State: " + $objJobsRunning.State + "`n"
				$RunningJobsOutput += "Job process PID: " + $JobPID + "`n"
				$RunningJobsOutput += "Job process time running: " +  $JobRunningTime + " min"
				$RunningJobsOutput += "`n`n"
			}
		}
	}
	$RunningJobsOutput | Out-File ".\RunningJobs.txt" -Force

	$intJobCount = $colJobsRunning.count
	if ($intJobCount -eq $null)
	{
		$intJobCount = "0"
	}

	$colJobsCompleted = @((Get-Job -State completed) | where-object {$null -ne $_.childjobs})
	foreach ($objJobsCompleted in $colJobsCompleted)
	{
		Remove-Job -Id $objJobsCompleted.id
		write-host "---- Finished job " $objJobsCompleted.name -ForegroundColor Green
	}

	if ($intJobCount -ge $JobCountMaxJobs)
	{
		write-host "** $intJobCount jobs still running.  Time: $((Get-Date).timeofday.tostring())" -ForegroundColor DarkYellow
		Start-Sleep -Seconds $JobCountPolling
		Update-O365DCJobCount $JobCountMaxJobs $JobCountPolling
	}
}

Function Get-ExchangeBoxStatus # See if any are checked
{
if (($chk_Org_Get_AcceptedDomain.checked -eq $true) -or
	($chk_Org_Get_ActiveSyncOrgSettings.checked -eq $true) -or
	($chk_Org_Get_AntiPhishPolicy.checked -eq $true) -or
	($chk_Org_Get_AntiSpoofingPolicy.checked -eq $true) -or
	($chk_Org_Get_AtpPolicyForO365.checked -eq $true) -or
	($chk_Org_Get_OnPremisesOrganization.checked -eq $true) -or
	($chk_Org_Get_SmimeConfig.checked -eq $true) -or
	($chk_Org_Get_CASMailboxPlan.checked -eq $true) -or
	($chk_Org_Get_Contact.checked -eq $true) -or
	($chk_Org_Get_MailboxPlan.checked -eq $true) -or
	($chk_Org_Get_MailUser.checked -eq $true) -or
	($chk_Org_Get_UnifiedGroup.checked -eq $true) -or
	($chk_Org_Get_DkimSigningConfig.checked -eq $true) -or
	($chk_Org_Get_MobileDevice.checked -eq $true) -or
	($chk_Org_Get_MobileDevicePolicy.checked -eq $true) -or
	($chk_Org_Get_AddressBookPolicy.checked -eq $true) -or
	($chk_Org_Get_AddressList.checked -eq $true) -or
	($chk_Org_Get_AvailabilityAddressSpace.checked -eq $true) -or
	($chk_Org_Get_CalendarProcessing.checked -eq $true) -or
	($chk_Org_Get_CASMailbox.checked -eq $true) -or
	($chk_Org_Get_DistributionGroup.checked -eq $true) -or
	($chk_Org_Get_DynamicDistributionGroup.checked -eq $true) -or
	($chk_Org_Get_EmailAddressPolicy.checked -eq $true) -or
	($chk_Org_Get_GlobalAddressList.checked -eq $true) -or
	($chk_Org_Get_Mailbox.checked -eq $true) -or
	($chk_Org_Get_MailboxFolderStatistics.checked -eq $true) -or
	($chk_Org_Get_MailboxPermission.checked -eq $true) -or
	($chk_Org_Get_MailboxStatistics.checked -eq $true) -or
	($chk_Org_Get_OfflineAddressBook.checked -eq $true) -or
	($chk_Org_Get_OrgConfig.checked -eq $true) -or
	($chk_Org_Get_OwaMailboxPolicy.checked -eq $true) -or
	($chk_Org_Get_PublicFolder.checked -eq $true) -or
	($chk_Org_Get_PublicFolderStatistics.checked -eq $true) -or
	($chk_Org_Get_InboundConnector.checked -eq $true) -or
	($chk_Org_Get_RemoteDomain.checked -eq $true) -or
	($chk_Org_Get_Rbac.checked -eq $true) -or
	($chk_Org_Get_RetentionPolicy.checked -eq $true) -or
	($chk_Org_Get_RetentionPolicyTag.checked -eq $true) -or
	($chk_Org_Get_OutboundConnector.checked -eq $true) -or
	($chk_Org_Get_TransportConfig.checked -eq $true) -or
	($chk_Org_Get_TransportRule.checked -eq $true) -or
	($chk_Org_Get_UmAutoAttendant.checked -eq $true) -or
	($chk_Org_Get_UmDialPlan.checked -eq $true) -or
	($chk_Org_Get_UmIpGateway.checked -eq $true) -or
	($chk_Org_Get_UmMailbox.checked -eq $true) -or
	#($chk_Org_Get_UmMailboxConfiguration.checked -eq $true) -or
	#($chk_Org_Get_UmMailboxPin.checked -eq $true) -or
	($chk_Org_Get_UmMailboxPolicy.checked -eq $true) -or
	($chk_Org_Get_UmServer.checked -eq $true) -or
	($chk_Org_Quota.checked -eq $true) -or
	($chk_Org_Get_AdminGroups.checked -eq $true))	{
		$true
	}
}

Function Get-ExchangeMbxBoxStatus # See if any are checked
{
if (($chk_Org_Get_CalendarProcessing.checked -eq $true) -or
	($chk_Org_Get_CASMailbox.checked -eq $true) -or
	($chk_Org_Get_Mailbox.checked -eq $true) -or
	($chk_Org_Get_MailboxFolderStatistics.checked -eq $true) -or
	($chk_Org_Get_MailboxPermission.checked -eq $true) -or
	($chk_Org_Get_MailboxStatistics.checked -eq $true) -or
	($chk_Org_Get_UmMailbox.checked -eq $true) -or
	#($chk_Org_Get_UmMailboxConfiguration.checked -eq $true) -or
	#($chk_Org_Get_UmMailboxPin.checked -eq $true) -or
	($chk_Org_Quota.checked -eq $true))
	{
		$true
	}
}

Function Get-AzureBoxStatus # See if any are checked
{
if (
	($chk_Azure_Get_AzureADApplication.checked -eq $true) -or
	($chk_Azure_Get_AzureADContact.checked -eq $true) -or
	($chk_Azure_Get_AzureADDevice.checked -eq $true) -or
	($chk_Azure_Get_AzureADDeviceRegisteredOwner.checked -eq $true) -or
	($chk_Azure_Get_AzureADDeviceRegisteredUser.checked -eq $true) -or
	($chk_Azure_Get_AzureADDirectoryRole.checked -eq $true) -or
	($chk_Azure_Get_AzureADDomain.checked -eq $true) -or
	($chk_Azure_Get_AzureAdDomainServiceConfigurationRecord.checked -eq $true) -or
	($chk_Azure_Get_AzureAdDomainVerificationDnsRecord.checked -eq $true) -or
	($chk_Azure_Get_AzureADGroup.checked -eq $true) -or
	($chk_Azure_Get_AzureADGroupMember.checked -eq $true) -or
	($chk_Azure_Get_AzureADGroupOwner.checked -eq $true) -or
	($chk_Azure_Get_AzureADSubscribedSku.checked -eq $true) -or
	($chk_Azure_Get_AzureADTenantDetail.checked -eq $true) -or
	($chk_Azure_Get_AzureADUser.checked -eq $true) -or
	($chk_Azure_Get_AzureADUserLicenseDetail.checked -eq $true) -or
	($chk_Azure_Get_AzureADUserMembership.checked -eq $true) -or
	($chk_Azure_Get_AzureADUserOwnedDevice.checked -eq $true) -or
	($chk_Azure_Get_AzureADUserRegisteredDevice.checked -eq $true)
	)
		{$true}
}

Function Get-AzureAdUserBoxStatus # See if any are checked
{
if (
	($chk_Azure_Get_AzureADUser.checked -eq $true) -or
	($chk_Azure_Get_AzureADUserLicenseDetail.checked -eq $true) -or
	($chk_Azure_Get_AzureADUserMembership.checked -eq $true) -or
	($chk_Azure_Get_AzureADUserOwnedDevice.checked -eq $true) -or
	($chk_Azure_Get_AzureADUserRegisteredDevice.checked -eq $true)
	)
		{$true}
}

Function Get-AzureAdDeviceBoxStatus # See if any are checked
{
if (
	($chk_Azure_Get_AzureADDevice.checked -eq $true) -or
	($chk_Azure_Get_AzureADDeviceRegisteredOwner.checked -eq $true) -or
	($chk_Azure_Get_AzureADDeviceRegisteredUser.checked -eq $true)
	)
		{$true}
}



Function Import-TargetsMailboxes
{
	Disable-AllTargetsButtons
    $status_Step1.Text = "Step 1 Status: Running"
	$File_Location = $location + "\mailbox.txt"
    if ((Test-Path $File_Location) -eq $true)
	{
		$EventLog = New-Object System.Diagnostics.EventLog('Application')
		$EventLog.MachineName = "."
		$EventLog.Source = "O365DC"
		try{$EventLog.WriteEntry("Starting O365DC Step 1 - Populate","Information", 10)} catch{}
	    $array_Mailboxes = @(([System.IO.File]::ReadAllLines($File_Location)) | sort-object -Unique)
		$global:intMailboxTotal = 0
	    $clb_Step1_Mailboxes_List.items.clear()
		foreach ($member_Mailbox in $array_Mailboxes | where-object {$_ -ne ""})
	    {
	        $clb_Step1_Mailboxes_List.items.add($member_Mailbox)
			$global:intMailboxTotal++
	    }
		For ($i=0;$i -le ($intMailboxTotal - 1);$i++)
		{
			$clb_Step1_Mailboxes_List.SetItemChecked($i,$true)
		}
		$EventLog = New-Object System.Diagnostics.EventLog('Application')
		$EventLog.MachineName = "."
		$EventLog.Source = "O365DC"
		try{$EventLog.WriteEntry("Ending O365DC Step 1 - Populate","Information", 11)} catch{}
		$txt_MailboxesTotal.Text = "Mailbox count = " + $intMailboxTotal
		$txt_MailboxesTotal.visible = $true
	    $status_Step1.Text = "Step 2 Status: Idle"
	}
	else
	{
		write-host	"The file mailbox.txt is not present.  Run Discover to create the file."
		$status_Step1.Text = "Step 1 Status: Failed - mailbox.txt file not found.  Run Discover to create the file."
	}
	Enable-AllTargetsButtons
}

Function Enable-TargetsMailbox
{
	For ($i=0;$i -le ($intMailboxTotal - 1);$i++)
	{
		$clb_Step1_Mailboxes_List.SetItemChecked($i,$true)
	}
}

Function Disable-TargetsMailbox
{
	For ($i=0;$i -le ($intMailboxTotal - 1);$i++)
	{
		$clb_Step1_Mailboxes_List.SetItemChecked($i,$False)
	}
}

Function Import-TargetsAzureAdUser
{
	Disable-AllTargetsButtons
    $status_Step1.Text = "Step 1 Status: Running"
	$File_Location = $location + "\AzureAdUser.txt"
    if ((Test-Path $File_Location) -eq $true)
	{
		$EventLog = New-Object System.Diagnostics.EventLog('Application')
		$EventLog.MachineName = "."
		$EventLog.Source = "O365DC"
		try{$EventLog.WriteEntry("Starting O365DC Step 1 - Populate","Information", 10)} catch{}
	    $array_AzureAdUser = @(([System.IO.File]::ReadAllLines($File_Location)) | sort-object -Unique)
		$global:intAzureAdUserTotal = 0
	    $clb_Step1_AzureAdUser_List.items.clear()
		foreach ($member_AzureAdUser in $array_AzureAdUser | where-object {$_ -ne ""})
	    {
	        $clb_Step1_AzureAdUser_List.items.add($member_AzureAdUser)
			$global:intAzureAdUserTotal++
	    }
		For ($i=0;$i -le ($intAzureAdUserTotal - 1);$i++)
		{
			$clb_Step1_AzureAdUser_List.SetItemChecked($i,$true)
		}
		$EventLog = New-Object System.Diagnostics.EventLog('Application')
		$EventLog.MachineName = "."
		$EventLog.Source = "O365DC"
		try{$EventLog.WriteEntry("Ending O365DC Step 1 - Populate","Information", 11)} catch{}
		$txt_AzureAdUserTotal.Text = "AzureAdUser count = " + $intAzureAdUserTotal
		$txt_AzureAdUserTotal.visible = $true
	    $status_Step1.Text = "Step 2 Status: Idle"
	}
	else
	{
		write-host	"The file AzureAdUser.txt is not present.  Run Discover to create the file."
		$status_Step1.Text = "Step 1 Status: Failed - AzureAdUser.txt file not found.  Run Discover to create the file."
	}
	Enable-AllTargetsButtons
}

Function Enable-TargetsAzureAdUser
{
	For ($i=0;$i -le ($intAzureAdUserTotal - 1);$i++)
	{
		$clb_Step1_AzureAdUser_List.SetItemChecked($i,$true)
	}
}

Function Disable-TargetsAzureAdUser
{
	For ($i=0;$i -le ($intAzureAdUserTotal - 1);$i++)
	{
		$clb_Step1_AzureAdUser_List.SetItemChecked($i,$False)
	}
}


Function Set-AllFunctionsClientAccess
{
	Param([boolean]$Check)
	$chk_Org_Get_ActiveSyncOrgSettings.checked = $Check
	$chk_Org_Get_MobileDevice.Checked = $Check
	$chk_Org_Get_MobileDevicePolicy.Checked = $Check
	$chk_Org_Get_AvailabilityAddressSpace.Checked = $Check
	$chk_Org_Get_OwaMailboxPolicy.Checked = $Check
}

Function Set-AllFunctionsGlobal
{
	Param([boolean]$Check)
	$chk_Org_Get_AntiPhishPolicy.checked  = $Check
	$chk_Org_Get_AntiSpoofingPolicy.checked  = $Check
	$chk_Org_Get_AtpPolicyForO365.checked  = $Check
	$chk_Org_Get_OnPremisesOrganization.checked = $Check
	$chk_Org_Get_SmimeConfig.checked  = $Check
	$chk_Org_Get_AddressBookPolicy.Checked = $Check
	$chk_Org_Get_AddressList.Checked = $Check
	$chk_Org_Get_EmailAddressPolicy.Checked = $Check
	$chk_Org_Get_GlobalAddressList.Checked = $Check
	$chk_Org_Get_OfflineAddressBook.Checked = $Check
	$chk_Org_Get_OrgConfig.Checked = $Check
	$chk_Org_Get_Rbac.Checked = $Check
	$chk_Org_Get_RetentionPolicy.Checked = $Check
	$chk_Org_Get_RetentionPolicyTag.Checked = $Check
}

Function Set-AllFunctionsRecipient
{
	Param([boolean]$Check)
	$chk_Org_Get_CASMailboxPlan.checked  = $Check
	$chk_Org_Get_Contact.checked  = $Check
	$chk_Org_Get_MailboxPlan.checked  = $Check
	$chk_Org_Get_MailUser.checked  = $Check
	$chk_Org_Get_UnifiedGroup.checked = $Check
	$chk_Org_Get_CalendarProcessing.Checked = $Check
	$chk_Org_Get_CASMailbox.Checked = $Check
	$chk_Org_Get_DistributionGroup.Checked = $Check
	$chk_Org_Get_DynamicDistributionGroup.Checked = $Check
	$chk_Org_Get_Mailbox.Checked = $Check
	$chk_Org_Get_MailboxFolderStatistics.Checked = $Check
	$chk_Org_Get_MailboxPermission.Checked = $Check
	$chk_Org_Get_MailboxStatistics.Checked = $Check
	$chk_Org_Get_PublicFolder.Checked = $Check
	$chk_Org_Get_PublicFolderStatistics.Checked = $Check
	$chk_Org_Quota.Checked = $Check
}

Function Set-AllFunctionsTransport
{
	Param([boolean]$Check)

	$chk_Org_Get_DkimSigningConfig.checked  = $Check
	$chk_Org_Get_AcceptedDomain.Checked = $Check
	$chk_Org_Get_InboundConnector.Checked = $Check
	$chk_Org_Get_RemoteDomain.Checked = $Check
	$chk_Org_Get_OutboundConnector.Checked = $Check
	$chk_Org_Get_TransportConfig.Checked = $Check
	$chk_Org_Get_TransportRule.Checked = $Check
}

Function Set-AllFunctionsUm
{
    Param([boolean]$Check)
	$chk_Org_Get_UmAutoAttendant.Checked = $Check
	$chk_Org_Get_UmDialPlan.Checked = $Check
	$chk_Org_Get_UmIpGateway.Checked = $Check
	$chk_Org_Get_UmMailbox.Checked = $Check
	#$chk_Org_Get_UmMailboxConfiguration.Checked = $Check
	#$chk_Org_Get_UmMailboxPin.Checked = $Check
	$chk_Org_Get_UmMailboxPolicy.Checked = $Check
}

Function Set-AllFunctionsMisc
{
    Param([boolean]$Check)
	$chk_Org_Get_AdminGroups.Checked = $Check
}

Function Set-AllFunctionsAzureAd
{
    Param([boolean]$Check)
	$chk_Azure_Get_AzureADApplication.Checked = $Check
	$chk_Azure_Get_AzureADContact.Checked = $Check
	$chk_Azure_Get_AzureADDevice.Checked = $Check
	$chk_Azure_Get_AzureADDeviceRegisteredOwner.Checked = $Check
	$chk_Azure_Get_AzureADDeviceRegisteredUser.Checked = $Check
	$chk_Azure_Get_AzureADDirectoryRole.Checked = $Check
	$chk_Azure_Get_AzureADDomain.Checked = $Check
	$chk_Azure_Get_AzureAdDomainServiceConfigurationRecord.Checked = $Check
	$chk_Azure_Get_AzureAdDomainVerificationDnsRecord.Checked = $Check
	$chk_Azure_Get_AzureADGroup.Checked = $Check
	$chk_Azure_Get_AzureADGroupMember.Checked = $Check
	$chk_Azure_Get_AzureADGroupOwner.Checked = $Check
	$chk_Azure_Get_AzureADSubscribedSku.Checked = $Check
	$chk_Azure_Get_AzureADTenantDetail.Checked = $Check
	$chk_Azure_Get_AzureADUser.Checked = $Check
	$chk_Azure_Get_AzureADUserLicenseDetail.Checked = $Check
	$chk_Azure_Get_AzureADUserMembership.Checked = $Check
	$chk_Azure_Get_AzureADUserOwnedDevice.Checked = $Check
	$chk_Azure_Get_AzureADUserRegisteredDevice.Checked = $Check
}

Function Set-AllFunctionsSpo
{
    Param([boolean]$Check)
	$chk_Org_Get_SPODeletedSite.Checked = $Check
	$chk_Org_Get_SPOExternalUser.Checked = $Check
	$chk_Org_Get_SPOGeoStorageQuota.Checked = $Check
	$chk_Org_Get_SPOSite.Checked = $Check
	$chk_Org_Get_SPOTenant.Checked = $Check
	$chk_Org_Get_SPOTenantSyncClientRestriction.Checked = $Check
	$chk_Org_Get_SPOUser.Checked = $Check
	$chk_Org_Get_SPOWebTemplate.Checked = $Check
}

Function Start-O365DCJob
{
    param(  [string]$server,`
            [string]$Job,`              # e.g. "Win32_ComputerSystem"
            [boolean]$JobType,`             # 0=WMI, 1=Exchange
            [string]$Location,`
            [string]$JobScriptName,`    # e.g. "dc_w32_cs.ps1"
            [int]$i,`                   # Number or $null
            [string]$PSSession)

    If ($JobType -eq 0) #WMI
        {Limit-O365DCJob $intWMIJobs $intWMIPolling}
    else                #Exchange
        {Limit-O365DCJob $intExchangeJobs $intExchangePolling}
    $strJobName = "$Job job for $server"
    write-host "-- Starting " $strJobName
    $PS_Loc = "$location\O365DC_Scripts\$JobScriptName"
    Start-Job -ScriptBlock {param($a,$b,$c,$d,$e) Powershell.exe -NoProfile -file $a $b $c $d $e} -ArgumentList @($PS_Loc,$location,$server,$i,$PSSession) -Name $strJobName
    start-sleep 1 # Allow time for child job to spawn
}

Function Check-CurrentPSSession
{
	$O365Session = $false
	$PSSession = Get-PSSession
	Foreach ($session in $PSSession)
	{
		# Check for $PSSession to commerical o365
		if (($session.computername -eq "outlook.office365.com") -and ($session.state -eq "opened") -and ($session.ConfigurationName -eq "Microsoft.Exchange"))
			{$O365Session = $true}
	}
	return $o365session
}

Function Split-List10
{
	Param(	$InputFile, `
			$OutputFile, `
			$Text
	)

	write-host "-- Splitting the list of checked $Text... "
	#$File_Location = $location + "\CheckedMailbox.txt"
	$File_Location = $location + "\$InputFile.txt"
	If ((Test-Path $File_Location) -eq $false)
	{
		# Create empty Mailbox.txt file if not present
		write-host "No $Text appear to be selected.  $Text tests will produce no output." -ForegroundColor Red
		"" | Out-File $File_Location
	}
	#$CheckedMailbox = [System.IO.File]::ReadAllLines($File_Location)
	#$CheckedMailboxCount = $CheckedMailbox.count
	#$CheckedMailboxCountSplit = [int]$CheckedMailboxCount/10
	$CheckedList = [System.IO.File]::ReadAllLines($File_Location)
	$CheckedListCount = $CheckedList.count
	$CheckedListCountSplit = [int]$CheckedListCount/10
	$Count = 0
	For ($FileCount = 1; $FileCount -le 10; $FileCount++)
	{
		if ((Test-Path ".\$OutputFile.Set$FileCount.txt") -eq $true) {Remove-Item ".\$OutputFile.Set$FileCount.txt" -Force}
		For (;$Count -lt ($FileCount*$CheckedListCountSplit);$Count++)
		{$CheckedList[$Count] | Out-File ".\$OutputFile.Set$FileCount.txt" -Append -Force}
	}
}

Function Test-CheckBoxAndRun
{
	Param (	$chkBox,` 	#$chk_Azure_Get_AzureAdUser.Checked
			$Text,`		#Get-AzureADUser
			$Script		#Azure_GetAzureAdUser
	)

	If ($chkbox -eq $true)
	{
		write-host "Starting $Text" -foregroundcolor green
		try
			{. $location\O365DC_Scripts\$Script.ps1 -location $location}
		catch [System.Management.Automation.CommandNotFoundException]
			{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
	}
}

#endregion *** Custom Functions ***

# Check Powershell version
$PowershellVersionNumber = $null
$powershellVersion = get-host
# Teminate if Powershell is less than version 2
if ($powershellVersion.Version.Major -lt "2")
{
    write-host "Unsupported Powershell version detected."
    write-host "Powershell v2 is required."
    end
}
# Powershell v2 or later required for Ex2010 environments
elseif ($powershellVersion.Version.Major -lt "3")
{
    $PowershellVersionNumber = 2
	write-host "Powershell version 2 detected" -ForegroundColor Green
}
# Powershell v3 or later required for Ex2013 or later environments
elseif ($powershellVersion.Version.Major -lt "4")
{
    $PowershellVersionNumber = 3
    write-host "Powershell version 3 detected" -ForegroundColor Green
}
elseif ($powershellVersion.Version.Major -lt "5")
{
    $PowershellVersionNumber = 4
    write-host "Powershell version 4 detected" -ForegroundColor Green
}
elseif ($powershellVersion.Version.Major -lt "6")
{
    $PowershellVersionNumber = 5
    write-host "Powershell version 5 detected" -ForegroundColor Green
}

# Check for presence of Powershell Profile and warn if present
if ((test-path $PROFILE) -eq $true)
{
	write-host "WARNING: Powershell profile detected." -ForegroundColor Red
	write-host "WARNING: All jobs will be executed using the -NoProfile switch" -ForegroundColor Red
}
else
{
	write-host "No Powershell profile detected." -ForegroundColor Green
}


# Connecting to all Office 365 services
# Prereq:
# Azure Active Directory V2 - https://docs.microsoft.com/en-us/office365/enterprise/powershell/connect-to-office-365-powershell##connect-with-the-azure-active-directory-powershell-for-graph-module
# SharePoint Online Management Shell - https://go.microsoft.com/fwlink/p/?LinkId=255251
# Skype for Business Online, Windows PowerShell Module - https://go.microsoft.com/fwlink/p/?LinkId=532439
# Exchange Online Remote Powershell that supports MFA - Hybrid blade of EAC - https://cmdletpswmodule.blob.core.windows.net/exopsmodule/Microsoft.Online.CSE.PSModule.Client.application
# Security and Compliance Center - https://docs.microsoft.com/en-us/powershell/exchange/office-365-scc/connect-to-scc-powershell/mfa-connect-to-scc-powershell?view=exchange-ps

$TestGUI = $true
if ($TestGUI -eq $true)
{}
else {
# Try to re-use an existing connection
$CurrentPSSession = Check-CurrentPsSession
If (($CurrentPSSession -eq $false) -or ($ForceNewConnection -eq $True))
{
	$domainHost = read-host "Please enter the domain host name.  (For example, if the domain is litware.onmicrosoft.com then enter 'litware'): "

	# Connection Urls
	# Hard-code these for initial testing
	# Commercial
	$SharepointAdminCenter = "https://$domainhost-admin.sharepoint.com"
	$ExoConnectioUri = "https://outlook.office365.com/powershell-liveid/"
	$SccConnectionUri = "https://ps.compliance.protection.outlook.com/powershell-liveid/"

	# 21Vianet

	# Office365 Germany
	# SccConnectionUri = https://ps.compliance.protection.outlook.de/powershell-liveid/
	# Connect-AzureAD -AzureEnvironment "AzureGermanyCloud"

	# US Government GCC

	# US Government DOD



	If ($MFA -ne $true)
	{
		write-host "Since MFA is not in use, we can store the credentials for re-use." -foregroundcolor green
		# Blank does not work and canceling the cred box doesn't work
		# Use switch for multiple credentials?
		$O365Cred = get-credential
	}
	elseif ($MFA -eq $true)
	{
		$O365Upn = Read-host "Please enter the user principal name with access to the tenant: "
	}

	# Connect to Azure AD
	if ($MFA -eq $true)
	{
		write-host "Multi-factor authentication is enabled" -foregroundcolor yellow
		write-host "Connecting to AzureAD" -ForegroundColor Green
		Connect-AzureAD -AccountId $O365Upn
	}
	else
	{
		#$AzureCredential = Get-Credential -UserName $AzureAdmin
		write-host "Connecting to AzureAD" -ForegroundColor Green
		Connect-AzureAD -Credential $O365Cred
	}

	# Connect to Sharepoint
	# Need to check this with non-MFA
	Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
	if ($MFA -eq $true)
	{
		write-host "Multi-factor authentication is enabled" -foregroundcolor yellow
		write-host "Connecting to Sharepoint Online" -ForegroundColor Green
		Connect-SPOService -url $SharepointAdminCenter -UserName $O365Upn
	}
	else
	{
		write-host "Connecting to Sharepoint Online" -ForegroundColor Green
		Connect-SPOService -url $SharepointAdminCenter -Credential $O365Cred
	}

	#Connect to Skype
	# Need to check this with non-MFA
	Import-Module SkypeOnlineConnector
	if ($MFA -eq $true)
	{
		write-host "Multi-factor authentication is enabled" -foregroundcolor yellow
		write-host "Connecting to Skype for Business Online" -ForegroundColor Green
		$CsSession = New-CsOnlineSession -UserName $O365Upn
	}
	else
	{
		write-host "Connecting to Skype for Business Online" -ForegroundColor Green
		$CsSession = New-CsOnlineSession -Credential $O365Cred
	}
	Import-PSSession $CsSession

	#Connect to Exchange Online
	If ($MFA -eq $true)
	{
		write-host "Multi-factor authentication is enabled" -foregroundcolor yellow
		write-host "Connecting to Exchange Online" -foregroundcolor green
		$ModuleLocation = "$($env:LOCALAPPDATA)\Apps\2.0"
		$ExoModuleLocation = @(Get-ChildItem -Path $ModuleLocation -Filter "Microsoft.Exchange.Management.ExoPowershellModule.manifest" -Recurse )
		If ($ExoModuleLocation.Count -ge 1)
		{
			write-host "ExoPowershellModule.manifest found.  Trying to load the dll." -foregroundcolor green
			$FullExoModulePath =  $ExoModuleLocation[0].Directory.tostring() + "\Microsoft.Exchange.Management.ExoPowershellModule.dll"
			Import-Module $FullExoModulePath  -Force
			$ExoSession	= New-ExoPSSession
		}
	}
	else
	{
		write-host "Connecting to Exchange Online" -foregroundcolor green
		$ExoSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExoConnectioUri -Credential $O365Cred -Authentication Basic -AllowRedirection
	}
	Import-PSSession $ExoSession

	#Connect to Security and Compliance Center
	If ($MFA -eq $true)
	{
		write-host "Multi-factor authentication is enabled" -foregroundcolor yellow
		write-host "Connecting to Security and Compliance Center" -foregroundcolor green
		Connect-IPPSSession -UserPrincipalName $AzureAdmin
	}
	else
	{
		write-host "Connecting to Security and Compliance Center" -foregroundcolor green
		$SccSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $SccConnectionUri -Credential $O365Cred -Authentication Basic -AllowRedirection
		Import-PSSession $SccSession -Prefix CC
	}
}
else
{
	#write-host "check session: "$CurrentPSSession
	#write-host "Force new connection: " $ForceNewConnection
	write-host "Existing connection to Microsoft.Exchange on outlook.office365.com detected."
}

# end TestGUI
}

#----------------------------------------------
# Write to Event Log
#----------------------------------------------
$EventLog = New-Object System.Diagnostics.EventLog('Application')
$EventLog.MachineName = "."
$EventLog.Source = "O365DC"
try{$EventLog.WriteEntry("Starting O365DC Run","Information", 1)} catch{}

#----------------------------------------------
# Initialize Arrays and Variables
#----------------------------------------------
Set-Variable -name intExchangeJobs -Scope global
Set-Variable -name intExchangePolling -Scope global
Set-Variable -name intExchJobTimeout -Scope global
Set-Variable -name INI -Scope global
Set-Variable -name intMailboxTotal -Scope global

$array_Mailboxes = @()
$UM=$true

if ($JobCount_Exchange -eq 0) 			{$intExchangeJobs = 10}
	else 							{$intExchangeJobs = $JobCount_Exchange}
if ($JobPolling_Exchange -eq 0) 		{$intExchangePolling = 5}
	else 							{$intExchangePolling = $JobPolling_Exchange}
if ($Timeout_Exchange_Job -eq 0)		{$intExchJobTimeout = 3600} 			# 3600 sec = 60 min
	else 							{$intExchJobTimeout = $Timeout_Exchange_Job}

#Set timestamp
$StartTime = Get-Date -UFormat %s
$append = $StartTime
$append = "v4_0_2." + $append
#----------------------------------------------
# Misc Code
#----------------------------------------------
$ScriptLoc = Split-Path -parent $MyInvocation.MyCommand.Definition
Set-Location $ScriptLoc
$location = [string]((get-Location).path)
$testfolder = test-path output
if ($testfolder -eq $false)
{
	new-item -name "output" -type directory -force | Out-Null
}
#Call the Function
write-host "Starting Office 365 Data Collector (O365DC) v4 with the following parameters: " -ForegroundColor Cyan
$EventText = "Starting Office 365 Data Collector (O365DC) v4 with the following parameters: `n"
if ($NoEMS -eq $false)
{
	write-host "`tIni Settings`t" -ForegroundColor Cyan
	$EventText += "`tIni Settings:`t" + $INI + "`n"
	write-host "`t`tExchange Ini:`t" $INI_Exchange -ForegroundColor Cyan
	$EventText += "`t`tExchange Ini:`t" + $INI_Exchange + "`n"
	write-host "`tNon-Exchange cmdlet jobs" -ForegroundColor Cyan
	$EventText += "`tNon-Exchange cmdlet jobs`n"
	write-host "`tExchange cmdlet jobs" -ForegroundColor Cyan
	$EventText += "`tExchange cmdlet jobs`n"
	write-host "`t`tMax jobs:`t" $intExchangeJobs -ForegroundColor Cyan
	$EventText += "`t`tMax jobs:`t" + $intExchangeJobs + "`n"
	write-host "`t`tPolling: `t" $intExchangePolling " seconds" -ForegroundColor Cyan
	$EventText += "`t`tPolling: `t`t" + $intExchangePolling + "`n"
	write-host "`t`tTimeout:`t" $intExchJobTimeout " seconds" -ForegroundColor Cyan
	$EventText += "`t`tTimeout:`t" + $intExchJobTimeout + "`n"

	$EventLog = New-Object System.Diagnostics.EventLog('Application')
	$EventLog.MachineName = "."
	$EventLog.Source = "O365DC"
	try{$EventLog.WriteEntry($EventText,"Information", 4)} catch{}
}
else
{
	write-host "`tNoEMS switch used" -ForegroundColor Cyan
	$EventLog = New-Object System.Diagnostics.EventLog('Application')
	$EventLog.MachineName = "."
	$EventLog.Source = "O365DC"
	try{$EventLog.WriteEntry("NoEMS switch used.","Information", 4)} catch{}
}

# Let's start the party
New-O365DCForm