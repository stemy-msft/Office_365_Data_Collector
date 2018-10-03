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
.PARAMETER JobCount_ExOrg
	Max number of jobs for Exchange cmdlet functions (Default = 10)
	Caution: The OOB throttling policy sets PowershellMaxConcurrency at 18 sessions per user per server
	Modifying this value without increasing the throttling policy can cause ExOrg jobs to immediately fail.
.PARAMETER JobPolling_ExOrg
	Polling interval for job completion for Exchange cmdlet functions (Default = 5 sec)
.PARAMETER Timeout_ExOrg_Job
	Job timeout for Exchange functions  (Default = 3600 sec)
	The default value is 3600 seconds but should be adjusted for organizations with a large number of mailboxes or servers over slow connections.
.PARAMETER ServerForPSSession
	Exchange server used for Powershell sessions
.PARAMETER INI_ExOrg
	Specify INI file for ExOrg Tests configuration
.PARAMETER NoEMS
	Use this switch to launch the tool in Powershell (No Exchange cmdlets)
.PARAMETER MFA
	Use this switch if Multi-Factor Authentication is required for the environment.
	If is recommended that the Trusted IPs be set in Azure AD Conditional Access to allow the admin account to use traditional user name
	and password when run from a trusted IP.  If this switch is set, multi-threading will be disabled.
.PARAMETER ForceNewConnection
	Use this switch to force Powershell to make a new connection to Office 365 instead of trying to re-use an existing session.
.EXAMPLE
	.\O365DC.v4.ps1 -JobCount_ExOrg 12
	This results in O365DC using 12 active ExOrg jobs instead of the default of 10.
.EXAMPLE
	.\O365DC.v4.ps1 -JobPolling_ExOrg 30
	This results in O365DC polling for completed ExOrg jobs every 30 seconds.
.EXAMPLE
	.\O365DC.v4.ps1 -Timeout_ExOrg_Job 7200
	This results in O365DC killing ExOrg jobs that have exceeded 7200 seconds at the next polling interval.
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

Param(	[int]$JobCount_ExOrg = 2,`
		[int]$JobPolling_ExOrg = 5,`
		[int]$Timeout_ExOrg_Job = 3600,`
		[string]$ServerForPSSession = $null,`
		[string]$INI_ExOrg,`
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

#region Step3 ExOrg Tier2
$tab_Step3_ExOrg = New-Object System.Windows.Forms.TabPage
$tab_Step3_ExOrg_Tier2 = New-Object System.Windows.Forms.TabControl
#endregion Step3 ExOrg Tier2

#region Step3 Client Access tab
$tab_Step3_ClientAccess = New-Object System.Windows.Forms.TabPage
$bx_ClientAccess_Functions = New-Object System.Windows.Forms.GroupBox
$btn_Step3_ClientAccess_CheckAll = New-Object System.Windows.Forms.Button
$btn_Step3_ClientAccess_UncheckAll = New-Object System.Windows.Forms.Button
$chk_Org_Get_MobileDevice = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_MobileDevicePolicy = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_AvailabilityAddressSpace = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_OWAMailboxPolicy = New-Object System.Windows.Forms.CheckBox

#endregion Step3 Client Access tab

#region Step3 Global tab
$tab_Step3_Global = New-Object System.Windows.Forms.TabPage
$bx_Global_Functions = New-Object System.Windows.Forms.GroupBox
$btn_Step3_Global_CheckAll = New-Object System.Windows.Forms.Button
$btn_Step3_Global_UncheckAll = New-Object System.Windows.Forms.Button
$chk_Org_Get_AddressBookPolicy  = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_AddressList  = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_EmailAddressPolicy = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_GlobalAddressList = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_OfflineAddressBook = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_OrgConfig = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_Rbac = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_RetentionPolicy = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_RetentionPolicyTag = New-Object System.Windows.Forms.CheckBox
#endregion Step3 Global tab

#region Step3 Recipient Tab
$tab_Step3_Recipient = New-Object System.Windows.Forms.TabPage
$bx_Recipient_Functions = New-Object System.Windows.Forms.GroupBox
$btn_Step3_Recipient_CheckAll = New-Object System.Windows.Forms.Button
$btn_Step3_Recipient_UncheckAll = New-Object System.Windows.Forms.Button
$chk_Org_Get_CalendarProcessing = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_CASMailbox = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_DistributionGroup = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_DynamicDistributionGroup = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_Mailbox = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_MailboxFolderStatistics = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_MailboxPermission = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_MailboxStatistics = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_PublicFolder = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_PublicFolderStatistics = New-Object System.Windows.Forms.CheckBox
$chk_Org_Quota = New-Object System.Windows.Forms.CheckBox
#endregion Step3 Recipient Tab

#region Step3 Transport Tab
$tab_Step3_Transport = New-Object System.Windows.Forms.TabPage
$bx_Transport_Functions = New-Object System.Windows.Forms.GroupBox
$btn_Step3_Transport_CheckAll = New-Object System.Windows.Forms.Button
$btn_Step3_Transport_UncheckAll = New-Object System.Windows.Forms.Button
$chk_Org_Get_AcceptedDomain = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_InboundConnector = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_RemoteDomain = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_OutboundConnector = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_TransportConfig = New-Object System.Windows.Forms.CheckBox
$chk_Org_Get_TransportRule = New-Object System.Windows.Forms.CheckBox
#endregion Step3 Transport Tab

#region Step3 Unified Messaging tab
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

#region Step3 Misc Tab
$tab_Step3_Misc = New-Object System.Windows.Forms.TabPage
$bx_Misc_Functions = New-Object System.Windows.Forms.GroupBox
$btn_Step3_Misc_CheckAll = New-Object System.Windows.Forms.Button
$btn_Step3_Misc_UncheckAll = New-Object System.Windows.Forms.Button
$chk_Org_Get_AdminGroups = New-Object System.Windows.Forms.CheckBox
#endregion Step3 Misc Tab

#EndRegion Step3 - Tests

#region Step4 - Reporting
$tab_Step4 = New-Object System.Windows.Forms.TabPage
$btn_Step4_Assemble = New-Object System.Windows.Forms.Button
$lbl_Step4_Assemble = New-Object System.Windows.Forms.Label
$bx_Step4_Functions = New-Object System.Windows.Forms.GroupBox
#$chk_Step4_DC_Report = New-Object System.Windows.Forms.CheckBox
#$chk_Step4_Ex_Report = New-Object System.Windows.Forms.CheckBox
$chk_Step4_ExOrg_Report = New-Object System.Windows.Forms.CheckBox
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
	$Message_About = $Message_About += "Release Date: September 2018 `n`n"
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
		get-mailbox -resultsize unlimited | ForEach-Object `
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
		try{& ".\O365DC_Scripts\Core_Parse_Ini_File.ps1" -IniFile ".\Templates\Template1_INI_ExOrg.ini"} catch{}
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
		try{& ".\O365DC_Scripts\Core_Parse_Ini_File.ps1" -IniFile ".\Templates\Template2_INI_ExOrg.ini"} catch{}
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
		try{& ".\O365DC_Scripts\Core_Parse_Ini_File.ps1" -IniFile ".\Templates\Template3_INI_ExOrg.ini"} catch {}
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
		try{& ".\O365DC_Scripts\Core_Parse_Ini_File.ps1" -IniFile ".\Templates\Template4_INI_ExOrg.ini"} catch{}
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
	$EventLogText = "Starting O365DC Step 3`nDomain Controllers: $intDCTotal`nExchange Servers: $intExTotal`nMailboxes: $intMailboxTotal"
	try{$EventLog.WriteEntry($EventLogText,"Information", 30)} catch{}
	#send the form to the back to expose the Powershell window when starting Step 3
	$form1.WindowState = "minimized"
	write-host "O365DC Form minimized." -ForegroundColor Green

	#Region Executing Exchange Organization Tests
	write-host "Starting Exchange Organization..." -ForegroundColor Green
	If (Get-ExOrgBoxStatus = $true)
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

		If (Get-ExOrgMbxBoxStatus = $true)
		{
			# Avoid this path if we're not running mailbox tests
			# Splitting CheckedMailboxes file 10 times
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
		}

		# First we start the jobs that query the organization instead of the Exchange server
		#Region ExOrg Non-server Functions
		If ($chk_Org_Get_AcceptedDomain.checked -eq $true)
			{
				write-host "Starting Get-AcceptedDomain" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetAcceptedDomain.ps1 -location $location}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		If ($chk_Org_Get_AddressBookPolicy.checked -eq $true)
			{
				write-host "Starting Get-AddressBookPolicy" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetAddressBookPolicy.ps1 -location $location}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		If ($chk_Org_Get_AddressList.checked -eq $true)
			{
				write-host "Starting Get-AddressList" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetAddressList.ps1 -location $location}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		If ($chk_Org_Get_AdminGroups.checked -eq $true)
			{
				write-host "Starting Get-AdminGroups" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_Misc_AdminGroups.ps1 -location $location}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		If ($chk_Org_Get_AvailabilityAddressSpace.checked -eq $true)
			{
				write-host "Starting Get-AvailabilityAddressSpace" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetAvailabilityAddressSpace.ps1 -location $location}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		If ($chk_Org_Get_CalendarProcessing.checked -eq $true)
		{
			For ($i = 1;$i -lt 11;$i++)
			{
				write-host "Starting Get-CalendarProcessing job $i" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetCalendarProcessing.ps1 -location $location -i $i}
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
					{.\O365DC_Scripts\ExOrg_GetCASMailbox.ps1 -location $location -i $i}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		}
		If ($chk_Org_Get_DistributionGroup.checked -eq $true)
			{
				write-host "Starting Get-DistributionGroup" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetDistributionGroup.ps1 -location $location}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		If ($chk_Org_Get_DynamicDistributionGroup.checked -eq $true)
			{
				write-host "Starting Get-DynamicDistributionGroup" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetDynamicDistributionGroup.ps1 -location $location}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		If ($chk_Org_Get_EmailAddressPolicy.checked -eq $true)
			{
				write-host "Starting Get-EmailAddressPolicy" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetEmailAddressPolicy.ps1 -location $location}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		If ($chk_Org_Get_GlobalAddressList.checked -eq $true)
			{
				write-host "Starting Get-GlobalAddressList job" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetGlobalAddressList.ps1 -location $location -i $i}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		If ($chk_Org_Get_Mailbox.checked -eq $true)
		{
			For ($i = 1;$i -lt 11;$i++)
			{
				write-host "Starting Get-Mailbox job $i" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetMbx.ps1 -location $location -i $i}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		}
		If ($chk_Org_Get_InboundConnector.checked -eq $true)
			{
				write-host "Starting Get-InboundConnector" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetInboundConnector.ps1 -location $location}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		If ($chk_Org_Get_MailboxFolderStatistics.checked -eq $true)
		{
			For ($i = 1;$i -lt 11;$i++)
			{
				write-host "Starting Get-MailboxFolderStatistics job $i" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetMbxFolderStatistics.ps1 -location $location -i $i}
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
					{.\O365DC_Scripts\ExOrg_GetMbxPermission.ps1 -location $location -i $i}
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
					{.\O365DC_Scripts\ExOrg_GetMbxStatistics.ps1 -location $location -i $i}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		}
		If ($chk_Org_Get_MobileDevice.checked -eq $true)
			{
				write-host "Starting Get-MobileDevice" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetMobileDevice.ps1 -location $location}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		If ($chk_Org_Get_MobileDevicePolicy.checked -eq $true)
			{
				write-host "Starting Get-MobileDevicePolicy" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetMobileDeviceMbxPolicy.ps1 -location $location}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		If ($chk_Org_Get_OfflineAddressBook.checked -eq $true)
			{
				write-host "Starting Get-OfflineAddressBook" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetOfflineAddressBook.ps1 -location $location}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		If ($chk_Org_Get_OrgConfig.checked -eq $true)
			{
				write-host "Starting Get-OrgConfig" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetOrgConfig.ps1 -location $location}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		If ($chk_Org_Get_OutboundConnector.checked -eq $true)
			{
				write-host "Starting Get-OutboundConnector" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetOutboundConnector.ps1 -location $location}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		If ($chk_Org_Get_OwaMailboxPolicy.checked -eq $true)
			{
				write-host "Starting Get-OwaMailboxPolicy" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetOwaMailboxPolicy.ps1 -location $location}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		If ($chk_Org_Get_PublicFolder.checked -eq $true)
			{
				write-host "Starting Get-PublicFolder" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetPublicFolder.ps1 -location $location}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		If ($chk_Org_Get_PublicFolderStatistics.checked -eq $true)
			{
				write-host "Starting Get-PublicFolderStatistics" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetPublicFolderStats.ps1 -location $location}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		If ($chk_Org_Quota.checked -eq $true)
			{
				For ($i = 1;$i -lt 11;$i++)
				{
					write-host "Starting Get-Quota" -foregroundcolor green
					try
						{.\O365DC_Scripts\ExOrg_Quota.ps1 -location $location -i $i}
					catch [System.Management.Automation.CommandNotFoundException]
						{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
				}
			}
		If ($chk_Org_Get_Rbac.checked -eq $true)
			{
				write-host "Starting Get-Rbac" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetRbac.ps1 -location $location}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		If ($chk_Org_Get_RemoteDomain.checked -eq $true)
			{
				write-host "Starting Get-RemoteDomain" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetRemoteDomain.ps1 -location $location}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		If ($chk_Org_Get_RetentionPolicy.checked -eq $true)
			{
				write-host "Starting Get-RetentionPolicy" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetRetentionPolicy.ps1 -location $location}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		If ($chk_Org_Get_RetentionPolicyTag.checked -eq $true)
			{
				write-host "Starting Get-RetentionPolicyTag" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetRetentionPolicyTag.ps1 -location $location}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		If ($chk_Org_Get_TransportConfig.checked -eq $true)
			{
				write-host "Starting Get-TransportConfig" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetTransportConfig.ps1 -location $location}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		If ($chk_Org_Get_TransportRule.checked -eq $true)
			{
				write-host "Starting Get-TransportRule" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetTransportRule.ps1 -location $location}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		If ($chk_Org_Get_UmAutoAttendant.checked -eq $true)
			{
				write-host "Starting Get-UmAutoAttendant" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetUmAutoAttendant.ps1 -location $location}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		If ($chk_Org_Get_UmDialPlan.checked -eq $true)
			{
				write-host "Starting Get-UmDialPlan" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetUmDialPlan.ps1 -location $location}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		If ($chk_Org_Get_UmIpGateway.checked -eq $true)
			{
				write-host "Starting Get-UmIpGateway" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetUmIpGateway.ps1 -location $location}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		If ($chk_Org_Get_UmMailbox.checked -eq $true)
		{
			For ($i = 1;$i -lt 11;$i++)
			{
				write-host "Starting Get-UmMailbox job $i" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetUmMailbox.ps1 -location $location -i $i}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		}
			#			If ($chk_Org_Get_UmMailboxConfiguration.checked -eq $true)
#			{
#				For ($i = 1;$i -lt 11;$i++)
#				{Start-O365DCJob -server $server -job "Get-UmMailboxConfiguration - Set $i" -JobType 0 -Location $location -JobScriptName "ExOrg_GetUmMailboxConfiguration.ps1" -i $i -PSSession $session_0}
#			}
#			If ($chk_Org_Get_UmMailboxPin.checked -eq $true)
#			{
#				For ($i = 1;$i -lt 11;$i++)
#				{Start-O365DCJob -server $server -job "Get-UmMailboxPin - Set $i" -JobType 0 -Location $location -JobScriptName "ExOrg_GetUmMailboxPin.ps1" -i $i -PSSession $session_0}
#			}
		If ($chk_Org_Get_UmMailboxPolicy.checked -eq $true)
			{
				write-host "Starting Get-UmMailboxPolicy" -foregroundcolor green
				try
					{.\O365DC_Scripts\ExOrg_GetUmMailboxPolicy.ps1 -location $location}
				catch [System.Management.Automation.CommandNotFoundException]
					{write-host "Cmdlet is not available in this PSSession." -foregroundcolor red}
			}
		#EndRegion ExOrg Non-Server Functions
	}
	else
	{
		write-host "---- No Exchange Organization Functions selected"
	}
	#EndRegion Executing Exchange Organization Tests

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
 #>		if ($chk_Step4_ExOrg_Report.checked -eq $true)
		{
			write-host "-- Starting to assemble the Exchange Organization Spreadsheet"
				.\O365DC_Scripts\Core_assemble_exorg_Excel.ps1 -RunLocation $location
				write-host "---- Completed the Exchange Organization Spreadsheet" -ForegroundColor Green
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
	$System_Drawing_Size_buttons = New-Object System.Drawing.Size
	$System_Drawing_Size_buttons.Height = 38
	$System_Drawing_Size_buttons.Width = 110
# Reusable status
	$System_Drawing_Size_Status = New-Object System.Drawing.Size
	$System_Drawing_Size_Status.Height = 22
	$System_Drawing_Size_Status.Width = 651
	$System_Drawing_Point_Status = New-Object System.Drawing.Point
	$System_Drawing_Point_Status.X = 3
	$System_Drawing_Point_Status.Y = 653
# Reusable tabs
	$System_Drawing_Size_tab_1 = New-Object System.Drawing.Size
	$System_Drawing_Size_tab_1.Height = 678
	$System_Drawing_Size_tab_1.Width = 700 #657
	$System_Drawing_Size_tab_2 = New-Object System.Drawing.Size
	$System_Drawing_Size_tab_2.Height = 678
	$System_Drawing_Size_tab_2.Width = 1000
# Reusable checkboxes
	$System_Drawing_Size_Reusable_chk = New-Object System.Drawing.Size
	$System_Drawing_Size_Reusable_chk.Height = 20
	$System_Drawing_Size_Reusable_chk.Width = 225
	$System_Drawing_Size_Reusable_chk_long = New-Object System.Drawing.Size
	$System_Drawing_Size_Reusable_chk_long.Height = 20
	$System_Drawing_Size_Reusable_chk_long.Width = 400

# Main Form
$form1.BackColor = [System.Drawing.Color]::FromArgb(255,169,169,169)
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 718
	$System_Drawing_Size.Width = 665
	$form1.ClientSize = $System_Drawing_Size
	$form1.MaximumSize = $System_Drawing_Size
	$form1.Font = $font_Calibri_10pt_normal
	$form1.FormBorderStyle = 2
	$form1.MaximizeBox = $False
	$form1.Name = "form1"
	$form1.ShowIcon = $False
	$form1.StartPosition = 1
	$form1.Text = "Office 365 Data Collector v4.0.2"

# Main Tabs
$tab_Master.Appearance = 2
	$tab_Master.Dock = 5
	$tab_Master.Font = $font_Calibri_14pt_normal
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 32
	$System_Drawing_Size.Width = 100
	$tab_Master.ItemSize = $System_Drawing_Size
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 0
	$System_Drawing_Point.Y = 0
	$tab_Master.Location = $System_Drawing_Point
	$tab_Master.Name = "tab_Master"
	$tab_Master.SelectedIndex = 0
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 718
	$System_Drawing_Size.Width = 665
	$tab_Master.Size = $System_Drawing_Size
	$tab_Master.SizeMode = "filltoright"
	$tab_Master.TabIndex = 12
	$form1.Controls.Add($tab_Master)

# Menu Strip
$Menu_Main.Location = new-object System.Drawing.Point(0, 0)
	$Menu_Main.Name = "MainMenu"
	$Menu_Main.Size = new-object System.Drawing.Size(1151, 24)
	$Menu_Main.TabIndex = 0
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
#EndRegion Form Main

#Region "Step1 - Targets"

#Region Step1 Main
# Reusable text box in Step1
	$System_Drawing_Size_Step1_text_box = New-Object System.Drawing.Size
	$System_Drawing_Size_Step1_text_box.Height = 27
	$System_Drawing_Size_Step1_text_box.Width = 400
# Reusable label in Step1
	$System_Drawing_Size_Step1_label = New-Object System.Drawing.Size
	$System_Drawing_Size_Step1_label.Height = 20
	$System_Drawing_Size_Step1_label.Width = 200
# Reusable Listbox in Step1
	$System_Drawing_Size_Step1_Listbox = New-Object System.Drawing.Size
	$System_Drawing_Size_Step1_Listbox.Height = 384
	$System_Drawing_Size_Step1_Listbox.Width = 200
# Reusable boxes in Step1 Tabs
	$System_Drawing_Size_Step1_box = New-Object System.Drawing.Size
	$System_Drawing_Size_Step1_box.Height = 482
	$System_Drawing_Size_Step1_box.Width = 536
# Reusable check buttons in Step1 tabs
	$System_Drawing_Size_Step1_btn = New-Object System.Drawing.Size
	$System_Drawing_Size_Step1_btn.Height = 25
	$System_Drawing_Size_Step1_btn.Width = 150
# Reusable check list boxes in Step1 tabs
	$System_Drawing_Size_Step1_clb = New-Object System.Drawing.Size
	$System_Drawing_Size_Step1_clb.Height = 350
	$System_Drawing_Size_Step1_clb.Width = 400
	$System_Drawing_Point_Step1_clb = New-Object System.Drawing.Point
	$System_Drawing_Point_Step1_clb.X = 50
	$System_Drawing_Point_Step1_clb.Y = 50
# Reusable Discover/populate buttons in Step1 tabs
	$System_Drawing_Point_Step1_Discover = New-Object System.Drawing.Point
	$System_Drawing_Point_Step1_Discover.X = 50
	$System_Drawing_Point_Step1_Discover.Y = 15
	$System_Drawing_Point_Step1_Populate = New-Object System.Drawing.Point
	$System_Drawing_Point_Step1_Populate.X = 300
	$System_Drawing_Point_Step1_Populate.Y = 15
# Reusable check/uncheck buttons in Step1 tabs
	$System_Drawing_Point_Step1_CheckAll = New-Object System.Drawing.Point
	$System_Drawing_Point_Step1_CheckAll.X = 50
	$System_Drawing_Point_Step1_CheckAll.Y = 450
	$System_Drawing_Point_Step1_UncheckAll = New-Object System.Drawing.Point
	$System_Drawing_Point_Step1_UncheckAll.X = 300
	$System_Drawing_Point_Step1_UncheckAll.Y = 450
$tab_Step1.BackColor = [System.Drawing.Color]::FromArgb(0,255,255,255)
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 4
	$System_Drawing_Point.Y = 36
	$tab_Step1.Location = $System_Drawing_Point
	$tab_Step1.Name = "tab_Step1"
	$tab_Step1.Padding = $System_Windows_Forms_Padding_Reusable
	$tab_Step1.TabIndex = 0
	$tab_Step1.Text = "  Targets  "
	$tab_Step1.Size = $System_Drawing_Size_tab_1
	$tab_Master.Controls.Add($tab_Step1)
$btn_Step1_Discover.Font = $font_Calibri_14pt_normal
	$btn_Step1_Discover.ForeColor = [System.Drawing.Color]::FromArgb(255,0,0,0)
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 20
	$System_Drawing_Point.Y = 15
	$btn_Step1_Discover.Location = $System_Drawing_Point
	$btn_Step1_Discover.Name = "btn_Step1_Discover"
	$btn_Step1_Discover.Size = $System_Drawing_Size_buttons
	$btn_Step1_Discover.TabIndex = 0
	$btn_Step1_Discover.Text = "Discover"
	$btn_Step1_Discover.Visible = $false
	$btn_Step1_Discover.UseVisualStyleBackColor = $True
	$btn_Step1_Discover.add_Click($handler_btn_Step1_Discover_Click)
	$tab_Step1.Controls.Add($btn_Step1_Discover)
$btn_Step1_Populate.Font = $font_Calibri_14pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 200
	$System_Drawing_Point.Y = 15
	$btn_Step1_Populate.Location = $System_Drawing_Point
	$btn_Step1_Populate.Name = "btn_Step1_Populate"
	$btn_Step1_Populate.Size = $System_Drawing_Size_buttons
	$btn_Step1_Populate.TabIndex = 9
	$btn_Step1_Populate.Text = "Load from File"
	$btn_Step1_Populate.Visible = $false
	$btn_Step1_Populate.UseVisualStyleBackColor = $True
	$btn_Step1_Populate.add_Click($handler_btn_Step1_Populate_Click)
	$tab_Step1.Controls.Add($btn_Step1_Populate)
$tab_Step1_Master.Font = $font_Calibri_12pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 20
	$System_Drawing_Point.Y = 60
	$tab_Step1_Master.Location = $System_Drawing_Point
	$tab_Step1_Master.Name = "tab_Step1_Master"
	$tab_Step1_Master.SelectedIndex = 0
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 525
	$System_Drawing_Size.Width = 550
	$tab_Step1_Master.Size = $System_Drawing_Size
	$tab_Step1_Master.TabIndex = 11
	$tab_Step1.Controls.Add($tab_Step1_Master)
$status_Step1.Font = $font_Calibri_10pt_normal
	$status_Step1.Location = $System_Drawing_Point_Status
	$status_Step1.Name = "status_Step1"
	$status_Step1.Size = $System_Drawing_Size_Status
	$status_Step1.TabIndex = 2
	$status_Step1.Text = "Step 1 Status"
	$tab_Step1.Controls.Add($status_Step1)
#EndRegion Step1 Main

#Region Step1 Mailboxes tab
$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 4
	$System_Drawing_Point.Y = 33
$tab_Step1_Mailboxes.Location = $System_Drawing_Point
	$tab_Step1_Mailboxes.Name = "tab_Step1_Mailboxes"
	$tab_Step1_Mailboxes.Padding = $System_Windows_Forms_Padding_Reusable
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 488
	$System_Drawing_Size.Width = 542
	$tab_Step1_Mailboxes.Size = $System_Drawing_Size
	$tab_Step1_Mailboxes.TabIndex = 1
	$tab_Step1_Mailboxes.Text = "Mailboxes"
	$tab_Step1_Mailboxes.UseVisualStyleBackColor = $True
	$tab_Step1_Master.Controls.Add($tab_Step1_Mailboxes)
$btn_Step1_Mailboxes_Discover.Font = $font_Calibri_10pt_normal
	$btn_Step1_Mailboxes_Discover.Location = $System_Drawing_Point_Step1_Discover
	$btn_Step1_Mailboxes_Discover.Name = "btn_Step1_Mailboxes_Discover"
	$btn_Step1_Mailboxes_Discover.Size = $System_Drawing_Size_Step1_btn
	$btn_Step1_Mailboxes_Discover.TabIndex = 9
	$btn_Step1_Mailboxes_Discover.Text = "Discover"
	$btn_Step1_Mailboxes_Discover.UseVisualStyleBackColor = $True
	$btn_Step1_Mailboxes_Discover.add_Click($handler_btn_Step1_Mailboxes_Discover)
	$bx_Mailboxes_List.Controls.Add($btn_Step1_Mailboxes_Discover)
$btn_Step1_Mailboxes_Populate.Font = $font_Calibri_10pt_normal
	$btn_Step1_Mailboxes_Populate.Location = $System_Drawing_Point_Step1_Populate
	$btn_Step1_Mailboxes_Populate.Name = "btn_Step1_Mailboxes_Populate"
	$btn_Step1_Mailboxes_Populate.Size = $System_Drawing_Size_Step1_btn
	$btn_Step1_Mailboxes_Populate.TabIndex = 10
	$btn_Step1_Mailboxes_Populate.Text = "Load from File"
	$btn_Step1_Mailboxes_Populate.UseVisualStyleBackColor = $True
	$btn_Step1_Mailboxes_Populate.add_Click($handler_btn_Step1_Mailboxes_Populate)
$bx_Mailboxes_List.Controls.Add($btn_Step1_Mailboxes_Populate)
	$bx_Mailboxes_List.Dock = 5
	$bx_Mailboxes_List.Font = $font_Calibri_10pt_bold
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 3
	$System_Drawing_Point.Y = 3
	$bx_Mailboxes_List.Location = $System_Drawing_Point
	$bx_Mailboxes_List.Name = "bx_Mailboxes_List"
	$bx_Mailboxes_List.Size = $System_Drawing_Size_Step1_box
	$bx_Mailboxes_List.TabIndex = 7
	$bx_Mailboxes_List.TabStop = $False
	$tab_Step1_Mailboxes.Controls.Add($bx_Mailboxes_List)
$clb_Step1_Mailboxes_List.Font = $font_Calibri_10pt_normal
	$clb_Step1_Mailboxes_List.Location = $System_Drawing_Point_Step1_clb
	$clb_Step1_Mailboxes_List.Name = "clb_Step1_Mailboxes_List"
	$clb_Step1_Mailboxes_List.Size = $System_Drawing_Size_Step1_clb
	$clb_Step1_Mailboxes_List.TabIndex = 10
	$clb_Step1_Mailboxes_List.horizontalscrollbar = $true
	$clb_Step1_Mailboxes_List.CheckOnClick = $true
	$bx_Mailboxes_List.Controls.Add($clb_Step1_Mailboxes_List)
$txt_MailboxesTotal.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 50
	$System_Drawing_Point.Y = 410
	$txt_MailboxesTotal.Location = $System_Drawing_Point
	$txt_MailboxesTotal.Name = "txt_MailboxesTotal"
	$txt_MailboxesTotal.Size = $System_Drawing_Size_Step1_text_box
	$txt_MailboxesTotal.TabIndex = 11
	$txt_MailboxesTotal.Visible = $False
	$bx_Mailboxes_List.Controls.Add($txt_MailboxesTotal)
$btn_Step1_Mailboxes_CheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step1_Mailboxes_CheckAll.Location = $System_Drawing_Point_Step1_CheckAll
	$btn_Step1_Mailboxes_CheckAll.Name = "btn_Step1_Mailboxes_CheckAll"
	$btn_Step1_Mailboxes_CheckAll.Size = $System_Drawing_Size_Step1_btn
	$btn_Step1_Mailboxes_CheckAll.TabIndex = 9
	$btn_Step1_Mailboxes_CheckAll.Text = "Check all on this tab"
	$btn_Step1_Mailboxes_CheckAll.UseVisualStyleBackColor = $True
	$btn_Step1_Mailboxes_CheckAll.add_Click($handler_btn_Step1_Mailboxes_CheckAll)
	$bx_Mailboxes_List.Controls.Add($btn_Step1_Mailboxes_CheckAll)
$btn_Step1_Mailboxes_UncheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step1_Mailboxes_UncheckAll.Location = $System_Drawing_Point_Step1_UncheckAll
	$btn_Step1_Mailboxes_UncheckAll.Name = "btn_Step1_Mailboxes_UncheckAll"
	$btn_Step1_Mailboxes_UncheckAll.Size = $System_Drawing_Size_Step1_btn
	$btn_Step1_Mailboxes_UncheckAll.TabIndex = 10
	$btn_Step1_Mailboxes_UncheckAll.Text = "Uncheck all on this tab"
	$btn_Step1_Mailboxes_UncheckAll.UseVisualStyleBackColor = $True
	$btn_Step1_Mailboxes_UncheckAll.add_Click($handler_btn_Step1_Mailboxes_UncheckAll)
	$bx_Mailboxes_List.Controls.Add($btn_Step1_Mailboxes_UncheckAll)
#EndRegion Step1 Mailboxes tab

#Endregion "Step1 - Targets"

#Region "Step2"
$tab_Step2.BackColor = [System.Drawing.Color]::FromArgb(0,255,255,255)
	$tab_Step2.Font = $font_Calibri_8pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 4
	$System_Drawing_Point.Y = 36
	$tab_Step2.Location = $System_Drawing_Point
	$tab_Step2.Name = "tab_Step2"
	$tab_Step2.Padding = $System_Windows_Forms_Padding_Reusable
	$tab_Step2.TabIndex = 3
	$tab_Step2.Text = "  Templates  "
	$tab_Step2.Size = $System_Drawing_Size_tab_1
	$tab_Master.Controls.Add($tab_Step2)
$bx_Step2_Templates.Font = $font_Calibri_10pt_bold
	$System_Drawing_Point_bx_Step2 = New-Object System.Drawing.Point
	$System_Drawing_Point_bx_Step2.X = 27	# 96-69
	$System_Drawing_Point_bx_Step2.Y = 91
	$bx_Step2_Templates.Location = $System_Drawing_Point_bx_Step2
	$bx_Step2_Templates.Name = "bx_Step2_Templates"
	$System_Drawing_Size_bx_Step2 = New-Object System.Drawing.Size
	$System_Drawing_Size_bx_Step2.Height = 487 #482 to short
	$System_Drawing_Size_bx_Step2.Width = 536
	$bx_Step2_Templates.Size = $System_Drawing_Size_bx_Step2
	$bx_Step2_Templates.TabIndex = 0
	$bx_Step2_Templates.TabStop = $False
	$bx_Step2_Templates.Text = "Select a data collection template"
	$tab_Step2.Controls.Add($bx_Step2_Templates)
$rb_Step2_Template_1.Checked = $False
	$rb_Step2_Template_1.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 50
	$System_Drawing_Point.Y = 25
	$rb_Step2_Template_1.Location = $System_Drawing_Point
	$rb_Step2_Template_1.Name = "rb_Step2_Template_1"
	$rb_Step2_Template_1.Size = $System_Drawing_Size_Reusable_chk_long
	$rb_Step2_Template_1.TabIndex = 0
	$rb_Step2_Template_1.Text = "Recommended tests"
	$rb_Step2_Template_1.UseVisualStyleBackColor = $True
	$rb_Step2_Template_1.add_Click($handler_rb_Step2_Template_1)
	$bx_Step2_Templates.Controls.Add($rb_Step2_Template_1)
$rb_Step2_Template_2.Checked = $False
	$rb_Step2_Template_2.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 50
	$System_Drawing_Point.Y = 50
	$rb_Step2_Template_2.Location = $System_Drawing_Point
	$rb_Step2_Template_2.Name = "rb_Step2_Template_2"
	$rb_Step2_Template_2.Size = $System_Drawing_Size_Reusable_chk_long
	$rb_Step2_Template_2.TabIndex = 0
	$rb_Step2_Template_2.Text = "All tests"
	$rb_Step2_Template_2.UseVisualStyleBackColor = $True
	$rb_Step2_Template_2.add_Click($handler_rb_Step2_Template_2)
	$bx_Step2_Templates.Controls.Add($rb_Step2_Template_2)
$rb_Step2_Template_3.Checked = $False
	$rb_Step2_Template_3.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 50
	$System_Drawing_Point.Y = 75
	$rb_Step2_Template_3.Location = $System_Drawing_Point
	$rb_Step2_Template_3.Name = "rb_Step2_Template_3"
	$rb_Step2_Template_3.Size = $System_Drawing_Size_Reusable_chk_long
	$rb_Step2_Template_3.TabIndex = 0
	$rb_Step2_Template_3.Text = "Minimum tests for Environmental Document"
	$rb_Step2_Template_3.UseVisualStyleBackColor = $True
	$rb_Step2_Template_3.add_Click($handler_rb_Step2_Template_3)
	$bx_Step2_Templates.Controls.Add($rb_Step2_Template_3)
$rb_Step2_Template_4.Checked = $False
	$rb_Step2_Template_4.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 50
	$System_Drawing_Point.Y = 100
	$rb_Step2_Template_4.Location = $System_Drawing_Point
	$rb_Step2_Template_4.Name = "rb_Step2_Template_4"
	$rb_Step2_Template_4.Size = $System_Drawing_Size_Reusable_chk_long
	$rb_Step2_Template_4.TabIndex = 0
	$rb_Step2_Template_4.Text = "Custom Template 1"
	$rb_Step2_Template_4.UseVisualStyleBackColor = $True
	$rb_Step2_Template_4.add_Click($handler_rb_Step2_Template_4)
	$bx_Step2_Templates.Controls.Add($rb_Step2_Template_4)
$Status_Step2.Font = $font_Calibri_10pt_normal
	$Status_Step2.Location = $System_Drawing_Point_Status
	$Status_Step2.Name = "Status_Step2"
	$Status_Step2.Size = $System_Drawing_Size_Status
	$Status_Step2.TabIndex = 12
	$Status_Step2.Text = "Step 2 Status"
	$tab_Step2.Controls.Add($Status_Step2)
#Endregion "Step2"

#Region "Step3 - Tests"
#Region Step3 Main
# Reusable boxes in Step3 Tabs
	$System_Drawing_Size_Step3_box = New-Object System.Drawing.Size
	$System_Drawing_Size_Step3_box.Height = 400
	$System_Drawing_Size_Step3_box.Width = 536
# Reusable check buttons in Step3 tabs
	$System_Drawing_Size_Step3_check_btn = New-Object System.Drawing.Size
	$System_Drawing_Size_Step3_check_btn.Height = 25
	$System_Drawing_Size_Step3_check_btn.Width = 150
# Reusable check/uncheck buttons in Step3 tabs
	$System_Drawing_Point_Step3_Check = New-Object System.Drawing.Point
	$System_Drawing_Point_Step3_Check.X = 50
	$System_Drawing_Point_Step3_Check.Y = 400
	$System_Drawing_Point_Step3_Uncheck = New-Object System.Drawing.Point
	$System_Drawing_Point_Step3_Uncheck.X = 300
	$System_Drawing_Point_Step3_Uncheck.Y = 400
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 4
	$System_Drawing_Point.Y = 36
$tab_Step3.Location = $System_Drawing_Point
	$tab_Step3.Name = "tab_Step3"
	$tab_Step3.Padding = $System_Windows_Forms_Padding_Reusable
	$tab_Step3.TabIndex = 2
	$tab_Step3.Text = "   Tests   "
	$tab_Step3.Size = $System_Drawing_Size_tab_1
	$tab_Master.Controls.Add($tab_Step3)
$tab_Step3_Master.Font = $font_Calibri_12pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 20
	$System_Drawing_Point.Y = 60
	$tab_Step3_Master.Location = $System_Drawing_Point
	$tab_Step3_Master.Name = "tab_Step3_Master"
	$tab_Step3_Master.SelectedIndex = 0
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 525
	$System_Drawing_Size.Width = 550
	$tab_Step3_Master.Size = $System_Drawing_Size
	$tab_Step3_Master.TabIndex = 11
	$tab_Step3.Controls.Add($tab_Step3_Master)
$btn_Step3_Execute.Font = $font_Calibri_14pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 20
	$System_Drawing_Point.Y = 15
	$btn_Step3_Execute.Location = $System_Drawing_Point
	$btn_Step3_Execute.Name = "btn_Step3_Execute"
	$btn_Step3_Execute.Size = $System_Drawing_Size_buttons
	$btn_Step3_Execute.TabIndex = 4
	$btn_Step3_Execute.Text = "Execute"
	$btn_Step3_Execute.UseVisualStyleBackColor = $True
	$btn_Step3_Execute.add_Click($handler_btn_Step3_Execute_Click)
	$tab_Step3.Controls.Add($btn_Step3_Execute)
$lbl_Step3_Execute.BackColor = [System.Drawing.Color]::FromArgb(0,255,255,255)
	$lbl_Step3_Execute.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 138
	$System_Drawing_Point.Y = 15
	$lbl_Step3_Execute.Location = $System_Drawing_Point
	$lbl_Step3_Execute.Name = "lbl_Step3"
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 38
	$System_Drawing_Size.Width = 510
	$lbl_Step3_Execute.Size = $System_Drawing_Size
	$lbl_Step3_Execute.TabIndex = 5
	$lbl_Step3_Execute.Text = "Select the functions below and click on the Execute button."
	$lbl_Step3_Execute.TextAlign = 16
	$tab_Step3.Controls.Add($lbl_Step3_Execute)
$status_Step3.Font = $font_Calibri_10pt_normal
	$status_Step3.Location = $System_Drawing_Point_Status
	$status_Step3.Name = "status_Step3"
	$status_Step3.Size = $System_Drawing_Size_Status
	$status_Step3.TabIndex = 10
	$status_Step3.Text = "Step 3 Status"
	$tab_Step3.Controls.Add($status_Step3)
#EndRegion Step3 Main

#Region Step3 ExOrg - Tier 2
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 4
	$System_Drawing_Point.Y = 33
$tab_Step3_ExOrg.Location = $System_Drawing_Point
	$tab_Step3_ExOrg.Name = "tab_Step3_ExOrg"
	$tab_Step3_ExOrg.Padding = $System_Windows_Forms_Padding_Reusable
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 488
	$System_Drawing_Size.Width = 542
	$tab_Step3_ExOrg.Size = $System_Drawing_Size
	$tab_Step3_ExOrg.TabIndex = 0
	$tab_Step3_ExOrg.Text = "Exchange Functions"
	$tab_Step3_ExOrg.UseVisualStyleBackColor = $True
	$tab_Step3_Master.Controls.Add($tab_Step3_ExOrg)

# ExOrg Tab Control
$tab_Step3_ExOrg_Tier2.Appearance = 2
	$tab_Step3_ExOrg_Tier2.Dock = 5
	$tab_Step3_ExOrg_Tier2.Font = $font_Calibri_10pt_normal
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 32
	$System_Drawing_Size.Width = 100
	$tab_Step3_ExOrg_Tier2.ItemSize = $System_Drawing_Size
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 0
	$System_Drawing_Point.Y = 0
	$tab_Step3_ExOrg_Tier2.Location = $System_Drawing_Point
	$tab_Step3_ExOrg_Tier2.Name = "tab_Step3_ExOrg_Tier2"
	$tab_Step3_ExOrg_Tier2.SelectedIndex = 0
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 718
	$System_Drawing_Size.Width = 665
	$tab_Step3_ExOrg_Tier2.Size = $System_Drawing_Size
	$tab_Step3_ExOrg_Tier2.TabIndex = 12
	$tab_Step3_ExOrg.Controls.Add($tab_Step3_ExOrg_Tier2)
#EndRegion Step3 ExOrg - Tier 2

#Region Step3 Client Access tab
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 4
	$System_Drawing_Point.Y = 33
$tab_Step3_ClientAccess.Location = $System_Drawing_Point
	$tab_Step3_ClientAccess.Name = "tab_Step3_ClientAccess"
	$tab_Step3_ClientAccess.Padding = $System_Windows_Forms_Padding_Reusable
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 488
	$System_Drawing_Size.Width = 542
	$tab_Step3_ClientAccess.Size = $System_Drawing_Size
	$tab_Step3_ClientAccess.TabIndex = 3
	$tab_Step3_ClientAccess.Text = "Client Access"
	$tab_Step3_ClientAccess.UseVisualStyleBackColor = $True
	$tab_Step3_ExOrg_Tier2.Controls.Add($tab_Step3_ClientAccess)
$bx_ClientAccess_Functions.Dock = 5
	$bx_ClientAccess_Functions.Font = $font_Calibri_10pt_bold
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 3
	$System_Drawing_Point.Y = 3
	$bx_ClientAccess_Functions.Location = $System_Drawing_Point
	$bx_ClientAccess_Functions.Name = "bx_ClientAccess_Functions"
	$bx_ClientAccess_Functions.Size = $System_Drawing_Size_Step3_box
	$bx_ClientAccess_Functions.TabIndex = 9
	$bx_ClientAccess_Functions.TabStop = $False
	$tab_Step3_ClientAccess.Controls.Add($bx_ClientAccess_Functions)
$btn_Step3_ClientAccess_CheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_ClientAccess_CheckAll.Location = $System_Drawing_Point_Step3_Check
	$btn_Step3_ClientAccess_CheckAll.Name = "btn_Step3_ClientAccess_CheckAll"
	$btn_Step3_ClientAccess_CheckAll.Size = $System_Drawing_Size_Step3_check_btn
	$btn_Step3_ClientAccess_CheckAll.TabIndex = 28
	$btn_Step3_ClientAccess_CheckAll.Text = "Check all on this tab"
	$btn_Step3_ClientAccess_CheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_ClientAccess_CheckAll.add_Click($handler_btn_Step3_ClientAccess_CheckAll_Click)
	$bx_ClientAccess_Functions.Controls.Add($btn_Step3_ClientAccess_CheckAll)
$btn_Step3_ClientAccess_UncheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_ClientAccess_UncheckAll.Location = $System_Drawing_Point_Step3_Uncheck
	$btn_Step3_ClientAccess_UncheckAll.Name = "btn_Step3_ClientAccess_UncheckAll"
	$btn_Step3_ClientAccess_UncheckAll.Size = $System_Drawing_Size_Step3_check_btn
	$btn_Step3_ClientAccess_UncheckAll.TabIndex = 29
	$btn_Step3_ClientAccess_UncheckAll.Text = "Uncheck all on this tab"
	$btn_Step3_ClientAccess_UncheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_ClientAccess_UncheckAll.add_Click($handler_btn_Step3_ClientAccess_UncheckAll_Click)
	$bx_ClientAccess_Functions.Controls.Add($btn_Step3_ClientAccess_UncheckAll)
$Col_1_loc = 35
$Col_2_loc = 290
$Row_1_loc = 25
$Row_2_loc = 25
$chk_Org_Get_MobileDevice.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_MobileDevice.Location = $System_Drawing_Point
	$chk_Org_Get_MobileDevice.Name = "chk_Org_Get_MobileDevice"
	$chk_Org_Get_MobileDevice.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_MobileDevice.TabIndex = 0
	$chk_Org_Get_MobileDevice.Text = "Get-MobileDevice"
	$chk_Org_Get_MobileDevice.UseVisualStyleBackColor = $True
	$bx_ClientAccess_Functions.Controls.Add($chk_Org_Get_MobileDevice)
$chk_Org_Get_MobileDevicePolicy.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_MobileDevicePolicy.Location = $System_Drawing_Point
	$chk_Org_Get_MobileDevicePolicy.Name = "chk_Org_Get_MobileDevicePolicy"
	$chk_Org_Get_MobileDevicePolicy.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_MobileDevicePolicy.TabIndex = 2
	$chk_Org_Get_MobileDevicePolicy.Text = "Get-MobileDeviceMailboxPolicy"
	$chk_Org_Get_MobileDevicePolicy.UseVisualStyleBackColor = $True
	$bx_ClientAccess_Functions.Controls.Add($chk_Org_Get_MobileDevicePolicy)
$chk_Org_Get_AvailabilityAddressSpace.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_AvailabilityAddressSpace.Location = $System_Drawing_Point
	$chk_Org_Get_AvailabilityAddressSpace.Name = "chk_Org_Get_AvailabilityAddressSpace"
	$chk_Org_Get_AvailabilityAddressSpace.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_AvailabilityAddressSpace.TabIndex = 10
	$chk_Org_Get_AvailabilityAddressSpace.Text = "Get-AvailabilityAddressSpace"
	$chk_Org_Get_AvailabilityAddressSpace.UseVisualStyleBackColor = $True
	$bx_ClientAccess_Functions.Controls.Add($chk_Org_Get_AvailabilityAddressSpace)
$chk_Org_Get_OwaMailboxPolicy.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_OwaMailboxPolicy.Location = $System_Drawing_Point
	$chk_Org_Get_OwaMailboxPolicy.Name = "chk_Org_Get_OwaMailboxPolicy"
	$chk_Org_Get_OwaMailboxPolicy.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_OwaMailboxPolicy.TabIndex = 3
	$chk_Org_Get_OwaMailboxPolicy.Text = "Get-OwaMailboxPolicy"
	$chk_Org_Get_OwaMailboxPolicy.UseVisualStyleBackColor = $True
	$bx_ClientAccess_Functions.Controls.Add($chk_Org_Get_OwaMailboxPolicy)
#EndRegion Step3 Client Access tab

#Region Step3 Global tab
$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 4
	$System_Drawing_Point.Y = 33
$tab_Step3_Global.Location = $System_Drawing_Point
	$tab_Step3_Global.Name = "tab_Step3_Global"
	$tab_Step3_Global.Padding = $System_Windows_Forms_Padding_Reusable
$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 488
	$System_Drawing_Size.Width = 542
$tab_Step3_Global.Size = $System_Drawing_Size
	$tab_Step3_Global.TabIndex = 3
	$tab_Step3_Global.Text = "Global and Database"
	$tab_Step3_Global.UseVisualStyleBackColor = $True
	$tab_Step3_ExOrg_Tier2.Controls.Add($tab_Step3_Global)
$bx_Global_Functions.Dock = 5
	$bx_Global_Functions.Font = $font_Calibri_10pt_bold
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 3
	$System_Drawing_Point.Y = 3
	$bx_Global_Functions.Location = $System_Drawing_Point
	$bx_Global_Functions.Name = "bx_Global_Functions"
	$bx_Global_Functions.Size = $System_Drawing_Size_Step3_box
	$bx_Global_Functions.TabIndex = 9
	$bx_Global_Functions.TabStop = $False
	$tab_Step3_Global.Controls.Add($bx_Global_Functions)
$btn_Step3_Global_CheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_Global_CheckAll.Location = $System_Drawing_Point_Step3_Check
	$btn_Step3_Global_CheckAll.Name = "btn_Step3_Global_CheckAll"
	$btn_Step3_Global_CheckAll.Size = $System_Drawing_Size_Step3_check_btn
	$btn_Step3_Global_CheckAll.TabIndex = 28
	$btn_Step3_Global_CheckAll.Text = "Check all on this tab"
	$btn_Step3_Global_CheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_Global_CheckAll.add_Click($handler_btn_Step3_Global_CheckAll_Click)
	$bx_Global_Functions.Controls.Add($btn_Step3_Global_CheckAll)
$btn_Step3_Global_UncheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_Global_UncheckAll.Location = $System_Drawing_Point_Step3_Uncheck
	$btn_Step3_Global_UncheckAll.Name = "btn_Step3_Global_UncheckAll"
	$btn_Step3_Global_UncheckAll.Size = $System_Drawing_Size_Step3_check_btn
	$btn_Step3_Global_UncheckAll.TabIndex = 29
	$btn_Step3_Global_UncheckAll.Text = "Uncheck all on this tab"
	$btn_Step3_Global_UncheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_Global_UncheckAll.add_Click($handler_btn_Step3_Global_UncheckAll_Click)
	$bx_Global_Functions.Controls.Add($btn_Step3_Global_UncheckAll)
$Col_1_loc = 35
$Col_2_loc = 290
$Row_1_loc = 25
$Row_2_loc = 25
$chk_Org_Get_AddressBookPolicy.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_AddressBookPolicy.Location = $System_Drawing_Point
	$chk_Org_Get_AddressBookPolicy.Name = "chk_Org_Get_AddressBookPolicy"
	$chk_Org_Get_AddressBookPolicy.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_AddressBookPolicy.TabIndex = 13
	$chk_Org_Get_AddressBookPolicy.Text = "Get-AddressBookPolicy"
	$chk_Org_Get_AddressBookPolicy.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_AddressBookPolicy)
$chk_Org_Get_AddressList.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_AddressList.Location = $System_Drawing_Point
	$chk_Org_Get_AddressList.Name = "chk_Org_Get_AddressList"
	$chk_Org_Get_AddressList.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_AddressList.TabIndex = 13
	$chk_Org_Get_AddressList.Text = "Get-AddressList"
	$chk_Org_Get_AddressList.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_AddressList)
$chk_Org_Get_EmailAddressPolicy.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_EmailAddressPolicy.Location = $System_Drawing_Point
	$chk_Org_Get_EmailAddressPolicy.Name = "chk_Org_Get_EmailAddressPolicy"
	$chk_Org_Get_EmailAddressPolicy.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_EmailAddressPolicy.TabIndex = 18
	$chk_Org_Get_EmailAddressPolicy.Text = "Get-EmailAddressPolicy"
	$chk_Org_Get_EmailAddressPolicy.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_EmailAddressPolicy)
$chk_Org_Get_GlobalAddressList.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_GlobalAddressList.Location = $System_Drawing_Point
	$chk_Org_Get_GlobalAddressList.Name = "chk_Org_Get_GlobalAddressList"
	$chk_Org_Get_GlobalAddressList.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_GlobalAddressList.TabIndex = 1
	$chk_Org_Get_GlobalAddressList.Text = "Get-GlobalAddressList"
	$chk_Org_Get_GlobalAddressList.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_GlobalAddressList)
$chk_Org_Get_OfflineAddressBook.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_OfflineAddressBook.Location = $System_Drawing_Point
	$chk_Org_Get_OfflineAddressBook.Name = "chk_Org_Get_OfflineAddressBook"
	$chk_Org_Get_OfflineAddressBook.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_OfflineAddressBook.TabIndex = 1
	$chk_Org_Get_OfflineAddressBook.Text = "Get-OfflineAddressBook"
	$chk_Org_Get_OfflineAddressBook.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_OfflineAddressBook)
$chk_Org_Get_OrgConfig.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_OrgConfig.Location = $System_Drawing_Point
	$chk_Org_Get_OrgConfig.Name = "chk_Org_Get_OrgConfig"
	$chk_Org_Get_OrgConfig.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_OrgConfig.TabIndex = 1
	$chk_Org_Get_OrgConfig.Text = "Get-OrganizationConfig"
	$chk_Org_Get_OrgConfig.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_OrgConfig)
$chk_Org_Get_Rbac.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_2_loc
	$System_Drawing_Point.Y = $Row_2_loc
	$Row_2_loc += 25
	$chk_Org_Get_Rbac.Location = $System_Drawing_Point
	$chk_Org_Get_Rbac.Name = "chk_Org_Get_Rbac"
	$chk_Org_Get_Rbac.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_Rbac.TabIndex = 6
	$chk_Org_Get_Rbac.Text = "Get-Rbac"
	$chk_Org_Get_Rbac.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_Rbac)
$chk_Org_Get_RetentionPolicy.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_2_loc
	$System_Drawing_Point.Y = $Row_2_loc
	$Row_2_loc += 25
	$chk_Org_Get_RetentionPolicy.Location = $System_Drawing_Point
	$chk_Org_Get_RetentionPolicy.Name = "chk_Org_Get_RetentionPolicy"
	$chk_Org_Get_RetentionPolicy.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_RetentionPolicy.TabIndex = 6
	$chk_Org_Get_RetentionPolicy.Text = "Get-RetentionPolicy"
	$chk_Org_Get_RetentionPolicy.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_RetentionPolicy)
$chk_Org_Get_RetentionPolicyTag.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_2_loc
	$System_Drawing_Point.Y = $Row_2_loc
	$Row_2_loc += 25
	$chk_Org_Get_RetentionPolicyTag.Location = $System_Drawing_Point
	$chk_Org_Get_RetentionPolicyTag.Name = "chk_Org_Get_RetentionPolicyTag"
	$chk_Org_Get_RetentionPolicyTag.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_RetentionPolicyTag.TabIndex = 6
	$chk_Org_Get_RetentionPolicyTag.Text = "Get-RetentionPolicyTag"
	$chk_Org_Get_RetentionPolicyTag.UseVisualStyleBackColor = $True
	$bx_Global_Functions.Controls.Add($chk_Org_Get_RetentionPolicyTag)
#EndRegion Step3 Global tab

#Region Step3 Recipient tab
$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 4
	$System_Drawing_Point.Y = 33
$tab_Step3_Recipient.Location = $System_Drawing_Point
	$tab_Step3_Recipient.Name = "tab_Step3_Recipient"
	$tab_Step3_Recipient.Padding = $System_Windows_Forms_Padding_Reusable
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 488
	$System_Drawing_Size.Width = 542
	$tab_Step3_Recipient.Size = $System_Drawing_Size
	$tab_Step3_Recipient.TabIndex = 3
	$tab_Step3_Recipient.Text = "Recipient"
	$tab_Step3_Recipient.UseVisualStyleBackColor = $True
	$tab_Step3_ExOrg_Tier2.Controls.Add($tab_Step3_Recipient)
$bx_Recipient_Functions.Dock = 5
	$bx_Recipient_Functions.Font = $font_Calibri_10pt_bold
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 3
	$System_Drawing_Point.Y = 3
	$bx_Recipient_Functions.Location = $System_Drawing_Point
	$bx_Recipient_Functions.Name = "bx_Recipient_Functions"
	$bx_Recipient_Functions.Size = $System_Drawing_Size_Step3_box
	$bx_Recipient_Functions.TabIndex = 9
	$bx_Recipient_Functions.TabStop = $False
	$tab_Step3_Recipient.Controls.Add($bx_Recipient_Functions)
$btn_Step3_Recipient_CheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_Recipient_CheckAll.Location = $System_Drawing_Point_Step3_Check
	$btn_Step3_Recipient_CheckAll.Name = "btn_Step3_Recipient_CheckAll"
	$btn_Step3_Recipient_CheckAll.Size = $System_Drawing_Size_Step3_check_btn
	$btn_Step3_Recipient_CheckAll.TabIndex = 28
	$btn_Step3_Recipient_CheckAll.Text = "Check all on this tab"
	$btn_Step3_Recipient_CheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_Recipient_CheckAll.add_Click($handler_btn_Step3_Recipient_CheckAll_Click)
	$bx_Recipient_Functions.Controls.Add($btn_Step3_Recipient_CheckAll)
$btn_Step3_Recipient_UncheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_Recipient_UncheckAll.Location = $System_Drawing_Point_Step3_Uncheck
	$btn_Step3_Recipient_UncheckAll.Name = "btn_Step3_Recipient_UncheckAll"
	$btn_Step3_Recipient_UncheckAll.Size = $System_Drawing_Size_Step3_check_btn
	$btn_Step3_Recipient_UncheckAll.TabIndex = 29
	$btn_Step3_Recipient_UncheckAll.Text = "Uncheck all on this tab"
	$btn_Step3_Recipient_UncheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_Recipient_UncheckAll.add_Click($handler_btn_Step3_Recipient_UncheckAll_Click)
	$bx_Recipient_Functions.Controls.Add($btn_Step3_Recipient_UncheckAll)
$Col_1_loc = 35
$Col_2_loc = 290
$Row_1_loc = 25
$Row_2_loc = 25
$chk_Org_Get_CalendarProcessing.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_CalendarProcessing.Location = $System_Drawing_Point
	$chk_Org_Get_CalendarProcessing.Name = "chk_Org_Get_CalendarProcessing"
	$chk_Org_Get_CalendarProcessing.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_CalendarProcessing.TabIndex = 8
	$chk_Org_Get_CalendarProcessing.Text = "Get-CalendarProcessing"
	$chk_Org_Get_CalendarProcessing.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_CalendarProcessing)
$chk_Org_Get_CASMailbox.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_CASMailbox.Location = $System_Drawing_Point
	$chk_Org_Get_CASMailbox.Name = "chk_Org_Get_CASMailbox"
	$chk_Org_Get_CASMailbox.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_CASMailbox.TabIndex = 9
	$chk_Org_Get_CASMailbox.Text = "Get-CASMailbox"
	$chk_Org_Get_CASMailbox.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_CASMailbox)
$chk_Org_Get_DistributionGroup.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_DistributionGroup.Location = $System_Drawing_Point
	$chk_Org_Get_DistributionGroup.Name = "chk_Org_Get_DistributionGroup"
	$chk_Org_Get_DistributionGroup.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_DistributionGroup.TabIndex = 15
	$chk_Org_Get_DistributionGroup.Text = "Get-DistributionGroup"
	$chk_Org_Get_DistributionGroup.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_DistributionGroup)
$chk_Org_Get_DynamicDistributionGroup.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_DynamicDistributionGroup.Location = $System_Drawing_Point
	$chk_Org_Get_DynamicDistributionGroup.Name = "chk_Org_Get_DynamicDistributionGroup"
	$chk_Org_Get_DynamicDistributionGroup.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_DynamicDistributionGroup.TabIndex = 16
	$chk_Org_Get_DynamicDistributionGroup.Text = "Get-DynamicDistributionGroup"
	$chk_Org_Get_DynamicDistributionGroup.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_DynamicDistributionGroup)
$chk_Org_Get_Mailbox.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_Mailbox.Location = $System_Drawing_Point
	$chk_Org_Get_Mailbox.Name = "chk_Org_Get_Mailbox"
	$chk_Org_Get_Mailbox.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_Mailbox.TabIndex = 21
	$chk_Org_Get_Mailbox.Text = "Get-Mailbox"
	$chk_Org_Get_Mailbox.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_Mailbox)
$chk_Org_Get_MailboxFolderStatistics.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_MailboxFolderStatistics.Location = $System_Drawing_Point
	$chk_Org_Get_MailboxFolderStatistics.Name = "chk_Org_Get_MailboxFolderStatistics"
	$chk_Org_Get_MailboxFolderStatistics.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_MailboxFolderStatistics.TabIndex = 24
	$chk_Org_Get_MailboxFolderStatistics.Text = "Get-MailboxFolderStatistics"
	$chk_Org_Get_MailboxFolderStatistics.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_MailboxFolderStatistics)
$chk_Org_Get_MailboxPermission.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_MailboxPermission.Location = $System_Drawing_Point
	$chk_Org_Get_MailboxPermission.Name = "chk_Org_Get_MailboxPermission"
	$chk_Org_Get_MailboxPermission.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_MailboxPermission.TabIndex = 25
	$chk_Org_Get_MailboxPermission.Text = "Get-MailboxPermission"
	$chk_Org_Get_MailboxPermission.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_MailboxPermission)
$chk_Org_Get_MailboxStatistics.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_MailboxStatistics.Location = $System_Drawing_Point
	$chk_Org_Get_MailboxStatistics.Name = "chk_Org_Get_MailboxStatistics"
	$chk_Org_Get_MailboxStatistics.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_MailboxStatistics.TabIndex = 27
	$chk_Org_Get_MailboxStatistics.Text = "Get-MailboxStatistics"
	$chk_Org_Get_MailboxStatistics.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_MailboxStatistics)
$chk_Org_Get_PublicFolder.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_PublicFolder.Location = $System_Drawing_Point
	$chk_Org_Get_PublicFolder.Name = "chk_Org_Get_PublicFolder"
	$chk_Org_Get_PublicFolder.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_PublicFolder.TabIndex = 5
	$chk_Org_Get_PublicFolder.Text = "Get-PublicFolder"
	$chk_Org_Get_PublicFolder.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_PublicFolder)
$chk_Org_Get_PublicFolderStatistics.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_PublicFolderStatistics.Location = $System_Drawing_Point
	$chk_Org_Get_PublicFolderStatistics.Name = "chk_Org_Get_PublicFolderStatistics"
	$chk_Org_Get_PublicFolderStatistics.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_PublicFolderStatistics.TabIndex = 7
	$chk_Org_Get_PublicFolderStatistics.Text = "Get-PublicFolderStatistics"
	$chk_Org_Get_PublicFolderStatistics.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Get_PublicFolderStatistics)
$chk_Org_Quota.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Quota.Location = $System_Drawing_Point
	$chk_Org_Quota.Name = "chk_Org_Quota"
	$chk_Org_Quota.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Quota.TabIndex = 15
	$chk_Org_Quota.Text = "Quota"
	$chk_Org_Quota.UseVisualStyleBackColor = $True
	$bx_Recipient_Functions.Controls.Add($chk_Org_Quota)
#EndRegion Step3 Recipient tab

#Region Step3 Transport tab
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 4
	$System_Drawing_Point.Y = 33
$tab_Step3_Transport.Location = $System_Drawing_Point
	$tab_Step3_Transport.Name = "tab_Step3_Transport"
	$tab_Step3_Transport.Padding = $System_Windows_Forms_Padding_Reusable
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 488
	$System_Drawing_Size.Width = 542
	$tab_Step3_Transport.Size = $System_Drawing_Size
	$tab_Step3_Transport.TabIndex = 3
	$tab_Step3_Transport.Text = "Transport"
	$tab_Step3_Transport.UseVisualStyleBackColor = $True
	$tab_Step3_ExOrg_Tier2.Controls.Add($tab_Step3_Transport)
$bx_Transport_Functions.Dock = 5
	$bx_Transport_Functions.Font = $font_Calibri_10pt_bold
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 3
	$System_Drawing_Point.Y = 3
	$bx_Transport_Functions.Location = $System_Drawing_Point
	$bx_Transport_Functions.Name = "bx_Transport_Functions"
	$bx_Transport_Functions.Size = $System_Drawing_Size_Step3_box
	$bx_Transport_Functions.TabIndex = 9
	$bx_Transport_Functions.TabStop = $False
	$tab_Step3_Transport.Controls.Add($bx_Transport_Functions)
$btn_Step3_Transport_CheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_Transport_CheckAll.Location = $System_Drawing_Point_Step3_Check
	$btn_Step3_Transport_CheckAll.Name = "btn_Step3_Transport_CheckAll"
	$btn_Step3_Transport_CheckAll.Size = $System_Drawing_Size_Step3_check_btn
	$btn_Step3_Transport_CheckAll.TabIndex = 28
	$btn_Step3_Transport_CheckAll.Text = "Check all on this tab"
	$btn_Step3_Transport_CheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_Transport_CheckAll.add_Click($handler_btn_Step3_Transport_CheckAll_Click)
	$bx_Transport_Functions.Controls.Add($btn_Step3_Transport_CheckAll)
$btn_Step3_Transport_UncheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_Transport_UncheckAll.Location = $System_Drawing_Point_Step3_Uncheck
	$btn_Step3_Transport_UncheckAll.Name = "btn_Step3_Transport_UncheckAll"
	$btn_Step3_Transport_UncheckAll.Size = $System_Drawing_Size_Step3_check_btn
	$btn_Step3_Transport_UncheckAll.TabIndex = 29
	$btn_Step3_Transport_UncheckAll.Text = "Uncheck all on this tab"
	$btn_Step3_Transport_UncheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_Transport_UncheckAll.add_Click($handler_btn_Step3_Transport_UncheckAll_Click)
	$bx_Transport_Functions.Controls.Add($btn_Step3_Transport_UncheckAll)
$Col_1_loc = 35
$Col_2_loc = 290
$Row_1_loc = 25
$Row_2_loc = 25
$chk_Org_Get_AcceptedDomain.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_AcceptedDomain.Location = $System_Drawing_Point
	$chk_Org_Get_AcceptedDomain.Name = "chk_Org_Get_AcceptedDomain"
	$chk_Org_Get_AcceptedDomain.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_AcceptedDomain.TabIndex = 0
	$chk_Org_Get_AcceptedDomain.Text = "Get-AcceptedDomain"
	$chk_Org_Get_AcceptedDomain.UseVisualStyleBackColor = $True
	$bx_Transport_Functions.Controls.Add($chk_Org_Get_AcceptedDomain)
$chk_Org_Get_InboundConnector.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_InboundConnector.Location = $System_Drawing_Point
	$chk_Org_Get_InboundConnector.Name = "chk_Org_Get_InboundConnector"
	$chk_Org_Get_InboundConnector.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_InboundConnector.TabIndex = 8
	$chk_Org_Get_InboundConnector.Text = "Get-InboundConnector"
	$chk_Org_Get_InboundConnector.UseVisualStyleBackColor = $True
	$bx_Transport_Functions.Controls.Add($chk_Org_Get_InboundConnector)
$chk_Org_Get_RemoteDomain.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_RemoteDomain.Location = $System_Drawing_Point
	$chk_Org_Get_RemoteDomain.Name = "chk_Org_Get_RemoteDomain"
	$chk_Org_Get_RemoteDomain.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_RemoteDomain.TabIndex = 8
	$chk_Org_Get_RemoteDomain.Text = "Get-RemoteDomain"
	$chk_Org_Get_RemoteDomain.UseVisualStyleBackColor = $True
	$bx_Transport_Functions.Controls.Add($chk_Org_Get_RemoteDomain)
$chk_Org_Get_OutboundConnector.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_OutboundConnector.Location = $System_Drawing_Point
	$chk_Org_Get_OutboundConnector.Name = "chk_Org_Get_OutboundConnector"
	$chk_Org_Get_OutboundConnector.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_OutboundConnector.TabIndex = 11
	$chk_Org_Get_OutboundConnector.Text = "Get-OutboundConnector"
	$chk_Org_Get_OutboundConnector.UseVisualStyleBackColor = $True
	$bx_Transport_Functions.Controls.Add($chk_Org_Get_OutboundConnector)
$chk_Org_Get_TransportConfig.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_TransportConfig.Location = $System_Drawing_Point
	$chk_Org_Get_TransportConfig.Name = "chk_Org_Get_TransportConfig"
	$chk_Org_Get_TransportConfig.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_TransportConfig.TabIndex = 12
	$chk_Org_Get_TransportConfig.Text = "Get-TransportConfig"
	$chk_Org_Get_TransportConfig.UseVisualStyleBackColor = $True
	$bx_Transport_Functions.Controls.Add($chk_Org_Get_TransportConfig)
$chk_Org_Get_TransportRule.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_TransportRule.Location = $System_Drawing_Point
	$chk_Org_Get_TransportRule.Name = "chk_Org_Get_TransportRule"
	$chk_Org_Get_TransportRule.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_TransportRule.TabIndex = 12
	$chk_Org_Get_TransportRule.Text = "Get-TransportRule"
	$chk_Org_Get_TransportRule.UseVisualStyleBackColor = $True
	$bx_Transport_Functions.Controls.Add($chk_Org_Get_TransportRule)
#EndRegion Step3 Transport tab

#Region Step3 UM tab
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 4
	$System_Drawing_Point.Y = 33
$tab_Step3_UM.Location = $System_Drawing_Point
	$tab_Step3_UM.Name = "tab_Step3_Misc"
	$tab_Step3_UM.Padding = $System_Windows_Forms_Padding_Reusable
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 488
	$System_Drawing_Size.Width = 300 #542
	$tab_Step3_UM.Size = $System_Drawing_Size
	$tab_Step3_UM.TabIndex = 4
	$tab_Step3_UM.Text = "    UM"
	$tab_Step3_UM.UseVisualStyleBackColor = $True
	$tab_Step3_ExOrg_Tier2.Controls.Add($tab_Step3_UM)
$bx_UM_Functions.Dock = 5
	$bx_UM_Functions.Font = $font_Calibri_10pt_bold
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 3
	$System_Drawing_Point.Y = 3
	$bx_UM_Functions.Location = $System_Drawing_Point
	$bx_UM_Functions.Name = "bx_Misc_Functions"
	$bx_UM_Functions.Size = $System_Drawing_Size_Step3_box
	$bx_UM_Functions.TabIndex = 9
	$bx_UM_Functions.TabStop = $False
	$tab_Step3_UM.Controls.Add($bx_UM_Functions)
$btn_Step3_UM_CheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_UM_CheckAll.Location = $System_Drawing_Point_Step3_Check
	$btn_Step3_UM_CheckAll.Name = "btn_Step3_Misc_CheckAll"
	$btn_Step3_UM_CheckAll.Size = $System_Drawing_Size_Step3_check_btn
	$btn_Step3_UM_CheckAll.TabIndex = 28
	$btn_Step3_UM_CheckAll.Text = "Check all on this tab"
	$btn_Step3_UM_CheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_UM_CheckAll.add_Click($handler_btn_Step3_UM_CheckAll_Click)
	$bx_UM_Functions.Controls.Add($btn_Step3_UM_CheckAll)
$btn_Step3_UM_UncheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_UM_UncheckAll.Location = $System_Drawing_Point_Step3_Uncheck
	$btn_Step3_UM_UncheckAll.Name = "btn_Step3_Misc_UncheckAll"
	$btn_Step3_UM_UncheckAll.Size = $System_Drawing_Size_Step3_check_btn
	$btn_Step3_UM_UncheckAll.TabIndex = 29
	$btn_Step3_UM_UncheckAll.Text = "Uncheck all on this tab"
	$btn_Step3_UM_UncheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_UM_UncheckAll.add_Click($handler_btn_Step3_UM_UncheckAll_Click)
	$bx_UM_Functions.Controls.Add($btn_Step3_UM_UncheckAll)
$Col_1_loc = 35
$Col_2_loc = 290
$Row_1_loc = 25
$Row_2_loc = 25
$chk_Org_Get_UmAutoAttendant.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_UmAutoAttendant.Location = $System_Drawing_Point
	$chk_Org_Get_UmAutoAttendant.Name = "chk_Org_Get_UmAutoAttendant"
	$chk_Org_Get_UmAutoAttendant.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_UmAutoAttendant.TabIndex = 0
	$chk_Org_Get_UmAutoAttendant.Text = "Get-UmAutoAttendant"
	$chk_Org_Get_UmAutoAttendant.UseVisualStyleBackColor = $True
	$bx_UM_Functions.Controls.Add($chk_Org_Get_UmAutoAttendant)
$chk_Org_Get_UmDialPlan.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_UmDialPlan.Location = $System_Drawing_Point
	$chk_Org_Get_UmDialPlan.Name = "chk_Org_Get_UmDialPlan"
	$chk_Org_Get_UmDialPlan.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_UmDialPlan.TabIndex = 0
	$chk_Org_Get_UmDialPlan.Text = "Get-UmDialPlan"
	$chk_Org_Get_UmDialPlan.UseVisualStyleBackColor = $True
	$bx_UM_Functions.Controls.Add($chk_Org_Get_UmDialPlan)
$chk_Org_Get_UmIpGateway.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_UmIpGateway.Location = $System_Drawing_Point
	$chk_Org_Get_UmIpGateway.Name = "chk_Org_Get_UmIpGateway"
	$chk_Org_Get_UmIpGateway.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_UmIpGateway.TabIndex = 0
	$chk_Org_Get_UmIpGateway.Text = "Get-UmIpGateway"
	$chk_Org_Get_UmIpGateway.UseVisualStyleBackColor = $True
	$bx_UM_Functions.Controls.Add($chk_Org_Get_UmIpGateway)
$chk_Org_Get_UmMailbox.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_UmMailbox.Location = $System_Drawing_Point
	$chk_Org_Get_UmMailbox.Name = "chk_Org_Get_UmMailbox"
	$chk_Org_Get_UmMailbox.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_UmMailbox.TabIndex = 0
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
#	$chk_Org_Get_UmMailboxConfiguration.Size = $System_Drawing_Size_Reusable_chk
#	$chk_Org_Get_UmMailboxConfiguration.TabIndex = 0
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
#	$chk_Org_Get_UmMailboxPin.Size = $System_Drawing_Size_Reusable_chk
#	$chk_Org_Get_UmMailboxPin.TabIndex = 0
#	$chk_Org_Get_UmMailboxPin.Text = "Get-UmMailboxPin"
#	$chk_Org_Get_UmMailboxPin.UseVisualStyleBackColor = $True
#	$bx_UM_Functions.Controls.Add($chk_Org_Get_UmMailboxPin)
$chk_Org_Get_UmMailboxPolicy.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_UmMailboxPolicy.Location = $System_Drawing_Point
	$chk_Org_Get_UmMailboxPolicy.Name = "chk_Org_Get_UmMailboxPolicy"
	$chk_Org_Get_UmMailboxPolicy.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_UmMailboxPolicy.TabIndex = 0
	$chk_Org_Get_UmMailboxPolicy.Text = "Get-UmMailboxPolicy"
	$chk_Org_Get_UmMailboxPolicy.UseVisualStyleBackColor = $True
	$bx_UM_Functions.Controls.Add($chk_Org_Get_UmMailboxPolicy)
#EndRegion Step3 UM tab

#Region Step3 Misc tab
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 4
	$System_Drawing_Point.Y = 33
$tab_Step3_Misc.Location = $System_Drawing_Point
	$tab_Step3_Misc.Name = "tab_Step3_Misc"
	$tab_Step3_Misc.Padding = $System_Windows_Forms_Padding_Reusable
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 488
	$System_Drawing_Size.Width = 542
	$tab_Step3_Misc.Size = $System_Drawing_Size
	$tab_Step3_Misc.TabIndex = 4
	$tab_Step3_Misc.Text = "Misc"
	$tab_Step3_Misc.UseVisualStyleBackColor = $True
	$tab_Step3_ExOrg_Tier2.Controls.Add($tab_Step3_Misc)
$bx_Misc_Functions.Dock = 5
	$bx_Misc_Functions.Font = $font_Calibri_10pt_bold
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 3
	$System_Drawing_Point.Y = 3
	$bx_Misc_Functions.Location = $System_Drawing_Point
	$bx_Misc_Functions.Name = "bx_Misc_Functions"
	$bx_Misc_Functions.Size = $System_Drawing_Size_Step3_box
	$bx_Misc_Functions.TabIndex = 9
	$bx_Misc_Functions.TabStop = $False
	$tab_Step3_Misc.Controls.Add($bx_Misc_Functions)
$btn_Step3_Misc_CheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_Misc_CheckAll.Location = $System_Drawing_Point_Step3_Check
	$btn_Step3_Misc_CheckAll.Name = "btn_Step3_Misc_CheckAll"
	$btn_Step3_Misc_CheckAll.Size = $System_Drawing_Size_Step3_check_btn
	$btn_Step3_Misc_CheckAll.TabIndex = 28
	$btn_Step3_Misc_CheckAll.Text = "Check all on this tab"
	$btn_Step3_Misc_CheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_Misc_CheckAll.add_Click($handler_btn_Step3_Misc_CheckAll_Click)
	$bx_Misc_Functions.Controls.Add($btn_Step3_Misc_CheckAll)
$btn_Step3_Misc_UncheckAll.Font = $font_Calibri_10pt_normal
	$btn_Step3_Misc_UncheckAll.Location = $System_Drawing_Point_Step3_Uncheck
	$btn_Step3_Misc_UncheckAll.Name = "btn_Step3_Misc_UncheckAll"
	$btn_Step3_Misc_UncheckAll.Size = $System_Drawing_Size_Step3_check_btn
	$btn_Step3_Misc_UncheckAll.TabIndex = 29
	$btn_Step3_Misc_UncheckAll.Text = "Uncheck all on this tab"
	$btn_Step3_Misc_UncheckAll.UseVisualStyleBackColor = $True
	$btn_Step3_Misc_UncheckAll.add_Click($handler_btn_Step3_Misc_UncheckAll_Click)
	$bx_Misc_Functions.Controls.Add($btn_Step3_Misc_UncheckAll)
$Col_1_loc = 35
$Col_2_loc = 290
$Row_1_loc = 25
$Row_2_loc = 25
$chk_Org_Get_AdminGroups.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = $Col_1_loc
	$System_Drawing_Point.Y = $Row_1_loc
	$Row_1_loc += 25
	$chk_Org_Get_AdminGroups.Location = $System_Drawing_Point
	$chk_Org_Get_AdminGroups.Name = "chk_Org_Get_AdminGroups"
	$chk_Org_Get_AdminGroups.Size = $System_Drawing_Size_Reusable_chk
	$chk_Org_Get_AdminGroups.TabIndex = 0
	$chk_Org_Get_AdminGroups.Text = "Get memberships of admin groups"
	$chk_Org_Get_AdminGroups.UseVisualStyleBackColor = $True
	$bx_Misc_Functions.Controls.Add($chk_Org_Get_AdminGroups)
#EndRegion Step3 Misc tab

#EndRegion "Step3 - Tests"

#Region "Step4 - Reporting"
$tab_Step4.BackColor = [System.Drawing.Color]::FromArgb(0,255,255,255)
	$tab_Step4.Font = $font_Calibri_8pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 4
	$System_Drawing_Point.Y = 36
	$tab_Step4.Location = $System_Drawing_Point
	$tab_Step4.Name = "tab_Step4"
	$tab_Step4.Padding = $System_Windows_Forms_Padding_Reusable
	$tab_Step4.TabIndex = 3
	$tab_Step4.Text = "  Reporting  "
	$tab_Step4.Size = $System_Drawing_Size_tab_1
	$tab_Master.Controls.Add($tab_Step4)
$btn_Step4_Assemble.Font = $font_Calibri_14pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 20
	$System_Drawing_Point.Y = 15
	$btn_Step4_Assemble.Location = $System_Drawing_Point
	$btn_Step4_Assemble.Name = "btn_Step4_Assemble"
	$btn_Step4_Assemble.Size = $System_Drawing_Size_buttons
	$btn_Step4_Assemble.TabIndex = 10
	$btn_Step4_Assemble.Text = "Execute"
	$btn_Step4_Assemble.UseVisualStyleBackColor = $True
	$btn_Step4_Assemble.add_Click($handler_btn_Step4_Assemble_Click)
	$tab_Step4.Controls.Add($btn_Step4_Assemble)
$lbl_Step4_Assemble.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 138
	$System_Drawing_Point.Y = 15
	$lbl_Step4_Assemble.Location = $System_Drawing_Point
	$lbl_Step4_Assemble.Name = "lbl_Step4"
	$System_Drawing_Size = New-Object System.Drawing.Size
	$System_Drawing_Size.Height = 38
	$System_Drawing_Size.Width = 510
	$lbl_Step4_Assemble.Size = $System_Drawing_Size
	$lbl_Step4_Assemble.TabIndex = 11
	$lbl_Step4_Assemble.Text = "If Office 2003 or later is installed, the Execute button can be used to assemble `nthe output from Tests into reports."
	$lbl_Step4_Assemble.TextAlign = 16
	$tab_Step4.Controls.Add($lbl_Step4_Assemble)
$bx_Step4_Functions.Font = $font_Calibri_10pt_bold
	$System_Drawing_Point_bx_Step4 = New-Object System.Drawing.Point
	$System_Drawing_Point_bx_Step4.X = 27	# 96-69
	$System_Drawing_Point_bx_Step4.Y = 91
	$bx_Step4_Functions.Location = $System_Drawing_Point_bx_Step4
	$bx_Step4_Functions.Name = "bx_Step4_Functions"
	$System_Drawing_Size_bx_Step4 = New-Object System.Drawing.Size
	$System_Drawing_Size_bx_Step4.Height = 487 #482 to short
	$System_Drawing_Size_bx_Step4.Width = 536
	$bx_Step4_Functions.Size = $System_Drawing_Size_bx_Step4
	$bx_Step4_Functions.TabIndex = 0
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
	$chk_Step4_DC_Report.Size = $System_Drawing_Size_Reusable_chk_long
	$chk_Step4_DC_Report.TabIndex = 0
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
	$chk_Step4_Ex_Report.Size = $System_Drawing_Size_Reusable_chk_long
	$chk_Step4_Ex_Report.TabIndex = 1
	$chk_Step4_Ex_Report.Text = "Generate Excel for Exchange servers"
	$chk_Step4_Ex_Report.UseVisualStyleBackColor = $True
	$bx_Step4_Functions.Controls.Add($chk_Step4_Ex_Report)
 #>$chk_Step4_ExOrg_Report.Checked = $True
	$chk_Step4_ExOrg_Report.CheckState = 1
	$chk_Step4_ExOrg_Report.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 50
	$System_Drawing_Point.Y = 75
	$chk_Step4_ExOrg_Report.Location = $System_Drawing_Point
	$chk_Step4_ExOrg_Report.Name = "chk_Step4_ExOrg_Report"
	$chk_Step4_ExOrg_Report.Size = $System_Drawing_Size_Reusable_chk_long
	$chk_Step4_ExOrg_Report.TabIndex = 2
	$chk_Step4_ExOrg_Report.Text = "Generate Excel for Exchange Organization"
	$chk_Step4_ExOrg_Report.UseVisualStyleBackColor = $True
	$bx_Step4_Functions.Controls.Add($chk_Step4_ExOrg_Report)
<# $chk_Step4_Exchange_Environment_Doc.Checked = $True
	$chk_Step4_Exchange_Environment_Doc.CheckState = 1
	$chk_Step4_Exchange_Environment_Doc.Font = $font_Calibri_10pt_normal
	$System_Drawing_Point = New-Object System.Drawing.Point
	$System_Drawing_Point.X = 50
	$System_Drawing_Point.Y = 100
	$chk_Step4_Exchange_Environment_Doc.Location = $System_Drawing_Point
	$chk_Step4_Exchange_Environment_Doc.Name = "chk_Step4_Exchange_Environment_Doc"
	$chk_Step4_Exchange_Environment_Doc.Size = $System_Drawing_Size_Reusable_chk_long
	$chk_Step4_Exchange_Environment_Doc.TabIndex = 2
	$chk_Step4_Exchange_Environment_Doc.Text = "Generate Word for Exchange Documention"
	$chk_Step4_Exchange_Environment_Doc.UseVisualStyleBackColor = $True
	$bx_Step4_Functions.Controls.Add($chk_Step4_Exchange_Environment_Doc)
 #>$Status_Step4.Font = $font_Calibri_10pt_normal
	$Status_Step4.Location = $System_Drawing_Point_Status
	$Status_Step4.Name = "Status_Step4"
	$Status_Step4.Size = $System_Drawing_Size_Status
	$Status_Step4.TabIndex = 12
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
$tab_Step5.TabIndex = 3
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
$bx_Step5_Functions.TabIndex = 0
$bx_Step5_Functions.TabStop = $False
$bx_Step5_Functions.Text = "If you're having trouble collecting data..."
$tab_Step5.Controls.Add($bx_Step5_Functions)

$Status_Step5.Font = $font_Calibri_10pt_normal
$Status_Step5.Location = $System_Drawing_Point_Status
$Status_Step5.Name = "Status_Step5"
$Status_Step5.Size = $System_Drawing_Size_Status
$Status_Step5.TabIndex = 12
$Status_Step5.Text = "Step 5 Status"
$tab_Step5.Controls.Add($Status_Step5)

#EndRegion "Step5 - Having Trouble?"
#>

#Region Set Tests Checkbox States
if (($INI_ExOrg -ne ""))
{
	# Code to parse INI
	write-host "Importing INI settings"
	write-host "ExOrg INI settings: " $ini_ExOrg
	# ExOrg INI
	write-host $ini_ExOrg
	if (($ini_ExOrg -ne "") -and ((Test-Path $ini_ExOrg) -eq $true))
	{
		write-host "File specified using the -INI_ExOrg switch" -ForegroundColor Green
		& ".\O365DC_Scripts\Core_Parse_Ini_File.ps1" -IniFile $INI_ExOrg
	}
	elseif (($ini_ExOrg -ne "") -and ((Test-Path $ini_ExOrg) -eq $false))
	{
		write-host "File specified using the -INI_ExOrg switch was not found" -ForegroundColor Red
	}
}
else
{
	# ExOrg Functions
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
					-or (($objJobsRunning.childjobs[0].output[1] -eq "ExOrg") -and ($JobRunningTime -gt ($intExchJobTimeout/60))))
				{
					try
					{
						(Get-Process | where-object {$_.id -eq $JobPID}).kill()
						write-host "Timer expired.  Killing job process $JobPID - " + $objJobsRunning.name -ForegroundColor Red
						$ErrorText = $objJobsRunning.name + "`n"
						$ErrorText += "Process $JobPID killed`n"
						if ($objJobsRunning.childjobs[0].output[1] -eq "WMI") {$ErrorText += "Timeout $intWMIJobTimeout seconds exceeded"}
						if ($objJobsRunning.childjobs[0].output[1] -eq "ExOrg") {$ErrorText += "Timeout $intExchJobTimeout seconds exceeded"}
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
					        -or (($objJobsRunning.childjobs[0].output[1] -eq "ExOrg") -and ($JobRunningTime -gt ($intExchJobTimeout/60))))
				        {
					        try
					        {
						        (Get-Process | where-object {$_.id -eq $JobPID}).kill()
						        write-host "Timer expired.  Killing job process $JobPID - " + $objJobsRunning.name -ForegroundColor Red
						        $ErrorText = $objJobsRunning.name + "`n"
						        $ErrorText += "Process $JobPID killed`n"
						        if ($objJobsRunning.childjobs[0].output[1] -eq "WMI") {$ErrorText += "Timeout $intWMIJobTimeout seconds exceeded"}
						        if ($objJobsRunning.childjobs[0].output[1] -eq "ExOrg") {$ErrorText += "Timeout $intExchJobTimeout seconds exceeded"}
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
					-or (($objJobsRunning.childjobs[0].output[1] -eq "ExOrg") -and ($JobRunningTime -gt ($intExchJobTimeout/60))))
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
					if ($objJobsRunning.childjobs[0].output[1] -eq "ExOrg") {$ErrorText += "Timeout $intExchJobTimeout seconds exceeded"}
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

Function Get-ExOrgBoxStatus # See if any are checked
{
if (($chk_Org_Get_AcceptedDomain.checked -eq $true) -or
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

Function Get-ExOrgMbxBoxStatus # See if any are checked
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

Function Set-AllFunctionsClientAccess
{
    Param([boolean]$Check)
	$chk_Org_Get_MobileDevice.Checked = $Check
	$chk_Org_Get_MobileDevicePolicy.Checked = $Check
	$chk_Org_Get_AvailabilityAddressSpace.Checked = $Check
	$chk_Org_Get_OwaMailboxPolicy.Checked = $Check
}

Function Set-AllFunctionsGlobal
{
    Param([boolean]$Check)
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

Function Start-O365DCJob
{
    param(  [string]$server,`
            [string]$Job,`              # e.g. "Win32_ComputerSystem"
            [boolean]$JobType,`             # 0=WMI, 1=ExOrg
            [string]$Location,`
            [string]$JobScriptName,`    # e.g. "dc_w32_cs.ps1"
            [int]$i,`                   # Number or $null
            [string]$PSSession)

    If ($JobType -eq 0) #WMI
        {Limit-O365DCJob $intWMIJobs $intWMIPolling}
    else                #ExOrg
        {Limit-O365DCJob $intExOrgJobs $intExOrgPolling}
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

# Connection Urls
# Hard-code these for initial testing
$SharepointAdminCenter = "https://sweetwafflefarm-admin.sharepoint.com"
$ExoConnectioUri = "https://outlook.office365.com/powershell-liveid/"
$SccConnectionUri = "https://ps.compliance.protection.outlook.com/powershell-liveid/"

# Try to re-use an existing connection
$CurrentPSSession = Check-CurrentPsSession
If (($CurrentPSSession -eq $false) -or ($ForceNewConnection -eq $True))
{
	If ($MFA -eq $null)
	{
		write-host "Since MFA is not in use, we can store the credentials for re-use."
		$O365Cred = get-credential
	}
	elseif ($MFA -eq $true)
	{
		#$O365Upn = Read-host "Please enter the user principal name with access to the tenant: "
	}

	<#
	# Connect to Azure AD
	if ($MFA -eq $true)
	{
		write-host "Multi-factor authentication is enabled" -foreground color yellow
		Connect-AzureAD -AccountId $O365Upn
	}
	else
	{
		$AzureCredential = Get-Credential -UserName $AzureAdmin
		Connect-AzureAD -Credential $O365Cred
	}

	# Connect to Sharepoint
	# Need to check this with non-MFA
	Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking
	if ($MFA -eq $true)
	{
		write-host "Multi-factor authentication is enabled" -foreground color yellow
		Connect-SPOService $SharepointAdminCenter -UserName $O365Upn
	}
	else
	{
		Connect-SPOService $SharepointAdminCenter -Credential $O365Cred
	}

	#Connect to Skype
	# Need to check this with non-MFA
	Import-Module SkypeOnlineConnector
	if ($MFA -eq $true)
	{
		write-host "Multi-factor authentication is enabled" -foreground color yellow
		$CsSession = New-CsOnlineSession -UserName $O365Upn
	}
	else
	{
		$CsSession = New-CsOnlineSession -Credential $O365Cred
	}
	Import-PSSession $CsSession -AllowClobber
	#>

	#Connect to Exchange Online
	write-host "Connecting to Exchange Online" -foregroundcolor green
	If ($MFA -eq $true)
	{
		write-host "Multi-factor authentication is enabled" -foregroundcolor yellow
		$ModuleLocation = "$($env:LOCALAPPDATA)\Apps\2.0"
		$ExoModuleLocation = @(Get-ChildItem -Path $ModuleLocation -Filter "Microsoft.Exchange.Management.ExoPowershellModule.manifest" -Recurse )
		If ($ExoModuleLocation.Count -ge 1)
		{
			write-host "ExoPowershellModule.manifest found.  Trying to load the dll." -foregroundcolor green
			$FullExoModulePath =  $ExoModuleLocation[0].Directory.tostring() + "\Microsoft.Exchange.Management.ExoPowershellModule.dll"
			Import-Module $FullExoModulePath  -Force
			$ExoSession	= New-ExoPSSession
			Import-PSSession $ExoSession -AllowClobber
		}
	}
	else
	{
		$ExoSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExoConnectioUri -Credential $O365Cred -Authentication Basic -AllowRedirection
		Import-PSSession $ExoSession -DisableNameChecking -AllowClobber
	}

	#Connect to Security and Compliance Center
	<#
	write-host "Connecting to Security and Compliance Center" -foregroundcolor green
	If ($MFA -eq $true)
	{
		write-host "Multi-factor authentication is enabled" -foregroundcolor yellow
		Connect-IPPSSession -UserPrincipalName $AzureAdmin
	}
	else
	{
		$SccSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $SccConnectionUri -Credential $UserCredential -Authentication Basic -AllowRedirection
		Import-PSSession $SccSession -DisableNameChecking
	}
	#>
}
else
{
	#write-host "check session: "$CurrentPSSession
	#write-host "Force new connection: " $ForceNewConnection
	write-host "Existing connection to Microsoft.Exchange on outlook.office365.com detected."
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
Set-Variable -name intExOrgJobs -Scope global
Set-Variable -name intExOrgPolling -Scope global
Set-Variable -name intExchJobTimeout -Scope global
Set-Variable -name INI -Scope global
Set-Variable -name intMailboxTotal -Scope global

$array_Mailboxes = @()
$UM=$true

if ($JobCount_ExOrg -eq 0) 			{$intExOrgJobs = 10}
	else 							{$intExOrgJobs = $JobCount_ExOrg}
if ($JobPolling_ExOrg -eq 0) 		{$intExOrgPolling = 5}
	else 							{$intExOrgPolling = $JobPolling_ExOrg}
if ($Timeout_ExOrg_Job -eq 0)		{$intExchJobTimeout = 3600} 			# 3600 sec = 60 min
	else 							{$intExchJobTimeout = $Timeout_ExOrg_Job}

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
	write-host "`t`tExOrg Ini:`t" $INI_ExOrg -ForegroundColor Cyan
	$EventText += "`t`tExOrg Ini:`t" + $INI_ExOrg + "`n"
	write-host "`tNon-Exchange cmdlet jobs" -ForegroundColor Cyan
	$EventText += "`tNon-Exchange cmdlet jobs`n"
	write-host "`tExchange cmdlet jobs" -ForegroundColor Cyan
	$EventText += "`tExchange cmdlet jobs`n"
	write-host "`t`tMax jobs:`t" $intExOrgJobs -ForegroundColor Cyan
	$EventText += "`t`tMax jobs:`t" + $intExOrgJobs + "`n"
	write-host "`t`tPolling: `t" $intExOrgPolling " seconds" -ForegroundColor Cyan
	$EventText += "`t`tPolling: `t`t" + $intExOrgPolling + "`n"
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