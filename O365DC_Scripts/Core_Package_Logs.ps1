[array]$FileList = $null
write-host "Starting to package the log files" -ForegroundColor Green
$Now = Get-Date
$Append = [string]$Now.month + "_" + [string]$now.Day + "_" + `
    [string]$now.year + "_" + [string]$now.hour + "_" + [string]$now.minute `
    + "_" + [string]$now.second

$O365DCLogsFile = ".\O365DC_Events_" + $append + ".txt"
#$O365DCLogs = Get-WinEvent -ProviderName O365DC |fl
Write-Host "Getting the last 24 hours of O365DC events from the application log" -ForegroundColor Green
$O365DCLogs = Get-EventLog -LogName application -Source O365DC -After $Now.AddDays(-1) | Format-List -Property TimeGenerated,EntryType,Source,EventID,Message
$O365DCLogs | Out-File -FilePath $O365DCLogsFile -Force

write-host "Gathering..." -ForegroundColor Green
$FilesInCurrentFolder = Get-ChildItem
foreach ($a in $FilesInCurrentFolder)
{
	#If (($a.name -like ($O365DCLogsFile.replace('.\',''))) -or `
	If (($a.name -like "O365DC_Events*") -or `
		($a.name -like "O365DC_Step3*") -or `
		($a.name -like "Failed*"))
		{
			write-host $a.fullname
			$FileList += [string]$a.fullname
		}
}
$ZipFilename = (get-location).path + "\O365DCPackagedLogs_" + $append + ".zip"
if (-not (Test-Path -LiteralPath $ZipFilename))
{Set-Content -Path $ZipFilename ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))}
$ZipFile = (New-Object -ComObject shell.application).NameSpace($ZipFilename)
write-host "Packaging..." -ForegroundColor Green
ForEach ($File in $FileList)
{
	write-host "Zipping $File"
	$zipfile.CopyHere($File)
}
write-host "Finished collecting logs." -ForegroundColor Green
write-host "Output log is $O365DCLogsFile" -ForegroundColor Green
