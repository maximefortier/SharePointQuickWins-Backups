#*************************************************************************************************
#
#		Name: eDMSFarmFullBackupWithEmailReport.ps1
#		Description: This Powershell Script performs a full backup of the eDMS. 
#		Author: Maxime Fortier - www.maximefortier.com
#		Version 1.0
#
#*************************************************************************************************

Add-PsSnapin Microsoft.SharePoint.Powershell

$eDMSBackupDirectory = "H:\Backup\edmsFarmBackup"

Backup-SPFarm -Directory $eDMSBackupDirectory -BackupMethod Full

$BackupReportFile = "H:\Reports\Staging\BackupReport.html"
$SPBackupXMLFile = "H:\Backup\edmsFarmBackup\spbrtoc.xml"

#Read SharePoint Backup and Restore XML File
[xml]$XMLContent = Get-Content $SPBackupXMLFile


#Create an HTML Report of Backup and Restore Operations
$a = "<style>"
$a = $a + "BODY{font-family: Calibri, Arial, sans-serif;}"
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$a = $a + "TH{border-width: 1px;padding: 1px;border-style: solid;border-color: black;background-color:DimGray}"
$a = $a + "TD{border-width: 1px;padding: 1px;border-style: solid;border-color: black;background-color:LightGrey}"
$a = $a + "</style>"
$b = "<H2>eDMS Farm Backup History</H2>"
$GetDateValue = Get-Date
$b = $b + "<H6>Report generated on " + $GetDateValue + "</H6>" 



$XMLContent.SPBackupRestoreHistory.SPHistoryObject | Sort-Object SPFinishTime -Descending | Select-Object SPRequestedBy,SPBackupMethod,SPRestoreMethod,SPStartTime,SPFinishTime,SPWarningCount,SPErrorCount,SPBackupDirectory,SPDirectoryName,SPIsBackup,SPConfigurationOnly,SPDirectoryNumber,SPTopComponent,SPTopComponentId,SPId,SPParentID | ConvertTo-HTML -head $a -body $b | Out-File $BackupReportFile 

Start-Sleep -s 5

#Configure Email Settings
$emailFrom = "FromAddress@YourDomain.com"
$emailTo = "ToAddress@YourDomain.com"
$subject = "eDMS SharePoint Farm Backup"
$body = "eDMS Farm Backup Completed : Review attachment for completion status and backup details."
$smtpServer = "YourMailServerFQDN"
$filePath = $BackupReportFile  

Function sendEmail([string]$emailFrom, [string]$emailTo, [string]$subject,[string]$body,[string]$smtpServer,[string]$filePath)
{
	#Configuring Email
	$email = New-Object System.Net.Mail.MailMessage 
	$email.From = $emailFrom
	$email.To.Add($emailTo)
	$email.Subject = $subject
	$email.Body = $body
 	
	#Configuring Attachment
	$emailAttach = New-Object System.Net.Mail.Attachment $filePath
	$email.Attachments.Add($emailAttach) 

	#Sending Email 

	$smtp = New-Object Net.Mail.SmtpClient($smtpServer)
	$smtp.Send($email)
}

#Call Function 
sendEmail $emailFrom $emailTo $subject $body $smtpServer $filePath
