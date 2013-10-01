#*************************************************************************************************
#
#		Name: eDMSBackupRetentionAndArchiving.ps1
#		Description: This Powershell Script deletes the eDMS Backups that are expired. 
#		Author: Maxime Fortier - www.maximefortier.com
#		Version 1.0
#
#*************************************************************************************************


$BackupReportFile = "H:\Reports\Staging\BackupRetentionReport.html"
$SPBackupXMLFile = "H:\Backup\edmsFarmBackup\spbrtoc.xml"


#Add-PsSnapin Microsoft.SharePoint.Powershell

#Configure Email Settings
$emailFrom = "FromAddress@YourDomain.com"
$emailTo = "ToAddress@YourDomain.com"
$subject = "eDMS Backup Retention and Archiving"
$body = "eDMS Backup Retention and Archiving : "
$smtpServer = "YourMailServerFQDN"
$filePath = $BackupReportFile

# Location of spbrtoc.xml
$spbrtoc = "H:\Backup\edmsFarmBackup\spbrtoc.xml" 


#Create an HTML Report of Backup and Restore Operations
$a = "<style>"
$a = $a + "BODY{font-family: Calibri, Arial, sans-serif;}"
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$a = $a + "TH{border-width: 1px;padding: 1px;border-style: solid;border-color: black;background-color:DimGray}"
$a = $a + "TD{border-width: 1px;padding: 1px;border-style: solid;border-color: black;background-color:LightGrey}"
$a = $a + "</style>"
$b = "<H2>eDMS Farm Backup Retention</H2>"
$GetDateValue = Get-Date
$b = $b + "<H6>Report generated on " + $GetDateValue + "</H6>" 

# Days of backup that will be remaining after backup cleanup.
$Days = 15 


# Import the Sharepoint backup report xml file
[xml]$XMLContent = Get-Content $spbrtoc 


# Find the old backups in spbrtoc.xml
#$old = $XMLContent.SPBackupRestoreHistory.SPHistoryObject | Where-Object { $_.SPStartTime -lt ((Get-Date).adddays(-$Days))  }
$ExpiredBackups = $XMLContent.SPBackupRestoreHistory.SPHistoryObject | Where-Object {($_.SPIsBackup -eq "True") -and ($_.SPStartTime -lt ((Get-Date).adddays(-$Days)))}


if ($ExpiredBackups -eq $Null) 
{ 
	$body = $body + "No reports of backups older than $days days found in spbrtoc.xml.`nspbrtoc.xml isn't changed and no files are removed.`n"
}
else
{
	# Delete the old backups from the Sharepoint backup report xml file
	$ExpiredBackups | % { $XMLContent.SPBackupRestoreHistory.RemoveChild($_) } 


	# Delete the physical folders in which the old backups were located
	$ExpiredBackups | % { Remove-Item $_.SPBackupDirectory -Recurse } 


	# Save the new Sharepoint backup report xml file
	$XMLContent.Save($spbrtoc)
	$body = $body + "Backup(s) entries older than $Days days were removed from spbrtoc.xml and from the filesystem."
}

$XMLContent.SPBackupRestoreHistory.SPHistoryObject | Where-Object {$_.SPIsBackup -eq "True"} | Sort-Object SPFinishTime -Descending | Select-Object SPBackupMethod,SPStartTime,SPFinishTime,SPBackupDirectory,SPDirectoryName,SPIsBackup | ConvertTo-HTML -head $a -body $b | Out-File $BackupReportFile 

#

Start-Sleep -s 5


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