#*************************************************************************************************
#
#		Name: eDMSDatabaseSizingReport.ps1
#		Description: This Powershell Script sends capacity information of the eDMS' databases. 
#		Author: Maxime Fortier - www.maximefortier.com
#		Version 1.0
#
#*************************************************************************************************


$eDMSDatabaseSizingReport = "H:\Reports\Staging\DBSizingReport.html"
$eDMSDiskSizingReport = "H:\Reports\Staging\DiskSizingReport.html"
$SQLServerInstance = "localhost\edms"


$a = "<style>"
$a = $a + "BODY{font-family: Calibri, Arial, sans-serif;}"
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$a = $a + "TH{border-width: 1px;padding: 1px;border-style: solid;border-color: black;background-color:DimGray}"
$a = $a + "TD{border-width: 1px;padding: 1px;border-style: solid;border-color: black;background-color:LightGrey}"
$a = $a + "</style>"
$b = "<H2>eDMS Disk Sizing</H2>"
$GetDateValue = Get-Date
$b = $b + "<H6>Report generated on " + $GetDateValue + "</H6>" 

#********************************************************************
# Code snippet from JAKOB BINDSLET, http://mspowershell.blogspot.ca
$outData = @("")
$dataFromServer = Get-WmiObject Win32_Volume -ComputerName localhost | Select-Object SystemName,Label,Name,DriveLetter,DriveType,Capacity,Freespace

foreach ($currline in $dataFromServer) {
    if ((-not $currline.name.StartsWith("\\")) -and ($currline.Drivetype -ne 5)) {
        [float]$tempfloat = ($currline.Freespace / 1000000) / ($currline.Capacity / 1000000)
        $temppercent = [math]::round(($tempfloat * 100),2)
        add-member -InputObject $currline -MemberType NoteProperty -name FreePercent -value "$temppercent %"
        $outData = $outData + $currline
    }
}
$outData | Select-Object SystemName,Name,Label,Capacity,Freespace, FreePercent | sort-object -property FreePercent | ConvertTo-HTML -head $a -body $b | Out-File $eDMSDiskSizingReport


# End of code snippet from JAKOB BINDSLET, http://mspowershell.blogspot.ca
#********************************************************************

$a = "<style>"
$a = $a + "BODY{font-family: Calibri, Arial, sans-serif;}"
$a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
$a = $a + "TH{border-width: 1px;padding: 1px;border-style: solid;border-color: black;background-color:DimGray}"
$a = $a + "TD{border-width: 1px;padding: 1px;border-style: solid;border-color: black;background-color:LightGrey}"
$a = $a + "</style>"
$b = "<H2>eDMS Database Sizing</H2>"
$GetDateValue = Get-Date
$b = $b + "<H6>Report generated on " + $GetDateValue + "</H6>" 


Invoke-Sqlcmd -Query "EXEC ManagementTools.dbo.usp_Sizing" -ServerInstance $SQLServerInstance | ConvertTo-HTML -head $a -body $b | Out-File $eDMSDatabaseSizingReport

Start-Sleep -s 5

#Configure Email Settings
$emailFrom = "FromAddress@YourDomain.com"
$emailTo = "ToAddress@YourDomain.com"
$subject = "eDMS SharePoint Database Sizing Report"
$body = "eDMS Database Sizing Report : Review attachment for detailed database sizing information."
$smtpServer = "YourMailServerFQDN"
#$filePath = $eDMSDatabaseSizingReport 


#Configuring Email
$email = New-Object System.Net.Mail.MailMessage 
$email.From = $emailFrom
$email.To.Add($emailTo)
$email.Subject = $subject
$email.Body = $body
	
#Configuring Attachment
$emailAttachDBSizing = New-Object System.Net.Mail.Attachment $eDMSDatabaseSizingReport
$emailAttachDiskSizing = New-Object System.Net.Mail.Attachment $eDMSDiskSizingReport
$email.Attachments.Add($emailAttachDBSizing) 
$email.Attachments.Add($emailAttachDiskSizing) 

#Sending Email 
$smtp = New-Object Net.Mail.SmtpClient($smtpServer)
$smtp.Send($email)


