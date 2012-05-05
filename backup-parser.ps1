[Reflection.Assembly]::LoadFile("c:\program files\microsoft\exchange\web services\1.1\Microsoft.Exchange.WebServices.dll")
$s = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1)
$s.Credentials = new-object net.networkcredential('xxxxxxx', 'xxxxxx', 'domainname.com')
$s.AutoDiscoverUrl("mailbox@domainname.com")
$inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($s, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
$softdel = [Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete

$properties = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$properties.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::HTML

$CVLogix = @"
public class vLogix {
	public string backupName;
	public string[][] jobInfo;
	public int successCount;
	public int errorCount;
}
"@

Add-Type -TypeDefinition $CVLogix

$CAsigra = @"
public class vAsigra {
		public string backupName;
		public string[][] jobInfo;
		public int successCount;
		public int errorCount;
}
"@
Add-Type -TypeDefinition $CAsigra

$CVeeam = @"
public class veeam {
	public string backupName;
	public string[][] jobInfo;
	public int successCount;
	public int errorCount;
}
"@	

Add-Type -TypeDefinition $CVeeam

$vault = @{}
$veeam = @{}
$asigra = @{}

$time = Get-WmiObject win32_Localtime

$backupName = ""
$newFilesBackedUp = ""
$startDate = ""
$errorMessage = ""
$warningMessage = ""

$wasSuccessful = 0

$atLeastOneSuccessful = 0
$atLeastOneFailed = 0

$inbox.FindItems(1000) | % {

	$_.load($properties)
	
	#Check to see if VaultLogix Vaulting Status
	if ($_.Subject -match "Vaultlogix Vaulting Status")
	{
		write-host $_.Body.Text
		add-content c:\scripts\vault.txt $_.Body.Text
		$_.Body -match "Backup set name: ([^<]+)"
		$backupName = $matches[1]

		if ($_.Body -match "Success") {
			$wasSuccessful = 1
		} else { 
			$wasSuccessful = 0 
		} 
		
		if (!($vault.containskey($backupName))) {
			write-host "`$vault doesn't have a job named: " $backupName
			$vaultArray = new-object vLogix
			$vaultArray.backupName = $backupName

			#initialize new backup job in vault hash
			$vault[$backupName] = $vaultArray
			$vault[$backupName].successCount = 0
			$vault[$backupName].errorCount = 0 
		}

		$_.Body -match "new files backed up: ([^<]+)"
		$newFilesBackedUp = $matches[1]

		#Check Start Date
		$_.Body -match "Started at: ([^<]+)"
		$startDate = $matches[1]

		if ($wasSuccessful) {
			#$vault[$backupName].startDatesSuccess += $matches[1]

			#[0] - job success boolean
			#[1] - successful job start date
			#[2] - new files backed up 

			$vault[$backupName].jobInfo += ,@(1, $startDate, $newFilesBackedUp)
			$vault[$backupName].successCount += 1
		}

		else  {	
			#$vault[$backupName].startDatesError += $matches[1]

			$vault[$backupName].errorCount += 1
		
			$_.Body -match "Backup status: ([^<]+)"
			$errorMessage = $matches[1]

			if ($_.Body -match "Warnings:.+\s\n([^<]+)") {
				#$vault[$backupName].warningMessage += $matches[1]
				$warningMessage = $matches[1]
			} else {
                $warningMessage = ""
            }

			$vault[$backupName].jobInfo += ,@(0, $startDate, $errorMessage, $warningMessage)
		}
	}

	#Check to see if Veeam Backup and Replication
	if ($_.Subject -match "Job\s(.+) completed: (\w+)") {
		write-host "Veeam Backup Report Detected.`n"
		write-host $_.Body.Text
		add-content c:\scripts\veeam.txt $_.Body.Text
		write-host "`n"

		$backupName = $matches[1]

		if ($matches[2] -eq "Success") {
			$wasSuccessful = 1
		} else {
			$wasSuccessful = 0
		}


		if (!($veeam.containsKey($backupName))) {
			write-host "`$veeam doesn't have a job named: " $backupName
			$veeamArray = new-object veeam
			$veeamArray.backupName = $backupName

			#initialize new backup job in veeam hash
			$veeam[$backupName] = $veeamArray
			$veeam[$backupName].successCount = 0
			$veeam[$backupName].errorCount = 0
		}

		$_.Body -match "<td nowrap=`"`">([^<]+)"
		$startDate = $matches[1]

		$_.Body -match "<td nowrap=`"`">(\d+\.\d+ \w+)"	
		$totalSize = $matches[1]

		if ($wasSuccessful) {
			$veeam[$backupName].jobInfo += ,@(1, $startdate, $totalSize)
			$veeam[$backupName].successCount += 1
		} else {
			$veeam[$backupName].jobInfo += ,@(0, $startdate, $totalSize)
			$veeam[$backupName].errorCount += 1
		}


	}

	#Check to see if Asigra Backup job
	if ($_.Subject -match "Backup Job") {

		write-host $_.Body.Text
		add-content c:\scripts\asigra.txt $_.Body.Text


		$_.Body -match "Backup Set:.+\\\\([^\\]+)"
		$backupName = $matches[1]
        write-host "Asigra backup set name: " $backupName


		if (!($asigra.containsKey($backupName))) {
			write-host "`$asigra doesn't have a job named: " $backupName
			$asigraArray = new-object vAsigra
			$asigraArray.backupName = $backupName
			
			#initialize new backup job in asigra hash
			$asigra[$backupName] = $asigraArray
			$asigra[$backupname].successCount = 0
			$asigra[$backupName].errorCount = 0
		}


		$_.Body -match "Errors:(&nbsp;)+\s(\d+)"
		$numErrors = $matches[2]	
		$_.Body -match "Warnings:(&nbsp;)+\s(\d+)"
		$numWarnings = $matches[2]
		$_.Body -match "Backed up files:([^<]+)"
		$newFilesBackedUp = $matches[1]
		write-host "Found " $newFilesBackedUp + " new backed up files.`n"
		$_.Body -match "Started at:(&nbsp;)+([^<]+)"
		$startdate = $matches[2]

		if ($_.Body -match "Backup (Successful|Completed)") {
			$wasSuccessful = 1	
			$asigra[$backupName].jobInfo += ,@(1, $startDate, $newFilesBackedUp, $numErrors, $numWarnings)
			$asigra[$backupName].successCount += 1
		}
		else {
			$wasSuccessful = 0
			$_.Body -match "Completion:(&nbsp;)+([^<]+)"
			$message = $matches[2]
			$asigra[$backupName].jobInfo += ,@(0, $startDate, $newFilesBackedUp, $numErrors, $numWarnings, $message)
			$asigra[$backupName].errorCount += 1
		}
	}

$_.Delete($softdel)
}

$mail = New-object microsoft.exchange.webservices.data.emailmessage($s)
$mail.Subject = "Backup Statistics - Beta v2"
$Body = "<head><style type=`"text/css`">body { font-family: Calibri; }</style></head><H2>Vaultlogix Report</H2>" 

$vault.getenumerator() | % {

	$jobCounter = 0

	$backupDates = $_.Value.jobInfo
	write-host "VaultLogix - BackupDates count: " $backupDates.count

	if ($backupDates.count -gt 0) {

		$Body += "&nbsp;&nbsp;<b>Backup Job: " + $_.Value.backupName + "</b><BR>"
		$Body += "&nbsp;&nbsp;&nbsp;&nbsp;Total Successful Jobs: " + $_.Value.successCount + "<BR>"
		$Body += "&nbsp;&nbsp;&nbsp;&nbsp;Total Failed Jobs: " + $_.Value.errorCount + "<BR>"
		$Body += "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Job Details:<BR>"
		

		while ($jobCounter -le $backupDates.count) {
			if ($backupDates[$jobCounter][0] -eq 1) {
				#Successful job
				$Body += "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color=`"green`"><b>Success</b></font> - Date: " + $backupDates[$jobCounter][1] + " - Files backed up: " + $backupDates[$jobCounter][2] + "<BR>"
			}
			else {
				#Failed job
				$Body += "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font
                color=`"red`"><b>Failed</b></font> - Date: " +
                $backupDates[$jobCounter][1] + " - <b>Error message:</b> " + $backupDates[$jobCounter][2] + "</font><BR>"
				if ($backupDates[$jobCounter][3] -ne "") {
					$Body += "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>Warning message:</b> " + $backupDates[$jobCounter][3] + "<BR>"
				}
			}

		$jobCounter += 1
		}

		#$Body += "&nbsp;&nbsp;&nbsp;&nbsp;New Files Backed up: " + $_.newFilesBackedUp + "<BR>"

	}
	else {
		#$Body += "<font color=`"red`"><H3>No successful jobs were reported!</H3></font>"
	}
	
	$Body += "<BR>"
}

$Body += "<BR><H2>Asigra Backup Report</H2>"

$asigra.getenumerator() | % {
	$jobCounter = 0

	$backupDates = $_.Value.jobInfo
	write-host "ASIGRA - BackupDates count: " $backupDates.count

	if ($backupDates.count -gt 0) {

		$Body += "&nbsp;&nbsp;<b>Backup Job: " + $_.Value.backupName + "</b><BR>"
		$Body += "&nbsp;&nbsp;&nbsp;&nbsp;Total Successful jobs: " + $_.Value.successCount + "<BR>"
		$Body += "&nbsp;&nbsp;&nbsp;&nbsp;Total Failed Jobs: " + $_.Value.errorCount + "<BR>"
		$Body += "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Job Details:<BR>"	

		while ($jobCounter -le $backupDates.count) {
			if ($backupDates[$jobCounter][0] -eq 1) {
				#Successful Job
				$Body += "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color=`"green`"><b>Success</b></font> - Date: " + $backupDates[$jobCounter][1] + " - Files Backed up: " + $backupDates[$jobCounter][2] + " - Errors: " + $backupDates[$jobCounter][3] + " - Warnings: " + $backupDates[$jobCounter][4] + "<BR>"
			}
			else {
				#Failed job
				$Body += "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color=`"red`"><b>Failed</b></font> - Date: " + $backupDates[$jobCounter][1] + " - Errors: " + $backupDates[$jobCounter][3] + " - Warnings: " + $backupDates[$jobCounter][4] + "<BR>"
				$Body += "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>Message:</b> " + $backupDates[$jobCounter][5] + "<BR>"
			}

		$jobCounter += 1
		
		}

		$Body += "<BR>"
	}

}

$Body += "<BR>"
$Body += "<BR><H2>Veeam Backup Report</H2>"

$veeam.getenumerator() | % {
	$jobCounter = 0

	$backupDates = $_.Value.jobInfo
	write-host "Veeam - BackupDates count: " $backupDates.count

	if ($backupDates.count -gt 0) {
		$Body += "&nbsp;&nbsp;<b>Backup Job: " + $_.Value.backupName + "</b><BR>"
		$Body += "&nbsp;&nbsp;&nbsp;&nbsp;Total Successful Jobs: " + $_.Value.successcount + "<BR>"
		$Body += "&nbsp;&nbsp;&nbsp;&nbsp;Total Failed Jobs: " + $_.Value.errorcount + "<BR>"
		$Body += "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Job Details:<BR>"

		while ($jobCounter -le $backupDates.count) {
			if ($backupDates[$jobCounter][0] -eq 1) {
				#Successful Job
				$Body += "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color=`"green`"><b>Success</b></font> - Date: " + $backupDates[$jobCounter][1] + " - Total Size: " + $backupDates[$jobCounter][2] + "<BR>"
			} else {
				#Failed Job
				$Body += "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color=`"red`"><b>Failed</b></font> - Date: " + $backupDates[$jobCounter][1] + " - Total Size: " + $backupDates[$jobCounter][2] + "<BR>"
			}

			$jobCounter += 1
		}

		$Body += "<BR>"

	}
}

$mail.Body = $Body
[Void] $mail.ToRecipients.Add("email@domainname.com")
$mail.SendAndSaveCopy()
