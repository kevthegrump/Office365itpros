# ReportSharedMailboxLicenses.PS1
# a script to analyze shared mailboxes and report if they need licenses
# https://github.com/12Knocksinna/Office365itpros/blob/master/ReportSharedMailboxLicenses.PS1

# KL Added Connection cmdlets 
# Added multiple Output options i.e. CSV to Clipboard (Default) ,File or GridView
# Added logic for >50GB Mailboxes with EXOPLAN2 License
# Added LicenseReason to Output

# KL Added Section
if (Get-PSSession | Where-Object { $_.computername -eq 'outlook.office365.com' -and $_.Availability -eq 'Available' } ) {
    write-output "Already connected to Exchange Online"
    }
    ELSE 
    {
        connect-exchangeonline
    }
Connect-MgGraph


# Check that we have the necessary Exchange Online module loaded
$ModulesLoaded = Get-Module | Select Name
If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script"; break}

# Define some variables
$ExoPlan1 = "9aaf7827-d63c-4b61-89c3-182f06f82e5c"
$ExoArchiveAddOn = "176a09a6-7ec5-4039-ac02-b2791c6ba793"
$ExoPlan2 = "efb87545-963c-4e0d-99df-69c6916d9eb0" # See https://docs.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
$MailboxLimit = 50GB
$MailboxLimitText=[string]([math]::Round(($MailboxLimit / 1GB),5)) +  "GB" # KL Added

Write-Host "Finding shared mailboxes..."
[array]$Mbx = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -Properties ProhibitSendReceiveQuota, ArchiveGuid, ArchiveName, ArchiveStatus, AutoExpandingArchiveEnabled, PrimarySmtpAddress, LitigationHoldEnabled -ResultSize Unlimited
If (!($Mbx)) {Write-Host "Can't find any shared mailboxes - exiting." ; break}

$i = 0
$SharedMbxReport = [System.Collections.Generic.List[Object]]::new()
ForEach ($M in $Mbx) {
   $i++
   $reason=@()
   Write-Host ("Processing mailbox {0} ({1} of {2})" -f $M.DisplayName, $i, $Mbx.count)
   $NeedsLicense = $False; $ArchiveStatus = $Null; $ExoArchiveLicense = $False; $ExoPlan2License = $False; $LicenseStatus = "OK"; $ArchiveStats = $Null
   $MailboxOverSize = $False; $ExoPlan1License = $False; $ArchiveMbxSize = $Null
   $NeedsLicenseArchive=$false #KL Added

   $MbxStats = Get-ExoMailboxStatistics -Identity $M.ExternalDirectoryObjectId -Properties LastLogonTime
   $MbxSize = [math]::Round(($MbxStats.TotalItemSize.Value.toBytes() / 1GB),5)
   If ($M.ArchiveStatus -ne "None") { #Mailbox has an archive
      $ArchiveStats = Get-ExoMailboxStatistics -Archive -Identity $M.ExternalDirectoryObjectId 
      IF ($ArchiveStats) {       
          $ArchiveMbxSize = [math]::Round(($ArchiveStats.TotalItemSize.Value.toBytes() / 1GB),5)}
   }
   $Licenses = Get-MgUserLicenseDetail -UserId $M.ExternalDirectoryObjectId | Select-Object -ExpandProperty ServicePlans | Where-Object {$_.ProvisioningStatus -eq "Success"} | Sort ServicePlanId -Unique
   If ($Licenses) { # The mailbox has some licenses
     If ($ExoArchiveAddOn -in $Licenses.ServicePlanId) { $ExoArchiveLicense = $True }
     If ($ExoPlan2 -in $Licenses.ServicePlanId) { $ExoPlan2License = $True }
     If ($ExoPlan1 -in $Licenses.ServicePlanId) { $ExpPlan1License = $True }
  }

  # Mailbox has an archive and it doesn't have an Exchange Online Plan 2 license, unless it has Exchange Online Plan 1 and the
  # archive add-on
  If ($M.ArchiveStatus -eq "Active") {
    If ($ExoPlan2License -eq $False) { 
        $NeedsLicenseArchive = $True 
        }
    If ($ExoPlan1License -eq $True -and $ExoArchiveLicense -eq $True) { 
        $NeedsLicenseArchive = $False  
    }
    #KL Added Reason - prob more efficient way to do this
    if ($NeedsLicenseArchive) {
        $NeedsLicense=$NeedsLicenseArchive
        $reason +="Archive"
    }
  }
  # Mailbox is on litigation hold and it doesn't have an Exchange Online Plan 2 license
  If ($M.LitigationHoldEnabled -eq $True -and $ExoPlan2License -eq $False)  { $NeedsLicense = $True 
    $reason +="Litigation Hold" #KL Added
  }
  # Mailbox is over the 50GB limit for unlicensed shared mailboxes
  If ($MbxStats.TotalItemSize.value -gt $MailboxLimit) { # Exceeds mailbox size for unlicensed shared mailboxes
      if ($ExoPlan2License  -eq $false) { #KL Added Check for ExOPLAN2
          $MailboxOverSize = $True
          $NeedsLicense = $True
          $Reason+="Mailbox > $MailboxLimitText" #KL ADDED
       }
      }

  $ReportLine = [PSCustomObject][Ordered]@{  # Write out details of the private channel and its members
       User                  = $M.DisplayName
       UPN                   = $M.PrimarySmtpAddress
       NeedsLicense          = $NeedsLicense
       LicenseReason         = $reason
       MailboxSizeGB         = $MbxSize
       MailboxItems          = $MbxStats.ItemCount
       MailboxOverSize       = $MailboxOverSize
       Archive               = $M.ArchiveStatus
       ArchiveSizeGB         = $ArchiveMbxSize
       ArchiveItems          = $ArchiveStats.ItemCount
       LitigationHoldEnabled = $M.LitigationHoldEnabled
       ExoPlan2              = $ExoPlan2License
       ExoArchiveAddOn       = $ExoArchiveLicense
       LastAccessedMailbox   = $mbxstats.lastlogontime #KL Added
       }       
     $SharedMbxReport.Add($ReportLine) 
 } #End ForEach $Mbx

[array]$NeedsLicenseList = $SharedMbxReport | Where-Object {$_.NeedsLicense -eq $True}
Write-Host " "
Write-Host ("{0} shared mailboxes scanned" -f $Mbx.count)
If ($NeedsLicenseList) {
 Write-Host ("{0} mailbox(es) need an Exchange Online Plan 2 license" -f $NeedsLicenseList.count)
 Write-Host ("Licenses must be assigned to the following mailboxes: {0}" -f $NeedsLicenseList.User -join ", " )
 $NeedsLicenseList | Format-Table User, MailboxOverSize, Archive, LitigationHoldEnabled
}
Else {
 Write-Host "Congratulations! All of your shared mailboxes are properly licensed." }

#KL Added Output Section
#Format CSV Output
$csvoutput=$SharedMbxReport | Select-object USer, UPN, NeedsLicense, @{n="Reason";e={$_.licenseReason -join "," }},MailboxSizeGB, MailboxItems, MailboxOverSize,Archive,ArchiveSizeGB,ArchiveItems,LitigationHoldEnabled,ExoPlan2,ExoArchiveAddOn,LastAccessedMailbox | convertto-csv -NoTypeInformation

#KL ADDED
#OUTPUT SECTION UNCOMMENT REQUIRED OPTION

#Display GridView
#$SharedMbxReport | Out-GridView 

#Copy CSV output to CLipboard
$csvOuput | set-clipboard 

#CSV Output to File
$csvfile="./ShareMailboxReport.csv"
$csvoutput | Out-File $csvfile

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
