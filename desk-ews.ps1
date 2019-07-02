#=======================================================================================================================
# Desk Hammer                                   
# See https://github.com/ciawhatsthat/ESA-MTL   
# For deatils                                   
# v1.1  --  July 2019                           
#
# Pre-reqs
# -A shared mailbox that the SOC will have acces to
# -2 folders in shared mailbox (SOC/COMPLETED)
# -A service account with access to that mailbox 
# -Service account with permissions to run eDiscovery with full permissions
#  *Add-MailboxPermission -Identity "phish@mydomain.com" -User svc_serviceaccount -AccessRights FullAccess
#
#=======================================================================================================================

#=======================================================================================================================#
# Global Definitions                            #
#=======================================================================================================================#
#SOC mailbox
$mb = "phish@mydomain.com"
#see https://www.undocumented-features.com/2015/10/01/storing-powershell-credentials-in-the-local-user-registry/
$hkcupath = "HKCU:\Software\mydomain\Credentials\o365"
#where the SOC will move emails to be proccessed
$startpath = '/SOC'
#where the emails will be moved after processing
$endpath = '/COMPLETED'
$secureCredUserName = (Get-ItemProperty -Path $hkcupath).UserName
$secureCredPassword = (Get-ItemProperty -Path $hkcupath).Password
$securePassword = ConvertTo-SecureString $secureCredPassword
#ensure that MSFT EWS is installed (https://www.microsoft.com/en-us/download/details.aspx?id=42951)
Import-Module -name "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
#set Exchange Version (https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/ews-schema-versions-in-exchange)
$exchVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
$exchService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($exchVersion)
$exchService.Url = "https://outlook.office365.com/EWS/Exchange.asmx"
$cred = New-Object System.Net.NetworkCredential($secureCredUserName, $securePassword)
$exchService.Credentials = $cred

#use FindTargetFolder get the 2 folder locations for EWS processing
$startfolder = FindTargetFolder($startpath)
$completedfolder = FindTargetFolder($endpath)
#create the query - this looks only for 1 email with an attachment
$sfattachment = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::HasAttachments, $true)
$sfcollection = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And);
$sfcollection.add($sfattachment)
$view = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList 1
#pull down the email into an object from the the startfolder location
$foundemails = $startfolder.FindItems($sfcollection,$view)

#loop through each email, pull the attachment (must be in .msg format, sadly the knowbe4 button sends as .eml and it doesn't work)
#from the attachment grab the sender, subject and date and pass that on to the ninja to start eDiscovery
foreach ($email in $foundemails.Items){
        $email.Load()
        $attachments = $email.Attachments
        foreach ($attachment in $attachments){
        $attachment.Load()
        $attachmentname = $attachment.FileName
        $psender = $attachment.item.Sender.Address
        $psubject = $attachment.Item.Subject
        #This is specific to our environment, since we tag external with our email security appliance, YMMV
        $psubject = $psubject.Replace("[EXTERNAL]","")
        $psubject = $psubject.Replace("[MARKETING]","")
        $psubject = $psubject.Replace("[BULK]","")
        $psubject = $psubject.trim()
        $psenton = $attachment.Item.DateTimeSent.ToString('M/d/y')
        }
    #mark email as read and move it to the completed folder
    $email.IsRead = $true
	$email.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
	[VOID]$email.Move($completedfolder.Id)        
    
    #pass it on to the deskninja for processing
    deskninja -senton $psenton -sender $psender -subject $psubject
    }
#clean up the variables
Get-Variable -Exclude Session,banner | Remove-Variable -EA 0

#this function uses the folder path to find the right folder and use it for processing
Function FindTargetFolder($folderpath){
    $tftargetidroot = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$mb)
    $tftargetfolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchService,$tftargetidroot)
    $pfarray = $folderpath.Split("/")

    # Loop processed folders path until target folder is found
    for ($i = 1; $i -lt $pfarray.Length; $i++){
    $fvfolderview = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1)
    $sfsearchfilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,$pfarray[$i])
        $findfolderresults = $exchService.FindFolders($tftargetfolder.Id,$sfsearchfilter,$fvfolderview)

    if ($findfolderresults.TotalCount -gt 0){
    foreach ($folder in $findfolderresults.Folders){
    $tftargetfolder = $folder
            }
        }
    
    }
    return $tfTargetFolder
    }

function deskninja {
    
    [cmdletbinding()]            
    param ( 
    [string]$sender,
    [string]$senton,
    [string]$subject
    )
# ====================================================================================================================
# DEFINE Function PARAMETERS
# ====================================================================================================================

# Mask errors
$ErrorActionPreference = "silentlycontinue"
$warningPreference = "silentlyContinue"

# Location of the MTL https://github.com/ciawhatsthat/ESA-MTL/blob/master/esa-mtl.ps1
$traceLog = "C:\location\of\esa-mtl.csv"
# ====================================================================================================================
# Location of the active users text file
# This ensures that when eDiscovery is run its only run against current/active users
# $list = @()
# $list += Get-ADUser -filter * -SearchBase 'OU=Where Active Users Live' -Properties proxyaddresses | select -ExpandProperty proxyaddresses
# # had to do some substing here to remove all x500 and ensure they contain the o365 address string @mydomain.onmicrosoft.com
# $list | Where-Object { $_.Substring(0,4) -ne "X500" } | Where-Object {!($_.Contains("@mydomain"))} | ForEach-Object {$_.substring(5)} | Set-Content -Path c:\path\to\file.csv -Force
# remove-variable list
#
# This is then scheduled to run every morning
# ====================================================================================================================
$activeusers = Get-Content -path "C:\location\of\activeusers.txt"

#see https://www.undocumented-features.com/2015/10/01/storing-powershell-credentials-in-the-local-user-registry/
$hkcupath = "HKCU:\Software\mydomain\Credentials\O365"

# SMTP relay and addresses for failure alerting
$smtprelay = "x.x.x.x"
$esamtlEmail = "Email Address <email@mydomain.com>"
$esamtlAdmin = "admins@mydomain.com"

# ====================================================================================================================
# Get list of users, compare to list of active users, create eDiscovery search
# ====================================================================================================================

        #The Do/while will continue to search for the email for 10 minutes, in case the email hasn't shown up in the MTL yet
        
        $Checkrecips = 0
        $recips = $null
        
        Do
        {
          if ($checkrecips -gt 10) {
            Send-MailMessage -from $esamtlEmail -to $esamtlAdmin -Subject "Desk Hammer timed out looking for email in MTL" -SmtpServer $smtprelay
            break
            }
          #import the tracking log for manipulation and search for the sender found in the PHISHING mailbox
          $recips = Import-Csv -path $tracelog | Where-Object {($_.sender -like "*$sender*") -or ($_.replyTo -like "*$sender*")}
          #search for the subject from the PHISHING mailbox
          $recips = $recips | Where-Object Subject -like "*$subject*"
          # A forced check to break the loop before the 1st sleep if the email is found on first try
          if ($recips -ne $null) {break}
          
          $Checkrecips++
          start-sleep -s 60
          
        } until ($recips -ne $null)
       
        #create an object of active recipiencts to parse through and pass to the eDiscovery
        $recips = $recips | Select-Object recipients
        $recips = $recips | Where-Object {$activeusers -contains $_.recipients} | Select-Object -ExpandProperty recipients
        #create the query to pass to the eDiscovery for the initial focused stop the bleeding search 
        $ContentMatchQuery = "sent=$senton AND participants:$sender AND subject:`"$subject`""
        #create the query for the straggler search eg someone in initial blast forwarded email to another internal user
        $stragglerMatchQUery = "sent=$senton AND subject:`"$subject`""
        #create a name with timestamp for the eDiscovery to stop the bleeding and straggler
        $search_name = $sender + ' - ' + (Get-Date).ToString()
        $straggler_name = 'Straggler - ' + $sender + ' - ' + (Get-Date).ToString()

# ====================================================================================================================
# Office 365 Protection Authentication
# ====================================================================================================================

# Credentials stored in registry 
    
    try {
        $secureCredUserName1 = (Get-ItemProperty -Path $hkcupath).UserName
        $secureCredPassword1 = (Get-ItemProperty -Path $hkcupath).Password
        $securePassword = ConvertTo-SecureString $secureCredPassword
        $credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $secureCredUserName1, $securePassword1
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://ps.compliance.protection.outlook.com/powershell-liveid" -Credential $credential -Authentication Basic -AllowRedirection -erroraction stop
        Import-PSSession $Session -AllowClobber -DisableNameChecking
    }
     Catch {
        #send an email to admins, disable the script.
        Remove-PSSession $Session
        Disable-ScheduledTask -TaskName "Desk-Hammer"
        Send-MailMessage -from $esamtlEmail -to $esamtlAdmin -Subject "Desk Hammer failed at connecting to O365, task has been disabled" -SmtpServer $smtprelay
        break;
    }

# ====================================================================================================================
# Create and run the eDsicovery
# ====================================================================================================================
       
 
        try {

            New-ComplianceSearch -Name "$search_name" -ExchangeLocation $recips -ContentMatchQuery $contentmatchquery -erroraction stop
            
        }catch{

            $ErrorMessage = $_.Exception
            Remove-PSSession $session
            Send-MailMessage -from $esamtlEmail -to $esamtlAdmin -Subject "Desk hammer failed during new-compliancesearch" -body $ErrorMessage -SmtpServer $smtprelay
            break 
        
        }
        
        try {

        Start-ComplianceSearch -identity "$search_name" -erroraction stop
            
        }catch{

            $ErrorMessage = $_.Exception
            Remove-PSSession $session
            Send-MailMessage -from $esamtlEmail -to $esamtlAdmin -Subject "Desk hammer failed during start-compliancesearch" -body $ErrorMessage -SmtpServer $smtprelay
            break 
        
        }
    
        while ( (get-compliancesearch $search_name | Select-Object -expand status) -ne "Completed" ) { Start-Sleep 5; write-host "." -NoNewline }
    
        try {
            $spamcount = (get-compliancesearch -identity $search_name).items        
            New-ComplianceSearchAction -SearchName $search_name -purge -purgetype softdelete -Confirm:$false -erroraction stop
            
        }catch{

            $ErrorMessage = $_.Exception
            Remove-PSSession $session
            Send-MailMessage -from $esamtlEmail -to $esamtlAdmin -Subject "Desk hammer failed during soft delete action" -body $ErrorMessage -SmtpServer $smtprelay
            break 
        
        }

        while ( (get-compliancesearchaction ($search_name + "_purge") | Select-Object -expand status) -ne "Completed" ) { Start-Sleep 1; write-host "." -NoNewline }

    
        # Start full content straggler search and send email to admins
        
        try {

            New-ComplianceSearch -Name $straggler_name -ExchangeLocation all -ExchangeLocationExclusion phish@bch.org -ContentMatchQuery $stragglerMatchQUery -erroraction stop
            
        }catch{

            $ErrorMessage = $_.Exception
            Remove-PSSession $session
            Send-MailMessage -from $esamtlEmail -to $esamtlAdmin -Subject "Desk hammer failed during straggler new-compliancesearch" -body $ErrorMessage -SmtpServer $smtprelay
            
        }
        
        try {

        Start-ComplianceSearch -identity $straggler_name -erroraction stop
            
        }catch{

            $ErrorMessage = $_.Exception
            Remove-PSSession $session
            Send-MailMessage -from $esamtlEmail -to $esamtlAdmin -Subject "Desk hammer failed during straggler start-compliancesearch" -body $ErrorMessage -SmtpServer $smtprelay
            
        }

        $straggler_body = @"

        The desk hammer found and deleted $spamcount emails.

        The straggler search was started to look for any emails that may have been forwared and not in the initial tracking log.
        
        If the subject: ($subject), is not common, then run:


            # Create a session to protection.office.com
            `$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://ps.compliance.protection.outlook.com/powershell-liveid" -Credential (get-Credential) -Authentication Basic -AllowRedirection -erroraction stop
            Import-PSSession `$Session -AllowClobber -DisableNameChecking
        
            # Check to see if the search is still running, and wait until it's done before starting purge
            while ( (get-compliancesearch -identity "$straggler_name" | Select-Object -expand status) -ne "Completed" ) { Start-Sleep 5; write-host "." -NoNewline }
        
            # Start the purge of the found emails
            New-ComplianceSearchAction -SearchName "$straggler_name" -Purge -PurgeType SoftDelete
        
            # Wait until the purge is complete
            while ( (get-compliancesearchaction -identity ($straggler_name + "_purge") | Select-Object -expand status) -ne "Completed" ) { Start-Sleep 1; write-host "." -NoNewline }

            # Remove the session
            Remove-PSSession `$Session
            

        If, however, the subject is common (i.e. FYI, Invoice, FYSA) and there is a fear that a soft delete will delete non malicious emails:

        Please log in to https://protection.office.com go to Search -> Content Search

        Look for "$straggler_name"

        Run an export on the search and verify the emails.  At that point a delete action can be built to remove any straggler emails.

"@
      
        Send-MailMessage -from $esamtlEmail -to $esamtlAdmin -Subject "Desk hammer has completed, however please see below" -body $straggler_body -SmtpServer $smtprelay
               
# clean up and clear all variables

Remove-PSSession $Session
Get-Variable -Exclude Session,banner | Remove-Variable -EA 0
}
