#==================================++++++++++++#
# Desk Hammer                                  #
# See https://github.com/ciawhatsthat/ESA-MTL  #
# For deatils                                  #
# v1.0  --  June 2019                          #
#==============================================#

#==============================================#
# Global Definitions                           #
#==============================================#

# Name of folder containing reported phishing emails
$folderName = "CHANGME"
# Name of folder to move processed messages to
$completed = "CHANGME"
# Create the Outlook object to access mailboxes
$Outlook = New-Object -ComObject Outlook.Application;
$namespace = $Outlook.GetNameSpace("MAPI")
# Folder where phishing messaged will be moved to to be processed
$Folder = $namespace.Folders.Item(1).Folders('Inbox').Folders.Item($folderName)
#folder to move processed emails too
$MoveTarget = $namespace.Folders.Item(1).Folders('Inbox').Folders.Item($completed)
# Path to save attachments to
$filepath = "C:\CHANGEME\"
# Run through each email in the folder
for ($c = $Folder.Items.Count; $c -ge 1; $c--) {
    $email = $Folder.Items($c)
    # The number of email attachments
    $intCount =  $email.Attachments.Count
    # If the email has attachments, let's open the .msg email
    if($intCount -gt 0) {
        # Let's go through those attachments
        for($i=1; $i -le $intCount; $i++) {

            # The attachment being looked at
            $attachment =  $email.Attachments.Item($i)

            # If this is a .msg, let's open it
            if($attachment.FileName -like "*.msg"){
                $attachmentPath = $filepath+$attachment.FileName
                $attachment.SaveAsFile($attachmentPath)
                Get-ChildItem $attachmentPath |
                    ForEach-Object {
                        $msg = $Outlook.Session.OpenSharedItem($_.FullName)
                        $pspammer = $msg.SenderEmailAddress
                        $psubject = $msg.subject
                        #trim any internally added tags
                        $psubject = $psubject.Replace("[CHANGE REMOVE AS NEEDED]","")
                        $psubject = $psubject.Replace("[EXTERNAL]","")
                        $psubject = $psubject.Replace("[MARKETING]","")
                        $psubject = $psubject.Replace("[BULK]","")
                        $psubject = $psubject.trim()
                        #get the date the email was sent on
                        $pdate    = $msg.SentOn.ToString('M/d/y')
                        #release the email so PS can continue
                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($msg) | Out-Null
                        #send gathered info (sender and subject to the deskninja to search and then delete the emails)
                        deskninja -senton $pdate -sender $pspammer -subject $psubject
                    }
                    # delete the saved attachment from the temp location
                    Remove-Item $attachmentPath
            }
        }
    }
    # Move email to Completed folder after processing
    $email.UnRead = $false
    [void] $email.Move($MoveTarget) 
}
Get-Variable -Exclude Session,banner | Remove-Variable -EA 0

function deskninja {
    
    [cmdletbinding()]            
    param ( 
    [string]$sender,
    [string]$senton,
    [string]$subject
    )
# ================================================================================
# DEFINE Function PARAMETERS
# ================================================================================

# Mask errors
$ErrorActionPreference = "silentlycontinue"
$warningPreference = "silentlyContinue"

# Location of the MTL
$traceLog = "C:\CHANGEME\esa-mtl.csv"
# ================================================================================
# Location of the active users text file
# $list = @()
# $list += Get-ADUser -filter * -SearchBase 'OU=Where Active Users Live' -Properties proxyaddresses | select -ExpandProperty proxyaddresses
# # had to do some substing here to remove all x500 and ensure they contain the o365 address string @mydomain.onmicrosoft.com
# $list | Where-Object { $_.Substring(0,4) -ne "X500" } | Where-Object {!($_.Contains("@mydomain"))} | ForEach-Object {$_.substring(5)} | Set-Content -Path c:\path\to\file.csv -Force
# remove-variable list
#
# This is then scheduled to run every morning
# ================================================================================
$activeusers = Get-Content -path "C:\CHANGEME\activeusers.txt"

#see https://www.undocumented-features.com/2015/10/01/storing-powershell-credentials-in-the-local-user-registry/
$hkcupath = "HKCU:\Software\CHANGEME\Credentials\CHANGEME"

# SMTP relay and addresses for failure alerting
$smtprelay = "CHANGEME"
$esamtlEmail = "Change Me <changeme@changeme.com>"
$esamtlAdmin = "change@changeme.com"
$socMailbox = "phishing@changeme.com"


# ================================================================================
# Get list of users, compare to list of active users, create eDiscovery search
# ================================================================================

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

# ================================================================================
# Office 365 Protection Authentication
# ================================================================================

# Credentials stored in registry 
    
    try {
        $secureCredUserName = (Get-ItemProperty -Path $hkcupath).UserName
        $secureCredPassword = (Get-ItemProperty -Path $hkcupath).Password
        $securePassword = ConvertTo-SecureString $secureCredPassword
        $credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $secureCredUserName, $securePassword
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

# ================================================================================
# Create and run the eDsicovery
# ================================================================================
       
 
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

            New-ComplianceSearch -Name $straggler_name -ExchangeLocation all -ExchangeLocationExclusion $socMailbox -ContentMatchQuery $stragglerMatchQUery -erroraction stop
            
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
            while ( (get-compliancesearchaction -identity ("$straggler_name + '_purge'") | Select-Object -expand status) -ne "Completed" ) { Start-Sleep 1; write-host "." -NoNewline }

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
