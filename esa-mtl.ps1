#===================================++++=#
# ESA Message Tracking Logs via SMA API  #
# v1.0  --  June 2019                    #
#========================================#
#========================================#
# Global Definitions                     #
#========================================#
# Enter your SMA address and port
$esa = 'sma.whaterver.com:nnnn'
# Enter your auth connection string in base64 
# $b  = [System.Text.Encoding]::UTF8.GetBytes("json:somepassword")
# [System.Convert]::ToBase64String($b)
# anNvbjpzb21lcGFzc3dvcmQ=
$authstring = 'CHANGEME'
# SMTP relay for failure alerting
$smtprelay = "Internal email relay for error reporting"
$esamtlEmail = "From email for error reporting"
$esamtlAdmin = "Admin email that would fix issues with this script"
$taskname = "Enter name of scheduled task"
# Enter location to store the tracking log
$outfile = "c:\changme\pathto.csv"
$inceptionDate = (Get-Date).AddMinutes(-16)
# Enter location for the last log date text file
$lastLogDateFile = "c:\changme\last-log-date.txt"

$startdate = Get-Content -Path $lastLogDateFile
# Get current date/time in correct ESA API Format, sadly only available down to the minute not seconds.
$enddate = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:00.000Z")

#========================================#
# Functions                              #
#========================================#

#esatime for the message query which subtracts 5 minutes from last log file to ensure that any overlap is caught.
function esatime($passtime){
    $passtime = (($passtime.split("("))[0]).trim()
    $passtime = $passtime | Get-Date
    $passtime = $passtime.AddMinutes(-5)
    $passtime = $passtime.ToUniversalTime()
    $passtime = $passtime.ToString("yyyy-MM-ddTHH:mm:00.000Z")
    return $passtime
}
#datetime stamp for the csv for sorting
function mtltime($passtime){
    $passtime = (($passtime.split("("))[0]).trim()
    $passtime = $passtime | Get-Date
    $passtime = $passtime.ToString("yyyy-MM-ddTHH:mm:ss")
    return $passtime
}

# ================================================================================
# LOG ROTATION
# ================================================================================

# Log rotation script stolen from:
#      https://gallery.technet.microsoft.com/scriptcenter/PowerShell-Script-to-Roll-a96ec7d4
function Reset-Log 
{ 
    #function checks to see if file in question is larger than the paramater specified if it is it will roll a log and delete the oldes log if there are more than x logs. 
    param([string]$fileName, [int64]$filesize = 1mb , [int] $logcount = 5) 
     
    $logRollStatus = $true 
    if(test-path $filename) 
    { 
        $file = Get-ChildItem $filename 
        if((($file).length) -ige $filesize) #this starts the log roll 
        { 
            $fileDir = $file.Directory 
            $fn = $file.name #this gets the name of the file we started with 
            $files = Get-ChildItem $filedir | ?{$_.name -like "$fn*"} | Sort-Object lastwritetime 
            $filefullname = $file.fullname #this gets the fullname of the file we started with 
            #$logcount +=1 #add one to the count as the base file is one more than the count 
            for ($i = ($files.count); $i -gt 0; $i--) 
            {  
                #[int]$fileNumber = ($f).name.Trim($file.name) #gets the current number of the file we are on 
                $files = Get-ChildItem $filedir | ?{$_.name -like "$fn*"} | Sort-Object lastwritetime 
                $operatingFile = $files | ?{($_.name).trim($fn) -eq $i} 
                if ($operatingfile) 
                 {$operatingFilenumber = ($files | ?{($_.name).trim($fn) -eq $i}).name.trim($fn)} 
                else 
                {$operatingFilenumber = $null} 
 
                if(($operatingFilenumber -eq $null) -and ($i -ne 1) -and ($i -lt $logcount)) 
                { 
                    $operatingFilenumber = $i 
                    $newfilename = "$filefullname.$operatingFilenumber" 
                    $operatingFile = $files | ?{($_.name).trim($fn) -eq ($i-1)} 
                    write-host "moving to $newfilename" 
                    move-item ($operatingFile.FullName) -Destination $newfilename -Force 
                } 
                elseif($i -ge $logcount) 
                { 
                    if($operatingFilenumber -eq $null) 
                    {  
                        $operatingFilenumber = $i - 1 
                        $operatingFile = $files | ?{($_.name).trim($fn) -eq $operatingFilenumber} 
                        
                    } 
                    write-host "deleting " ($operatingFile.FullName) 
                    remove-item ($operatingFile.FullName) -Force 
                } 
                elseif($i -eq 1) 
                { 
                    $operatingFilenumber = 1 
                    $newfilename = "$filefullname.$operatingFilenumber" 
                    write-host "moving to $newfilename" 
                    move-item $filefullname -Destination $newfilename -Force 
                } 
                else 
                { 
                    $operatingFilenumber = $i +1  
                    $newfilename = "$filefullname.$operatingFilenumber" 
                    $operatingFile = $files | ?{($_.name).trim($fn) -eq ($i-1)} 
                    write-host "moving to $newfilename" 
                    move-item ($operatingFile.FullName) -Destination $newfilename -Force    
                } 
                     
            } 
 
                     
          } 
         else 
         { $logRollStatus = $false} 
    } 
    else 
    { 
        $logrollStatus = $false 
    } 
    $LogRollStatus 
}

#========================================#
# API Creation and Call                  #
#========================================#

# Create the header object with Auth for connecting to the ESA
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Authorization", "Basic $authstring" )
$headers.Add("Content-Type", 'application/json')
# Create the query given enddate of now and start date from last known log in the MTL (-5 minutes for the possible 4 minute lag)
$restquery = "http://$esa/sma/api/v2.0/message-tracking/messages?startDate=$startdate&endDate=$enddate&searchOption=messages&deliveryStatus=DELIVERED&messageDirection=incoming"

# Make the API Call and read it into an object
try {
$response = Invoke-RestMethod $restquery -Headers $headers -erroraction stop
}
catch {
    #if API call fails, send alert email and stop the script from running in scheduled tasks
    Send-MailMessage -from $esamtlEmail -to $esamtlAdmin -Subject "ESA MTL API call failed - fix and re-enable Scheduled task" -SmtpServer $smtprelay
    Disable-ScheduledTask -TaskName "ESA-MTL"
}
# Create a log for each email recipient
$Data = $response.data.attributes | ForEach-Object {
     foreach ($recipients in $_.recipient){
         [PSCustomObject]@{
             timestamp      = mtltime($_.timestamp)
             sender         = $_.sender
             senderreplyTo  = $_.senderreplyTo
             subject        = $_.subject
             recipients     = $recipients
         }
     }
 }

#========================================#
# MTL Sorting and Writing                #
#========================================#

 # Sort the logs by date
$messageTracesSorted = $Data | Sort-Object timestamp
# Append the logs to the MTL
$messageTracesSorted | Export-Csv $outfile -NoTypeInformation -Append
# Kludge - read back entire file to deduplicate given that the 5 minute overlap will create them
$fixit = Import-Csv $outfile | Sort-Object * -Unique | Sort-Object timestamp
# Write out the final logfile
$fixit | Export-Csv $outfile -NoTypeInformation
# Get the last log in the file and get the timestamp to write to a file to be read for next script execution
esatime(($messageTracesSorted | Select-Object -Last 1).timestamp) | Out-File -FilePath $lastLogDateFile -Force -NoNewline

#========================================#
# Cleanup                                #
#========================================#

#Check the logfile size and run the reset-log function if needed.
$traceSize = Get-Item $outfile
if ($traceSize.Length -gt 49MB ) {
    Start-Sleep -Seconds 30
    Reset-Log -fileName $traceLog -filesize 50mb -logcount 10
}

# Clear all the variables
Get-Variable -Exclude Session,banner | Remove-Variable -EA 0
