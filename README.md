# ESA-MTL
Powershell to get message tracking logs from Cisco ESA via SMA API

This Powershell will connect to a Security Management Appliance (SMA) and pull the email tracking from however many Email Security Appliances (ESA) are connected to it and save it to a .csv for use with https://github.com/LogRhythm-Labs/PIE instead of the default O365 Message Tracking.

Please see https://www.cisco.com/c/en/us/td/docs/security/esa/esa_all/esa_api/esa_api_12-0/b_ESA_API_Getting_Started_Guide_12-0/b_ESA_API_Getting_Started_Guide_chapter_00.html for information on the ESA/SMA API

See https://www.cisco.com/c/dam/en/us/td/docs/security/security_management/sma/sma12-0/AsyncOS-API-Addendum-GD_General_Deployment.xlsx for additional information on what APIs are available.

This script is intended to be run via a scheduled task every 5 minutes or so depending on environment.

This script can also be used in conjunction with the deskhammer.ps1 script to nuke an email from specific recipient’s inboxes 

# Desk Hammer
Powershell script to stop the bleeding from a phishing email
This script is loaded on the profile of the soc mailbox (eg phishing@whatever.com) it connects to outlook where 2 Inbox folders need to exist; SERVICEDESK (or whatever) and COMPLETED.

The script will check for any emails in the SERVICEDESK folder, search the MTL for any recipients that received the email, generate and start an eDisocvery search with https://protection.office.com with a scope of the just the recipients found to speed up the process.  Then run the purge soft delete from the results of the eDiscovery.  This essentially stops the bleeding of a phishing email, assuming that the SOC moves the email in the right folder. 

It will then start another eDiscovery search with a full scope to search all of exchange looking for any stragglers, eg forwarded to other internal users and send a copy/paste script for an analyst to run that will delete any straggler messages found.

Some caveats:
The profile must be logged in, as the scheduled task running from session 0 (ie the user isn't logged in) cannot access the outlook folder
The task should be scheduled to run every minute, but set not start unless the previous iteration has completed.
Test that wherever this is setup it can use internal relay for error reporting

# Future
Instead of having to have a local outlook installed and the user needing to be logged in, connect to o365 inbox directly via powershell.  This should also fix the issue of needing the user to be logged in.
