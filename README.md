# ESA-MTL
Powershell to get message tracking logs from Cisco ESA via SMA API

This Powershell will connect to a Security Management Appliance (SMA) and pull the email tracking from however many Email Security Appliances (ESA) are connected to it and save it to a .csv for use with https://github.com/LogRhythm-Labs/PIE instead of the default O365 Message Tracking.

Please see https://www.cisco.com/c/en/us/td/docs/security/esa/esa_all/esa_api/esa_api_12-0/b_ESA_API_Getting_Started_Guide_12-0/b_ESA_API_Getting_Started_Guide_chapter_00.html for information on the ESA/SMA API

See https://www.cisco.com/c/dam/en/us/td/docs/security/security_management/sma/sma12-0/AsyncOS-API-Addendum-GD_General_Deployment.xlsx for additional information on what APIs are available.

This script is inteaded to be run via a scheduled task every 5 minutes or so depending on environment.

This script can also be used in conjuction with the ESA-Desk.ps1 script to nuke an email from specific recipients inboxes without having to go through the O365 Audit search process instead using search-mailbox.
