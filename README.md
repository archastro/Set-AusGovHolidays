# Set-AusGovHolidays
This script creates or updates - based on whether they exist - a Response Group Holiday set for each Australian state, based on data from the Australian Government website.

## NOTES
v1.5 - Migrate XML file to a host managed by me so that it can be modified, and update the script to use the additional fields. 
This is due to the fact that some holidays are marked as only running from 7pm until midnight, and one holiday (Easter Tuesday in Tasmania) is Government only.
As a bonus, this may allow the XML file to be extended to include additional jurisdictions (e.g. New Zealand) in future.
v1.4 - Rewrote main program loop and pulled some tasks out in to functions
V1.3 - Added the option to specify a poolFQDN for large deployments as the automatic detection script could have performance issues. Added -DisplayOutput option to show an Out-Gridview of the created holidays.
V1.2 - Script now detects date format (dd/MM/yyyy or MM/dd/yyyy) and rewrites the dates listed in the XML file to match the server's current date format.
V1.1 - Modify popups to be more descriptive, target only pools which host workflows with a +61 LineURI
V1.0 - Updated script tested and verified in a production environment
V0.9 - Initial script tested and verified in a lab environment

## DESCRIPTION
On first run, this script will:
    *Download a modified version of the [Australian Government's XML holiday list](http://www.australia.gov.au/about-australia/special-dates-and-events/public-holidays) from [http://code.wespeakbinary.com.au/holidays.html](http://code.wespeakbinary.com.au/holidays.html)
    *Create a holiday set for each state, for example, "NSW Holidays (AusGov)." AusGov is used as a marker to denote sets created by this script.
    *Create holidays in each set based on the content in the XML file
Once created, holiday sets can be assigned to Response Group workflows to handle holidays.
As the XML file offered by the Australian Government is updated periodically, this script is designed to update Holiday Sets with newly published information.
When run subsequently, either manually or via a scheduled task, the script will:
    *Download the latest copy of the Australian Government's holiday list
    *Clear the contents of the existing "AusGov" Holiday Sets
    *Populate the existing sets with the updated holidays
Following subsequent runs, Response Group Workflows do not need to be updated as the names of the Holiday Sets does not change.

## Parameters
`-Action [Create/Update]`
    Declare whether the script should create new sets or Update existing sets
`-PoolFQDN [String] `
    Specify a pool on which to update holiday sets
`-XMLFile [File location] `
    If your server does not have access to download the XML file, you may provide it here
`-HideOutput [switch] `
    Once complete, the script will display a new IE window with a list of all created holidays unless this switch is invoked
`-MailOutput [switch] `
    Once complete, the script will mail an HTML file with holiday set details if this switch is invoked. Requires -notificationemail and -smtpserver for sending
`-NotificationEmail `
    When -Mailoutput is selected, email address to which the script will send its HTML output
`-SMTPServer `
    When -MailOutput is selected, this unauthenticated SMTP Relay will be used to send the email
    
## Examples
#### EXAMPLE
`.\Set-AusGovHolidays.ps1 -Action Create`
Detect pools where Skype for Business hosts Australian Response Group Workflows and Create new sets
#### EXAMPLE
`.\Set-AusGovHolidays.ps1 -Action Create -poolFQDN pool.contoso.com`
Create new holiday sets on a specified pool
#### EXAMPLE
`.\Set-AusGovHolidays.ps1 -Action Update -poolFQDN pool.contoso.com -XMLFile c:\file.xml`
Update existing holiday sets on a specified pool and use an offline XML file
#### EXAMPLE
`.\Set-AusGovHolidays.ps1 -Action Update -HideOutput -Mailoutput -NotificationEmail leigh@wespeakbinary.com.au -SMTPServer relay.wespeakbinary.com.au`
Update existing holiday sets on all auto-detected pools and send output via email to leigh@wespeakbinary.com.au
### INPUTS
The script does not support piped input
### OUTPUTS
The script produces a grid view output of the currently deployed holiday sets, including any custom (non-AusGov) sets
### LINK
[http://wespeakbinary.com.au](http://wespeakbinary.com.au)
