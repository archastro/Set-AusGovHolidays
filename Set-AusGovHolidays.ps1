<#
.SYNOPSIS
This script creates or updates - based on whether they exist - a Response Group Holiday set for each Australian state, based on data from the Australian Government website.

.NOTES
v1.5 - Migrate XML file to a host managed by me so that it can be modified, and update the script to use the additional fields. 
This is due to the fact that some holidays are marked as only running from 7pm until midnight, and one holiday (Easter Tuesday in Tasmania) is Government only.
As a bonus, this may allow the XML file to be extended to include additional jurisdictions (e.g. New Zealand) in future.

v1.4 - Rewrote main program loop and pulled some tasks out in to functions

V1.3 - Added the option to specify a poolFQDN for large deployments as the automatic detection script could have performance issues. Added -DisplayOutput option to show an Out-Gridview of the created holidays.

V1.2 - Script now detects date format (dd/MM/yyyy or MM/dd/yyyy) and rewrites the dates listed in the XML file to match the server's current date format.

V1.1 - Modify popups to be more descriptive, target only pools which host workflows with a +61 LineURI

V1.0 - Updated script tested and verified in a production environment

V0.9 - Initial script tested and verified in a lab environment

.DESCRIPTION
On first run, this script will:
    •Download a modified version of the Australian Government's XML holiday list (http://www.australia.gov.au/about-australia/special-dates-and-events/public-holidays) from http://code.wespeakbinary.com.au/holidays.html
    •Create a holiday set for each state, for example, "NSW Holidays (AusGov)." AusGov is used as a marker to denote sets created by this script.
    •Create holidays in each set based on the content in the XML file

Once created, holiday sets can be assigned to Response Group workflows to handle holidays.

As the XML file offered by the Australian Government is updated periodically, this script is designed to update Holiday Sets with newly published information.

When run subsequently, either manually or via a scheduled task, the script will:
    •Download the latest copy of the Australian Government's holiday list
    •Clear the contents of the existing "AusGov" Holiday Sets
    •Populate the existing sets with the updated holidays

Following subsequent runs, Response Group Workflows do not need to be updated as the names of the Holiday Sets does not change.

.PARAMETER -Action [Create/Update]
    Declare whether the script should create new sets or Update existing sets

.PARAMETER -PoolFQDN [String] 
    Specify a pool on which to update holiday sets

.PARAMETER -XMLFile [File location] 
    If your server does not have access to download the XML file, you may provide it here

.PARAMETER -HideOutput [switch] 
    Once complete, the script will display a new IE window with a list of all created holidays unless this switch is invoked

.PARAMETER -MailOutput [switch] 
    Once complete, the script will mail an HTML file with holiday set details if this switch is invoked. Requires -notificationemail and -smtpserver for sending

.PARAMETER -NotificationEmail 
    When -Mailoutput is selected, email address to which the script will send its HTML output

.PARAMETER -SMTPServer 
    When -MailOutput is selected, this unauthenticated SMTP Relay will be used to send the email

.EXAMPLE
.\Set-AusGovHolidays.ps1 -Action Create
Detect pools where Skype for Business hosts Australian Response Group Workflows and Create new sets

.EXAMPLE
.\Set-AusGovHolidays.ps1 -Action Create -poolFQDN pool.contoso.com
Create new holiday sets on a specified pool

.EXAMPLE
.\Set-AusGovHolidays.ps1 -Action Update -poolFQDN pool.contoso.com -XMLFile c:\file.xml
Update existing holiday sets on a specified pool and use an offline XML file

.EXAMPLE
.\Set-AusGovHolidays.ps1 -Action Update -HideOutput -Mailoutput -NotificationEmail leigh@wespeakbinary.com.au -SMTPServer relay.wespeakbinary.com.au
Update existing holiday sets on all auto-detected pools and send output via email to leigh@wespeakbinary.com.au

.INPUTS
The script does not support piped input

.OUTPUTS
The script produces a grid view output of the currently deployed holiday sets, including any custom (non-AusGov) sets

.LINK
http://wespeakbinary.com.au

#>
[cmdletbinding()]
param (
    [Parameter(Mandatory=$true,position=0)]
    [ValidateSet('Create','Update')]
    [string]$Action,
    [Parameter(Mandatory=$false,Position=1)]
    [string]$poolFQDN,
    [Parameter(Mandatory=$false,Position=2)]
    [string]$XMLFile,
    [Parameter(Mandatory=$false,Position=3)]
    [switch]$HideOutput,
    [switch]$mailoutput,
    [string]$notificationemail,
    [string]$smtpserver
)

# If $poolFQDN was provided, use that. Otherwise, detect pools where Australian response groups exist

if ($poolFQDN -eq [String]::Empty) {
    Write-Host "No poolFQDN specified. Detecting pools which host Australian response groups." -ForegroundColor Yellow
    # Find all pools which host Response Groups
    $workflows = Get-CsApplicationEndpoint -Filter {Lineuri -like "*+61*" -and OwnerURN -like "*RGS"}
        if (($workflows | measure).count -eq 0) {
        throw "No Response Groups have been defined in this environment. Please create some."
        } else {
        $pools = $workflows.Registrarpool
        $pools = $pools.friendlyname | select -Unique
        }
    } else {
    write-host "Proceeding with user-specified pool $poolFQDN." -ForegroundColor Green
    $pools = $poolFQDN
    }


# Check to see if the user has provided an XML file, or if we need to download it

if ($xmlfile -eq [String]::Empty) {
#    if ($xmlfile -eq $null) {
    # Download and import the latest version of the Australian Government's holiday list
    Write-Host "No user specified XML file provided. Downloading the latest version of the Australian Government's holiday list" -ForegroundColor Yellow
    [xml]$XmlFile = (New-Object System.Net.WebClient).DownloadString("http://code.wespeakbinary.com.au/holidays.xml")
    # Test to make sure the XML file is valid
    if ($xmlfile -ne $null) {
        write-host "XML File downloaded successfully. Proceeding." -ForegroundColor Green
        } else {
        throw "XML File Invalid. Please check http://www.australia.gov.au/about-australia/special-dates-and-events/public-holidays."
        }
    } else {
    write-host "Proceeding with user-specified XML file $xmlfile" -ForegroundColor Yellow
    [xml]$xmlfile = Get-Content -Path $XMLFile
    }


# Check the date format of the current server

$international = Get-ItemProperty -Path "HKCU:\Control Panel\International"
If ($international.sShortDate -like "*d/M*") {
    $dateformat = "dd/MM/yyyy"
    Write-Host "Current Server Date format is Australian" -ForegroundColor Green
    } elseif ($international.sShortDate -like "*M/d*") {
    $dateformat = "MM/dd/yyyy"
    Write-Host "Current Server Data format is American" -ForegroundColor Green
    } else {
    throw "Date Format Invalid. Please set the ShortDate format to either dd/MM/yyyy or MM/dd/yyyy"
    }


# Function to check if AusGov holiday sets already exist
$allsets = Get-CsRgsHolidaySet
#function Check-Sets {
#    foreach ($pool in $pools) {
#        ($allsets | Where-Object {$_.Name -like "*AusGov*"}| measure).count -ne 0)
#    }
#}

# Function to create new holiday sets if they don't exist
function Create-Holidays {
    foreach ($pool in $pools) {
        Write-host "Creating holidays for $pool" -ForegroundColor Green
        foreach ($jurisdiction in $xmlfile.ausgovEvents.jurisdiction) {
            $holidays = @()
            $state = $jurisdiction.jurisdictionName
            $events = $jurisdiction.events.event

            foreach ($event in $events) {
                $id = $event.rawdate
                $date = ([datetime]::parseexact($event.date,"dd/MM/yyyy",[System.Globalization.CultureInfo]::InvariantCulture)).ToString($dateformat)
                $enddate = ([datetime]::parseexact($event.date,"dd/MM/yyyy",[System.Globalization.CultureInfo]::InvariantCulture)).AddDays(1).ToString($dateformat)
                $starttime = $event.starttime
                $endtime = $event.endtime
                $HolidayName = [regex]::replace($event.holidaytitle, "`t", "")
                $name = $event.jurisdiction.toupper() + " " + $HolidayName + " " + $event.year
                Write-Host "`t`tAdding $name ($date $starttime to $enddate $endtime)" -ForegroundColor DarkGreen
                $i = new-csrgsholiday -startdate "$date $starttime" -enddate "$date $endtime" -name "$name"
                $holidays += $i
            }
        # Create the new set
        Write-Host "Writing set $state Holidays (AusGov) to $pool" -ForegroundColor Green
        New-CsRgsHolidaySet -Parent $pool -Name "$state Holidays (AusGov)" -HolidayList($holidays) -Verbose
        }

    }
}

# Function to update existing holiday sets
function Update-Holidays {
    foreach ($pool in $pools) {
        Write-Host "Updating holidays for $pool" -ForegroundColor Green
        foreach ($jurisdiction in $xmlfile.ausgovEvents.jurisdiction) {
            $holidays = @()
            $state = $jurisdiction.jurisdictionName
            $events = $jurisdiction.events.event
            $set = Get-CsRgsHolidaySet -Name "$state Holidays (AusGov)"
            $setname = $set.name

        # Clear existing holidays
            Write-Host "`n`tClearing existing holidays from $setname on pool $pool" -ForegroundColor Green
            $Set.HolidayList.Clear()
            Set-CsRgsHolidaySet -Instance $set
            Write-host "`t`t$setname is now empty." -ForegroundColor Yellow

        # Add new events to holiday array
            Write-Host "`n`tAdding new Holidays to $setname on pool $pool" -ForegroundColor Green
            foreach ($event in $events) {
                $id = $event.rawdate
                $date = ([datetime]::parseexact($event.date,"dd/MM/yyyy",[System.Globalization.CultureInfo]::InvariantCulture)).ToString($dateformat)
                $enddate = ([datetime]::parseexact($event.date,"dd/MM/yyyy",[System.Globalization.CultureInfo]::InvariantCulture)).AddDays(1).ToString($dateformat)
                $starttime = $event.starttime
                $endtime = $event.endtime
                $HolidayName = [regex]::replace($event.holidaytitle, "`t", "")
                $name = $event.jurisdiction.toupper() + " " + $HolidayName + " " + $event.year
                Write-Host "`t`t>>$name ($date $starttime to $enddate $endtime), " -ForegroundColor DarkGreen
                $i = new-csrgsholiday -startdate "$date $starttime" -enddate "$enddate $endtime" -name "$name"
                $holidays += $i
            }
        
        # Add the holidays to the set
            foreach ($holiday in $holidays) {
                $set.HolidayList.Add($holiday)
            }
            Set-CsRgsHolidaySet -Instance $set
        }
        Write-Host "Holidays completed for $pool" -ForegroundColor Green
    }
Write-Host "All pools completed" -ForegroundColor Green
}

$logoutput | ConvertTo-Html | out-file .\logoutput.html
#Invoke-Expression .\logoutput.html

# Function to display sets in a pretty grid output

function Display-Holidays {
    Foreach ($set in Get-CsRgsHolidaySet) {
        Foreach ($holiday in $Set.HolidayList){
            [pscustomobject] @{
            Pool = $set.OwnerPool
            Set = $set.Name
            Name = $holiday.Name
            StartDate = $holiday.StartDate
            EndDate = $holiday.EndDate
            }
        }
    } 
}


# Construct HTML components for display/email
if ($mailoutput -eq $true or $Hideoutput -eq $false) {
    $a = "<style>"
    $a = $a + "BODY{background-color:#ffffff;}"
    $a = $a + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
    $a = $a + "TH{border-width: 1px;padding: 1px;border-style: solid;border-color: black;}"
    $a = $a + "TD{border-width: 1px;padding: 5px;border-style: solid;border-color: black;}"
    $a = $a + "</style>"
}

# Main program loop

if ($Action.Equals('Update') -eq $true) {
        Update-Holidays
        if ($HideOutput -ne $true) {
            Display-Holidays | Sort-Object StartDate | ConvertTo-HTML -head $a | Out-File .\AusGovHolidays.htm
            Invoke-Expression .\AusGovHolidays.htm
        }
    } else {
        Create-Holidays
            Display-Holidays | Sort-Object StartDate | ConvertTo-HTML -head $a | Out-File .\AusGovHolidays.htm
            Invoke-Expression .\AusGovHolidays.htm
        }

if ($mailoutput -eq $true) {
            $mailfrom = "holidays@" + (Get-CsSipDomain | ? {$_.IsDefault -eq $true}).identity
            $body = get-content .\AusGovHolidays.htm -raw
            Send-MailMessage -To $notificationemail -From $mailfrom -SmtpServer $smtpserver `
            -Body $body -BodyAsHtml -Subject "Set-AusGovHolidays Completion Notification"
}
