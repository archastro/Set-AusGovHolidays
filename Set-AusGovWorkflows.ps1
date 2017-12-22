<#
.SYNOPSIS
This script updates existing workflows with new holiday sets based on Set-AusGovHolidays.ps1
.NOTES
V1.0 - Initial script tested and verified in a lab environment
.DESCRIPTION
Looks for response group workflows in your environment which have an existing name of "New South Wales," for example, and replaces
the holiday set assigned by that workflow with "NSW Public Holidays (AusGov)" as per the Set-AusGovHolidays.ps1 script

.PARAMETERS None
    At this stage, the script has no parameters

.EXAMPLE
.\Set-AusGovWorkflows.ps1
Update workflows with new sets

.INPUTS
The script does not support piped input
.OUTPUTS
The script produces a grid view output of what's been deployed in HTML format. Failures are highlighted to make them easy to spot.
.LINK
http://wespeakbinary.com.au
#>
$flows = Get-CsRgsWorkflow
$flows = $flows | Sort-Object Name
$sets = Get-CsRgsHolidaySet
    $newsets = $sets | ? {$_.Name -like "*AusGov*"} 
        $newnsw = $newsets | ? {$_.Name -like "*NSW*"}
        $newvic = $newsets | ? {$_.Name -like "*VIC*"}
        $newqld = $newsets | ? {$_.Name -like "*QLD*"}
        $newtas = $newsets | ? {$_.Name -like "*TAS*"}
        $newact = $newsets | ? {$_.Name -like "*ACT*"}
        $newnt = $newsets | ? {$_.Name -like "*NT*"}
        $newsa = $newsets | ? {$_.Name -like "*SA*"}
        $newwa = $newsets | ? {$_.Name -like "*WA*"}

    $oldsets = $sets | ? {$_.Name -notlike "*AusGov*"} 
        $oldnsw = $oldsets | ? {$_.Name -like "*New South Wales*"}
        $oldvic = $oldsets | ? {$_.Name -like "*Victoria*"}
        $oldqld = $oldsets | ? {$_.Name -like "2017-Queensland"}
        $oldtas = $oldsets | ? {$_.Name -like "*Tasmani*"}
        $oldact = $oldsets | ? {$_.Name -like "*ACT*"}
        $oldnt = $oldsets | ? {$_.Name -like "*Northern*"}
        $oldsa = $oldsets | ? {$_.Name -like "*South Australia*"}
        $oldwa = $oldsets | ? {$_.Name -like "*Western*"}

$outputarray = @()


function update-output {
    $output | Add-Member -Type NoteProperty -Name Result -Value "Success"
    $output | Add-Member -type NoteProperty -Name Workflow -Value $flow.name
    $output | Add-Member -type NoteProperty -Name SipAddress -Value $flow.primaryuri
    $output | Add-Member -type NoteProperty -Name LineURI -Value $flow.lineuri
    $output | Add-Member -type NoteProperty -Name OldSet -Value $old.name
    $output | Add-Member -type NoteProperty -Name NewSet -Value $new.name
}

function update-workflow {
    $flow.HolidaySetIDList.remove($old.Identity) | Out-Null
    $flow.HolidaySetIDList.Add($new.identity)
    Set-CsRgsWorkflow -Instance $flow
}

foreach ($flow in $flows) {
    $output = $null
    $output = New-Object system.object
    if ($flow.HolidaySetIDList -like $oldnsw.Identity) {
        $old = $oldnsw
        $new = $newnsw
        update-workflow
        update-output

    } elseif ($flow.HolidaySetIDList -like $oldvic.Identity) {
        $old = $oldvic
        $new = $newvic
        update-workflow
        update-output

    } elseif ($flow.HolidaySetIDList -like $oldqld.Identity) {
        $old = $oldqld
        $new = $newqld
        update-workflow
        update-output

    } elseif ($flow.HolidaySetIDList -like $oldact.Identity) {
        $old = $oldact
        $new = $newact
        update-workflow
        update-output

    } elseif ($flow.HolidaySetIDList -like $oldtas.Identity) {
        $old = $oldtas
        $new = $newtas
        update-workflow
        update-output

    } elseif ($flow.HolidaySetIDList -like $oldsa.Identity) {
        $old = $oldsa
        $new = $newsa
        update-workflow
        update-output

    } elseif ($flow.HolidaySetIDList -like $oldwa.Identity) {
        $old = $oldwa
        $new = $newwa
        update-workflow
        update-output

    } elseif ($flow.HolidaySetIDList -like $oldnt.Identity) {
        $old = $oldnt
        $new = $newnt
        update-workflow
        update-output

    } else {
        Write-Host $flow.name " has an unknown or empty holiday set. Please check!" -ForegroundColor DarkRed
        $output | Add-Member -Type NoteProperty -Name Result -Value '<span style="background-color:#FF0000;">ERROR</span>'
        $output | Add-Member -type NoteProperty -Name Workflow -Value $flow.name
        $output | Add-Member -type NoteProperty -Name SipAddress -Value $flow.primaryuri
        $output | Add-Member -type NoteProperty -Name LineURI -Value $flow.lineuri
        $output | Add-Member -type NoteProperty -Name OldSet -Value '<span style="background-color:#FF0000;">Empty</span>'
        $output | Add-Member -type NoteProperty -Name NewSet -Value '<span style="background-color:#FF0000;">No action required</span>'
        }
    $outputarray += $output
}

$a = "<style>"
$a = $a + 'table{font-family: monospace;border-collapse:collapse;width:100%}td,th{border:1px solid #ddd;padding:8px}tr:nth-child(even){background-color:#f2f2f2}tr:hover{background-color:#ddd}th{padding-top:12px;padding-bottom:12px;text-align:left;background-color:#4CAF50;color:#fff}'
$a = $a + "</style>"

$htmlout = $outputarray | Sort-Object Workflow | ConvertTo-HTML -head $a
$htmlout -replace '&gt;','>' -replace '&lt;','<' -replace '&quot;','"' | Out-File .\AusGovWorkflows.htm
Invoke-Expression .\AusGovWorkflows.htm

$finalsets = get-csrgsholidayset | ? {$_.Name -notlike "*Ausgov*"}
#$removesets = @()
$removesets = [System.Collections.ArrayList]($finalsets.name)
$flows = Get-CsRgsWorkflow
Write-Host "The following old sets exist in the environment:" -ForegroundColor Yellow
$removesets
Write-Host "Testing if any of these sets are still in use:" -ForegroundColor Yellow
foreach ($finalset in $finalsets) {
    foreach ($flow in $flows) {
        if ($flow.HolidaySetIDList -contains $finalset.identity) {
            Write-Host "`t" $flow.name "contains" $finalset.name -ForegroundColor Green
            $removesets.remove($finalset.name)
            }
    }
}

Write-Host "Deleting unused old sets:" -ForegroundColor Yellow
$removesets | fl name

foreach ($set in $removesets) {
    get-csrgsholidayset -name $set | Remove-CsRgsHolidaySet
    }
