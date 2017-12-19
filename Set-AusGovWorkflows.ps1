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
        $oldqld = $oldsets | ? {$_.Name -like "*Queensland*"}
        $oldtas = $oldsets | ? {$_.Name -like "*Tasmani*"}
        $oldact = $oldsets | ? {$_.Name -like "*Australian*"}
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
        $output | Add-Member -type NoteProperty -Name OldSet -Value '<span style="background-color:#FF0000;">Empty or Unknown</span>'
        $output | Add-Member -type NoteProperty -Name NewSet -Value '<span style="background-color:#FF0000;">Check and assign manually</span>'
        }
    $outputarray += $output
}

$a = "<style>"
$a = $a + 'table{font-family: monospace;border-collapse:collapse;width:100%}td,th{border:1px solid #ddd;padding:8px}tr:nth-child(even){background-color:#f2f2f2}tr:hover{background-color:#ddd}th{padding-top:12px;padding-bottom:12px;text-align:left;background-color:#4CAF50;color:#fff}'
$a = $a + "</style>"

$htmlout = $outputarray | Sort-Object Workflow | ConvertTo-HTML -head $a
$htmlout -replace '&gt;','>' -replace '&lt;','<' -replace '&quot;','"' | Out-File .\AusGovWorkflows.htm
Invoke-Expression .\AusGovWorkflows.htm

