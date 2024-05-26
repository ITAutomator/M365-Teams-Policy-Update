#####
## To enable scrips, Run powershell 'as admin' then type
## Set-ExecutionPolicy Unrestricted
#####
# Transcript Open
$Transcript = [System.IO.Path]::GetTempFileName()               
Start-Transcript -path $Transcript | Out-Null
# Main function header - Put ITAutomator.psm1 in same folder as script
$scriptFullname = $PSCommandPath ; if (!($scriptFullname)) {$scriptFullname =$MyInvocation.InvocationName }
$scriptXML      = $scriptFullname.Substring(0, $scriptFullname.LastIndexOf('.'))+ ".xml"  ### replace .ps1 with .xml
$scriptCSV      = $scriptFullname.Substring(0, $scriptFullname.LastIndexOf('.'))+ ".csv"  ### replace .ps1 with .csv
$scriptDir      = Split-Path -Path $scriptFullname -Parent
$scriptName     = Split-Path -Path $scriptFullname -Leaf
$scriptBase     = $scriptName.Substring(0, $scriptName.LastIndexOf('.'))
$psm1="$($scriptDir)\ITAutomator.psm1";if ((Test-Path $psm1)) {Import-Module $psm1 -Force} else {write-output "Err 99: Couldn't find '$(Split-Path $psm1 -Leaf)'";Start-Sleep -Seconds 10;Exit(99)}
$psm1="$($scriptDir)\ITAutomator M365.psm1";if ((Test-Path $psm1)) {Import-Module $psm1 -Force} else {write-output "Err 99: Couldn't find '$(Split-Path $psm1 -Leaf)'";Start-Sleep -Seconds 10;Exit(99)}
Write-Host "-----------------------------------------------------------------------------"
Write-Host ("$scriptName        Computer:$env:computername User:$env:username PSver:"+($PSVersionTable.PSVersion.Major))
Write-Host ""
Write-Host "Bulk actions in M365"
Write-Host ""
Write-Host ""
Write-Host "-----------------------------------------------------------------------------"
PressEnterToContinue
$no_errors = $true
$error_txt = ""
$results = @()
#region modules
<#
(prereqs: as admin run these powershell commands)
Install-Module Microsoft.Graph.Authentication
Install-Module Microsoft.Graph.Identity.DirectoryManagement
Install-Module Microsoft.Graph.Users
#>
$modules=@()
$modules+="Microsoft.Graph.Authentication"
$modules+="MicrosoftTeams"
ForEach ($module in $modules)
{ 
    Write-Host "Loadmodule $($module)..." -NoNewline ; $lm_result=LoadModule $module -checkver $false; Write-Host $lm_result
    if ($lm_result.startswith("ERR")) {
        Write-Host "ERR: Load-Module $($module) failed. Suggestion: Open PowerShell $($PSVersionTable.PSVersion.Major) as admin and run: Install-Module $($module)";Start-sleep  3; Exit
    }
}
#endregion modules
# Connect
$myscopes=@()
$myscopes+="User.ReadWrite.All"
$myscopes+="GroupMember.ReadWrite.All"
$myscopes+="Group.ReadWrite.All"
$connected_ok = ConnectMgGraph -myscopes $myscopes
$domain_mg = Get-MgDomain -ErrorAction Ignore| Where-object IsDefault -eq $True | Select-object -ExpandProperty Id
if (-not ($connected_ok))
{ # connect failed
    Write-Host "Connection failed."
}
else
{ # connect ok
    $connected_ok = ConnectMicrosoftTeams $domain_mg
    if (-not ($connected_ok))
    { # connect failed
        Write-Host "Connection failed.";Start-sleep  3; Exit
    }
    Write-Host "CONNECTED"
    Write-Host "-------------------- Get-CsExternalAccessPolicy (Shows this org's policies)"
    # show current policies
    $pol_curr = Get-CsExternalAccessPolicy
    $pol_curr | Select-Object Identity,EnableFederationAccess,EnableTeamsConsumerAccess | Format-Table | Out-String | Write-Host
    ####### Retrieve User list
    $mg_properties = @(
        'id'
        ,'UserPrincipalName'
        ,'AccountEnabled'
        ,'DisplayName'
        ,'mail'
        ,'UserType'
    )
    $mgusers = Get-MGuser -All -Property $mg_properties
    Write-Host "User Count: $($mgusers.count) [All users]"
    $mgusers = $mgusers | Where-Object UserType -EQ Member
    Write-Host "User Count: $($mgusers.count) [UserType=Members (vs Guests)]"
    $mgusers = $mgusers | Where-Object AccountEnabled -eq $true
    Write-Host "User Count: $($mgusers.count) [AccountEnabled=True]"
    # sort
    $entries = $mgusers | Sort-Object DisplayName
    ##### Retrieve User list
    $processed=0
    $i=0
    $entriescount = $entries.Count
    $DateSnap = get-date -format "yyyy-MM-dd"
    $rows = @()
    foreach ($x in $entries)
    { # each entry
        $i++
        $processed++
        $user_pol = Get-CsUserPolicyAssignment -User $x.id -PolicyType ExternalAccessPolicy | Select-Object PolicyName -ExpandProperty PolicyName
        if ($null -eq $user_pol) {$user_pol = "<none>"}
        Write-host "$($i) of $($entriescount): $($x.displayName), ExternalAccessPolicy: $($user_pol)"
        $row_obj=[pscustomobject][ordered]@{
            DateSnap             = $DateSnap
            DisplayName          = $x.DisplayName
            Mail                 = $x.Mail
            UserPrincipalName    = $x.UserPrincipalName
            Id                   = $x.Id
            ExternalAccessPolicy = $user_pol
        }
        $rows += $row_obj
    } # each entry
    Write-Host "------------------------------------------------------------------------------------"
    Write-host "Exporting info to CSV..."
    $date = get-date -format "yyyy-MM-dd_HH-mm-ss"
    $scriptCSVdated1= $scriptCSV.Replace(".csv"," $($date) Users.csv")
    $scriptCSVdated2= $scriptCSV.Replace(".csv"," $($date) Policies.csv")
    if ($PSVersionTable.PSVersion.Major -lt 7)
    { # ps 5 (Excel likes UTF8-Bom CSVs, PS5 defaults utf8 to BOM)
        $rows | Export-Csv $scriptCSVdated1 -NoTypeInformation -Encoding utf8
        $pol_curr | Export-Csv $scriptCSVdated2 -NoTypeInformation -Encoding utf8
    }
    else
    { # ps 7 (Excel likes UTF8-Bom CSVs, PS7 changed utf8 to be NOBOM, so use utf8BOM)
        $rows | Export-Csv $scriptCSVdated1 -NoTypeInformation -Encoding utf8BOM
        $pol_curr | Export-Csv $scriptCSVdated2 -NoTypeInformation -Encoding utf8BOM
    }
	#################### Transcript Save
    Stop-Transcript | Out-Null
    $date = get-date -format "yyyy-MM-dd_HH-mm-ss"
    New-Item -Path (Join-Path (Split-Path $scriptFullname -Parent) ("\Logs")) -ItemType Directory -Force | Out-Null #Make Logs folder
    $TranscriptTarget = Join-Path (Split-Path $scriptFullname -Parent) ("Logs\"+[System.IO.Path]::GetFileNameWithoutExtension($scriptFullname)+"_"+$date+"_log.txt")
    If (Test-Path $TranscriptTarget) {Remove-Item $TranscriptTarget -Force}
    Move-Item $Transcript $TranscriptTarget -Force
    #################### Transcript Save
} # connect ok
PressEnterToContinue