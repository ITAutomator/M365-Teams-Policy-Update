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
# see if there's a more recent CSV with this naming scheme
$scriptCSV_latest = Get-ChildItem -Path $scriptCSV.Replace(".csv","*.csv") | Sort-Object LastWriteTime | Select-Object -Last 1 | Select-Object FullName -ExpandProperty FullName
if (-not $scriptCSV_latest)
{
    ######### Template
    "GroupNameOrEmail,GroupPolicyName,NonGroupPolicyName,GlobalExternalAccessDefault " | Add-Content $scriptCSV
    "Allow External TeamsMessaging,TeamsExternalAccess,FederationAndPICDefault,False" | Add-Content $scriptCSV
    ######### 
	$ErrOut=201; Write-Host "Err $ErrOut : Couldn't find '$(Split-Path $scriptCSV -leaf)'. Template CSV created. Edit CSV and run again.";Pause; Exit($ErrOut)
}
# ----------Fill $entries with contents of file or something
$entries=@(import-csv $scriptCSV_latest)
$entriescount = $entries.count
Write-Host "-----------------------------------------------------------------------------"
Write-Host ("$scriptName        Computer:$env:computername User:$env:username PSver:"+($PSVersionTable.PSVersion.Major))
Write-Host ""
Write-Host "Bulk actions in M365"
Write-Host "See Readme for more info."
Write-Host ""
Write-Host "CSV: $(Split-Path $scriptCSV_latest -leaf) ($($entriescount) entries)"
$entries | Format-Table
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
    # show orgwide settings
    Write-Host "-------------------- Get-CsTenantFederationConfiguration (Shows this org's master setting)"
    Write-Host "Adjust here: https://admin.teams.microsoft.com/company-wide-settings/external-communications"
    Write-Host "Note: If these settings block external access, no users can get around it via policies. In this case the script only stages policies."
    $tenant_config = Get-CsTenantFederationConfiguration
    $tenant_config | Format-Table
    # show current policies
    Write-Host "-------------------- Get-CsExternalAccessPolicy (Shows this org's policies)"
    $pol_curr = Get-CsExternalAccessPolicy
    $pol_curr | Select-Object Identity,EnableFederationAccess,EnableTeamsConsumerAccess | Format-Table | Out-String | Write-Host
    # get entries from CSV
    $x = $entries[0]
    # Check policy names
    $pol_name_group = $x.GroupPolicyName
    if ($x.NonGroupPolicyName -eq "Global") {
        $pol_name_nongroup = $null
    }
    elseif ($x.NonGroupPolicyName -eq "Blocked") {
        $pol_name_nongroup = "$($pol_name_group)Blocked"
    }
    else {
        $pol_name_nongroup = $x.NonGroupPolicyName
    }
    # Check policy settings
    foreach ($entrycount in 1..3)
    { # check policies: global, allow, block
        if ($entrycount -eq 1) {
            # the Global setting
            $pol_name = "Global"
            $pol_value = [System.Convert]::ToBoolean($x.GlobalExternalAccessDefault)
        }
        if ($entrycount -eq 2) {
            # the allow policy
            $pol_name = $x.GroupPolicyName
            $pol_value = $true
        }
        if ($entrycount -eq 3) {
            # the block policy
            $pol_name = "$($x.GroupPolicyName)Blocked"
            $pol_value = $false
        }
        # Retrieve policy
        $pol_curr = Get-CsExternalAccessPolicy -Identity $pol_name -ErrorAction Ignore
        # Create if missing
        if (-not $pol_curr)
        { # policy doesn't exist
            Write-Host "Policy '$($pol_name)' Not found" -ForegroundColor Red
            if (AskForChoice("Create this policy: $($pol_name)?")) {
                # Create a new policy for this purpose
                New-CsExternalAccessPolicy -Identity $pol_name -EnableFederationAccess $pol_value -EnableTeamsConsumerAccess $pol_value
            } # Created
            else {
                Write-Host "Policy creation aborted.";Start-sleep  3; Exit
            } # Aborted
        } # policy doesn't exist
        else
        { # policy exists
            Write-Host "Checking policy '$($pol_name)' ... " -NoNewline
            if (($pol_curr.EnableFederationAccess -eq $pol_value) -and ($pol_curr.EnableTeamsConsumerAccess -eq $pol_value))
            { # settings OK
                Write-Host $pol_value -NoNewline
                Write-Host " (Already OK)" -ForegroundColor Green
            } # settings OK
            else
            { # settings Bad
                Write-Host "Incorrect" -ForegroundColor Red
                Write-Host "   EnableFederationAccess: $($pol_curr.EnableFederationAccess)"
                Write-Host "EnableTeamsConsumerAccess: $($pol_curr.EnableTeamsConsumerAccess)"
                if (AskForChoice("Set these to $($x.GlobalExternalAccessDefault )?"))
                {
                    Set-CsExternalAccessPolicy -Identity $pol_name -EnableFederationAccess $pol_value
                    Set-CsExternalAccessPolicy -Identity $pol_name -EnableTeamsConsumerAccess $pol_value
                }
            } # settings Bad
        } # policy exists
    }  # check 2 policies: global plus named policy
    #region check group
    Write-Host "Checking Group '$($x.GroupNameOrEmail)'..." -NoNewline
    $group = Get-MgGroup -Filter "(mail eq '$($x.GroupNameOrEmail)') or (displayname eq '$($x.GroupNameOrEmail)')"
    if (-not $group) 
    { # group bad
        Write-Host "Group not found: $($x.GroupNameOrEmail) ERR"  -ForegroundColor Red
        Write-Host "Aborted: no group. Create a mail-enabled security group.";Start-sleep  3; Exit
    } # group bad
    else
    { # group ok
        Write-Host "Checking members ... " -NoNewline
        $grp_children = @(GroupChildren -DirectoryObjectId $group.Id -Recurse $false)
        Write-Host "$($grp_children.Count) Members" -ForegroundColor Green
    } # group ok
    #endregion Check Group
    foreach ($entrycount in 1..2)
    { # members, nonmembers
        if ($entrycount -eq 1) {
            $pol_action="allow"
            $entries = $grp_children
            $pol_name_test = $pol_name_group
        } # allow entries
        else {
            $pol_action="block"
            $pol_name_test = $pol_name_nongroup
            ####### Retrieve Azure AD User list
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
            $mgusers = $mgusers | Where-Object Id -NotIn ($grp_children.id)
            Write-Host "User Count: $($mgusers.count) [Not group members]"
            # sort
            $mgusers = $mgusers | Sort-Object DisplayName
            # filter for licensed
            $entries = @()
            ForEach ($mguser in $mgusers)
            {
                $licinfo = UserLicenseInfo $mguser.UserPrincipalName
                if ($licinfo.Count -gt 0)
                {
                    $entries+=$mguser
                }
            }
            Write-Host "User Count: $($entries.count) [Licensed]"
        } # block entries
        # entries
        $processed=0
        $choiceLoop=0
        $i=0
        $entriescount = $entries.Count
        foreach ($x in $entries)
        { # each entry
            $i++
            write-host "Pass $($entrycount) of 2 [$($pol_action) policy] ----- $($i) of $($entriescount): $($x.displayName)"
            if ($choiceLoop -ne 1)
            { # Process all not selected yet, Ask
                $choices = @("&Yes","Yes to &All","&No","No and E&xit") 
                $choiceLoop = AskforChoice -Message "Process entry $($i)?" -Choices $choices -DefaultChoice 1
            } # Process all not selected yet, Ask
            if (($choiceLoop -eq 0) -or ($choiceLoop -eq 1))
            { # Process
                $processed++
                #######
                ####### Start code for object $x
                #region Object X
                $user_pol = Get-CsUserPolicyAssignment -User $x.id -PolicyType ExternalAccessPolicy | Select-Object PolicyName -ExpandProperty PolicyName
                if ($null -eq $user_pol) {$user_pol = "<none>"}
                Write-host "ExternalAccessPolicy for $($x.displayName): $($user_pol) - " -NoNewline
                if (($user_pol -eq $pol_name_test) -or (($pol_action -eq "remove") -and $user_pol -eq "<none>"))
                {
                    Write-Host "Already OK" -ForegroundColor Green
                }
                else 
                {
                    Write-Host "Adjusted to $($pol_name_test)" -ForegroundColor Yellow
                    # Grant-CsExternalAccessPolicy -PolicyName $pol_name_test -Identity $x.Id
                    New-CsBatchPolicyAssignmentOperation -PolicyType ExternalAccessPolicy -PolicyName $pol_name_test -Identity $x.Id
                }
                #endregion Object X
                ####### End code for object $x
                #######
            } # Process
            if ($choiceLoop -eq 2)
            {
                write-host ("Entry "+$i+" skipped.")
            }
            if ($choiceLoop -eq 3)
            {
                write-host "Aborting."
                break
            }
        } # each entry
    } # members, nonmembers
    Write-Host "------------------------------------------------------------------------------------"
    #region   Adjust NonGroupMembers back to default
    #endregionAdjust NonGroupMembers back to default
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