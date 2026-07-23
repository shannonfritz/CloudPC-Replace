#Requires -Version 5.1

using module .\CloudPCReplace.psm1

<#
.SYNOPSIS
    Cloud PC Replace WPF GUI - Modern UI for replacing Windows 365 Cloud PCs

.PARAMETER MockMode
    Runs the tool with simulated data instead of real Graph API calls.
    Auto-starts a self-driving demo on launch.

    NOTE: Mock mode is a BEST-EFFORT DEMO/development aid, not a fully
    supported code path. Its fake data (e.g. Get-ProvisioningPoliciesForGroup
    returning the same policy for every group) can trip validation in the later
    swap stages. Mock is fully isolated in the '#region Module API Overrides for
    Mock Mode' block below (it swaps leaf Graph functions at startup); the
    production module CloudPCReplace.psm1 contains no mock logic, so mock quirks
    never affect real runs. Fix on a best-effort basis only.

.NOTES
    Author: Cloud PC Replace Tool
    Version: 5.1.1
    Date: 2026-07-23
    Requires: CloudPCReplace.psm1 in same directory

.EXAMPLE
    .\Start-CloudPCReplaceWPF.ps1
    Runs normally, connects to a real tenant.

.EXAMPLE
    .\Start-CloudPCReplaceWPF.ps1 -MockMode
    Runs with simulated data for testing/demo purposes.
#>
param(
    [switch]$MockMode
)

$script:MockMode = $MockMode.IsPresent

$script:ToolVersion = "5.1.1"

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms   # For SaveFileDialog

$modulePath = Join-Path $PSScriptRoot "CloudPCReplace.psm1"
Import-Module $modulePath -Force

#region Global State
$script:cancellationToken  = $false
$script:replaceRunning     = $false
$script:pausing            = $false
$script:sourceGroupId      = $null
$script:sourceGroupName    = $null
$script:targetGroupId      = $null
$script:targetGroupName    = $null
$script:allSourceUsers     = @()
$script:groupMemberCache   = @{}
$script:userStates         = @{}
$script:maxConcurrent      = 2
$script:verboseLogging     = $false
$script:LogFilePath        = $null
$script:allPolicyGroups    = @()
$script:scheduledStartTime = $null
$script:scheduledStopTime  = $null
$script:scheduleArmed      = $false
$script:lastKeepAlive      = $null

$script:gracePeriodTimeoutMinutes        = 15
$script:endingGracePeriodTimeoutMinutes  = 30
$script:deprovisionTimeoutMinutes        = 60
$script:provisioningTimeoutMinutes       = 90
$script:mockPollInterval                 = 10   # Seconds between polls in MOCK MODE only (production uses 60s / 180s)
#endregion

#region Mock Data
$script:MockGroups = @(
    @{ id = 'aaaaaaaa-0001-0001-0001-aaaaaaaaaaaa'; displayName = 'w365 CPC in US Central';   onPremisesSyncEnabled = $false; groupTypes = @(); policyName = 'W365 Enterprise 2vCPU/8GB/256GB US' }
    @{ id = 'aaaaaaaa-0002-0002-0002-aaaaaaaaaaaa'; displayName = 'w365 CPC in US East';      onPremisesSyncEnabled = $false; groupTypes = @(); policyName = 'W365 Enterprise 2vCPU/8GB/256GB US' }
    @{ id = 'aaaaaaaa-0003-0003-0003-aaaaaaaaaaaa'; displayName = 'w365 CPC in US West';      onPremisesSyncEnabled = $false; groupTypes = @(); policyName = 'W365 Enterprise 4vCPU/16GB/256GB US' }
    @{ id = 'aaaaaaaa-0004-0004-0004-aaaaaaaaaaaa'; displayName = 'w365 CPC in EU West';      onPremisesSyncEnabled = $false; groupTypes = @(); policyName = 'W365 Enterprise 2vCPU/8GB/256GB EU' }
    @{ id = 'aaaaaaaa-0005-0005-0005-aaaaaaaaaaaa'; displayName = 'w365 CPC Hybrid Joined';   onPremisesSyncEnabled = $true;  groupTypes = @(); policyName = 'W365 Enterprise Hybrid 2vCPU/8GB' }
    @{ id = 'aaaaaaaa-0006-0006-0006-aaaaaaaaaaaa'; displayName = 'w365 CPC Dynamic Test';    onPremisesSyncEnabled = $false; groupTypes = @('DynamicMembership'); policyName = 'W365 Enterprise 2vCPU/8GB/256GB US' }
)

$script:MockUsers = @(
    @{ id = 'u001'; userPrincipalName = 'AdeleV@M365x.onmicrosoft.com';     displayName = 'Adele Vance' }
    @{ id = 'u002'; userPrincipalName = 'AlexW@M365x.onmicrosoft.com';      displayName = 'Alex Wilber' }
    @{ id = 'u003'; userPrincipalName = 'BrianJ@M365x.onmicrosoft.com';     displayName = 'Brian Johnson' }
    @{ id = 'u004'; userPrincipalName = 'ChristieC@M365x.onmicrosoft.com';  displayName = 'Christie Cline' }
    @{ id = 'u005'; userPrincipalName = 'DebraB@M365x.onmicrosoft.com';     displayName = 'Debra Berger' }
    @{ id = 'u006'; userPrincipalName = 'GradyA@M365x.onmicrosoft.com';     displayName = 'Grady Archie' }
    @{ id = 'u007'; userPrincipalName = 'HenriettaM@M365x.onmicrosoft.com'; displayName = 'Henrietta Mueller' }
    @{ id = 'u008'; userPrincipalName = 'IsaiahL@M365x.onmicrosoft.com';    displayName = 'Isaiah Langer' }
    @{ id = 'u009'; userPrincipalName = 'JohannaL@M365x.onmicrosoft.com';   displayName = 'Johanna Lorenz' }
    @{ id = 'u010'; userPrincipalName = 'JoniS@M365x.onmicrosoft.com';      displayName = 'Joni Sherman' }
    @{ id = 'u011'; userPrincipalName = 'LeeG@M365x.onmicrosoft.com';       displayName = 'Lee Gu' }
    @{ id = 'u012'; userPrincipalName = 'LidiaH@M365x.onmicrosoft.com';     displayName = 'Lidia Holloway' }
    @{ id = 'u013'; userPrincipalName = 'LynneR@M365x.onmicrosoft.com';     displayName = 'Lynne Robbins' }
    @{ id = 'u014'; userPrincipalName = 'MeganB@M365x.onmicrosoft.com';     displayName = 'Megan Bowen' }
    @{ id = 'u015'; userPrincipalName = 'MiriamG@M365x.onmicrosoft.com';    displayName = 'Miriam Graham' }
    @{ id = 'u016'; userPrincipalName = 'NestorW@M365x.onmicrosoft.com';    displayName = 'Nestor Wilke' }
    @{ id = 'u017'; userPrincipalName = 'PattiF@M365x.onmicrosoft.com';     displayName = 'Patti Fernandez' }
    @{ id = 'u018'; userPrincipalName = 'PradeepG@M365x.onmicrosoft.com';   displayName = 'Pradeep Gupta' }
    @{ id = 'u019'; userPrincipalName = 'SallyR@M365x.onmicrosoft.com';     displayName = 'Sally Reyes' }
    @{ id = 'u020'; userPrincipalName = 'TamaraB@M365x.onmicrosoft.com';    displayName = 'Tamara Bryan' }
)

# Tracks mock CPC state per user during demo
$script:MockCPCState = @{}

function Get-MockCPCForUser {
    param([string]$UserId)
    if (-not $script:MockCPCState.ContainsKey($UserId)) {
        # Initial state - user has a provisioned CPC
        $script:MockCPCState[$UserId] = @{
            phase  = 'provisioned'   # provisioned → inGracePeriod → deprovisioning → gone → provisioning → provisioned2
            ticks  = 0
        }
    }
    $state = $script:MockCPCState[$UserId]
    $state.ticks++

    switch ($state.phase) {
        'provisioned' {
            return @(@{ id = 'cpc-mock-0001'; managedDeviceName = 'CPC-Mock-OLD01'; status = 'provisioned';
                        servicePlanName = 'Cloud PC Enterprise 2vCPU/8GB/256GB'; provisioningPolicyId = 'pol-source-001' })
        }
        'inGracePeriod' {
            return @(@{ id = 'cpc-mock-0001'; managedDeviceName = 'CPC-Mock-OLD01'; status = 'inGracePeriod';
                        servicePlanName = 'Cloud PC Enterprise 2vCPU/8GB/256GB'; provisioningPolicyId = 'pol-source-001' })
        }
        'deprovisioning' {
            if ($state.ticks -ge 3) { $state.phase = 'gone'; $state.ticks = 0 }
            return @(@{ id = 'cpc-mock-0001'; managedDeviceName = 'CPC-Mock-OLD01'; status = 'deprovisioning';
                        servicePlanName = 'Cloud PC Enterprise 2vCPU/8GB/256GB'; provisioningPolicyId = 'pol-source-001' })
        }
        'gone'         { return @() }
        'provisioning' {
            if ($state.ticks -ge 4) { $state.phase = 'provisioned2'; $state.ticks = 0 }
            return @(@{ id = 'cpc-mock-0001'; managedDeviceName = 'CPC-Mock-NEW01'; status = 'provisioning';
                        servicePlanName = 'Cloud PC Enterprise 2vCPU/8GB/256GB'; provisioningPolicyId = 'pol-target-001' })
        }
        'provisioned2' {
            return @(@{ id = 'cpc-mock-0001'; managedDeviceName = 'CPC-Mock-NEW01'; status = 'provisioned';
                        servicePlanName = 'Cloud PC Enterprise 2vCPU/8GB/256GB'; provisioningPolicyId = 'pol-target-001' })
        }
    }
    return @()
}

function Invoke-MockGraceEnd {
    param([string]$UserId)
    if ($script:MockCPCState.ContainsKey($UserId)) {
        $script:MockCPCState[$UserId].phase = 'deprovisioning'
        $script:MockCPCState[$UserId].ticks = 0
    }
}

function Invoke-MockGroupRemove {
    param([string]$UserId)
    if ($script:MockCPCState.ContainsKey($UserId)) {
        $script:MockCPCState[$UserId].phase = 'inGracePeriod'
    }
}

function Invoke-MockGroupAdd {
    param([string]$UserId)
    if ($script:MockCPCState.ContainsKey($UserId)) {
        $script:MockCPCState[$UserId].phase = 'provisioning'
        $script:MockCPCState[$UserId].ticks = 0
    }
}
#endregion

#region Logging

function Write-StatusLog { param([string]$Message, [string]$Color = 'White') Write-Log $Message 'Info  ' $Color }
function Write-InfoLog   { param([string]$Message, [string]$Color = 'Cyan')  if ($script:verboseLogging) { Write-Log $Message 'Info  ' $Color } elseif ($script:LogFilePath) { Add-Content -Path $script:LogFilePath -Value "[$( Get-Date -Format 'HH:mm:ss')] $Message" -ErrorAction SilentlyContinue } }
function Write-DebugLog  { param([string]$Message, [string]$Color = 'Gray')  if ($script:verboseLogging) { Write-Log $Message 'Debug ' $Color } elseif ($script:LogFilePath) { Add-Content -Path $script:LogFilePath -Value "[$( Get-Date -Format 'HH:mm:ss')] $Message" -ErrorAction SilentlyContinue } }
function Write-PollingLog{ param([string]$Message, [string]$Color = 'Gray')  if ($script:verboseLogging) { Write-Log $Message 'Poll  ' $Color } elseif ($script:LogFilePath) { Add-Content -Path $script:LogFilePath -Value "[$( Get-Date -Format 'HH:mm:ss')] $Message" -ErrorAction SilentlyContinue } }
#endregion

#region Module API Overrides for Mock Mode
# BEST-EFFORT DEMO ONLY. This entire block runs only when -MockMode is passed and
# swaps the leaf Graph API functions for fakes. Production code paths and the
# CloudPCReplace.psm1 state machine are untouched by this, so it is safe to edit
# freely here. Known limitation: the stubbed data can trip validation in the later
# swap stages (e.g. same provisioning policy returned for all groups). Not a
# supported code path — do not invest heavily.
if ($script:MockMode) {
    # Alias mock state globally so module-scope overrides can access it (hashtable = reference type)
    $global:MockCPCState = $script:MockCPCState
    $global:MockUsers    = $script:MockUsers

    # Override Graph API functions at MODULE scope so Invoke-CloudPCReplaceStep intercepts them
    $module = Get-Module -Name CloudPCReplace
    if ($module) {
        & $module {
            function script:Get-CloudPCForUser {
                param([string]$UserId)
                if (-not $global:MockCPCState.ContainsKey($UserId)) {
                    $global:MockCPCState[$UserId] = @{ phase = 'provisioned'; ticks = 0 }
                }
                $state = $global:MockCPCState[$UserId]
                $state.ticks++
                switch ($state.phase) {
                    'provisioned'    { return @(@{ id = 'cpc-mock-0001'; managedDeviceName = 'CPC-Mock-OLD01'; status = 'provisioned';    servicePlanName = 'Cloud PC Enterprise 2vCPU/8GB/256GB'; provisioningPolicyId = 'pol-source-001' }) }
                    'inGracePeriod'  { return @(@{ id = 'cpc-mock-0001'; managedDeviceName = 'CPC-Mock-OLD01'; status = 'inGracePeriod';  servicePlanName = 'Cloud PC Enterprise 2vCPU/8GB/256GB'; provisioningPolicyId = 'pol-source-001' }) }
                    'deprovisioning' {
                        if ($state.ticks -ge 3) { $state.phase = 'gone'; $state.ticks = 0 }
                        return @(@{ id = 'cpc-mock-0001'; managedDeviceName = 'CPC-Mock-OLD01'; status = 'deprovisioning'; servicePlanName = 'Cloud PC Enterprise 2vCPU/8GB/256GB'; provisioningPolicyId = 'pol-source-001' })
                    }
                    'gone'           { return @() }
                    'provisioning'   {
                        if ($state.ticks -ge 2) { $state.phase = 'provisioned2'; $state.ticks = 0 }
                        return @(@{ id = 'cpc-mock-0001'; managedDeviceName = 'CPC-Mock-NEW01'; status = 'provisioning'; servicePlanName = 'Cloud PC Enterprise 2vCPU/8GB/256GB'; provisioningPolicyId = 'pol-target-001' })
                    }
                    'provisioned2'   { return @(@{ id = 'cpc-mock-0001'; managedDeviceName = 'CPC-Mock-NEW01'; status = 'provisioned';    servicePlanName = 'Cloud PC Enterprise 2vCPU/8GB/256GB'; provisioningPolicyId = 'pol-target-001' }) }
                }
                return @()
            }

            function script:Remove-UserFromGroup {
                param([string]$UserId, [string]$GroupId, [string]$UserPrincipalName)
                $key = if ($UserPrincipalName) { $UserPrincipalName } else { $UserId }
                if (-not $global:MockCPCState.ContainsKey($key)) { $global:MockCPCState[$key] = @{ phase = 'provisioned'; ticks = 0 } }
                $global:MockCPCState[$key].phase = 'inGracePeriod'
                return $true
            }

            function script:Add-UserToGroup {
                param([string]$UserId, [string]$GroupId, [string]$UserPrincipalName)
                $key = if ($UserPrincipalName) { $UserPrincipalName } else { $UserId }
                if (-not $global:MockCPCState.ContainsKey($key)) { $global:MockCPCState[$key] = @{ phase = 'provisioned'; ticks = 0 } }
                $global:MockCPCState[$key].phase = 'provisioning'
                $global:MockCPCState[$key].ticks = 0
                return $true
            }

            function script:Stop-CloudPCGracePeriod {
                param([string]$CloudPCId, [string]$UserPrincipalName)
                $key = if ($UserPrincipalName) { $UserPrincipalName } else { $CloudPCId }
                if (-not $global:MockCPCState.ContainsKey($key)) { $global:MockCPCState[$key] = @{ phase = 'provisioned'; ticks = 0 } }
                $global:MockCPCState[$key].phase = 'deprovisioning'
                $global:MockCPCState[$key].ticks = 0
                return $true
            }

            function script:Get-UserGroupMemberships {
                param([string]$UserId)
                return @(@{ id = 'aaaaaaaa-0004-0004-0004-aaaaaaaaaaaa'; displayName = 'w365 CPC in EU West' })
            }

            function script:Get-UserInfo {
                param([string]$UserPrincipalName)
                $user = $global:MockUsers | Where-Object { $_.userPrincipalName -eq $UserPrincipalName } | Select-Object -First 1
                if (-not $user) { return @{ id = 'mock-id-unknown'; userPrincipalName = $UserPrincipalName; displayName = $UserPrincipalName } }
                return @{ id = $user.id; userPrincipalName = $user.userPrincipalName; displayName = $user.displayName }
            }

            function script:Get-ProvisioningPoliciesForGroup {
                param([string]$GroupId)
                return @('pol-source-001')
            }
        }
    }

    # Script-scope overrides for functions called directly from GUI event handlers
    function Find-EntraIDGroups {
        param([string]$SearchTerm)
        return $script:MockGroups | Where-Object { $_.displayName -like "*$SearchTerm*" }
    }
    function Get-GroupMembers {
        param([string]$GroupId)
        return $script:MockUsers
    }
}
#endregion

#region Helper Functions
function Update-NextPollDisplay {
    param([UserReplaceState]$State)
    if ($State.Status -in @("Success","Success (Warnings)","Failed","Warning")) {
        $State.NextPollDisplay = "-"; return
    }
    $now = Get-Date
    $secondsSinceLastPoll = ($now - $State.LastPollTime).TotalSeconds
    $pollInterval = if ($State.Stage -eq "Waiting for Provisioning") { 180 } else { 60 }
    $secondsUntilNextPoll = [math]::Max(0, $pollInterval - $secondsSinceLastPoll)
    $isWaiting = $State.Stage -in @("Waiting for Grace Period","Ending Grace Period","Waiting for Deprovision","Waiting for Provisioning")
    $isImmediate = $State.Stage -in @("Getting User Info","Getting Current Cloud PC","Removing from Source","Adding to Target")
    if ($isImmediate) { $State.NextPollDisplay = "Now" }
    elseif ($isWaiting) {
        $elapsed = ($now - $State.StageStartTime)
        $elapsedMin = [math]::Floor($elapsed.TotalMinutes)
        $timeout = switch ($State.Stage) {
            "Waiting for Grace Period"  { $script:gracePeriodTimeoutMinutes }
            "Ending Grace Period"       { $script:endingGracePeriodTimeoutMinutes }
            "Waiting for Deprovision"   { $script:deprovisionTimeoutMinutes }
            "Waiting for Provisioning"  { $script:provisioningTimeoutMinutes }
        }
        $State.NextPollDisplay = "$([math]::Round($secondsUntilNextPoll))s $elapsedMin/$($timeout)m"
    }
    else { $State.NextPollDisplay = "$([math]::Round($secondsUntilNextPoll))s" }
}

function Save-JobSummaryAuto {
    param([UserReplaceState]$State)
    $timestamp = Get-Date -Format "yyyyMMdd"
    $csvPath = Join-Path $PSScriptRoot "CloudPC-Replace-Summary-$timestamp.csv"
    $row = [PSCustomObject]@{
        User        = $State.UserPrincipalName
        SourceGroup = $State.SourceGroupName
        OldCPC      = if ($State.OldCPCName) { $State.OldCPCName } else { "" }
        TargetGroup = $State.TargetGroupName
        NewCPC      = if ($State.NewCPCName) { $State.NewCPCName } else { "" }
        Status      = $State.Status
        Stage       = $State.Stage
        StartTime   = if ($State.StartTime -gt [DateTime]::MinValue) { $State.StartTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "" }
        EndTime     = if ($State.EndTime   -gt [DateTime]::MinValue) { $State.EndTime.ToString("yyyy-MM-dd HH:mm:ss")   } else { "" }
        Message     = if ($State.FinalMessage) { $State.FinalMessage } else { $State.ErrorMessage }
    }
    if (Test-Path $csvPath) { $row | Export-Csv -Path $csvPath -NoTypeInformation -Append }
    else                    { $row | Export-Csv -Path $csvPath -NoTypeInformation }
}

function Get-StatusBadge {
    param([string]$Status, [string]$Stage)
    switch ($Status) {
        "Success"            { return @{ Icon="✓";  Bg="#DCFCE7"; Fg="#15803D"; Icon2="✓" } }
        "Success (Warnings)" { return @{ Icon="⚠";  Bg="#FEF9C3"; Fg="#92400E"; Icon2="⚠" } }
        "Failed"             { return @{ Icon="✗";  Bg="#FEE2E2"; Fg="#991B1B"; Icon2="✗" } }
        "Warning"            { return @{ Icon="⚠";  Bg="#FEF3C7"; Fg="#B45309"; Icon2="⚠" } }
        "InProgress" {
            if ($Stage -eq "Waiting for Provisioning") {
                return @{ Icon="◉"; Bg="#E0F2FE"; Fg="#0369A1"; Icon2="◉" }
            }
            return @{ Icon="↻"; Bg="#DBEAFE"; Fg="#1D4ED8"; Icon2="↻" }
        }
        "Queued"             { return @{ Icon="⏱"; Bg="#F0F0F0"; Fg="#555555"; Icon2="⏱" } }
        default              { return @{ Icon="·";  Bg="#F0F0F0"; Fg="#888888"; Icon2="·" } }
    }
}

function Get-StatusPriority([string]$Status) {
    # Lower = sorts earlier (top of grid): Failed > Completed/Success > Monitoring > Active > Queued
    switch -Wildcard ($Status) {
        'Failed'       { return 0 }
        'Success*'     { return 1 }
        'Completed'    { return 1 }
        'Warning'      { return 1 }
        'Monitoring'   { return 2 }
        'Active'       { return 3 }
        default        { return 4 }   # Queued
    }
}

function Sort-JobItems {
    # Sort $script:JobItems in-place by (StatusPriority, QueueOrder) using ObservableCollection.Move()
    $n = $script:JobItems.Count
    if ($n -lt 2) { return }
    $targetOrder = @(0..($n-1) | ForEach-Object {
        $item = $script:JobItems[$_]
        [PSCustomObject]@{
            UPN        = $item.UPN
            Pri        = Get-StatusPriority $item.Status
            QueueOrder = if ($script:userStates[$item.UPN]) { $script:userStates[$item.UPN].QueueOrder } else { 9999 }
        }
    } | Sort-Object Pri, QueueOrder | ForEach-Object { $_.UPN })
    for ($i = 0; $i -lt $n; $i++) {
        if ($script:JobItems[$i].UPN -ne $targetOrder[$i]) {
            for ($j = $i + 1; $j -lt $n; $j++) {
                if ($script:JobItems[$j].UPN -eq $targetOrder[$i]) {
                    $script:JobItems.Move($j, $i); break
                }
            }
        }
    }
}

function Get-QueuedRange {
    $min = -1; $max = -1
    for ($i = 0; $i -lt $script:JobItems.Count; $i++) {
        if ($script:JobItems[$i].Status -eq 'Queued') {
            if ($min -eq -1) { $min = $i }
            $max = $i
        }
    }
    return @{ Min = $min; Max = $max }
}


function Format-Duration {
    param([TimeSpan]$Span)
    if ($Span.Ticks -lt 0) { $Span = [TimeSpan]::Zero }
    $h = [math]::Floor($Span.TotalHours)
    $m = $Span.Minutes
    $s = $Span.Seconds
    if ($h -gt 0)     { return ('{0}h{1:00}m{2:00}s' -f $h, $m, $s) }
    elseif ($m -gt 0) { return ('{0}m{1:00}s' -f $m, $s) }
    else              { return ('{0}s' -f $s) }
}

function Update-JobGrid {
    param([UserReplaceState]$State)
    $script:Window.Dispatcher.Invoke([Action]{
        # Find the matching row in the ItemsSource
        $item = $script:JobItems | Where-Object { $_.UPN -eq $State.UserPrincipalName }
        if ($item) {
            $displayStatus = if ($State.Status -eq "InProgress") {
                if ($State.Stage -eq "Waiting for Provisioning") { "Monitoring" } else { "Active" }
            } else { $State.Status }
            $item.Stage      = Get-StageDisplay $State.Stage
            $item.Status     = $displayStatus
            $item.NextPoll   = $State.NextPollDisplay
            $item.Messages   = if ($State.Status -in @('Success','Success (Warnings)') -and $State.StartTime -gt [DateTime]::MinValue -and $State.EndTime -gt [DateTime]::MinValue) {
                $dur = "Completed in $(Format-Duration ($State.EndTime - $State.StartTime))"
                if ($State.Status -eq 'Success (Warnings)') { "$dur — provisioned with warnings, review in Intune" } else { $dur }
            } elseif ($State.Status -eq 'InProgress' -and $State.StatusNote) { $State.StatusNote
            } elseif ($State.ErrorMessage) { $State.ErrorMessage } else { "" }
            $badge = Get-StatusBadge -Status $State.Status -Stage $State.Stage
            $item.BadgeColor = $badge.Bg
            $item.BadgeFg    = $badge.Fg
            $item.BadgeIcon  = $badge.Icon2
            Sort-JobItems
        }
    })
}
#endregion

#region Job Item (Observable)
Add-Type @"
using System.ComponentModel;
public class JobItem : INotifyPropertyChanged {
    public event PropertyChangedEventHandler PropertyChanged;
    private void OnProp(string n) { if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs(n)); }

    private string _upn;      public string UPN       { get { return _upn; }       set { _upn = value;       OnProp("UPN"); } }
    private string _source;   public string Source    { get { return _source; }    set { _source = value;    OnProp("Source"); } }
    private string _target;   public string Target    { get { return _target; }    set { _target = value;    OnProp("Target"); } }
    private string _stage;    public string Stage     { get { return _stage; }     set { _stage = value;     OnProp("Stage"); } }
    private string _status;   public string Status    { get { return _status; }    set { _status = value;    OnProp("Status"); } }
    private string _nextpoll; public string NextPoll  { get { return _nextpoll; }  set { _nextpoll = value;  OnProp("NextPoll"); } }
    private string _messages; public string Messages  { get { return _messages; }  set { _messages = value;  OnProp("Messages"); } }
    private string _badgebg;  public string BadgeColor{ get { return _badgebg; }   set { _badgebg = value;   OnProp("BadgeColor"); } }
    private string _badgefg;  public string BadgeFg   { get { return _badgefg; }   set { _badgefg = value;   OnProp("BadgeFg"); } }
    private string _badgeico; public string BadgeIcon { get { return _badgeico; }  set { _badgeico = value;  OnProp("BadgeIcon"); } }
}
"@

$script:JobItems = New-Object System.Collections.ObjectModel.ObservableCollection[JobItem]
#endregion

#region XAML
[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Cloud PC Replace Tool"
    Height="820" Width="1280" MinHeight="700" MinWidth="1100"
    WindowStartupLocation="CenterScreen"
    FontFamily="Segoe UI" FontSize="13"
    Background="#F3F4F6">

    <Window.Resources>
        <!-- Base button style -->
        <Style TargetType="Button" x:Key="BaseBtn">
            <Setter Property="Height"      Value="32"/>
            <Setter Property="Padding"     Value="12,0"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor"      Value="Hand"/>
            <Setter Property="FontFamily"  Value="Segoe UI"/>
            <Setter Property="FontSize"    Value="13"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}"
                                CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Opacity" Value="0.85"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="bd" Property="Opacity" Value="0.70"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="bd" Property="Opacity" Value="0.4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="Button" x:Key="PrimaryBtn" BasedOn="{StaticResource BaseBtn}">
            <Setter Property="Background"  Value="#0078D4"/>
            <Setter Property="Foreground"  Value="White"/>
            <Setter Property="FontWeight"  Value="SemiBold"/>
        </Style>
        <Style TargetType="Button" x:Key="SuccessBtn" BasedOn="{StaticResource BaseBtn}">
            <Setter Property="Background"  Value="#107C10"/>
            <Setter Property="Foreground"  Value="White"/>
            <Setter Property="FontWeight"  Value="SemiBold"/>
        </Style>
        <Style TargetType="Button" x:Key="DangerBtn" BasedOn="{StaticResource BaseBtn}">
            <Setter Property="Background"  Value="#C42B1C"/>
            <Setter Property="Foreground"  Value="White"/>
            <Setter Property="FontWeight"  Value="SemiBold"/>
        </Style>
        <Style TargetType="Button" x:Key="NeutralBtn" BasedOn="{StaticResource BaseBtn}">
            <Setter Property="Background"  Value="#E0E0E0"/>
            <Setter Property="Foreground"  Value="#1F1F1F"/>
        </Style>
        <!-- Spinner style for RepeatButton (same look as NeutralBtn) -->
        <Style TargetType="RepeatButton" x:Key="SpinnerBtn">
            <Setter Property="Background"      Value="#E0E0E0"/>
            <Setter Property="Foreground"      Value="#1F1F1F"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor"          Value="Hand"/>
            <Setter Property="FontFamily"      Value="Segoe UI"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="RepeatButton">
                        <Border x:Name="bd" Background="{TemplateBinding Background}"
                                CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Opacity" Value="0.85"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="bd" Property="Opacity" Value="0.70"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="bd" Property="Opacity" Value="0.4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="Button" x:Key="GhostBtn" BasedOn="{StaticResource BaseBtn}">
            <Setter Property="Background"  Value="Transparent"/>
            <Setter Property="Foreground"  Value="#0078D4"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="Transparent" CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#EBF3FB"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="bd" Property="Opacity" Value="0.4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- TextBox style -->
        <Style TargetType="TextBox">
            <Setter Property="Height"            Value="32"/>
            <Setter Property="Padding"           Value="8,4"/>
            <Setter Property="BorderBrush"       Value="#CCCCCC"/>
            <Setter Property="BorderThickness"   Value="1"/>
            <Setter Property="Background"        Value="White"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TextBox">
                        <Border Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6">
                            <ScrollViewer x:Name="PART_ContentHost" Margin="0" Padding="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsFocused" Value="True">
                    <Setter Property="BorderBrush" Value="#0078D4"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- Card panel style -->
        <Style TargetType="Border" x:Key="Card">
            <Setter Property="Background"      Value="White"/>
            <Setter Property="CornerRadius"    Value="8"/>
            <Setter Property="Padding"         Value="12"/>
            <Setter Property="Effect">
                <Setter.Value>
                    <DropShadowEffect BlurRadius="8" ShadowDepth="1" Opacity="0.08" Color="Black"/>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Section header -->
        <Style TargetType="TextBlock" x:Key="SectionHeader">
            <Setter Property="FontSize"    Value="12"/>
            <Setter Property="FontWeight"  Value="SemiBold"/>
            <Setter Property="Foreground"  Value="#555555"/>
            <Setter Property="Margin"      Value="0,0,0,8"/>
        </Style>

        <!-- DataGrid row selected style -->
        <Style TargetType="DataGridRow">
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#CCE4F7"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <!-- Collapsible panel Expander style -->
        <Style TargetType="Expander" x:Key="PanelExpander">
            <Setter Property="IsExpanded"    Value="True"/>
            <Setter Property="Background"    Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding"       Value="0"/>
        </Style>
    </Window.Resources>

    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*" MinHeight="200"/>
            <RowDefinition Height="6"/>
            <RowDefinition Height="180" MinHeight="80"/>
        </Grid.RowDefinitions>

        <!-- ═══════ TOP BAR ═══════ -->
        <Border Grid.Row="0" Style="{StaticResource Card}" Margin="0,0,0,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="☁  Cloud PC Replace" FontSize="16" FontWeight="Bold" Foreground="#0078D4" VerticalAlignment="Center"/>
                    <TextBlock x:Name="lblVersion" Text=" v$($script:ToolVersion)" FontSize="12" Foreground="#888" VerticalAlignment="Center" BaselineOffset="0" Margin="0,4,0,0"/>
                    <Border x:Name="mockBadge" Background="#FF8C00" CornerRadius="4" Padding="6,3" Margin="10,0,0,0"
                            VerticalAlignment="Center"
                            Visibility="$(if($script:MockMode){'Visible'}else{'Collapsed'})">
                        <TextBlock Text="MOCK MODE" FontSize="11" FontWeight="Bold" Foreground="White" VerticalAlignment="Center"/>
                    </Border>
                </StackPanel>

                <TextBlock x:Name="lblConnectionStatus" Grid.Column="1" Text="Not connected"
                           VerticalAlignment="Center" Foreground="#888" Margin="0,0,16,0"/>

                <Button x:Name="btnConnect" Grid.Column="2" Content="Connect to Graph"
                        Style="{StaticResource PrimaryBtn}" Width="150"/>
            </Grid>
        </Border>

        <!-- ═══════ BOTTOM: LOG ═══════ -->
        <Border Grid.Row="3" Style="{StaticResource Card}" Margin="0,4,0,0">
            <DockPanel>
                <DockPanel DockPanel.Dock="Top" Margin="0,0,0,6">
                    <TextBlock Text="Activity Log" Style="{StaticResource SectionHeader}" DockPanel.Dock="Left"/>
                    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center">
                        <CheckBox x:Name="chkVerbose" Content="Verbose" Foreground="#555" FontSize="12" Margin="0,0,14,0"/>
                        <CheckBox x:Name="chkAutoScroll" Content="Auto-scroll" IsChecked="True" Foreground="#555" FontSize="12"/>
                    </StackPanel>
                </DockPanel>
                <RichTextBox x:Name="rtbLog" IsReadOnly="True" Background="#FAFAFA"
                             BorderBrush="#E0E0E0" BorderThickness="1"
                             FontFamily="Consolas" FontSize="12" Foreground="#1F1F1F"
                             VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                             Padding="6"/>
            </DockPanel>
        </Border>

        <!-- ═══════ MAIN CONTENT ═══════ -->
        <GridSplitter Grid.Row="2" HorizontalAlignment="Stretch" Height="8"
                      Background="Transparent" Cursor="SizeNS" ShowsPreview="True">
            <GridSplitter.Template>
                <ControlTemplate>
                    <Grid Background="Transparent">
                        <TextBlock Text="&#xE76F;" FontFamily="Segoe MDL2 Assets" FontSize="10"
                                   Foreground="#BBBBBB" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Grid>
                </ControlTemplate>
            </GridSplitter.Template>
        </GridSplitter>

        <Grid Grid.Row="1" Margin="0,0,0,4">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>

            <!-- TOP PANELS: [Source + Users] | [Target] — collapsible -->
            <Expander Grid.Row="0" Style="{StaticResource PanelExpander}" Margin="0,0,0,6">
                <Expander.Header>
                    <TextBlock Text="GROUPS &amp; USERS" FontSize="12" FontWeight="SemiBold"
                               Foreground="#555555" VerticalAlignment="Center" Margin="4,0,0,0"/>
                </Expander.Header>
                <Grid Height="210" Margin="0,8,0,0">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="2*"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>

                <!-- Combined SOURCE GROUP + USERS card -->
                <Border Grid.Column="0" Style="{StaticResource Card}" Margin="0,0,6,0">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>

                        <!-- Source Group (left half) -->
                        <DockPanel Grid.Column="0">
                            <Grid DockPanel.Dock="Top" Margin="0,0,0,8">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="SOURCE" FontSize="11" FontWeight="SemiBold"
                                           Foreground="#555" VerticalAlignment="Center" Margin="0,0,8,0"/>
                                <Grid Grid.Column="1" Margin="0,0,6,0">
                                    <TextBox x:Name="txtSearchSource"
                                             ToolTip="Current provisioning profile group"/>
                                    <TextBlock Text="filter…" IsHitTestVisible="False"
                                               Foreground="#AAAAAA" FontStyle="Italic" FontSize="12"
                                               VerticalAlignment="Center" Margin="6,0,0,0">
                                        <TextBlock.Style>
                                            <Style TargetType="TextBlock">
                                                <Setter Property="Visibility" Value="Collapsed"/>
                                                <Style.Triggers>
                                                    <MultiDataTrigger>
                                                        <MultiDataTrigger.Conditions>
                                                            <Condition Binding="{Binding Text, ElementName=txtSearchSource}" Value=""/>
                                                            <Condition Binding="{Binding IsFocused, ElementName=txtSearchSource}" Value="False"/>
                                                        </MultiDataTrigger.Conditions>
                                                        <Setter Property="Visibility" Value="Visible"/>
                                                    </MultiDataTrigger>
                                                </Style.Triggers>
                                            </Style>
                                        </TextBlock.Style>
                                    </TextBlock>
                                </Grid>
                                <Button x:Name="btnSearchSource" Grid.Column="2" Content="&#xE72C;"
                                        FontFamily="Segoe MDL2 Assets" FontSize="13"
                                        Style="{StaticResource PrimaryBtn}" Width="36" Padding="0"
                                        ToolTip="Refresh group list" IsEnabled="False"/>
                            </Grid>
                            <TextBlock x:Name="lblSourcePolicy" DockPanel.Dock="Bottom"
                                               Text="Policy: (click to select, Ctrl/Shift for multiple)" FontSize="11"
                                               Foreground="#888" Margin="2,4,0,0"/>
                            <ListBox x:Name="lstSourceGroups" BorderBrush="#E0E0E0" BorderThickness="1"
                                     SelectionMode="Extended"
                                     ScrollViewer.VerticalScrollBarVisibility="Auto"/>
                        </DockPanel>

                        <!-- Thin vertical divider -->
                        <Border Grid.Column="1" Width="1" Background="#E0E0E0" Margin="10,0"/>

                        <!-- Users (right half) -->
                        <DockPanel Grid.Column="2">
                            <Grid DockPanel.Dock="Top" Margin="0,0,0,8">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="0" Text="USERS" FontSize="11" FontWeight="SemiBold"
                                           Foreground="#555" VerticalAlignment="Center" Margin="0,0,8,0"/>
                                <Grid Grid.Column="1" Margin="0,0,6,0">
                                    <TextBox x:Name="txtSearchUsers"/>
                                    <TextBlock Text="filter…" IsHitTestVisible="False"
                                               Foreground="#AAAAAA" FontStyle="Italic" FontSize="12"
                                               VerticalAlignment="Center" Margin="6,0,0,0">
                                        <TextBlock.Style>
                                            <Style TargetType="TextBlock">
                                                <Setter Property="Visibility" Value="Collapsed"/>
                                                <Style.Triggers>
                                                    <MultiDataTrigger>
                                                        <MultiDataTrigger.Conditions>
                                                            <Condition Binding="{Binding Text, ElementName=txtSearchUsers}" Value=""/>
                                                            <Condition Binding="{Binding IsFocused, ElementName=txtSearchUsers}" Value="False"/>
                                                        </MultiDataTrigger.Conditions>
                                                        <Setter Property="Visibility" Value="Visible"/>
                                                    </MultiDataTrigger>
                                                </Style.Triggers>
                                            </Style>
                                        </TextBlock.Style>
                                    </TextBlock>
                                </Grid>
                                <Button x:Name="btnSelectAll"  Grid.Column="2" Content="All"  Style="{StaticResource GhostBtn}" Height="28" Padding="6,0" FontSize="12"/>
                                <Button x:Name="btnSelectNone" Grid.Column="3" Content="None" Style="{StaticResource GhostBtn}" Height="28" Padding="6,0" FontSize="12"/>
                            </Grid>
                            <DockPanel DockPanel.Dock="Bottom" Margin="0,6,0,0">
                                <TextBlock x:Name="lblUsersCount" Text="0 users" FontSize="11" Foreground="#888" DockPanel.Dock="Left" VerticalAlignment="Center"/>
                                <Button x:Name="btnAddToQueue" Content="+ Add to Queue" Style="{StaticResource SuccessBtn}"
                                        HorizontalAlignment="Right" IsEnabled="False"/>
                            </DockPanel>
                            <ListBox x:Name="lstUsers" BorderBrush="#E0E0E0" BorderThickness="1"
                                     SelectionMode="Extended" ScrollViewer.VerticalScrollBarVisibility="Auto"/>
                        </DockPanel>
                    </Grid>
                </Border>

                <!-- Target Group -->
                <Border Grid.Column="1" Style="{StaticResource Card}" Margin="6,0,0,0">
                    <DockPanel>
                        <Grid DockPanel.Dock="Top" Margin="0,0,0,8">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBlock Grid.Column="0" Text="TARGET" FontSize="11" FontWeight="SemiBold"
                                       Foreground="#555" VerticalAlignment="Center" Margin="0,0,8,0"/>
                            <Grid Grid.Column="1" Margin="0,0,6,0">
                                <TextBox x:Name="txtSearchTarget"
                                         ToolTip="New provisioning profile group"/>
                                <TextBlock Text="filter…" IsHitTestVisible="False"
                                           Foreground="#AAAAAA" FontStyle="Italic" FontSize="12"
                                           VerticalAlignment="Center" Margin="6,0,0,0">
                                    <TextBlock.Style>
                                        <Style TargetType="TextBlock">
                                            <Setter Property="Visibility" Value="Collapsed"/>
                                            <Style.Triggers>
                                                <MultiDataTrigger>
                                                    <MultiDataTrigger.Conditions>
                                                        <Condition Binding="{Binding Text, ElementName=txtSearchTarget}" Value=""/>
                                                        <Condition Binding="{Binding IsFocused, ElementName=txtSearchTarget}" Value="False"/>
                                                    </MultiDataTrigger.Conditions>
                                                    <Setter Property="Visibility" Value="Visible"/>
                                                </MultiDataTrigger>
                                            </Style.Triggers>
                                        </Style>
                                    </TextBlock.Style>
                                </TextBlock>
                            </Grid>
                            <Button x:Name="btnSearchTarget" Grid.Column="2" Content="&#xE72C;"
                                    FontFamily="Segoe MDL2 Assets" FontSize="13"
                                    Style="{StaticResource PrimaryBtn}" Width="36" Padding="0"
                                    ToolTip="Refresh group list" IsEnabled="False"/>
                        </Grid>
                        <TextBlock x:Name="lblTargetPolicy" DockPanel.Dock="Bottom"
                                           Text="Policy: (not selected)" FontSize="11"
                                           Foreground="#888" Margin="2,4,0,0"/>
                        <ListBox x:Name="lstTargetGroups" BorderBrush="#E0E0E0" BorderThickness="1"
                                 ScrollViewer.VerticalScrollBarVisibility="Auto"/>
                    </DockPanel>
                </Border>
                </Grid><!-- end panels inner grid -->
            </Expander>

            <!-- QUEUE/GRID PANEL -->
            <Border Grid.Row="1" Style="{StaticResource Card}">
                <DockPanel>
                    <!-- Toolbar -->
                    <DockPanel DockPanel.Dock="Top" Margin="0,0,0,10">
                        <TextBlock Text="JOB QUEUE" Style="{StaticResource SectionHeader}" DockPanel.Dock="Left" VerticalAlignment="Center" Margin="0,0,0,0"/>
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                            <!-- Concurrent jobs spinner -->
                            <TextBlock Text="Concurrent jobs:" VerticalAlignment="Center" Margin="0,0,6,0" Foreground="#555" FontSize="12"/>
                            <RepeatButton x:Name="btnConcurrentDown" Content="−" Width="28" Height="28" Delay="400" Interval="120"
                                          Style="{StaticResource SpinnerBtn}" FontSize="15" FontWeight="Bold" Padding="0"
                                          VerticalContentAlignment="Center" Margin="0,0,0,0"/>
                            <TextBox x:Name="txtConcurrent" Text="2" Width="38" TextAlignment="Center"
                                     BorderBrush="#CCCCCC" BorderThickness="0,1" Height="28" Margin="0"/>
                            <RepeatButton x:Name="btnConcurrentUp" Content="+" Width="28" Height="28" Delay="400" Interval="120"
                                          Style="{StaticResource SpinnerBtn}" FontSize="15" FontWeight="Bold" Padding="0"
                                          VerticalContentAlignment="Center" Margin="0,0,16,0"/>
                            <!-- Queue management -->
                            <Button x:Name="btnMoveTop"    Content="⇈ Top"    Style="{StaticResource NeutralBtn}" Margin="0,0,4,0" IsEnabled="False"/>
                            <Button x:Name="btnMoveUp"     Content="↑ Up"     Style="{StaticResource NeutralBtn}" Margin="0,0,4,0" IsEnabled="False"/>
                            <Button x:Name="btnMoveDown"   Content="↓ Down"   Style="{StaticResource NeutralBtn}" Margin="0,0,4,0" IsEnabled="False"/>
                            <Button x:Name="btnMoveBottom" Content="⇊ Bottom" Style="{StaticResource NeutralBtn}" Margin="0,0,12,0" IsEnabled="False"/>
                            <Button x:Name="btnRemove"     Style="{StaticResource NeutralBtn}" Margin="0,0,4,0"  IsEnabled="False">
                                <StackPanel Orientation="Horizontal"><TextBlock Text="&#xE74D;" FontFamily="Segoe MDL2 Assets" FontSize="11" Margin="0,0,5,0" VerticalAlignment="Center"/><TextBlock Text="Remove" VerticalAlignment="Center"/></StackPanel>
                            </Button>
                            <Button x:Name="btnClearQueue" Style="{StaticResource NeutralBtn}" Margin="0,0,12,0" IsEnabled="False">
                                <StackPanel Orientation="Horizontal"><TextBlock Text="&#xE894;" FontFamily="Segoe MDL2 Assets" FontSize="11" Margin="0,0,5,0" VerticalAlignment="Center"/><TextBlock Text="Clear" VerticalAlignment="Center"/></StackPanel>
                            </Button>
                            <Button x:Name="btnSchedule"   Content="⏱ Schedule" Style="{StaticResource NeutralBtn}" Margin="0,0,4,0" IsEnabled="False"/>
                            <Button x:Name="btnStart"      Content="▶ Start"  Style="{StaticResource SuccessBtn}" Margin="0,0,4,0" IsEnabled="False"/>
                            <Button x:Name="btnStop"       Content="⏹ Stop"   Style="{StaticResource DangerBtn}"  Margin="0,0,12,0" IsEnabled="False"/>
                            <Button x:Name="btnExport"     Content="Export CSV" Style="{StaticResource NeutralBtn}"/>
                        </StackPanel>
                    </DockPanel>

                    <!-- Stat chip strip -->
                    <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal" Margin="0,8,0,0">
                        <Border Background="#F0F0F0" CornerRadius="10" Padding="8,3" Margin="0,0,6,0">
                            <TextBlock x:Name="lblStatTotal"   Text="Total: 0"       FontSize="11" Foreground="#555555"/>
                        </Border>
                        <Border Background="#F0F0F0" CornerRadius="10" Padding="8,3" Margin="0,0,6,0">
                            <TextBlock x:Name="lblStatQueued"  Text="⏱ 0 Queued"     FontSize="11" Foreground="#666666"/>
                        </Border>
                        <Border Background="#DBEAFE" CornerRadius="10" Padding="8,3" Margin="0,0,6,0">
                            <TextBlock x:Name="lblStatActive"  Text="↻ 0 Active"     FontSize="11" Foreground="#1D4ED8"/>
                        </Border>
                        <Border Background="#E0F2FE" CornerRadius="10" Padding="8,3" Margin="0,0,6,0">
                            <TextBlock x:Name="lblStatMonitor" Text="◉ 0 Monitoring"  FontSize="11" Foreground="#0369A1"/>
                        </Border>
                        <Border Background="#DCFCE7" CornerRadius="10" Padding="8,3" Margin="0,0,6,0">
                            <TextBlock x:Name="lblStatSuccess" Text="✓ 0 Success"     FontSize="11" Foreground="#15803D"/>
                        </Border>
                        <Border x:Name="brdStatWarning" Background="#FEF9C3" CornerRadius="10" Padding="8,3" Margin="0,0,6,0">
                            <TextBlock x:Name="lblStatWarning" Text="⚠ 0 Warnings"    FontSize="11" Foreground="#92400E"/>
                        </Border>
                        <Border Background="#FEE2E2" CornerRadius="10" Padding="8,3">
                            <TextBlock x:Name="lblStatFailed"  Text="✗ 0 Failed"      FontSize="11" Foreground="#991B1B"/>
                        </Border>
                        <Border x:Name="brdStatStartsIn" Background="#EDE9FE" CornerRadius="10" Padding="8,3" Margin="6,0,0,0" Visibility="Collapsed">
                            <TextBlock x:Name="lblStatStartsIn" Text="⏱ Starts in --:--:--" FontSize="11" Foreground="#6D28D9"/>
                        </Border>
                        <Border x:Name="brdStatStopsIn" Background="#FFE4E6" CornerRadius="10" Padding="8,3" Margin="6,0,0,0" Visibility="Collapsed">
                            <TextBlock x:Name="lblStatStopsIn" Text="⏸ Pauses in --:--:--" FontSize="11" Foreground="#BE123C"/>
                        </Border>
                    </StackPanel>

                    <!-- DataGrid -->
                    <DataGrid x:Name="dgJobs"
                              AutoGenerateColumns="False"
                              CanUserAddRows="False" CanUserDeleteRows="False"
                              CanUserSortColumns="False"
                              IsReadOnly="True"
                              SelectionMode="Extended" SelectionUnit="FullRow"
                              GridLinesVisibility="Horizontal"
                              HorizontalGridLinesBrush="#EEEEEE"
                              BorderBrush="#E0E0E0" BorderThickness="1"
                              RowBackground="White" AlternatingRowBackground="#FAFAFA"
                              HeadersVisibility="Column"
                              ColumnHeaderHeight="34"
                              TextOptions.TextFormattingMode="Display"
                              RenderOptions.ClearTypeHint="Enabled"
                              RowHeight="36"
                              FontSize="12">
                        <DataGrid.CellStyle>
                            <Style TargetType="DataGridCell">
                                <Setter Property="BorderThickness" Value="0"/>
                                <Setter Property="Template">
                                    <Setter.Value>
                                        <ControlTemplate TargetType="DataGridCell">
                                            <Border Background="{TemplateBinding Background}" Padding="8,0">
                                                <ContentPresenter VerticalAlignment="Center"/>
                                            </Border>
                                        </ControlTemplate>
                                    </Setter.Value>
                                </Setter>
                            </Style>
                        </DataGrid.CellStyle>
                        <DataGrid.ColumnHeaderStyle>
                            <Style TargetType="DataGridColumnHeader">
                                <Setter Property="Background"   Value="#F3F4F6"/>
                                <Setter Property="Foreground"   Value="#555555"/>
                                <Setter Property="FontWeight"   Value="SemiBold"/>
                                <Setter Property="FontSize"     Value="12"/>
                                <Setter Property="Padding"      Value="8,0"/>
                                <Setter Property="BorderBrush"  Value="#E0E0E0"/>
                                <Setter Property="BorderThickness" Value="0,0,0,1"/>
                            </Style>
                        </DataGrid.ColumnHeaderStyle>
                        <DataGrid.RowStyle>
                            <Style TargetType="DataGridRow">
                                <Style.Triggers>
                                    <Trigger Property="IsSelected" Value="True">
                                        <Setter Property="Background" Value="#CCE4F7"/>
                                    </Trigger>
                                </Style.Triggers>
                            </Style>
                        </DataGrid.RowStyle>
                        <DataGrid.Columns>
                            <DataGridTemplateColumn Header="User"         Width="19*" MinWidth="120">
                                <DataGridTemplateColumn.CellTemplate><DataTemplate>
                                    <TextBlock Text="{Binding UPN}"      TextTrimming="CharacterEllipsis" ToolTip="{Binding UPN}"      VerticalAlignment="Center" Margin="4,0"/>
                                </DataTemplate></DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                            <DataGridTemplateColumn Header="Source Group" Width="14*" MinWidth="90" >
                                <DataGridTemplateColumn.CellTemplate><DataTemplate>
                                    <TextBlock Text="{Binding Source}"   TextTrimming="CharacterEllipsis" ToolTip="{Binding Source}"   VerticalAlignment="Center" Margin="4,0"/>
                                </DataTemplate></DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                            <DataGridTemplateColumn Header="Target Group" Width="14*" MinWidth="90" >
                                <DataGridTemplateColumn.CellTemplate><DataTemplate>
                                    <TextBlock Text="{Binding Target}"   TextTrimming="CharacterEllipsis" ToolTip="{Binding Target}"   VerticalAlignment="Center" Margin="4,0"/>
                                </DataTemplate></DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                            <DataGridTemplateColumn Header="Stage"        Width="20*" MinWidth="100">
                                <DataGridTemplateColumn.CellTemplate><DataTemplate>
                                    <TextBlock Text="{Binding Stage}"    TextTrimming="CharacterEllipsis" ToolTip="{Binding Stage}"    VerticalAlignment="Center" Margin="4,0"/>
                                </DataTemplate></DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                            <DataGridTemplateColumn Header="Status" Width="9*" MinWidth="80">
                                <DataGridTemplateColumn.CellTemplate>
                                    <DataTemplate>
                                        <Border Background="{Binding BadgeColor}" CornerRadius="10"
                                                Padding="6,2" HorizontalAlignment="Left"
                                                VerticalAlignment="Center" Margin="4,3">
                                            <StackPanel Orientation="Horizontal">
                                                <TextBlock Text="{Binding BadgeIcon}" FontSize="11"
                                                           Foreground="{Binding BadgeFg}"
                                                           VerticalAlignment="Center" Margin="0,0,4,0"/>
                                                <TextBlock Text="{Binding Status}" FontSize="11"
                                                           FontWeight="SemiBold" Foreground="{Binding BadgeFg}"
                                                           VerticalAlignment="Center"/>
                                            </StackPanel>
                                        </Border>
                                    </DataTemplate>
                                </DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                            <DataGridTextColumn Header="Next Poll"    Binding="{Binding NextPoll}" Width="9*"  MinWidth="70"/>
                            <DataGridTemplateColumn Header="Messages"     Width="21*" MinWidth="100">
                                <DataGridTemplateColumn.CellTemplate><DataTemplate>
                                    <TextBlock Text="{Binding Messages}" TextTrimming="CharacterEllipsis" ToolTip="{Binding Messages}" VerticalAlignment="Center" Margin="4,0"/>
                                </DataTemplate></DataGridTemplateColumn.CellTemplate>
                            </DataGridTemplateColumn>
                        </DataGrid.Columns>
                    </DataGrid>
                </DockPanel>
            </Border>
        </Grid>
    </Grid>
</Window>
"@
#endregion

#region Build Window
$reader = [System.Xml.XmlNodeReader]::new($xaml)
$script:Window = [System.Windows.Markup.XamlReader]::Load($reader)

# Wire up named elements
$btnConnect        = $script:Window.FindName('btnConnect')
$btnSearchSource   = $script:Window.FindName('btnSearchSource')
$btnSearchTarget   = $script:Window.FindName('btnSearchTarget')
$txtSearchSource   = $script:Window.FindName('txtSearchSource')
$txtSearchTarget   = $script:Window.FindName('txtSearchTarget')
$txtSearchUsers    = $script:Window.FindName('txtSearchUsers')
$lstSourceGroups   = $script:Window.FindName('lstSourceGroups')
$lstTargetGroups   = $script:Window.FindName('lstTargetGroups')
$lblSourcePolicy   = $script:Window.FindName('lblSourcePolicy')
$lblTargetPolicy   = $script:Window.FindName('lblTargetPolicy')
$lstUsers          = $script:Window.FindName('lstUsers')
$lblUsersCount     = $script:Window.FindName('lblUsersCount')
$lblConnectionStatus = $script:Window.FindName('lblConnectionStatus')
$btnAddToQueue     = $script:Window.FindName('btnAddToQueue')
$btnSchedule       = $script:Window.FindName('btnSchedule')
$btnStart          = $script:Window.FindName('btnStart')
$btnStop           = $script:Window.FindName('btnStop')
$btnClearQueue     = $script:Window.FindName('btnClearQueue')
$btnRemove         = $script:Window.FindName('btnRemove')
$btnMoveUp         = $script:Window.FindName('btnMoveUp')
$btnMoveDown       = $script:Window.FindName('btnMoveDown')
$btnMoveTop        = $script:Window.FindName('btnMoveTop')
$btnMoveBottom     = $script:Window.FindName('btnMoveBottom')
$btnExport         = $script:Window.FindName('btnExport')
$btnSelectAll      = $script:Window.FindName('btnSelectAll')
$btnSelectNone     = $script:Window.FindName('btnSelectNone')
$chkVerbose        = $script:Window.FindName('chkVerbose')
$txtConcurrent     = $script:Window.FindName('txtConcurrent')
$btnConcurrentUp   = $script:Window.FindName('btnConcurrentUp')
$btnConcurrentDown = $script:Window.FindName('btnConcurrentDown')
$dgJobs            = $script:Window.FindName('dgJobs')
$lblStatTotal   = $script:Window.FindName('lblStatTotal')
$lblStatQueued  = $script:Window.FindName('lblStatQueued')
$lblStatActive  = $script:Window.FindName('lblStatActive')
$lblStatMonitor = $script:Window.FindName('lblStatMonitor')
$lblStatSuccess = $script:Window.FindName('lblStatSuccess')
$lblStatWarning = $script:Window.FindName('lblStatWarning')
$brdStatWarning = $script:Window.FindName('brdStatWarning')
$lblStatFailed  = $script:Window.FindName('lblStatFailed')
$brdStatStartsIn = $script:Window.FindName('brdStatStartsIn')
$lblStatStartsIn = $script:Window.FindName('lblStatStartsIn')
$brdStatStopsIn  = $script:Window.FindName('brdStatStopsIn')
$lblStatStopsIn  = $script:Window.FindName('lblStatStopsIn')
$rtbLog            = $script:Window.FindName('rtbLog')
$chkAutoScroll     = $script:Window.FindName('chkAutoScroll')

# Bind DataGrid
$dgJobs.ItemsSource = $script:JobItems

# Console startup banner (visible in the terminal used to launch the script)
Write-Host ""
Write-Host "  Cloud PC Replace Tool v$script:ToolVersion$(if($script:MockMode){' [MOCK MODE]'})" -ForegroundColor Cyan
Write-Host "  ─────────────────────────────────────────" -ForegroundColor DarkGray

$script:Window.Add_ContentRendered({
    $script:LogFilePath = Initialize-Logging -LogPath (Join-Path $PSScriptRoot "CloudPCReplace_Logs")
    Write-Host "  Logging to: $script:LogFilePath" -ForegroundColor DarkGray
    Write-Host ""
    Write-StatusLog "[Info  ] Cloud PC Replace Tool v$script:ToolVersion started$(if($script:MockMode){' (MOCK MODE)'})"
})

# Custom dialog centered over the app window (replaces MessageBox for centering reliability)
function Show-AppDialog {
    param(
        [string]$Message,
        [string]$Title   = "Notice",
        [string]$Icon    = "Warning",   # Warning | Info | Error | Question
        [string]$Buttons = "OK"         # OK | YesNo
    )
    $iconGlyph = switch ($Icon) { 'Warning'{"⚠"}; 'Error'{"✖"}; 'Question'{"?"}; default{"ℹ"} }
    $iconColor = switch ($Icon) { 'Warning'{"#7A5700"}; 'Error'{"#C42B1C"}; 'Question'{"#0078D4"}; default{"#0078D4"} }
    $escapedMsg = [System.Security.SecurityElement]::Escape($Message)
    $buttonXaml = if ($Buttons -eq 'YesNo') {
        '<StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
             <Button x:Name="btnYes" Content="Yes" Width="90" Height="32" Margin="0,0,8,0" Background="#0078D4" Foreground="White" FontWeight="SemiBold" BorderThickness="0" Cursor="Hand">
                 <Button.Style><Style TargetType="Button"><Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button"><Border Background="{TemplateBinding Background}" CornerRadius="4" SnapsToDevicePixels="True"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border><ControlTemplate.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#106EBE"/></Trigger><Trigger Property="IsPressed" Value="True"><Setter Property="Background" Value="#005A9E"/></Trigger></ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter></Style></Button.Style>
             </Button>
             <Button x:Name="btnNo"  Content="No"  Width="90" Height="32" Background="White" Foreground="#1F1F1F" FontWeight="SemiBold" BorderBrush="#CCCCCC" BorderThickness="1" Cursor="Hand">
                 <Button.Style><Style TargetType="Button"><Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button"><Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4" SnapsToDevicePixels="True"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border><ControlTemplate.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#F0F0F0"/></Trigger><Trigger Property="IsPressed" Value="True"><Setter Property="Background" Value="#E0E0E0"/></Trigger></ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter></Style></Button.Style>
             </Button>
         </StackPanel>'
    } else {
        '<Button x:Name="btnOK" Content="OK" Width="90" HorizontalAlignment="Right" Height="32" Background="#0078D4" Foreground="White" FontWeight="SemiBold" BorderThickness="0" Cursor="Hand">
             <Button.Style><Style TargetType="Button"><Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button"><Border Background="{TemplateBinding Background}" CornerRadius="4" SnapsToDevicePixels="True"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border><ControlTemplate.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#106EBE"/></Trigger><Trigger Property="IsPressed" Value="True"><Setter Property="Background" Value="#005A9E"/></Trigger></ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter></Style></Button.Style>
         </Button>'
    }
    [xml]$dlgXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$Title" SizeToContent="WidthAndHeight"
        MinWidth="340" MaxWidth="520"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize"
        FontFamily="Segoe UI" FontSize="13" Background="White">
    <Border Padding="24,20">
        <StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,16">
                <TextBlock Text="$iconGlyph" FontSize="22" Foreground="$iconColor" VerticalAlignment="Top" Margin="0,0,12,0"/>
                <TextBlock Text="$escapedMsg" TextWrapping="Wrap" MaxWidth="380" VerticalAlignment="Center"/>
            </StackPanel>
            $buttonXaml
        </StackPanel>
    </Border>
</Window>
"@
    $dlgReader = [System.Xml.XmlNodeReader]::new($dlgXaml)
    $dlg = [System.Windows.Markup.XamlReader]::Load($dlgReader)
    $dlg.Owner = $script:Window
    $script:_dlgResult = 'OK'
    if ($Buttons -eq 'YesNo') {
        $dlg.FindName('btnYes').Add_Click([System.Windows.RoutedEventHandler]{
            param($s, $e); $script:_dlgResult = 'Yes'; [System.Windows.Window]::GetWindow($s).Close()
        })
        $dlg.FindName('btnNo').Add_Click([System.Windows.RoutedEventHandler]{
            param($s, $e); $script:_dlgResult = 'No';  [System.Windows.Window]::GetWindow($s).Close()
        })
    } else {
        $dlg.FindName('btnOK').Add_Click([System.Windows.RoutedEventHandler]{
            param($s, $e); [System.Windows.Window]::GetWindow($s).Close()
        })
    }
    $dlg.ShowDialog() | Out-Null
    return $script:_dlgResult
}

function Resolve-ScheduleTime {
    param([string]$Mode, [string]$AtText, [string]$InText, [string]$Unit, [datetime]$Base)
    if ($Mode -eq 'at') {
        $parsed = [datetime]::MinValue
        if (-not [DateTime]::TryParse($AtText, [ref]$parsed)) { return $null }
        $t = $Base.Date.AddHours($parsed.Hour).AddMinutes($parsed.Minute)
        if ($t -le $Base) { $t = $t.AddDays(1) }
        return $t
    } else {
        $n = 0.0
        if (-not [double]::TryParse($InText, [ref]$n)) { return $null }
        if ($n -le 0) { return $null }
        if ($Unit -eq 'Hours') { return $Base.AddHours($n) } else { return $Base.AddMinutes($n) }
    }
}

function Show-ScheduleDialog {
    [xml]$dlgXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Schedule Queue" SizeToContent="WidthAndHeight"
        MinWidth="440" WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize" FontFamily="Segoe UI" FontSize="13" Background="White">
    <Window.Resources>
        <Style x:Key="PrimaryBtn" TargetType="Button">
            <Setter Property="Background" Value="#0078D4"/><Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/><Setter Property="BorderThickness" Value="0"/><Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button">
                <Border Background="{TemplateBinding Background}" CornerRadius="4"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border>
                <ControlTemplate.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#106EBE"/></Trigger><Trigger Property="IsEnabled" Value="False"><Setter Property="Background" Value="#A0C7E8"/></Trigger></ControlTemplate.Triggers>
            </ControlTemplate></Setter.Value></Setter>
        </Style>
        <Style x:Key="SecondaryBtn" TargetType="Button">
            <Setter Property="Background" Value="White"/><Setter Property="Foreground" Value="#1F1F1F"/>
            <Setter Property="FontWeight" Value="SemiBold"/><Setter Property="BorderBrush" Value="#CCCCCC"/><Setter Property="BorderThickness" Value="1"/><Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button">
                <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="4"><ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border>
                <ControlTemplate.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter Property="Background" Value="#F0F0F0"/></Trigger></ControlTemplate.Triggers>
            </ControlTemplate></Setter.Value></Setter>
        </Style>
    </Window.Resources>
    <Border Padding="24,20">
        <StackPanel>
            <TextBlock Text="Schedule Queue" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,2"/>
            <TextBlock Text="The app must stay open until the scheduled time." FontSize="11" Foreground="#888" Margin="0,0,0,16"/>

            <TextBlock Text="START" FontSize="11" FontWeight="Bold" Foreground="#666" Margin="0,0,0,8"/>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                <RadioButton x:Name="rbStartIn" GroupName="startMode" Content="Start in" IsChecked="True" VerticalAlignment="Center" Width="80"/>
                <TextBox x:Name="txtStartIn" Text="30" Width="60" Margin="0,0,6,0" Height="26" VerticalContentAlignment="Center"/>
                <ComboBox x:Name="cmbStartUnit" Width="90" Height="26" SelectedIndex="0">
                    <ComboBoxItem Content="Minutes"/><ComboBoxItem Content="Hours"/>
                </ComboBox>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,16">
                <RadioButton x:Name="rbStartAt" GroupName="startMode" Content="Start at" VerticalAlignment="Center" Width="80"/>
                <TextBox x:Name="txtStartAt" Text="22:00" Width="80" Height="26" VerticalContentAlignment="Center"/>
                <TextBlock Text="(HH:mm, 24-hr)" FontSize="11" Foreground="#888" VerticalAlignment="Center" Margin="8,0,0,0"/>
            </StackPanel>

            <CheckBox x:Name="chkEnableStop" Content="Also schedule a pause" Margin="0,0,0,8" FontWeight="SemiBold"/>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                <RadioButton x:Name="rbStopIn" GroupName="stopMode" Content="Pause in" IsChecked="True" IsEnabled="False" VerticalAlignment="Center" Width="80"/>
                <TextBox x:Name="txtStopIn" Text="6" Width="60" Margin="0,0,6,0" Height="26" IsEnabled="False" VerticalContentAlignment="Center"/>
                <ComboBox x:Name="cmbStopUnit" Width="90" Height="26" SelectedIndex="1" IsEnabled="False">
                    <ComboBoxItem Content="Minutes"/><ComboBoxItem Content="Hours"/>
                </ComboBox>
                <TextBlock Text="(after start)" FontSize="11" Foreground="#888" VerticalAlignment="Center" Margin="8,0,0,0"/>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,16">
                <RadioButton x:Name="rbStopAt" GroupName="stopMode" Content="Pause at" IsEnabled="False" VerticalAlignment="Center" Width="80"/>
                <TextBox x:Name="txtStopAt" Text="04:00" Width="80" Height="26" IsEnabled="False" VerticalContentAlignment="Center"/>
                <TextBlock Text="(HH:mm, 24-hr)" FontSize="11" Foreground="#888" VerticalAlignment="Center" Margin="8,0,0,0"/>
            </StackPanel>

            <Border Background="#F3F4F6" CornerRadius="6" Padding="12,10" Margin="0,0,0,16">
                <TextBlock x:Name="lblSummary" TextWrapping="Wrap" FontSize="12" Foreground="#374151"/>
            </Border>

            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="btnSched" Content="Schedule" Width="100" Height="32" Margin="0,0,8,0" Style="{StaticResource PrimaryBtn}"/>
                <Button x:Name="btnCancel" Content="Cancel" Width="90" Height="32" Style="{StaticResource SecondaryBtn}"/>
            </StackPanel>
        </StackPanel>
    </Border>
</Window>
'@
    $reader = [System.Xml.XmlNodeReader]::new($dlgXaml)
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader)
    $dlg.Owner = $script:Window

    $rbStartIn = $dlg.FindName('rbStartIn'); $rbStartAt = $dlg.FindName('rbStartAt')
    $txtStartIn = $dlg.FindName('txtStartIn'); $cmbStartUnit = $dlg.FindName('cmbStartUnit'); $txtStartAt = $dlg.FindName('txtStartAt')
    $chkEnableStop = $dlg.FindName('chkEnableStop')
    $rbStopIn = $dlg.FindName('rbStopIn'); $rbStopAt = $dlg.FindName('rbStopAt')
    $txtStopIn = $dlg.FindName('txtStopIn'); $cmbStopUnit = $dlg.FindName('cmbStopUnit'); $txtStopAt = $dlg.FindName('txtStopAt')
    $lblSummary = $dlg.FindName('lblSummary')
    $btnSched = $dlg.FindName('btnSched'); $btnCancel = $dlg.FindName('btnCancel')

    $script:_schedResult = $null
    $script:_schedStart  = $null
    $script:_schedStop   = $null

    $recalc = {
        $now = Get-Date
        $startMode = if ($rbStartAt.IsChecked) { 'at' } else { 'in' }
        $startUnit = if ($cmbStartUnit.SelectedItem) { $cmbStartUnit.SelectedItem.Content } else { 'Minutes' }
        $startAt = Resolve-ScheduleTime -Mode $startMode -AtText $txtStartAt.Text -InText $txtStartIn.Text -Unit $startUnit -Base $now

        $stopEnabled = [bool]$chkEnableStop.IsChecked
        $stopAt = $null
        $stopErr = $false
        if ($stopEnabled -and $startAt) {
            $stopMode = if ($rbStopAt.IsChecked) { 'at' } else { 'in' }
            $stopUnit = if ($cmbStopUnit.SelectedItem) { $cmbStopUnit.SelectedItem.Content } else { 'Hours' }
            if ($stopMode -eq 'in') {
                $stopAt = Resolve-ScheduleTime -Mode 'in' -InText $txtStopIn.Text -Unit $stopUnit -Base $startAt
            } else {
                $parsed = [datetime]::MinValue
                if ([DateTime]::TryParse($txtStopAt.Text, [ref]$parsed)) {
                    $t = $now.Date.AddHours($parsed.Hour).AddMinutes($parsed.Minute)
                    while ($t -le $startAt) { $t = $t.AddDays(1) }
                    $stopAt = $t
                }
            }
            if (-not $stopAt -or $stopAt -le $startAt) { $stopErr = $true }
        }

        $script:_schedStart = $startAt
        $script:_schedStop  = if ($stopEnabled) { $stopAt } else { $null }

        if (-not $startAt) {
            $lblSummary.Text = "Enter a valid start time."
            $btnSched.IsEnabled = $false
        } elseif ($stopErr) {
            $lblSummary.Text = "Pause time must be after the start time."
            $btnSched.IsEnabled = $false
        } else {
            $s = "Start:    " + $startAt.ToString('ddd h:mm tt')
            if ($script:_schedStop) {
                $rt = $script:_schedStop - $startAt
                $rtStr = "{0}h {1}m" -f [int][math]::Floor($rt.TotalHours), $rt.Minutes
                $s += "`nPause:    " + $script:_schedStop.ToString('ddd h:mm tt')
                $s += "`nRuntime:  " + $rtStr
            } else {
                $s += "`nRuns until done"
            }
            $lblSummary.Text = $s
            $btnSched.IsEnabled = $true
        }
    }

    $toggleStop = {
        $en = [bool]$chkEnableStop.IsChecked
        $rbStopIn.IsEnabled = $en; $rbStopAt.IsEnabled = $en
        $txtStopIn.IsEnabled = $en; $cmbStopUnit.IsEnabled = $en; $txtStopAt.IsEnabled = $en
        & $recalc
    }

    $rbStartIn.Add_Checked($recalc); $rbStartAt.Add_Checked($recalc)
    $txtStartIn.Add_TextChanged($recalc); $txtStartAt.Add_TextChanged($recalc)
    $cmbStartUnit.Add_SelectionChanged($recalc)
    $chkEnableStop.Add_Checked($toggleStop); $chkEnableStop.Add_Unchecked($toggleStop)
    $rbStopIn.Add_Checked($recalc); $rbStopAt.Add_Checked($recalc)
    $txtStopIn.Add_TextChanged($recalc); $txtStopAt.Add_TextChanged($recalc)
    $cmbStopUnit.Add_SelectionChanged($recalc)

    $btnSched.Add_Click([System.Windows.RoutedEventHandler]{
        param($s, $e)
        $script:_schedResult = @{ StartAt = $script:_schedStart; StopAt = $script:_schedStop }
        [System.Windows.Window]::GetWindow($s).Close()
    })
    $btnCancel.Add_Click([System.Windows.RoutedEventHandler]{
        param($s, $e)
        $script:_schedResult = $null
        [System.Windows.Window]::GetWindow($s).Close()
    })

    & $recalc
    $dlg.ShowDialog() | Out-Null
    return $script:_schedResult
}

function Write-Log {
    param([string]$Message, [string]$Level = 'Info  ', [string]$Color = 'Default')
    $timestamp = Get-Date -Format "HH:mm:ss"
    $logMessage = "[$timestamp] $Message"
    if ($script:LogFilePath) { Add-Content -Path $script:LogFilePath -Value $logMessage -ErrorAction SilentlyContinue }
    if ($rtbLog) {
        # Map caller color names to WPF hex colors suitable for white background
        $hexColor = switch ($Color) {
            'Green'    { '#107C10' }
            'Cyan'     { '#0078D4' }
            'Blue'     { '#0078D4' }
            'Red'      { '#C42B1C' }
            'Yellow'   { '#7A5700' }
            'Magenta'  { '#7719AA' }
            'Gray'     { '#767676' }
            'DarkGray' { '#767676' }
            'White'    { '#1F1F1F' }
            default    { '#1F1F1F' }
        }
        # Also color by log level prefix if no explicit color set
        if ($Color -eq 'Default' -or $Color -eq 'White') {
            $hexColor = switch -Regex ($Message) {
                '^\[OK\s*\]'     { '#107C10' }
                '^\[FAIL\s*\]'   { '#C42B1C' }
                '^\[WARN\s*\]'   { '#7A5700' }
                '^\[Detect\]'    { '#107C10' }
                '^\[Action\]'    { '#0078D4' }
                '^\[Search\]'    { '#5C6BC0' }
                '^\[Group\s*\]'  { '#0078D4' }
                '^\[Poll\s*\]'   { '#767676' }
                '^\[Debug\s*\]'  { '#767676' }
                default          { '#1F1F1F' }
            }
        }
        $rtbLog.Dispatcher.Invoke([Action]{
            $para = [System.Windows.Documents.Paragraph]::new()
            $para.Margin = [System.Windows.Thickness]::new(0)
            $run  = [System.Windows.Documents.Run]::new($logMessage)
            $run.Foreground = [System.Windows.Media.SolidColorBrush][System.Windows.Media.ColorConverter]::ConvertFromString($hexColor)
            $para.Inlines.Add($run)
            $rtbLog.Document.Blocks.Add($para)
            if ($chkAutoScroll.IsChecked) { $rtbLog.ScrollToEnd() }
        })
    } else {
        Write-Host $logMessage
    }
}
#endregion

#region Event Handlers

# Verbose toggle
$chkVerbose.Add_Checked({   $script:verboseLogging = $true;  Write-Log "[Info  ] Verbose logging enabled" })
$chkVerbose.Add_Unchecked({ $script:verboseLogging = $false; Write-Log "[Info  ] Verbose logging disabled" })

# Concurrent spinner
function Set-ConcurrentValue { param([int]$Delta)
    $cur = 2; try { $cur = [int]$script:txtConcurrent.Text } catch {}
    $cur = [math]::Max(1, [math]::Min(40, $cur + $Delta))
    $script:txtConcurrent.Text = "$cur"
    $script:maxConcurrent = $cur
    Write-Log "[Info  ] Concurrent limit set to $cur"
}
$btnConcurrentUp.Add_Click(  { Set-ConcurrentValue  1 })
$btnConcurrentDown.Add_Click({ Set-ConcurrentValue -1 })
$script:txtConcurrent.Add_LostFocus({
    $val = 2
    if ([int]::TryParse($script:txtConcurrent.Text, [ref]$val)) {
        $val = [math]::Max(1, [math]::Min(40, $val))
        $script:txtConcurrent.Text = "$val"
        $script:maxConcurrent = $val
    } else {
        $script:txtConcurrent.Text = "$script:maxConcurrent"
    }
})
$script:txtConcurrent.Add_KeyDown({ param($s,$e)
    if ($e.Key -eq 'Return') { $script:txtConcurrent.MoveFocus([System.Windows.Input.TraversalRequest]::new([System.Windows.Input.FocusNavigationDirection]::Next)) | Out-Null }
})

function Update-GroupList {
    param($ListBox, [array]$Groups, [string]$Filter = '')
    $ListBox.Items.Clear()
    $arr      = [System.Collections.ArrayList]::new()
    $filtered = if ([string]::IsNullOrWhiteSpace($Filter)) { $Groups } else {
        $Groups | Where-Object { $_.displayName -like "*$Filter*" }
    }
    foreach ($g in ($filtered | Sort-Object displayName)) {
        $warn = @()
        if ($g.onPremisesSyncEnabled -eq $true)          { $warn += "AD-Synced" }
        if ($g.groupTypes -contains 'DynamicMembership') { $warn += "Dynamic" }
        $label = $g.displayName + $(if ($warn.Count) { " ⚠ [" + ($warn -join ', ') + "]" } else { "" })
        $ListBox.Items.Add($label) | Out-Null
        $arr.Add($g) | Out-Null
    }
    $ListBox.Tag = $arr.ToArray()
}

# Connect
$btnConnect.Add_Click({
    # Disconnect flow
    if ($btnConnect.Tag -eq 'connected') {
        try { Disconnect-MgGraph | Out-Null } catch {}
        $btnConnect.Content             = "Connect to Graph"
        $btnConnect.Background          = '#0078D4'
        $btnConnect.Tag                 = $null
        $btnConnect.IsEnabled           = $true
        $lblConnectionStatus.Text       = "Not connected"
        $lblConnectionStatus.Foreground = '#888'
        $btnSearchSource.IsEnabled      = $false
        $btnSearchTarget.IsEnabled      = $false
        $script:allPolicyGroups = @()
        $script:groupMemberCache = @{}
        $lstSourceGroups.Items.Clear(); $lstSourceGroups.Tag = $null
        $lstTargetGroups.Items.Clear(); $lstTargetGroups.Tag = $null
        $lblSourcePolicy.Text = "Policy: (click to select, Ctrl/Shift for multiple)"
        $lblTargetPolicy.Text = "Policy: (not selected)"
        $script:sourceGroupId   = $null; $script:sourceGroupName = $null
        $script:targetGroupId   = $null; $script:targetGroupName = $null
        $lstUsers.Items.Clear();  $script:allSourceUsers = @()
        Write-Log "[Info  ] Disconnected"
        return
    }

    # Connect flow
    $btnConnect.IsEnabled = $false
    $lblConnectionStatus.Text = "Connecting..."
    $lblConnectionStatus.Foreground = '#888'
    try {
        if ($script:MockMode) {
            Start-Sleep -Milliseconds 600
            $lblConnectionStatus.Text       = "Connected (Mock Tenant)"
            $lblConnectionStatus.Foreground = '#107C10'
            $btnConnect.Content             = "✓ Disconnect"
            $btnConnect.Background          = '#107C10'
            $btnConnect.Tag                 = 'connected'
            $btnConnect.IsEnabled           = $true
            $script:allPolicyGroups = @($script:MockGroups | ForEach-Object { $_ })
            Update-GroupList -ListBox $lstSourceGroups -Groups $script:allPolicyGroups
            Update-GroupList -ListBox $lstTargetGroups -Groups $script:allPolicyGroups
            $btnSearchSource.IsEnabled      = $true
            $btnSearchTarget.IsEnabled      = $true
            Write-Log "[OK    ] Connected to mock tenant — loaded $($script:allPolicyGroups.Count) policy group(s)"
        } else {
            $ok = Connect-MgGraphForReplace -AuthMethod 'Interactive'
            if ($ok) {
                $ctx = Get-MgContext
                $lblConnectionStatus.Text       = "Connected: $($ctx.Account)"
                $lblConnectionStatus.Foreground = '#107C10'
                $btnConnect.Content             = "✓ Disconnect"
                $btnConnect.Background          = '#107C10'
                $btnConnect.Tag                 = 'connected'
                $btnConnect.IsEnabled           = $true
                try {
                    $script:allPolicyGroups = @(Get-EnterprisePolicyGroups)
                } catch {
                    $script:allPolicyGroups = @()
                    Write-Log "[WARN  ] Could not load policy groups: $_"
                }
                Update-GroupList -ListBox $lstSourceGroups -Groups $script:allPolicyGroups
                Update-GroupList -ListBox $lstTargetGroups -Groups $script:allPolicyGroups
                $btnSearchSource.IsEnabled      = $true
                $btnSearchTarget.IsEnabled      = $true
                Write-Log "[OK    ] Loaded $($script:allPolicyGroups.Count) policy group(s)"
            } else {
                $btnConnect.IsEnabled = $true
            }
        }
    } catch {
        $lblConnectionStatus.Text       = "Connection failed"
        $lblConnectionStatus.Foreground = '#C42B1C'
        $btnConnect.IsEnabled           = $true
        Write-Log "[FAIL  ] Connection error: $_"
    }
})

# Refresh policy group list (both buttons reload from API and re-filter both lists)
$btnSearchSource.Add_Click({
    $btnSearchSource.IsEnabled = $false
    $btnSearchTarget.IsEnabled = $false
    try {
        if ($script:MockMode) {
            $script:allPolicyGroups = @($script:MockGroups | ForEach-Object { $_ })
        } else {
            $script:allPolicyGroups = @(Get-EnterprisePolicyGroups)
        }
        Write-Log "[Info  ] Refreshed — $($script:allPolicyGroups.Count) policy group(s) loaded"
    } catch { Write-Log "[FAIL  ] Refresh error: $_" }
    finally {
        Update-GroupList -ListBox $lstSourceGroups -Groups $script:allPolicyGroups -Filter $txtSearchSource.Text
        Update-GroupList -ListBox $lstTargetGroups -Groups $script:allPolicyGroups -Filter $txtSearchTarget.Text
        # Reset all selection state — visual selection is gone after repopulation
        $script:sourceGroupId   = $null; $script:sourceGroupName = $null
        $script:targetGroupId   = $null; $script:targetGroupName = $null
        $lblSourcePolicy.Text   = "Policy: (click to select, Ctrl/Shift for multiple)"
        $lblTargetPolicy.Text   = "Policy: (not selected)"
        $lstUsers.Items.Clear(); $script:allSourceUsers = @()
        $script:groupMemberCache = @{}
        $lblUsersCount.Text     = "0 users"
        Update-AddToQueueButton
        $btnSearchSource.IsEnabled = $true
        $btnSearchTarget.IsEnabled = $true
    }
})
$btnSearchTarget.Add_Click({
    $btnSearchSource.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
})

# Live filter as the user types
$txtSearchSource.Add_TextChanged({
    param($s, $e)
    $term = $s.Text.Trim()
    $lstSourceGroups.Items.Clear()
    $arr = [System.Collections.ArrayList]::new()
    $filtered = if ([string]::IsNullOrWhiteSpace($term)) { $script:allPolicyGroups } else {
        $script:allPolicyGroups | Where-Object { $_.displayName -like "*$term*" }
    }
    foreach ($g in ($filtered | Sort-Object displayName)) {
        $warn = @()
        if ($g.onPremisesSyncEnabled -eq $true)          { $warn += "AD-Synced" }
        if ($g.groupTypes -contains 'DynamicMembership') { $warn += "Dynamic" }
        $lstSourceGroups.Items.Add($g.displayName + $(if ($warn.Count) { " ⚠ [" + ($warn -join ', ') + "]" } else { "" })) | Out-Null
        $arr.Add($g) | Out-Null
    }
    $lstSourceGroups.Tag = $arr.ToArray()
})
$txtSearchTarget.Add_TextChanged({
    param($s, $e)
    $term = $s.Text.Trim()
    $lstTargetGroups.Items.Clear()
    $arr = [System.Collections.ArrayList]::new()
    $filtered = if ([string]::IsNullOrWhiteSpace($term)) { $script:allPolicyGroups } else {
        $script:allPolicyGroups | Where-Object { $_.displayName -like "*$term*" }
    }
    foreach ($g in ($filtered | Sort-Object displayName)) {
        $warn = @()
        if ($g.onPremisesSyncEnabled -eq $true)          { $warn += "AD-Synced" }
        if ($g.groupTypes -contains 'DynamicMembership') { $warn += "Dynamic" }
        $lstTargetGroups.Items.Add($g.displayName + $(if ($warn.Count) { " ⚠ [" + ($warn -join ', ') + "]" } else { "" })) | Out-Null
        $arr.Add($g) | Out-Null
    }
    $lstTargetGroups.Tag = $arr.ToArray()
})

# Source group(s) selected
$lstSourceGroups.Add_SelectionChanged({
    try {
        if ($null -eq $lstSourceGroups.Tag) { return }
        $groups     = [object[]]$lstSourceGroups.Tag
        $selectedGs = @($lstSourceGroups.SelectedItems | ForEach-Object {
            $label = $_
            $groups | Where-Object {
                $warn = @()
                if ($_.onPremisesSyncEnabled -eq $true)          { $warn += "AD-Synced" }
                if ($_.groupTypes -contains 'DynamicMembership') { $warn += "Dynamic" }
                ($_.displayName + $(if ($warn.Count) { " ⚠ [" + ($warn -join ', ') + "]" } else { "" })) -eq $label
            } | Select-Object -First 1
        } | Where-Object { $_ })

        if ($selectedGs.Count -eq 0) {
            $script:sourceGroupId   = $null
            $script:sourceGroupName = $null
            $lblSourcePolicy.Text   = "Policy: (click to select, Ctrl/Shift for multiple)"
            $lstUsers.Items.Clear(); $script:allSourceUsers = @()
            $lblUsersCount.Text     = "0 users"
            Update-AddToQueueButton
            return
        }

        # Reject if any selected group is invalid — deselect that item and warn
        foreach ($g in $selectedGs) {
            if ($g.onPremisesSyncEnabled -eq $true) {
                Show-AppDialog -Message "'$($g.displayName)' is an AD-Synced group and cannot be used.`n`nUse an Assigned Entra Cloud Group instead." -Title "AD-Synced Group" -Icon Warning
                $lstSourceGroups.SelectedItems.Remove($lstSourceGroups.SelectedItem) | Out-Null
                return
            }
            if ($g.groupTypes -contains 'DynamicMembership') {
                Show-AppDialog -Message "'$($g.displayName)' is a Dynamic group and cannot be used.`n`nUse an Assigned (static) Entra Cloud Group instead." -Title "Dynamic Group" -Icon Warning
                $lstSourceGroups.SelectedItems.Remove($lstSourceGroups.SelectedItem) | Out-Null
                return
            }
        }

        # Policy label
        $policyNames = @($selectedGs | ForEach-Object { $_.policyName } | Where-Object { $_ } | Select-Object -Unique)
        $lblSourcePolicy.Text = if ($selectedGs.Count -eq 1) {
            if ($policyNames.Count -gt 0) { "Policy: $($policyNames -join ', ')" } else { "Policy: (unknown)" }
        } else {
            if ($policyNames.Count -eq 1) { "Policy: $($policyNames[0]) ($($selectedGs.Count) groups)" } else { "Policy: (multiple)" }
        }

        # Load members for any group not yet cached
        foreach ($g in $selectedGs) {
            if (-not $script:groupMemberCache.ContainsKey($g.id)) {
                Write-Log "[Group ] Loading members from '$($g.displayName)'..."
                try {
                    $members = Get-GroupMembers -GroupId $g.id | Sort-Object displayName
                    $script:groupMemberCache[$g.id] = @($members)
                    Write-Log "[Group ] Loaded $($script:groupMemberCache[$g.id].Count) member(s) from '$($g.displayName)'"
                } catch {
                    Write-Log "[FAIL  ] Error loading members for '$($g.displayName)': $_"
                    $script:groupMemberCache[$g.id] = @()
                }
            }
        }

        # Merge members across selected groups — deduplicate by UPN, first group wins
        $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $merged = [System.Collections.ArrayList]::new()
        foreach ($g in $selectedGs) {
            foreach ($u in $script:groupMemberCache[$g.id]) {
                if ($seen.Add($u.userPrincipalName)) {
                    # Attach sourceGroupId/Name to each user entry
                    $entry = [PSCustomObject]@{
                        id                = $u.id
                        displayName       = $u.displayName
                        userPrincipalName = $u.userPrincipalName
                        sourceGroupId     = $g.id
                        sourceGroupName   = $g.displayName
                    }
                    $merged.Add($entry) | Out-Null
                }
            }
        }
        $script:allSourceUsers = @($merged | Sort-Object displayName)

        $lstUsers.Items.Clear()
        foreach ($u in $script:allSourceUsers) {
            $lstUsers.Items.Add("$($u.displayName)  ($($u.userPrincipalName))") | Out-Null
        }
        $lblUsersCount.Text = "$($script:allSourceUsers.Count) users"

        # sourceGroupId/Name now per-user; keep script vars as sentinel for Add button check
        $script:sourceGroupId   = if ($selectedGs.Count -eq 1) { $selectedGs[0].id } else { 'multi' }
        $script:sourceGroupName = if ($selectedGs.Count -eq 1) { $selectedGs[0].displayName } else { "(multiple)" }

        Update-AddToQueueButton
    } catch { Write-Log "[FAIL  ] Source SelectionChanged error: $_" }
})

# Target group selected
$lstTargetGroups.Add_SelectionChanged({
    $idx = $lstTargetGroups.SelectedIndex
    if ($idx -lt 0 -or $null -eq $lstTargetGroups.Tag) { return }
    $g = $lstTargetGroups.Tag[$idx]
    if ($g.onPremisesSyncEnabled -eq $true) {
        Show-AppDialog -Message "'$($g.displayName)' is an AD-Synced group and cannot be used." -Title "AD-Synced Group" -Icon Warning
        $lstTargetGroups.SelectedIndex = -1; $lblTargetPolicy.Text = "Policy: (not selected)"; return
    }
    if ($g.groupTypes -contains 'DynamicMembership') {
        Show-AppDialog -Message "'$($g.displayName)' is a Dynamic group and cannot be used." -Title "Dynamic Group" -Icon Warning
        $lstTargetGroups.SelectedIndex = -1; $lblTargetPolicy.Text = "Policy: (not selected)"; return
    }
    $script:targetGroupId   = $g.id
    $script:targetGroupName = $g.displayName
    $lblTargetPolicy.Text   = if ($g.policyName) { "Policy: $($g.policyName)" } else { "Policy: (unknown)" }
    Write-Log "[Group ] Target group selected: $($g.displayName)"
    Update-AddToQueueButton
})

# User search filter
$txtSearchUsers.Add_TextChanged({
    $term = $txtSearchUsers.Text.Trim()
    $lstUsers.Items.Clear()
    $filtered = if ([string]::IsNullOrWhiteSpace($term)) {
        $script:allSourceUsers
    } else {
        $script:allSourceUsers | Where-Object { $_.displayName -like "*$term*" -or $_.userPrincipalName -like "*$term*" }
    }
    foreach ($u in ($filtered | Sort-Object displayName)) { $lstUsers.Items.Add("$($u.displayName)  ($($u.userPrincipalName))") | Out-Null }
    $selected = ($lstUsers.SelectedItems | Measure-Object).Count
    $lblUsersCount.Text = "$($filtered.Count) users$(if($selected -gt 0){" | $selected selected"})"
})

$lstUsers.Add_SelectionChanged({
    $selected = ($lstUsers.SelectedItems | Measure-Object).Count
    $total    = $lstUsers.Items.Count
    $lblUsersCount.Text = "$total users$(if($selected -gt 0){" | $selected selected"})"
    Update-AddToQueueButton
})

function Update-AddToQueueButton {
    $hasSource  = -not [string]::IsNullOrEmpty($script:sourceGroupId)
    $hasTarget  = -not [string]::IsNullOrEmpty($script:targetGroupId)
    $hasUsers   = ($lstUsers.SelectedItems | Measure-Object).Count -gt 0
    $btnAddToQueue.IsEnabled = $hasSource -and $hasTarget -and $hasUsers
}

# Select All / None
$btnSelectAll.Add_Click({
    $lstUsers.SelectAll()
    Update-AddToQueueButton
})
$btnSelectNone.Add_Click({
    $lstUsers.UnselectAll()
    Update-AddToQueueButton
})

# Add to Queue
$btnAddToQueue.Add_Click({
    if ($script:sourceGroupId -eq $script:targetGroupId) {
        Show-AppDialog -Message "Source and Target groups must be different." -Title "Invalid Selection" -Icon Warning
        return
    }

    $selectedIndices = @($lstUsers.SelectedItems)
    if ($selectedIndices.Count -eq 0) { return }

    $nextOrder  = if ($script:userStates.Count -gt 0) { ($script:userStates.Values | Measure-Object QueueOrder -Maximum).Maximum + 1 } else { 0 }
    $added      = 0
    $skipped    = @()
    $sameGroup  = @()

    foreach ($item in $selectedIndices) {
        # Extract UPN from display string "Name  (upn@domain)"
        if ($item -match '\(([^)]+)\)$') {
            $upn = $matches[1]
        } else { continue }

        if ($script:userStates.ContainsKey($upn)) {
            $skipped += $upn
            continue
        }

        # Find user object (carries per-user sourceGroupId/Name)
        $userObj = $script:allSourceUsers | Where-Object { $_.userPrincipalName -eq $upn } | Select-Object -First 1
        if (-not $userObj) { continue }

        # Skip users whose source group IS the target group
        if ($userObj.sourceGroupId -eq $script:targetGroupId) {
            $sameGroup += $upn
            continue
        }

        $state = [UserReplaceState]::new()
        $state.UserPrincipalName = $upn
        $state.UserId            = $userObj.id
        $state.SourceGroupId     = $userObj.sourceGroupId
        $state.SourceGroupName   = $userObj.sourceGroupName
        $state.TargetGroupId     = $script:targetGroupId
        $state.TargetGroupName   = $script:targetGroupName
        $state.Status            = "Queued"
        $state.Stage             = "Queued"
        $state.QueueOrder        = $nextOrder++
        $state.StartTime         = [DateTime]::MinValue
        $state.EndTime           = [DateTime]::MinValue
        $state.StageStartTime    = Get-Date
        $state.LastPollTime      = Get-Date

        $script:userStates[$upn] = $state

        $jobItem = [JobItem]::new()
        $jobItem.UPN      = $upn
        $jobItem.Source   = $userObj.sourceGroupName
        $jobItem.Target   = $script:targetGroupName
        $jobItem.Stage    = "Queued"
        $jobItem.Status   = "Queued"
        $jobItem.NextPoll = "-"
        $jobItem.Messages = ""
        $jobItem.BadgeColor = "#F0F0F0"
        $jobItem.BadgeFg    = "#555555"
        $jobItem.BadgeIcon  = "⏱"
        $script:JobItems.Add($jobItem)

        $added++
        Write-Log "[Info  ] Queued: $upn"
    }

    if ($added -gt 0) {
        $btnStart.IsEnabled      = $true
        if (-not $script:scheduleArmed -and -not $script:replaceRunning) { $btnSchedule.IsEnabled = $true }
        $btnClearQueue.IsEnabled = $true
        Write-Log "[Info  ] Added $added user(s) to queue. Total: $($script:userStates.Count)"
        Update-SummaryLabel
    }
    if ($skipped.Count -gt 0) {
        $names = $skipped -join "`n"
        Show-AppDialog -Message "The following user(s) are already in the queue and were skipped:`n`n$names" -Title "Already Queued" -Icon Info
    }
    if ($sameGroup.Count -gt 0) {
        $names = $sameGroup -join "`n"
        Show-AppDialog -Message "The following user(s) are already in the target group and were skipped:`n`n$names" -Title "Same Source/Target" -Icon Warning
    }
    $lstUsers.UnselectAll()
    Update-AddToQueueButton
})

function Update-SummaryLabel {
    $total   = $script:userStates.Count
    $queued  = ($script:userStates.Values | Where-Object { $_.Status -eq 'Queued' } | Measure-Object).Count
    $active  = ($script:userStates.Values | Where-Object { $_.Status -eq 'InProgress' -and $_.Stage -ne 'Waiting for Provisioning' } | Measure-Object).Count
    $monitor = ($script:userStates.Values | Where-Object { $_.Status -eq 'InProgress' -and $_.Stage -eq 'Waiting for Provisioning' } | Measure-Object).Count
    $success  = ($script:userStates.Values | Where-Object { $_.Status -eq 'Success' } | Measure-Object).Count
    $warnings = ($script:userStates.Values | Where-Object { $_.Status -eq 'Success (Warnings)' } | Measure-Object).Count
    $failed   = ($script:userStates.Values | Where-Object { $_.Status -eq 'Failed' } | Measure-Object).Count
    $lblStatTotal.Text   = "Total: $total"
    $lblStatQueued.Text  = "⏱ $queued Queued"
    $lblStatActive.Text  = "↻ $active Active"
    $lblStatMonitor.Text = "◉ $monitor Monitoring"
    $lblStatSuccess.Text = "✓ $success Success"
    $lblStatWarning.Text = "⚠ $warnings Warnings"
    $lblStatFailed.Text  = "✗ $failed Failed"
}

# Queue management buttons
$dgJobs.Add_SelectionChanged({
    $selected = @($dgJobs.SelectedItems)
    if ($selected.Count -eq 0) {
        $btnRemove.IsEnabled = $false
        $btnMoveUp.IsEnabled = $false; $btnMoveDown.IsEnabled = $false
        $btnMoveTop.IsEnabled = $false; $btnMoveBottom.IsEnabled = $false
        return
    }
    $btnRemove.IsEnabled = ($selected | Where-Object { $_.Status -ne 'InProgress' }).Count -gt 0
    Update-MoveButtonStates
})

$btnRemove.Add_Click({
    $selected = @($dgJobs.SelectedItems)
    if ($selected.Count -eq 0) { return }
    $blocked = $selected | Where-Object { $_.Status -eq 'InProgress' }
    if ($blocked.Count -gt 0 -and $blocked.Count -eq $selected.Count) {
        Show-AppDialog -Message "Cannot remove jobs that are currently running." -Title "Cannot Remove" -Icon Warning
        return
    }
    $toRemove = $selected | Where-Object { $_.Status -ne 'InProgress' }
    foreach ($item in $toRemove) {
        $script:userStates.Remove($item.UPN)
        $script:JobItems.Remove($item)
        Write-Log "[Info  ] Removed $($item.UPN) from queue"
    }
    if ($blocked.Count -gt 0) { Write-Log "[WARN  ] Skipped $($blocked.Count) running job(s)" }
    Update-SummaryLabel
})

$btnClearQueue.Add_Click({
    # Clear all non-running jobs from the grid (Queued, Failed, Success, etc.); leave InProgress jobs alone.
    $clearable = @($script:JobItems | Where-Object { $_.Status -ne 'InProgress' })
    $running   = @($script:JobItems | Where-Object { $_.Status -eq 'InProgress' }).Count
    if ($clearable.Count -eq 0) {
        Show-AppDialog -Message "There are no jobs to clear. Running jobs cannot be cleared." -Title "Nothing to Clear" -Icon Info
        return
    }
    $msg = "Clear $($clearable.Count) job(s) from the grid?"
    if ($running -gt 0) { $msg += "`n`n$running running job(s) will be left in place." }
    $confirm = Show-AppDialog -Message $msg -Title "Confirm" -Icon Question -Buttons YesNo
    if ($confirm -ne 'Yes') { return }
    foreach ($item in $clearable) {
        $script:userStates.Remove($item.UPN)
        $script:JobItems.Remove($item)
    }
    # Re-evaluate button states based on what remains
    $remainingClearable = @($script:JobItems | Where-Object { $_.Status -ne 'InProgress' }).Count
    $btnClearQueue.IsEnabled = ($remainingClearable -gt 0)
    $stillQueued = @($script:userStates.Values | Where-Object { $_.Status -eq 'Queued' }).Count
    if ($stillQueued -eq 0 -and -not $script:replaceRunning -and -not $script:scheduleArmed) {
        $btnStart.IsEnabled = $false; $btnSchedule.IsEnabled = $false
    }
    Update-SummaryLabel
    Write-Log "[Info  ] Cleared $($clearable.Count) job(s) from the grid"
})

function Update-MoveButtonStates {
    $selected = @($dgJobs.SelectedItems | Where-Object { $_.Status -eq 'Queued' })
    if ($selected.Count -eq 0) {
        $btnMoveUp.IsEnabled = $false; $btnMoveDown.IsEnabled = $false
        $btnMoveTop.IsEnabled = $false; $btnMoveBottom.IsEnabled = $false
        return
    }
    $range   = Get-QueuedRange
    $indices = $selected | ForEach-Object { $script:JobItems.IndexOf($_) } | Sort-Object
    $btnMoveTop.IsEnabled    = ($indices[0] -gt $range.Min)
    $btnMoveUp.IsEnabled     = ($indices[0] -gt $range.Min)
    $btnMoveDown.IsEnabled   = ($indices[-1] -lt $range.Max)
    $btnMoveBottom.IsEnabled = ($indices[-1] -lt $range.Max)
}

function Move-Job {
    param([string]$Direction)
    $selected = @($dgJobs.SelectedItems | Where-Object { $_.Status -eq 'Queued' })
    if ($selected.Count -eq 0) { return }
    $range   = Get-QueuedRange
    if ($range.Min -eq -1) { return }

    # Capture UPNs to restore selection after move
    $selUpns = $selected | ForEach-Object { $_.UPN }

    # Sort selected by current position
    $sorted = $selected | ForEach-Object {
        [PSCustomObject]@{ Item = $_; Idx = $script:JobItems.IndexOf($_) }
    } | Sort-Object Idx

    switch ($Direction) {
        'Up' {
            if ($sorted[0].Idx -le $range.Min) { return }
            foreach ($entry in $sorted) {
                $cur = $script:JobItems.IndexOf($entry.Item)
                $script:JobItems.Move($cur, $cur - 1)
            }
        }
        'Down' {
            if ($sorted[-1].Idx -ge $range.Max) { return }
            foreach ($entry in ($sorted | Sort-Object Idx -Descending)) {
                $cur = $script:JobItems.IndexOf($entry.Item)
                $script:JobItems.Move($cur, $cur + 1)
            }
        }
        'Top' {
            $target = $range.Min
            foreach ($entry in $sorted) {
                $cur = $script:JobItems.IndexOf($entry.Item)
                $script:JobItems.Move($cur, $target)
                $target++
            }
        }
        'Bottom' {
            $target = $range.Max
            foreach ($entry in ($sorted | Sort-Object Idx -Descending)) {
                $cur = $script:JobItems.IndexOf($entry.Item)
                $script:JobItems.Move($cur, $target)
                $target--
            }
        }
    }

    # Sync QueueOrder
    for ($i = $range.Min; $i -le $range.Max; $i++) {
        $upn = $script:JobItems[$i].UPN
        if ($script:userStates[$upn]) { $script:userStates[$upn].QueueOrder = $i }
    }

    # Restore selection
    $dgJobs.SelectedItems.Clear()
    foreach ($upn in $selUpns) {
        $item = $script:JobItems | Where-Object { $_.UPN -eq $upn } | Select-Object -First 1
        if ($item) { $dgJobs.SelectedItems.Add($item) }
    }
    $first = $script:JobItems | Where-Object { $_.UPN -eq $selUpns[0] } | Select-Object -First 1
    if ($first) { $dgJobs.ScrollIntoView($first) }
    Update-MoveButtonStates
}
$btnMoveUp.Add_Click(     { Move-Job 'Up' })
$btnMoveDown.Add_Click(   { Move-Job 'Down' })
$btnMoveTop.Add_Click(    { Move-Job 'Top' })
$btnMoveBottom.Add_Click( { Move-Job 'Bottom' })

# Start / Pause / Stop / Schedule
# Centralized run-control button states. 'idle' = not running (Start begins/resumes),
# 'running' = jobs processing (Start acts as Pause), 'pausing' = graceful drain in progress.
function Set-RunButtons {
    param([ValidateSet('idle','running','pausing')][string]$Mode)
    switch ($Mode) {
        'idle' {
            $btnStart.Content   = "▶ Start"
            # Enabled if anything is resumable (Queued, or InProgress frozen by a Stop)
            $resumable = ($script:userStates.Values | Where-Object { $_.Status -in @('Queued','InProgress') } | Measure-Object).Count -gt 0
            $btnStart.IsEnabled = $resumable
            $btnStop.IsEnabled  = $false
            $hasQueued = ($script:userStates.Values | Where-Object { $_.Status -eq 'Queued' } | Measure-Object).Count -gt 0
            $btnSchedule.IsEnabled = ($hasQueued -and -not $script:scheduleArmed)
        }
        'running' {
            $btnStart.Content      = "⏸ Pause"
            $btnStart.IsEnabled     = $true
            $btnStop.IsEnabled      = $true
            $btnSchedule.IsEnabled  = $false
        }
        'pausing' {
            $btnStart.Content      = "⏸ Pausing…"
            $btnStart.IsEnabled     = $false   # can't re-pause; Stop still available to force-abort
            $btnStop.IsEnabled      = $true
            $btnSchedule.IsEnabled  = $false
        }
    }
}

function Start-Processing {
    $script:replaceRunning    = $true
    $script:cancellationToken = $false
    $script:pausing           = $false
    Set-RunButtons -Mode running
    Write-Log "[Action] Processing started"
    $script:ProcessTimer.Start()
}

# Graceful drain: stop starting NEW jobs, let in-progress jobs run to completion,
# then halt (leftover Queued jobs stay queued). The ProcessTimer keeps running until
# the tick's drain-complete check clears it.
function Suspend-Processing {
    if (-not $script:replaceRunning) { return }
    $script:cancellationToken = $true
    $script:pausing           = $true
    $inProgress = ($script:userStates.Values | Where-Object { $_.Status -eq 'InProgress' } | Measure-Object).Count
    if ($inProgress -gt 0) {
        Set-RunButtons -Mode pausing
        Write-Log "[Action] Pause requested - no new jobs will start; $inProgress in-progress job(s) will run to completion"
    } else {
        # Nothing in flight — halt immediately
        $script:ProcessTimer.Stop()
        $script:replaceRunning = $false
        $script:pausing        = $false
        Set-RunButtons -Mode idle
        Write-Log "[Action] Paused - queue halted; no jobs were in progress"
    }
}

# Hard stop: freeze immediately. In-progress jobs stay at their current step and do
# NOT continue until the queue is resumed with Start.
function Stop-Processing {
    $script:cancellationToken = $true
    $script:pausing           = $false
    $script:ProcessTimer.Stop()
    $script:replaceRunning = $false
    Set-RunButtons -Mode idle
    Write-Log "[WARN  ] Processing stopped - in-progress jobs frozen at their current step (press Start to resume)"
}

function Set-ScheduleChips {
    $now = Get-Date
    if ($script:scheduleArmed -and $script:scheduledStartTime) {
        $r = $script:scheduledStartTime - $now
        if ($r.TotalSeconds -lt 0) { $r = [TimeSpan]::Zero }
        $lblStatStartsIn.Text = ("⏱ Starts in {0:D2}:{1:D2}:{2:D2}" -f [int][math]::Floor($r.TotalHours), $r.Minutes, $r.Seconds)
        $brdStatStartsIn.Visibility = 'Visible'
    } else {
        $brdStatStartsIn.Visibility = 'Collapsed'
    }
    if ($script:scheduledStopTime -and ($script:replaceRunning -or $script:scheduleArmed)) {
        $r = $script:scheduledStopTime - $now
        if ($r.TotalSeconds -lt 0) { $r = [TimeSpan]::Zero }
        $lblStatStopsIn.Text = ("⏸ Pauses in {0:D2}:{1:D2}:{2:D2}" -f [int][math]::Floor($r.TotalHours), $r.Minutes, $r.Seconds)
        $brdStatStopsIn.Visibility = 'Visible'
    } else {
        $brdStatStopsIn.Visibility = 'Collapsed'
    }
}

function Set-Schedule {
    param([datetime]$StartAt, $StopAt)
    $script:scheduledStartTime = $StartAt
    $script:scheduledStopTime  = $StopAt
    $script:scheduleArmed      = $true
    $script:lastKeepAlive      = Get-Date
    $btnSchedule.IsEnabled = $false
    $btnStart.IsEnabled    = $true    # allows "start now"
    $btnStop.IsEnabled     = $true    # allows "cancel schedule"
    if (-not $script:ScheduleTimer.IsEnabled) { $script:ScheduleTimer.Start() }
    Set-ScheduleChips
    $stopTxt = if ($StopAt) { " · pause $($StopAt.ToString('ddd h:mm tt'))" } else { " · runs until done" }
    Write-Log "[Action] Queue scheduled: start $($StartAt.ToString('ddd h:mm tt'))$stopTxt"
}

function Clear-Schedule {
    param([string]$Reason = "cancelled")
    $script:scheduleArmed      = $false
    $script:scheduledStartTime = $null
    $script:scheduledStopTime  = $null
    if ($script:ScheduleTimer) { $script:ScheduleTimer.Stop() }
    $brdStatStartsIn.Visibility = 'Collapsed'
    $brdStatStopsIn.Visibility  = 'Collapsed'
    $hasQueued = ($script:userStates.Values | Where-Object { $_.Status -eq 'Queued' } | Measure-Object).Count -gt 0
    $btnStart.IsEnabled    = $hasQueued
    $btnSchedule.IsEnabled = $hasQueued
    $btnStop.IsEnabled     = $false
    Write-Log "[Info  ] Scheduled run $Reason"
}

function Clear-ScheduledStop {
    # Release a scheduled auto-pause "cap" once a run ends on its own (before the cap fires),
    # so a leftover stop time can't later pause an unrelated run or leave the ScheduleTimer ticking.
    if (-not $script:scheduledStopTime) { return }
    $script:scheduledStopTime = $null
    $brdStatStopsIn.Visibility = 'Collapsed'
    if (-not $script:scheduleArmed -and $script:ScheduleTimer) { $script:ScheduleTimer.Stop() }
}

$btnSchedule.Add_Click({
    $result = Show-ScheduleDialog
    if ($result) { Set-Schedule -StartAt $result.StartAt -StopAt $result.StopAt }
})

$btnStart.Add_Click({
    # While running, this button acts as Pause (graceful drain)
    if ($script:replaceRunning) {
        Suspend-Processing
        return
    }
    if ($script:scheduleArmed) {
        # "Start now" overrides the pending start but preserves any scheduled stop
        $script:scheduleArmed      = $false
        $script:scheduledStartTime = $null
        $brdStatStartsIn.Visibility = 'Collapsed'
        Write-Log "[Info  ] Scheduled start overridden - starting now"
        if (-not $script:scheduledStopTime) { $script:ScheduleTimer.Stop() }
    }
    Start-Processing
})

$btnStop.Add_Click({
    if ($script:scheduleArmed) {
        Clear-Schedule -Reason "cancelled by user"
        return
    }
    # Confirm hard stop only when jobs are actually mid-swap (freezing them has real consequences)
    $inProgress = ($script:userStates.Values | Where-Object { $_.Status -eq 'InProgress' } | Measure-Object).Count
    if ($inProgress -gt 0) {
        $msg = "Stop now? $inProgress job(s) in progress will be frozen at their current step and will not finish - a Cloud PC mid-swap could be left deprovisioned with no replacement.`n`n"
        if (-not $script:pausing) { $msg += "To let in-progress jobs finish first, use the Pause button instead.`n`n" }
        $msg += "Queued jobs stay queued either way. Stop anyway?"
        $confirm = Show-AppDialog -Message $msg -Title "Stop Processing?" -Icon Warning -Buttons YesNo
        if ($confirm -ne 'Yes') { return }
    }
    Stop-Processing
    if ($script:scheduledStopTime -or $script:ScheduleTimer.IsEnabled) {
        $script:scheduledStopTime = $null
        $script:ScheduleTimer.Stop()
        $brdStatStopsIn.Visibility = 'Collapsed'
    }
})

#region Schedule Timer (DispatcherTimer - drives armed start/stop + token keep-alive)
$script:ScheduleTimer = [System.Windows.Threading.DispatcherTimer]::new()
$script:ScheduleTimer.Interval = [TimeSpan]::FromSeconds(1)
$script:ScheduleTimer.Add_Tick({
    $now = Get-Date

    # Token keep-alive while waiting for a scheduled start
    if ($script:scheduleArmed -and -not $script:MockMode) {
        if (-not $script:lastKeepAlive -or ($now - $script:lastKeepAlive).TotalMinutes -ge 10) {
            $script:lastKeepAlive = $now
            try {
                Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization?$select=id&$top=1' -ErrorAction Stop | Out-Null
                if ($script:verboseLogging) { Write-Log "[Debug ] Token keep-alive OK" }
            } catch {
                Write-Log "[WARN  ] Token keep-alive failed - Graph connection may have expired. Reconnect before the scheduled start."
            }
        }
    }

    # Fire scheduled START
    if ($script:scheduleArmed -and $script:scheduledStartTime -and $now -ge $script:scheduledStartTime) {
        $script:scheduleArmed      = $false
        $script:scheduledStartTime = $null
        $brdStatStartsIn.Visibility = 'Collapsed'

        $preflightOk = $true
        if (-not $script:MockMode) {
            $ctx = $null
            try { $ctx = Get-MgContext } catch {}
            if (-not $ctx) { $preflightOk = $false }
        }

        if ($preflightOk) {
            Write-Log "[Action] Scheduled start time reached - starting queue"
            Start-Processing
        } else {
            Write-Log "[FAIL  ] Scheduled start aborted - Graph connection lost. Jobs remain queued."
            $script:scheduledStopTime = $null
            $script:ScheduleTimer.Stop()
            $brdStatStopsIn.Visibility = 'Collapsed'
            $hasQueued = ($script:userStates.Values | Where-Object { $_.Status -eq 'Queued' } | Measure-Object).Count -gt 0
            $btnStart.IsEnabled    = $hasQueued
            $btnSchedule.IsEnabled = $hasQueued
            $btnStop.IsEnabled     = $false
            Show-AppDialog -Message "The scheduled start could not begin because the Microsoft Graph connection was lost or expired.`n`nYour jobs are still queued. Please reconnect and start manually." -Title "Scheduled Start Failed" -Icon Error
        }
    }

    # Fire scheduled STOP
    if ($script:scheduledStopTime -and $script:replaceRunning -and $now -ge $script:scheduledStopTime) {
        $script:scheduledStopTime = $null
        Write-Log "[Action] Scheduled stop time reached - pausing queue (in-progress jobs will run to completion)"
        Suspend-Processing
        $brdStatStopsIn.Visibility = 'Collapsed'
    }

    Set-ScheduleChips

    # Nothing left to drive? Stop the timer.
    if (-not $script:scheduleArmed -and -not $script:scheduledStopTime) {
        $script:ScheduleTimer.Stop()
    }
})
#endregion


# Export
$btnExport.Add_Click({
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Filter   = "CSV files (*.csv)|*.csv"
    $dlg.FileName = "CloudPC-Replace-Summary-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
    if ($dlg.ShowDialog()) {
        $allJobs = $script:userStates.Values | Sort-Object QueueOrder
        if ($allJobs.Count -eq 0) { Write-Log "[WARN  ] No jobs to export"; return }
        # Expand jobs with multiple old CPCs into separate rows (matches WinForms behaviour)
        $rows = @()
        foreach ($job in $allJobs) {
            $oldCPCNames = if ($job.OldCPCName) { $job.OldCPCName -split ', ' } else { @('') }
            foreach ($oldCPC in $oldCPCNames) {
                $rows += [PSCustomObject]@{
                    User        = $job.UserPrincipalName
                    SourceGroup = $job.SourceGroupName
                    OldCPC      = $oldCPC
                    TargetGroup = $job.TargetGroupName
                    NewCPC      = $job.NewCPCName
                    Status      = $job.Status
                    Stage       = $job.Stage
                    StartTime   = if ($job.StartTime -and $job.StartTime -gt [DateTime]::MinValue) { $job.StartTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "" }
                    EndTime     = if ($job.EndTime   -and $job.EndTime   -gt [DateTime]::MinValue) { $job.EndTime.ToString("yyyy-MM-dd HH:mm:ss") }   else { "" }
                    Message     = if ($job.FinalMessage) { $job.FinalMessage } else { $job.ErrorMessage }
                }
            }
        }
        $rows | Export-Csv -Path $dlg.FileName -NoTypeInformation
        Write-Log "[OK    ] Exported $($rows.Count) row(s) from $($allJobs.Count) job(s) to $($dlg.FileName)"
    }
})
#endregion

#region Processing Timer (DispatcherTimer - runs on UI thread)
$script:ProcessTimer = [System.Windows.Threading.DispatcherTimer]::new()
$script:ProcessTimer.Interval = [TimeSpan]::FromSeconds(3)
$script:timerTickCount = 0

$script:ProcessTimer.Add_Tick({
    $script:timerTickCount++
    if ($script:verboseLogging) {
        Write-Log "[Debug ] [TICK #$($script:timerTickCount)] $(Get-Date -Format 'HH:mm:ss.fff')"
    }

    $now = Get-Date
    $activeCount    = ($script:userStates.Values | Where-Object { $_.Status -eq 'InProgress' -and $_.Stage -ne 'Waiting for Provisioning' } | Measure-Object).Count
    $queuedCount    = ($script:userStates.Values | Where-Object { $_.Status -eq 'Queued' } | Measure-Object).Count

    $currentMax = 2
    try { $currentMax = [int]$txtConcurrent.Text } catch {}

    # Start queued jobs up to concurrency limit
    if (-not $script:cancellationToken -and $activeCount -lt $currentMax -and $queuedCount -gt 0) {
        $toStart = $currentMax - $activeCount
        $toStart = [math]::Min($toStart, $queuedCount)
        $readyJobs = $script:userStates.Values | Where-Object { $_.Status -eq 'Queued' } | Sort-Object QueueOrder | Select-Object -First $toStart
        foreach ($s in $readyJobs) {
            $s.Status         = 'InProgress'
            $s.Stage          = 'Getting User Info'
            $s.ProgressPercent = 5
            $s.StageStartTime = $now
            $s.LastPollTime   = $now
            $s.StartTime      = $now
            $s.ErrorMessage   = "Started at $(Get-Date -Format 'HH:mm:ss')"
            Update-JobGrid $s
        }
    }

    # Process each active job
    foreach ($state in @($script:userStates.Values | Where-Object { $_.Status -eq 'InProgress' })) {
        $isImmediate = $state.Stage -in @('Getting User Info','Getting Current Cloud PC','Removing from Source','Adding to Target','Complete')
        $pollInterval = if ($script:MockMode) { $script:mockPollInterval } elseif ($state.Stage -eq 'Waiting for Provisioning') { 180 } else { 60 }
        $secondsSinceLastPoll = ($now - $state.LastPollTime).TotalSeconds

        if ($isImmediate -or $secondsSinceLastPoll -ge $pollInterval) {
            $timeouts = @{
                GracePeriodTimeout       = $script:gracePeriodTimeoutMinutes
                EndingGracePeriodTimeout = $script:endingGracePeriodTimeoutMinutes
                DeprovisionTimeout       = $script:deprovisionTimeoutMinutes
                ProvisioningTimeout      = $script:provisioningTimeoutMinutes
            }

            $onLog = {
                param($Message, $Level, $Color)
                if ($Level -in @('Debug', 'Polling') -and -not $script:verboseLogging) { return }
                Write-Log $Message -Color $Color
            }

            $onGridUpdate = {
                param($State, $ColumnName, $Value)
                $item = $script:JobItems | Where-Object { $_.UPN -eq $State.UserPrincipalName } | Select-Object -First 1
                if ($item) {
                    switch ($ColumnName) {
                        'Messages' { $item.Messages = $Value }
                        'Stage'    { $item.Stage    = $Value }
                        'Status'   { $item.Status   = $Value }
                    }
                }
            }

            try {
                Invoke-CloudPCReplaceStep -State $state -Timeouts $timeouts -OnLog $onLog -OnGridUpdate $onGridUpdate
            } catch {
                $state.Status       = 'Failed'
                $state.ErrorMessage = $_.ToString()
                Write-Log "[FAIL  ] $($state.UserPrincipalName): $_"
            }

            $state.LastPollTime = $now
            Update-JobGrid $state

            if ($state.Status -in @('Success','Success (Warnings)','Failed','Warning')) {
                Save-JobSummaryAuto -State $state
                Write-Log "[OK    ] $($state.UserPrincipalName): Job complete - $($state.Status)"
            }
        }

        Update-NextPollDisplay $state
        $item = $script:JobItems | Where-Object { $_.UPN -eq $state.UserPrincipalName } | Select-Object -First 1
        if ($item) { $item.NextPoll = $state.NextPollDisplay }
    }

    Update-SummaryLabel

    # Check for pause-drain completion or full completion
    $anyInProgress = ($script:userStates.Values | Where-Object { $_.Status -eq 'InProgress' } | Measure-Object).Count -gt 0
    $queuedCountNow = ($script:userStates.Values | Where-Object { $_.Status -eq 'Queued' } | Measure-Object).Count

    if ($script:pausing -and -not $anyInProgress) {
        # Graceful drain finished — halt, leaving any queued jobs in place to resume later
        $script:ProcessTimer.Stop()
        $script:replaceRunning    = $false
        $script:pausing           = $false
        $script:cancellationToken = $false
        Clear-ScheduledStop
        Set-RunButtons -Mode idle
        $qMsg = if ($queuedCountNow -gt 0) { " - $queuedCountNow job(s) still queued (press Start to resume)" } else { "" }
        Write-Log "[Action] Paused - all in-progress jobs finished$qMsg"
        Update-SummaryLabel
    }
    elseif ($script:replaceRunning -and -not $anyInProgress -and $queuedCountNow -eq 0) {
        $script:ProcessTimer.Stop()
        $script:replaceRunning = $false
        Clear-ScheduledStop
        Set-RunButtons -Mode idle
        Write-Log "[OK    ] === ALL JOBS COMPLETE ==="
        Update-SummaryLabel
    }
})
#endregion

#region Self-Driving Demo (Mock Mode only)
if ($script:MockMode) {
    $script:DemoTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:DemoStep  = 0
    $script:DemoTimer.Interval = [TimeSpan]::FromMilliseconds(800)

    $script:DemoTimer.Add_Tick({
        $script:DemoStep++
        switch ($script:DemoStep) {
            1  {
                Write-Log "[Info  ] [DEMO] Starting self-driving demo..."
                $btnConnect.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
            }
            3  {
                # Filter BOTH lists before selecting anything — the Search button does a
                # hard refresh that resets all selection state, so it must run before
                # any source/target group is picked (otherwise it wipes the selection).
                $txtSearchSource.Text = "w365"
                $txtSearchTarget.Text = "w365"
                $btnSearchSource.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
            }
            5  {
                # Select first non-warning source group (skip AD-Synced / Dynamic groups)
                for ($i = 0; $i -lt $lstSourceGroups.Items.Count; $i++) {
                    if ($lstSourceGroups.Items[$i] -notmatch '⚠') {
                        $lstSourceGroups.SelectedIndex = $i; break
                    }
                }
            }
            7  {
                # Select second non-warning target group (different from source)
                $picked = 0
                for ($i = 0; $i -lt $lstTargetGroups.Items.Count; $i++) {
                    if ($lstTargetGroups.Items[$i] -notmatch '⚠') {
                        $picked++
                        if ($picked -eq 2) { $lstTargetGroups.SelectedIndex = $i; break }
                    }
                }
            }
            9  {
                # Select first 2 users
                $lstUsers.UnselectAll()
                $lstUsers.SelectedIndex = 0
                if ($lstUsers.Items.Count -gt 1) {
                    $lstUsers.SelectedItems.Add($lstUsers.Items[1]) | Out-Null
                }
                Update-AddToQueueButton
            }
            11 {
                $btnAddToQueue.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
            }
            13 {
                Write-Log "[Info  ] [DEMO] Starting processing - watch the queue!"
                $btnStart.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
                $script:DemoTimer.Stop()  # Hand off to process timer
            }
        }
    })
}
#endregion

# Start demo on load (mock mode only)
$script:Window.Add_Loaded({
    if ($script:MockMode) {
        $script:DemoTimer.Start()
    }
})

# Clean up on close
$script:Window.Add_Closed({
    $script:ProcessTimer.Stop()
    if ($script:MockMode -and $script:DemoTimer) { $script:DemoTimer.Stop() }
})

# Disclaimer — auto-dismissed after 2s in Mock Mode
[xml]$discXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Cloud PC Replace Tool v$($script:ToolVersion) - Important Notice"
        Width="620" SizeToContent="Height"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        FontFamily="Segoe UI" FontSize="13" Background="White">
    <Window.Resources>
        <Style x:Key="AcceptBtn" TargetType="Button">
            <Setter Property="Background" Value="#107C10"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}" CornerRadius="3">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#0E6E0E"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#0A560A"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="bd" Property="Background" Value="#BFD9BF"/>
                                <Setter Property="Foreground" Value="#F0F5F0"/>
                                <Setter Property="Cursor" Value="Arrow"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Border Padding="24,20">
        <StackPanel>
            <TextBlock Text="IMPORTANT NOTICE — PLEASE READ CAREFULLY"
                       FontWeight="Bold" FontSize="14" Margin="0,0,0,14"/>
            <ScrollViewer Height="240" VerticalScrollBarVisibility="Auto"
                          BorderBrush="#E0E0E0" BorderThickness="1" Margin="0,0,0,14">
                <TextBlock Padding="10" TextWrapping="Wrap" Foreground="#1F1F1F" xml:space="preserve">This Cloud PC Replace Tool is provided AS-IS without warranty or support.

By using this tool, you acknowledge and agree that:

  • This is an UNOFFICIAL, COMMUNITY-DEVELOPED tool
  • It is NOT supported, endorsed, or warranted by anyone
  • You use this tool entirely AT YOUR OWN RISK
  • You are responsible for understanding what this tool does
  • You have had the opportunity to review the source code
  • You should TEST in a non-production environment first
  • You accept full responsibility for any consequences
  • Changes made by this tool affect production Cloud PCs
  • There is no built-in rollback mechanism
  • You have proper backups and disaster recovery plans

This tool automates Cloud PC Replace operations using the Graph API. It will:
  • Remove users from groups
  • End grace periods on Cloud PCs
  • Deprovision existing Cloud PCs
  • Add users to new groups
  • Trigger new Cloud PC provisioning

These are PERMANENT operations that affect user productivity.

Recommended: Review the source code in Start-CloudPCReplaceWPF.ps1 and CloudPCReplace.psm1 before proceeding.</TextBlock>
            </ScrollViewer>
            <CheckBox x:Name="chkAccept" Margin="0,0,0,14"
                      Content="I have read and understand the above. I accept all risks and responsibilities."
                      FontWeight="SemiBold"/>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="btnAccept" Content="Accept" Width="90" Height="32"
                        Style="{StaticResource AcceptBtn}" IsEnabled="False" Margin="0,0,8,0"/>
                <Button x:Name="btnDecline" Content="Decline" Width="90" Height="32"
                        Background="#E0E0E0" Foreground="#1F1F1F" BorderThickness="0"
                        Cursor="Hand"/>
            </StackPanel>
            <TextBlock x:Name="lblMockNote" Text="[MOCK MODE] Auto-accepting in 2 seconds..."
                       FontSize="11" Foreground="#FF8C00" HorizontalAlignment="Center"
                       Margin="0,10,0,0"
                       Visibility="$(if($script:MockMode){'Visible'}else{'Collapsed'})"/>
        </StackPanel>
    </Border>
</Window>
"@
$discReader = [System.Xml.XmlNodeReader]::new($discXaml)
$discWin    = [System.Windows.Markup.XamlReader]::Load($discReader)
$discAccept = $discWin.FindName('btnAccept')
$discDecline= $discWin.FindName('btnDecline')
$discCheck  = $discWin.FindName('chkAccept')
$discCheck.Add_Checked({   $discAccept.IsEnabled = $true  })
$discCheck.Add_Unchecked({ $discAccept.IsEnabled = $false })
$discAccept.Add_Click([System.Windows.RoutedEventHandler]{ param($s,$e)
    [System.Windows.Window]::GetWindow($s).DialogResult = $true })
$discDecline.Add_Click([System.Windows.RoutedEventHandler]{ param($s,$e)
    [System.Windows.Window]::GetWindow($s).DialogResult = $false })

if ($script:MockMode) {
    # Auto-accept after 2 seconds — use .Tag to pass references, avoiding closure capture
    $discWin.Add_ContentRendered({
        param($wnd, $e)
        $t = [System.Windows.Threading.DispatcherTimer]::new()
        $t.Interval = [TimeSpan]::FromSeconds(2)
        $t.Tag = $wnd   # store window on timer so Tick can reach it
        $t.Add_Tick({
            param($timer, $e)
            $timer.Stop()
            $w = $timer.Tag
            if ($w -and $w.IsLoaded) { $w.DialogResult = $true }
        })
        $t.Start()
    })
}

if (-not $discWin.ShowDialog()) {
    Write-Host "User declined terms. Exiting." -ForegroundColor Yellow
    exit
}

# Launch
[void]$script:Window.ShowDialog()
