#Requires -Version 5.1

using module .\CloudPCReplace.psm1

# Version - Single source of truth
# TO UPDATE VERSION: Change this one value and it updates everywhere (title bar, docs, etc.)
$script:ToolVersion = "4.3.0"

<#
.SYNOPSIS
    Cloud PC Replace GUI - Replace Windows 365 Cloud PCs between provisioning profiles

.DESCRIPTION
    Replaces users' Cloud PCs from one provisioning profile to another by managing
    group memberships. The tool deprovisions the old Cloud PC and provisions a new one
    with a different configuration (network, image, region, join type, etc.).
    
    Common use cases:
    - Hybrid Joined → Entra Joined (or reverse)
    - Different network configuration
    - Different region/data center
    - Different image or applications
    - Different size/performance tier

.NOTES
    Author: Cloud PC Replace Tool
    Version: 4.3.0
    Date: 2026-02-27
    
    Changelog v4.3 (Modular Refactoring):
    - Core business logic moved to module for code reuse
    - Unified 6-character logging tags (Action/Detect/Info/Debug/Poll)
    - Enhanced CPC identity tracking with OLD/NEW labels
    - Improved debug logging for troubleshooting
    - Module architecture ready for WPF GUI development
    - Deprovision skip bug fix (CRITICAL - license reuse)
    - Timeout adjustments (60/60/90 min with Warning state)
    - Configuration change logging
    - Button layout fixes (no overlap)
    - Smooth countdown displays (updates every 3s)
    - Unicode character fixes
    - Comprehensive logging audit and cleanup
    
    Changelog v4.0:
    - Queue-based architecture with independent jobs
    - Dynamic concurrency control
    - Queue management (reorder, remove)
    - Messages column with timeline tracking
    - Real-time user filtering
    - Layout improvements (Source/Target/Users panels)
    - API resilience (flip-flop handling)
    - Non-blocking monitoring stage
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Import module functions (using module above only loads classes)
$modulePath = Join-Path $PSScriptRoot "CloudPCReplace.psm1"
Import-Module $modulePath -Force

#region Global Variables
$script:cancellationToken = $false
$script:replaceRunning = $false
$script:sourceGroupId = $null
$script:targetGroupId = $null
$script:userStates = @{}
$script:maxConcurrent = 2
$script:verboseLogging = $false  # Toggle for DEBUG/INFO/POLLING messages

# Timeout settings (in minutes)
$script:gracePeriodTimeoutMinutes = 15
$script:endingGracePeriodTimeoutMinutes = 30
$script:deprovisionTimeoutMinutes = 60
$script:provisioningTimeoutMinutes = 90
#endregion

#region Helper Functions
function Write-VerboseLog {
    param(
        [string]$Message, 
        [string]$Color = "Gray"
    )
    $timestamp = Get-Date -Format "HH:mm:ss"
    $logMessage = "[$timestamp] $Message"
    
    # Always write to log file
    if ($script:LogFilePath) {
        Add-Content -Path $script:LogFilePath -Value $logMessage
    }
    
    # Only write to console if verbose logging is enabled
    if ($script:verboseLogging) {
        Write-Host $logMessage -ForegroundColor $Color
    }
}

function Write-DebugLog {
    param(
        [string]$Message,
        [string]$Color = "DarkGray"
    )
    $timestamp = Get-Date -Format "HH:mm:ss"
    $logMessage = "[$timestamp] $Message"
    
    # Always write to log file
    if ($script:LogFilePath) {
        Add-Content -Path $script:LogFilePath -Value $logMessage
    }
    
    # Only write to console if verbose logging is enabled
    if ($script:verboseLogging) {
        Write-Host $logMessage -ForegroundColor $Color
    }
}

function Write-InfoLog {
    param(
        [string]$Message,
        [string]$Color = "Cyan"
    )
    $timestamp = Get-Date -Format "HH:mm:ss"
    $logMessage = "[$timestamp] $Message"
    
    # Always write to log file
    if ($script:LogFilePath) {
        Add-Content -Path $script:LogFilePath -Value $logMessage
    }
    
    # Only write to console if verbose logging is enabled
    if ($script:verboseLogging) {
        Write-Host $logMessage -ForegroundColor $Color
    }
}

function Write-PollingLog {
    param(
        [string]$Message,
        [string]$Color = "DarkGray"
    )
    $timestamp = Get-Date -Format "HH:mm:ss"
    $logMessage = "[$timestamp] $Message"
    
    # Always write to log file
    if ($script:LogFilePath) {
        Add-Content -Path $script:LogFilePath -Value $logMessage
    }
    
    # Only write to console if verbose logging is enabled
    if ($script:verboseLogging) {
        Write-Host $logMessage -ForegroundColor $Color
    }
}

# Helper for important stage/status messages (always visible)
function Write-StatusLog {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    $timestamp = Get-Date -Format "HH:mm:ss"
    $logMessage = "[$timestamp] $Message"
    
    # Always write to log file
    if ($script:LogFilePath) {
        Add-Content -Path $script:LogFilePath -Value $logMessage
    }
    
    # Always write to console (not affected by verbose setting)
    Write-Host $logMessage -ForegroundColor $Color
}

# Export logging functions for module use
$script:WriteDebugLog = ${function:Write-DebugLog}
$script:WriteInfoLog = ${function:Write-InfoLog}
$script:WritePollingLog = ${function:Write-PollingLog}

function Get-NextPollDisplay {
    param(
        [string]$Stage,
        [datetime]$StageStartTime,
        [datetime]$LastPollTime
    )
    
    $now = Get-Date
    $secondsSinceLastPoll = ($now - $LastPollTime).TotalSeconds
    
    # Determine if this is a long-running stage with timeout display
    $timeout = $null
    if ($Stage -eq "Waiting for Grace Period") {
        $timeout = $script:gracePeriodTimeoutMinutes
    }
    elseif ($Stage -eq "Ending Grace Period") {
        $timeout = $script:endingGracePeriodTimeoutMinutes
    }
    elseif ($Stage -eq "Waiting for Deprovision") {
        $timeout = $script:deprovisionTimeoutMinutes
    }
    elseif ($Stage -eq "Waiting for Provisioning") {
        $timeout = $script:provisioningTimeoutMinutes
    }
    
    if ($timeout) {
        # Show both countdown to next poll AND elapsed time
        $secondsUntilNextPoll = [math]::Max(0, 60 - $secondsSinceLastPoll)
        $elapsed = ($now - $StageStartTime)
        $elapsedMin = [math]::Floor($elapsed.TotalMinutes)
        $result = "$([math]::Round($secondsUntilNextPoll))s $elapsedMin/$($timeout)m"
        
        # Get caller info for debugging
        $caller = (Get-PSCallStack)[1].Command
        Write-DebugLog "[DEBUG] NextPoll from $caller : Stage='$Stage', Elapsed=$([math]::Round($elapsed.TotalMinutes,1))m, Result='$result'" "DarkGray"
        return $result
    }
    else {
        # Normal countdown only
        $secondsUntilNextPoll = [math]::Max(0, 60 - $secondsSinceLastPoll)
        return "$([math]::Round($secondsUntilNextPoll))s"
    }
}

function Export-JobSummary {
    param(
        [string]$FilePath,
        [switch]$Append
    )
    
    # Get ALL jobs (queued, in progress, completed)
    $allJobs = $script:userStates.Values | Sort-Object QueueOrder
    
    if ($allJobs.Count -eq 0) {
        Write-InfoLog "[INFO] No jobs to export" "Yellow"
        return
    }
    
    # Prepare CSV data - expand jobs with multiple old CPCs into separate rows
    $csvData = @()
    foreach ($job in $allJobs) {
        # Split old CPC names if multiple
        $oldCPCNames = if ($job.OldCPCName) { $job.OldCPCName -split ', ' } else { @('') }
        
        # Create one row per old CPC (or one row if no old CPCs)
        foreach ($oldCPC in $oldCPCNames) {
            $csvData += [PSCustomObject]@{
                User = $job.UserPrincipalName
                SourceGroup = $job.SourceGroupName
                OldCPC = $oldCPC
                TargetGroup = $job.TargetGroupName
                NewCPC = $job.NewCPCName
                Status = $job.Status
                Stage = $job.Stage
                StartTime = if ($job.StartTime -and $job.StartTime -gt [DateTime]::MinValue) { $job.StartTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "" }
                EndTime = if ($job.EndTime -and $job.EndTime -gt [DateTime]::MinValue) { $job.EndTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "" }
                Message = if ($job.FinalMessage) { $job.FinalMessage } else { $job.ErrorMessage }
            }
        }
    }
    
    # Export to CSV
    if ($Append -and (Test-Path $FilePath)) {
        $csvData | Export-Csv -Path $FilePath -NoTypeInformation -Append
    } else {
        $csvData | Export-Csv -Path $FilePath -NoTypeInformation
    }
    
    Write-StatusLog "[SUCCESS] Exported $($csvData.Count) row(s) from $($allJobs.Count) job(s) to: $FilePath" -Color Green
}

function Save-JobSummaryAuto {
    param([UserReplaceState]$State)
    
    # Auto-save to timestamped file in script directory
    $timestamp = Get-Date -Format "yyyyMMdd"
    $csvPath = Join-Path $PSScriptRoot "CloudPC-Replace-Summary-$timestamp.csv"
    
    # Match old and new CPCs by service plan, creating one row per pair
    $csvRows = @()
    
    if ($State.OldCPCs.Count -gt 0 -and $State.NewCPCs.Count -gt 0) {
        # Group by service plan and match
        $oldByPlan = $State.OldCPCs | Group-Object ServicePlan
        $newByPlan = $State.NewCPCs | Group-Object ServicePlan
        
        foreach ($oldGroup in $oldByPlan) {
            $newGroup = $newByPlan | Where-Object { $_.Name -eq $oldGroup.Name }
            
            if ($newGroup) {
                # Pair old and new CPCs of the same plan (by index)
                for ($i = 0; $i -lt $oldGroup.Count; $i++) {
                    $oldCPC = $oldGroup.Group[$i]
                    $newCPC = if ($i -lt $newGroup.Count) { $newGroup.Group[$i] } else { $null }
                    
                    $csvRows += [PSCustomObject]@{
                        User = $State.UserPrincipalName
                        SourceGroup = $State.SourceGroupName
                        OldCPC = "$($oldCPC.Name) ($($oldCPC.ServicePlan))"
                        TargetGroup = $State.TargetGroupName
                        NewCPC = if ($newCPC) { "$($newCPC.Name) ($($newCPC.ServicePlan))" } else { "" }
                        Status = $State.Status
                        Stage = $State.Stage
                        StartTime = if ($State.StartTime -and $State.StartTime -gt [DateTime]::MinValue) { $State.StartTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "" }
                        EndTime = if ($State.EndTime -and $State.EndTime -gt [DateTime]::MinValue) { $State.EndTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "" }
                        Message = $State.FinalMessage
                    }
                }
            }
        }
    }
    elseif ($State.OldCPCs.Count -gt 0) {
        # Job failed before new CPCs provisioned - still export old CPCs
        foreach ($oldCPC in $State.OldCPCs) {
            $csvRows += [PSCustomObject]@{
                User = $State.UserPrincipalName
                SourceGroup = $State.SourceGroupName
                OldCPC = "$($oldCPC.Name) ($($oldCPC.ServicePlan))"
                TargetGroup = $State.TargetGroupName
                NewCPC = ""
                Status = $State.Status
                Stage = $State.Stage
                StartTime = if ($State.StartTime -and $State.StartTime -gt [DateTime]::MinValue) { $State.StartTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "" }
                EndTime = if ($State.EndTime -and $State.EndTime -gt [DateTime]::MinValue) { $State.EndTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "" }
                Message = if ($State.FinalMessage) { $State.FinalMessage } else { $State.ErrorMessage }
            }
        }
    }
    else {
        # Fallback - user had no old CPCs
        $csvRows += [PSCustomObject]@{
            User = $State.UserPrincipalName
            SourceGroup = $State.SourceGroupName
            OldCPC = ""
            TargetGroup = $State.TargetGroupName
            NewCPC = if ($State.NewCPCs.Count -gt 0) { $State.NewCPCs[0].Name } else { "" }
            Status = $State.Status
            Stage = $State.Stage
            StartTime = if ($State.StartTime -and $State.StartTime -gt [DateTime]::MinValue) { $State.StartTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "" }
            EndTime = if ($State.EndTime -and $State.EndTime -gt [DateTime]::MinValue) { $State.EndTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "" }
            Message = if ($State.FinalMessage) { $State.FinalMessage } else { $State.ErrorMessage }
        }
    }
    
    # Append to daily CSV file
    if (Test-Path $csvPath) {
        $csvRows | Export-Csv -Path $csvPath -NoTypeInformation -Append
    } else {
        $csvRows | Export-Csv -Path $csvPath -NoTypeInformation
    }
    
    Write-VerboseLog "[DEBUG] Auto-saved $($csvRows.Count) row(s) for job to: $csvPath" "Gray"
}
#endregion

#region Helper Functions
# Helper function to update NextPoll display value based on state
function Update-NextPollDisplay {
    param([UserReplaceState]$State)
    
    # Don't update if job is complete/failed - keep the "-" value
    if ($State.Status -in @("Success", "Success (Warnings)", "Failed", "Warning")) {
        $State.NextPollDisplay = "-"
        return
    }
    
    $now = Get-Date
    $secondsSinceLastPoll = ($now - $State.LastPollTime).TotalSeconds
    
    # Determine polling interval based on stage
    $pollInterval = if ($State.Stage -eq "Waiting for Provisioning") { 180 } else { 60 }
    $secondsUntilNextPoll = [math]::Max(0, $pollInterval - $secondsSinceLastPoll)
    
    # Check if this is a long-running stage with timeout tracking
    $isWaitingStage = $State.Stage -in @("Waiting for Grace Period", "Ending Grace Period", "Waiting for Deprovision", "Waiting for Provisioning")
    $isImmediateStage = $State.Stage -in @("Getting User Info", "Getting Current Cloud PC", "Removing from Source", "Adding to Target")
    
    if ($isImmediateStage) {
        $State.NextPollDisplay = "Now"
    }
    elseif ($isWaitingStage) {
        # Show countdown + elapsed/timeout
        $elapsed = ($now - $State.StageStartTime)
        $elapsedMin = [math]::Floor($elapsed.TotalMinutes)
        
        $timeout = switch ($State.Stage) {
            "Waiting for Grace Period" { $script:gracePeriodTimeoutMinutes }
            "Ending Grace Period" { $script:endingGracePeriodTimeoutMinutes }
            "Waiting for Deprovision" { $script:deprovisionTimeoutMinutes }
            "Waiting for Provisioning" { $script:provisioningTimeoutMinutes }
        }
        
        # Always show seconds countdown for smooth updates
        # Format: "177s 5/90m" or "45s 12/90m"
        $State.NextPollDisplay = "$([math]::Round($secondsUntilNextPoll))s $elapsedMin/$($timeout)m"
    }
    else {
        # Other stages (shouldn't normally see these, but fallback)
        $State.NextPollDisplay = "$([math]::Round($secondsUntilNextPoll))s"
    }
}
#endregion

# Helper function to update grid based on queue order
function Update-QueueGridOrder {
    param(
        [string]$KeepSelection  # UPN to re-select after rebuild
    )
    
    # Get all jobs sorted by queue order
    $sortedJobs = $script:userStates.Values | Sort-Object QueueOrder
    
    # Clear and rebuild grid in order
    $script:GridStatus.Rows.Clear()
    
    $rowToSelect = -1
    $rowIndex = 0
    
    foreach ($state in $sortedJobs) {
        # Add new row
        $rowIndex = $script:GridStatus.Rows.Add()
        
        # Translate internal status to display status
        $displayStatus = if ($state.Status -eq "InProgress") {
            if ($state.Stage -eq "Waiting for Provisioning") { "Monitoring" } else { "Active" }
        } else {
            $state.Status
        }
        
        $script:GridStatus.Rows[$rowIndex].Cells["User"].Value = $state.UserPrincipalName
        $script:GridStatus.Rows[$rowIndex].Cells["Source"].Value = $state.SourceGroupName
        $script:GridStatus.Rows[$rowIndex].Cells["Target"].Value = $state.TargetGroupName
        $script:GridStatus.Rows[$rowIndex].Cells["Stage"].Value = Get-StageDisplay $state.Stage
        $script:GridStatus.Rows[$rowIndex].Cells["Status"].Value = $displayStatus
        $script:GridStatus.Rows[$rowIndex].Cells["NextPoll"].Value = $state.NextPollDisplay
        $script:GridStatus.Rows[$rowIndex].Cells["Messages"].Value = $state.ErrorMessage
        
        # Set row color based on status
        $color = switch ($state.Status) {
            "Queued" { [System.Drawing.Color]::LightGray }
            "InProgress" { 
                if ($state.Stage -eq "Waiting for Provisioning") {
                    [System.Drawing.Color]::LightCyan
                } else {
                    [System.Drawing.Color]::LightBlue
                }
            }
            "Success" { [System.Drawing.Color]::LightGreen }
            "Failed" { [System.Drawing.Color]::LightCoral }
            default { [System.Drawing.Color]::White }
        }
        $script:GridStatus.Rows[$rowIndex].DefaultCellStyle.BackColor = $color
        
        # Update the GridRowIndex in state
        $state.GridRowIndex = $rowIndex
        
        # Track which row to re-select
        if ($KeepSelection -and $state.UserPrincipalName -eq $KeepSelection) {
            $rowToSelect = $rowIndex
        }
    }
    
    # Re-select the row if requested
    if ($rowToSelect -ge 0) {
        $script:GridStatus.ClearSelection()
        $script:GridStatus.Rows[$rowToSelect].Selected = $true
    }
}

function Show-ReplaceGUI {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Cloud PC Replace Tool v$script:ToolVersion"
    $form.Size = New-Object System.Drawing.Size(1200, 780)
    $form.MinimumSize = New-Object System.Drawing.Size(1200, 780)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "Sizable"
    $form.MaximizeBox = $true
    
    # Connection & Settings
    $connectionGroup = New-Object System.Windows.Forms.GroupBox
    $connectionGroup.Location = New-Object System.Drawing.Point(10, 10)
    $connectionGroup.Size = New-Object System.Drawing.Size(1160, 60)
    $connectionGroup.Anchor = 'Top,Left,Right'
    $connectionGroup.Text = "Connection & Settings"
    
    # Connection status label (will show tenant info after connection)
    $lblConnectionStatus = New-Object System.Windows.Forms.Label
    $lblConnectionStatus.Location = New-Object System.Drawing.Point(10, 25)
    $lblConnectionStatus.Size = New-Object System.Drawing.Size(650, 20)
    $lblConnectionStatus.Text = "Not connected"
    $lblConnectionStatus.ForeColor = [System.Drawing.Color]::Gray
    $connectionGroup.Controls.Add($lblConnectionStatus)
    
    $lblConcurrency = New-Object System.Windows.Forms.Label
    $lblConcurrency.Location = New-Object System.Drawing.Point(670, 25)
    $lblConcurrency.Size = New-Object System.Drawing.Size(100, 20)
    $lblConcurrency.Text = "Max Concurrent:"
    $connectionGroup.Controls.Add($lblConcurrency)
    
    $numConcurrency = New-Object System.Windows.Forms.NumericUpDown
    $numConcurrency.Location = New-Object System.Drawing.Point(775, 23)
    $numConcurrency.Size = New-Object System.Drawing.Size(60, 20)
    $numConcurrency.Minimum = 1
    $numConcurrency.Maximum = 40
    $numConcurrency.Value = 2
    $script:numConcurrency = $numConcurrency  # Store in script scope for timer access
    $numConcurrency.Add_ValueChanged({
        $oldValue = if ($script:lastConcurrentValue) { $script:lastConcurrentValue } else { 2 }
        $newValue = $numConcurrency.Value
        if ($oldValue -ne $newValue) {
            Write-StatusLog "[CONFIG] Max concurrent changed: $oldValue -> $newValue" -Color Cyan
            $script:lastConcurrentValue = $newValue
        }
    })
    $script:lastConcurrentValue = 2
    $connectionGroup.Controls.Add($numConcurrency)
    
    $chkVerboseLogging = New-Object System.Windows.Forms.CheckBox
    $chkVerboseLogging.Location = New-Object System.Drawing.Point(850, 23)
    $chkVerboseLogging.Size = New-Object System.Drawing.Size(140, 20)
    $chkVerboseLogging.Text = "Verbose Logging"
    $chkVerboseLogging.Checked = $false
    $script:verboseLogging = $false
    $chkVerboseLogging.Add_CheckedChanged({
        $script:verboseLogging = $chkVerboseLogging.Checked
        # Update module's verbose logging setting too
        (Get-Module CloudPCReplace).Invoke({ $script:verboseLogging = $args[0] }, $chkVerboseLogging.Checked)
        $status = if ($script:verboseLogging) { "enabled" } else { "disabled" }
        Write-StatusLog "[LOGGING] Verbose logging $status" -Color $(if ($script:verboseLogging) { "Yellow" } else { "Gray" })
    })
    $connectionGroup.Controls.Add($chkVerboseLogging)
    
    $btnConnect = New-Object System.Windows.Forms.Button
    $btnConnect.Location = New-Object System.Drawing.Point(1000, 20)
    $btnConnect.Size = New-Object System.Drawing.Size(150, 28)
    $btnConnect.Anchor = 'Top,Right'
    $btnConnect.Text = "Connect to Graph"
    $btnConnect.BackColor = [System.Drawing.Color]::LightBlue
    $connectionGroup.Controls.Add($btnConnect)
    
    $form.Controls.Add($connectionGroup)
    
    # Source Group
    $sourceGroup = New-Object System.Windows.Forms.GroupBox
    $sourceGroup.Location = New-Object System.Drawing.Point(10, 80)
    $sourceGroup.Size = New-Object System.Drawing.Size(575, 135)
    $sourceGroup.Anchor = 'Top,Left'
    $sourceGroup.Text = "Source Group (Current Provisioning Profile)"
    
    $txtSearchSource = New-Object System.Windows.Forms.TextBox
    $txtSearchSource.Location = New-Object System.Drawing.Point(10, 23)
    $txtSearchSource.Size = New-Object System.Drawing.Size(435, 20)
    $txtSearchSource.Anchor = 'Top,Left,Right'
    $txtSearchSource.Add_KeyDown({
        param($sender, $e)
        if ($e.KeyCode -eq 'Enter') {
            $e.SuppressKeyPress = $true
            $btnSearchSource.PerformClick()
        }
        elseif ($e.Control -and $e.KeyCode -eq 'A') {
            $sender.SelectAll()
            $e.SuppressKeyPress = $true
        }
    })
    $sourceGroup.Controls.Add($txtSearchSource)
    
    $btnSearchSource = New-Object System.Windows.Forms.Button
    $btnSearchSource.Location = New-Object System.Drawing.Point(455, 21)
    $btnSearchSource.Size = New-Object System.Drawing.Size(100, 25)
    $btnSearchSource.Anchor = 'Top,Right'
    $btnSearchSource.Text = "Search"
    $btnSearchSource.Enabled = $false
    $sourceGroup.Controls.Add($btnSearchSource)
    
    $lstGroupsSource = New-Object System.Windows.Forms.ListBox
    $lstGroupsSource.Location = New-Object System.Drawing.Point(10, 50)
    $lstGroupsSource.Size = New-Object System.Drawing.Size(545, 80)
    $lstGroupsSource.Anchor = 'Top,Left,Right,Bottom'
    $lstGroupsSource.DisplayMember = "displayName"
    $sourceGroup.Controls.Add($lstGroupsSource)
    
    $form.Controls.Add($sourceGroup)
    
    # Target Group  
    $targetGroup = New-Object System.Windows.Forms.GroupBox
    $targetGroup.Location = New-Object System.Drawing.Point(10, 225)
    $targetGroup.Size = New-Object System.Drawing.Size(575, 135)
    $targetGroup.Anchor = 'Top,Left'
    $targetGroup.Text = "Target Group (New Provisioning Profile)"
    
    $txtSearchTarget = New-Object System.Windows.Forms.TextBox
    $txtSearchTarget.Location = New-Object System.Drawing.Point(10, 23)
    $txtSearchTarget.Size = New-Object System.Drawing.Size(435, 20)
    $txtSearchTarget.Anchor = 'Top,Left,Right'
    $txtSearchTarget.Add_KeyDown({
        param($sender, $e)
        if ($e.KeyCode -eq 'Enter') {
            $e.SuppressKeyPress = $true
            $btnSearchTarget.PerformClick()
        }
        elseif ($e.Control -and $e.KeyCode -eq 'A') {
            $sender.SelectAll()
            $e.SuppressKeyPress = $true
        }
    })
    $targetGroup.Controls.Add($txtSearchTarget)
    
    $btnSearchTarget = New-Object System.Windows.Forms.Button
    $btnSearchTarget.Location = New-Object System.Drawing.Point(455, 21)
    $btnSearchTarget.Size = New-Object System.Drawing.Size(100, 25)
    $btnSearchTarget.Anchor = 'Top,Right'
    $btnSearchTarget.Text = "Search"
    $btnSearchTarget.Enabled = $false
    $targetGroup.Controls.Add($btnSearchTarget)
    
    $lstGroupsTarget = New-Object System.Windows.Forms.ListBox
    $lstGroupsTarget.Location = New-Object System.Drawing.Point(10, 50)
    $lstGroupsTarget.Size = New-Object System.Drawing.Size(545, 80)
    $lstGroupsTarget.Anchor = 'Top,Left,Right,Bottom'
    $lstGroupsTarget.DisplayMember = "displayName"
    $targetGroup.Controls.Add($lstGroupsTarget)
    
    $form.Controls.Add($targetGroup)
    
    # Users Group (right column)
    $usersGroup = New-Object System.Windows.Forms.GroupBox
    $usersGroup.Location = New-Object System.Drawing.Point(595, 80)
    $usersGroup.Size = New-Object System.Drawing.Size(575, 280)
    $usersGroup.Anchor = 'Top,Right'
    $usersGroup.Text = "Users (from Source Group)"
    
    $txtSearchUsers = New-Object System.Windows.Forms.TextBox
    $txtSearchUsers.Location = New-Object System.Drawing.Point(10, 23)
    $txtSearchUsers.Size = New-Object System.Drawing.Size(435, 20)
    $txtSearchUsers.Anchor = 'Top,Left,Right'
    $txtSearchUsers.Add_TextChanged({
        $searchText = $txtSearchUsers.Text.Trim()
        
        # Store current check states
        $checkedUsers = @{}
        for ($i = 0; $i -lt $chkUsersSource.Items.Count; $i++) {
            if ($chkUsersSource.GetItemChecked($i)) {
                $user = $chkUsersSource.Items[$i]
                $checkedUsers[$user.userPrincipalName] = $true
            }
        }
        
        # Clear and repopulate with filtered list
        $chkUsersSource.Items.Clear()
        
        if ([string]::IsNullOrWhiteSpace($searchText)) {
            # No filter - show all users
            if ($script:allSourceUsers) {
                foreach ($user in $script:allSourceUsers) {
                    $index = $chkUsersSource.Items.Add($user)
                    if ($checkedUsers.ContainsKey($user.userPrincipalName)) {
                        $chkUsersSource.SetItemChecked($index, $true)
                    }
                }
            }
        } else {
            # Filter users
            if ($script:allSourceUsers) {
                foreach ($user in $script:allSourceUsers) {
                    if ($user.displayName -like "*$searchText*" -or $user.userPrincipalName -like "*$searchText*") {
                        $index = $chkUsersSource.Items.Add($user)
                        if ($checkedUsers.ContainsKey($user.userPrincipalName)) {
                            $chkUsersSource.SetItemChecked($index, $true)
                        }
                    }
                }
            }
        }
        
        # Update count label
        $checkedCount = ($chkUsersSource.CheckedItems | Measure-Object).Count
        $totalCount = ($chkUsersSource.Items | Measure-Object).Count
        if ($searchText) {
            $lblUsersSource.Text = "Users: $totalCount (filtered from $($script:allSourceUsers.Count)) | Selected: $checkedCount"
        } else {
            $lblUsersSource.Text = "Users: $totalCount | Selected: $checkedCount"
        }
    })
    $txtSearchUsers.Add_KeyDown({
        param($sender, $e)
        if ($e.KeyCode -eq 'Enter') {
            $e.SuppressKeyPress = $true
            $btnSearchUsers.PerformClick()
        }
        elseif ($e.Control -and $e.KeyCode -eq 'A') {
            $sender.SelectAll()
            $e.SuppressKeyPress = $true
        }
    })
    $usersGroup.Controls.Add($txtSearchUsers)
    
    $btnSearchUsers = New-Object System.Windows.Forms.Button
    $btnSearchUsers.Location = New-Object System.Drawing.Point(455, 21)
    $btnSearchUsers.Size = New-Object System.Drawing.Size(100, 25)
    $btnSearchUsers.Anchor = 'Top,Right'
    $btnSearchUsers.Text = "Clear"
    $btnSearchUsers.Add_Click({
        $txtSearchUsers.Clear()
        $txtSearchUsers.Focus()
    })
    $usersGroup.Controls.Add($btnSearchUsers)
    
    $lblUsersSource = New-Object System.Windows.Forms.Label
    $lblUsersSource.Location = New-Object System.Drawing.Point(10, 50)
    $lblUsersSource.Size = New-Object System.Drawing.Size(300, 15)
    $lblUsersSource.Text = "Users: 0"
    $usersGroup.Controls.Add($lblUsersSource)
    
    # Select All/None as clickable labels (right-aligned)
    $lblSelectAll = New-Object System.Windows.Forms.Label
    $lblSelectAll.Location = New-Object System.Drawing.Point(430, 50)
    $lblSelectAll.Size = New-Object System.Drawing.Size(60, 15)
    $lblSelectAll.Anchor = 'Top,Right'
    $lblSelectAll.Text = "Select All"
    $lblSelectAll.ForeColor = [System.Drawing.Color]::Blue
    $lblSelectAll.Cursor = [System.Windows.Forms.Cursors]::Hand
    $lblSelectAll.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Underline)
    $usersGroup.Controls.Add($lblSelectAll)
    
    $lblSelectNone = New-Object System.Windows.Forms.Label
    $lblSelectNone.Location = New-Object System.Drawing.Point(490, 50)
    $lblSelectNone.Size = New-Object System.Drawing.Size(70, 15)
    $lblSelectNone.Anchor = 'Top,Right'
    $lblSelectNone.Text = "Select None"
    $lblSelectNone.ForeColor = [System.Drawing.Color]::Blue
    $lblSelectNone.Cursor = [System.Windows.Forms.Cursors]::Hand
    $lblSelectNone.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Underline)
    $usersGroup.Controls.Add($lblSelectNone)
    
    $chkUsersSource = New-Object System.Windows.Forms.CheckedListBox
    $chkUsersSource.Location = New-Object System.Drawing.Point(10, 68)
    $chkUsersSource.Size = New-Object System.Drawing.Size(545, 207)
    $chkUsersSource.Anchor = 'Top,Left,Right,Bottom'
    $chkUsersSource.CheckOnClick = $true
    $chkUsersSource.Add_ItemCheck({
        param($sender, $e)
        # ItemCheck fires BEFORE the state changes, so calculate what the count WILL be
        $currentChecked = ($chkUsersSource.CheckedItems | Measure-Object).Count
        $willBeChecked = if ($e.NewValue -eq 'Checked') { $currentChecked + 1 } else { $currentChecked - 1 }
        $totalCount = ($chkUsersSource.Items | Measure-Object).Count
        
        $searchText = $txtSearchUsers.Text.Trim()
        if ($searchText -and $script:allSourceUsers) {
            $lblUsersSource.Text = "Users: $totalCount (filtered from $($script:allSourceUsers.Count)) | Selected: $willBeChecked"
        } else {
            $lblUsersSource.Text = "Users: $totalCount | Selected: $willBeChecked"
        }
    })
    $usersGroup.Controls.Add($chkUsersSource)
    
    $form.Controls.Add($usersGroup)
    
    # Queue Management
    $queueMgmtGroup = New-Object System.Windows.Forms.GroupBox
    $queueMgmtGroup.Location = New-Object System.Drawing.Point(10, 370)
    $queueMgmtGroup.Size = New-Object System.Drawing.Size(1160, 60)
    $queueMgmtGroup.Anchor = 'Top,Left,Right'
    $queueMgmtGroup.Text = "Queue Management"
    
    # Standard button height and Y position for consistent alignment
    $btnHeight = 30
    $btnY = 15  # Centers 30px button in 60px groupbox
    
    # Group 1: Queue operations
    $btnAddToQueue = New-Object System.Windows.Forms.Button
    $btnAddToQueue.Location = New-Object System.Drawing.Point(10, $btnY)
    $btnAddToQueue.Size = New-Object System.Drawing.Size(95, $btnHeight)
    $btnAddToQueue.Text = "Add to Queue"
    $btnAddToQueue.Enabled = $false
    $btnAddToQueue.BackColor = [System.Drawing.Color]::LightBlue
    $btnAddToQueue.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $queueMgmtGroup.Controls.Add($btnAddToQueue)
    
    $btnRemoveSelected = New-Object System.Windows.Forms.Button
    $btnRemoveSelected.Location = New-Object System.Drawing.Point(115, $btnY)
    $btnRemoveSelected.Size = New-Object System.Drawing.Size(85, $btnHeight)
    $btnRemoveSelected.Text = "Remove"
    $btnRemoveSelected.Enabled = $false
    $btnRemoveSelected.BackColor = [System.Drawing.Color]::LightGray
    $btnRemoveSelected.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $queueMgmtGroup.Controls.Add($btnRemoveSelected)
    
    $btnClearQueue = New-Object System.Windows.Forms.Button
    $btnClearQueue.Location = New-Object System.Drawing.Point(210, $btnY)
    $btnClearQueue.Size = New-Object System.Drawing.Size(95, $btnHeight)
    $btnClearQueue.Text = "Clear Queue"
    $btnClearQueue.Enabled = $false
    $btnClearQueue.BackColor = [System.Drawing.Color]::LightYellow
    $btnClearQueue.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $queueMgmtGroup.Controls.Add($btnClearQueue)
    
    # Group 2: Reorder buttons (30px gap from Group 1)
    $btnMoveUp = New-Object System.Windows.Forms.Button
    $btnMoveUp.Location = New-Object System.Drawing.Point(335, $btnY)
    $btnMoveUp.Size = New-Object System.Drawing.Size(75, $btnHeight)
    $btnMoveUp.Text = "Move Up"
    $btnMoveUp.Enabled = $false
    $queueMgmtGroup.Controls.Add($btnMoveUp)
    
    $btnMoveDown = New-Object System.Windows.Forms.Button
    $btnMoveDown.Location = New-Object System.Drawing.Point(420, $btnY)
    $btnMoveDown.Size = New-Object System.Drawing.Size(90, $btnHeight)
    $btnMoveDown.Text = "Move Down"
    $btnMoveDown.Enabled = $false
    $queueMgmtGroup.Controls.Add($btnMoveDown)
    
    $btnMoveTop = New-Object System.Windows.Forms.Button
    $btnMoveTop.Location = New-Object System.Drawing.Point(520, $btnY)
    $btnMoveTop.Size = New-Object System.Drawing.Size(65, $btnHeight)
    $btnMoveTop.Text = "To Top"
    $btnMoveTop.Enabled = $false
    $queueMgmtGroup.Controls.Add($btnMoveTop)
    
    $btnMoveBottom = New-Object System.Windows.Forms.Button
    $btnMoveBottom.Location = New-Object System.Drawing.Point(595, $btnY)
    $btnMoveBottom.Size = New-Object System.Drawing.Size(85, $btnHeight)
    $btnMoveBottom.Text = "To Bottom"
    $btnMoveBottom.Enabled = $false
    $queueMgmtGroup.Controls.Add($btnMoveBottom)
    
    # Group 3: Processing controls (30px gap from Group 2)
    $btnStartProcessing = New-Object System.Windows.Forms.Button
    $btnStartProcessing.Location = New-Object System.Drawing.Point(710, $btnY)
    $btnStartProcessing.Size = New-Object System.Drawing.Size(120, $btnHeight)
    $btnStartProcessing.Text = "Start Processing"
    $btnStartProcessing.Enabled = $false
    $btnStartProcessing.BackColor = [System.Drawing.Color]::LightGreen
    $btnStartProcessing.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $queueMgmtGroup.Controls.Add($btnStartProcessing)
    
    $btnStopProcessing = New-Object System.Windows.Forms.Button
    $btnStopProcessing.Location = New-Object System.Drawing.Point(840, $btnY)
    $btnStopProcessing.Size = New-Object System.Drawing.Size(120, $btnHeight)
    $btnStopProcessing.Text = "Stop Processing"
    $btnStopProcessing.Enabled = $false
    $btnStopProcessing.BackColor = [System.Drawing.Color]::LightCoral
    $btnStopProcessing.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $queueMgmtGroup.Controls.Add($btnStopProcessing)
    
    # Group 4: Export (30px gap from Group 3)
    $btnExportSummary = New-Object System.Windows.Forms.Button
    $btnExportSummary.Location = New-Object System.Drawing.Point(990, $btnY)
    $btnExportSummary.Size = New-Object System.Drawing.Size(115, $btnHeight)
    $btnExportSummary.Text = "Export Summary"
    $btnExportSummary.Enabled = $true
    $btnExportSummary.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $queueMgmtGroup.Controls.Add($btnExportSummary)
    
    $form.Controls.Add($queueMgmtGroup)
    
    # Status Grid
    $statusGroup = New-Object System.Windows.Forms.GroupBox
    $statusGroup.Location = New-Object System.Drawing.Point(10, 440)
    $statusGroup.Size = New-Object System.Drawing.Size(1160, 290)
    $statusGroup.Anchor = 'Top,Bottom,Left,Right'
    $statusGroup.Text = "Replace Status (Updates every 3 seconds)"
    
    $gridStatus = New-Object System.Windows.Forms.DataGridView
    $script:GridStatus = $gridStatus  # Store in script scope for Update-QueueGridOrder function
    $gridStatus.Location = New-Object System.Drawing.Point(10, 25)
    $gridStatus.Size = New-Object System.Drawing.Size(1135, 225)
    $gridStatus.Anchor = 'Top,Bottom,Left,Right'
    $gridStatus.AllowUserToAddRows = $false
    $gridStatus.AllowUserToDeleteRows = $false
    $gridStatus.ReadOnly = $true
    $gridStatus.SelectionMode = "FullRowSelect"
    $gridStatus.MultiSelect = $false
    $gridStatus.AutoSizeColumnsMode = "Fill"
    $gridStatus.RowHeadersVisible = $false
    
    $colUser = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colUser.Name = "User"
    $colUser.HeaderText = "User"
    $colUser.FillWeight = 19
    $gridStatus.Columns.Add($colUser) | Out-Null
    
    $colSource = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colSource.Name = "Source"
    $colSource.HeaderText = "Source Group"
    $colSource.FillWeight = 14
    $gridStatus.Columns.Add($colSource) | Out-Null
    
    $colTarget = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colTarget.Name = "Target"
    $colTarget.HeaderText = "Target Group"
    $colTarget.FillWeight = 14
    $gridStatus.Columns.Add($colTarget) | Out-Null
    
    $colStage = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colStage.Name = "Stage"
    $colStage.HeaderText = "Stage"
    $colStage.FillWeight = 15
    $gridStatus.Columns.Add($colStage) | Out-Null
    
    $colStatus = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colStatus.Name = "Status"
    $colStatus.HeaderText = "Status"
    $colStatus.FillWeight = 7
    $gridStatus.Columns.Add($colStatus) | Out-Null
    
    $colNextPoll = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colNextPoll.Name = "NextPoll"
    $colNextPoll.HeaderText = "Next Poll"
    $colNextPoll.FillWeight = 10
    $gridStatus.Columns.Add($colNextPoll) | Out-Null
    
    $colMessages = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMessages.Name = "Messages"
    $colMessages.HeaderText = "Messages"
    $colMessages.FillWeight = 21
    $gridStatus.Columns.Add($colMessages) | Out-Null
    
    $statusGroup.Controls.Add($gridStatus)
    
    $lblProgressSummary = New-Object System.Windows.Forms.Label
    $lblProgressSummary.Location = New-Object System.Drawing.Point(10, 258)
    $lblProgressSummary.Size = New-Object System.Drawing.Size(1135, 25)
    $lblProgressSummary.Anchor = 'Bottom,Left,Right'
    $lblProgressSummary.Text = "Ready"
    $lblProgressSummary.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $statusGroup.Controls.Add($lblProgressSummary)
    
    $form.Controls.Add($statusGroup)
    
    
    # Timer - Runs every 3 seconds
    $processTimer = New-Object System.Windows.Forms.Timer
    $processTimer.Interval = 3000
    $processTimer.Enabled = $false  # Don't start until replace begins
    
    Write-Host "Timer created with interval: $($processTimer.Interval)ms" -ForegroundColor Cyan
    
    # Timer tick - runs every 3 seconds
    $script:timerTickCount = 0
    $processTimer.Add_Tick({
        $script:timerTickCount++
        # Timer tick - console only, don't log to file (fires every 3 seconds)
        if ($script:verboseLogging) {
            $timestamp = Get-Date -Format "HH:mm:ss"
            Write-Host "[$timestamp] [TIMER TICK #$($script:timerTickCount)] Starting at $(Get-Date -Format 'HH:mm:ss.fff') | Interval: $($processTimer.Interval)ms" -ForegroundColor Magenta
        }
        
        # Guard against timer firing before replace starts or after completion
        if (-not $script:replaceRunning -or -not $script:ProgressLabel) {
            return
        }
        
        $now = Get-Date
        $activeCount = 0
        $queuedCount = 0
        $monitoringCount = 0
        $successCount = 0
        $failedCount = 0
        
        # Count status
        foreach ($state in $script:userStates.Values) {
            switch ($state.Status) {
                "InProgress" { 
                    # Jobs in "Waiting for Provisioning" don't consume concurrency slots
                    # They're just monitoring - all admin work is done
                    if ($state.Stage -eq "Waiting for Provisioning") {
                        $monitoringCount++
                    }
                    else {
                        $activeCount++
                    }
                }
                "Queued" { $queuedCount++ }
                "Success" { $successCount++ }
                "Success (Warnings)" { $successCount++ }
                "Failed" { $failedCount++ }
            }
        }
        
        # Start new users if under concurrency limit
        # Read current concurrency setting from control (allows dynamic adjustment)
        $currentMaxConcurrent = if ($script:numConcurrency) { $script:numConcurrency.Value } else { 1 }
        # Only count jobs doing active admin work (not monitoring provisioning)
        if ($activeCount -lt $currentMaxConcurrent -and $queuedCount -gt 0 -and -not $script:cancellationToken) {
            $toStart = $currentMaxConcurrent - $activeCount
            # Select queued jobs in order (by QueueOrder property)
            $queuedUsers = $script:userStates.Values | Where-Object { $_.Status -eq "Queued" } | Sort-Object QueueOrder | Select-Object -First $toStart
            
            foreach ($state in $queuedUsers) {
                $state.Status = "InProgress"
                $state.Stage = "Getting User Info"
                $state.ProgressPercent = 5
                $state.StageStartTime = $now
                $state.LastPollTime = $now
                
                # Add "Started at" message
                $startTime = Get-Date -Format "HH:mm:ss"
                $state.ErrorMessage = "Started at $startTime"
                Write-VerboseLog "[DEBUG] Setting 'Started at $startTime' message for $($state.UserPrincipalName) at row $($state.GridRowIndex)" "Cyan"
                $script:GridStatus.Rows[$state.GridRowIndex].Cells["Messages"].Value = $state.ErrorMessage
                
                $script:GridStatus.Rows[$state.GridRowIndex].Cells["Status"].Value = "Active"
                $script:GridStatus.Rows[$state.GridRowIndex].Cells["Stage"].Value = Get-StageDisplay $state.Stage
                $script:GridStatus.Rows[$state.GridRowIndex].DefaultCellStyle.BackColor = [System.Drawing.Color]::LightBlue
            }
        }
        
        # Process each ACTIVE user - check if it's time to poll
        $gridStatus.SuspendLayout()
        try {
            foreach ($state in $script:userStates.Values) {
                if ($state.Status -ne "InProgress") { continue }
                
                $secondsSinceLastPoll = ($now - $state.LastPollTime).TotalSeconds
                
                # Immediate stages don't need polling delay
                $isImmediateStage = $state.Stage -in @("Getting User Info", "Getting Current Cloud PC", "Removing from Source", "Adding to Target")
                
                # Determine polling interval based on stage
                # Monitoring (provisioning) uses 3-minute interval since it's very slow (40-90 min)
                # Other waiting stages use 1-minute interval for more responsive feedback
                $pollInterval = if ($state.Stage -eq "Waiting for Provisioning") { 180 } else { 60 }
                
                # Is it time to process this user?
                if ($isImmediateStage -or $secondsSinceLastPoll -ge $pollInterval) {
                    # Call module function with callbacks
                    $timeouts = @{
                        GracePeriodTimeout = $script:gracePeriodTimeoutMinutes
                        EndingGracePeriodTimeout = $script:endingGracePeriodTimeoutMinutes
                        DeprovisionTimeout = $script:deprovisionTimeoutMinutes
                        ProvisioningTimeout = $script:provisioningTimeoutMinutes
                    }
                    
                    $onLog = {
                        param($Message, $Level, $Color)
                        switch ($Level) {
                            "Status" { Write-StatusLog $Message -Color $Color }
                            "Info" { Write-InfoLog $Message $Color }
                            "Debug" { Write-DebugLog $Message $Color }
                            "Polling" { Write-PollingLog $Message $Color }
                            "Verbose" { Write-VerboseLog $Message $Color }
                            default { Write-Host $Message -ForegroundColor $Color }
                        }
                    }
                    
                    $onGridUpdate = {
                        param($State, $ColumnName, $Value)
                        if ($State.GridRowIndex -ge 0 -and $State.GridRowIndex -lt $script:GridStatus.Rows.Count) {
                            $script:GridStatus.Rows[$State.GridRowIndex].Cells[$ColumnName].Value = $Value
                        }
                    }
                    
                    Invoke-CloudPCReplaceStep -State $state -Timeouts $timeouts -OnLog $onLog -OnGridUpdate $onGridUpdate
                    $state.LastPollTime = $now
                    
                    # Update grid with new state values (stage, status, messages)
                    $displayStatus = if ($state.Status -eq "InProgress") {
                        if ($state.Stage -eq "Waiting for Provisioning") { "Monitoring" } else { "Active" }
                    } else {
                        $state.Status
                    }
                    
                    $script:GridStatus.Rows[$state.GridRowIndex].Cells["Stage"].Value = Get-StageDisplay $state.Stage
                    $script:GridStatus.Rows[$state.GridRowIndex].Cells["Status"].Value = $displayStatus
                    
                    if ($state.ErrorMessage) {
                        $script:GridStatus.Rows[$state.GridRowIndex].Cells["Messages"].Value = $state.ErrorMessage
                    }
                    
                    # Update row color based on status
                    $color = switch ($state.Status) {
                        "Success" { [System.Drawing.Color]::LightGreen }
                        "Success (Warnings)" { [System.Drawing.Color]::LightYellow }
                        "Failed" { [System.Drawing.Color]::LightCoral }
                        "Warning" { [System.Drawing.Color]::Orange }
                        "InProgress" {
                            if ($state.Stage -eq "Waiting for Provisioning") {
                                [System.Drawing.Color]::LightCyan
                            } else {
                                [System.Drawing.Color]::LightBlue
                            }
                        }
                        default { [System.Drawing.Color]::White }
                    }
                    $script:GridStatus.Rows[$state.GridRowIndex].DefaultCellStyle.BackColor = $color
                    
                    # If job completed, save summary
                    if ($state.Status -in @("Success", "Success (Warnings)", "Failed", "Warning")) {
                        Save-JobSummaryAuto -State $state
                    }
                }
                
                # Update the NextPoll display value based on current state
                Update-NextPollDisplay -State $state
                $script:GridStatus.Rows[$state.GridRowIndex].Cells["NextPoll"].Value = $state.NextPollDisplay
            }
        }
        finally {
            $gridStatus.ResumeLayout()
        }
        
        # Update summary
        $totalCount = $script:userStates.Count
        $completedCount = $successCount + $failedCount
        $currentMaxConcurrent = if ($script:numConcurrency) { $script:numConcurrency.Value } else { 1 }
        $script:ProgressLabel.Text = "Total: $totalCount | Queued: $queuedCount | Active: $activeCount (max: $currentMaxConcurrent) | Monitoring: $monitoringCount | Complete: $completedCount | Success: $successCount | Failed: $failedCount"
        
        # Check if done (monitoring jobs still need to complete)
        if ($completedCount -eq $totalCount -and $queuedCount -eq 0 -and $monitoringCount -eq 0) {
            $script:ProcessTimer.Stop()
            $script:replaceRunning = $false
            $btnStartProcessing.Enabled = $false
            $btnStopProcessing.Enabled = $false
            
            # Update progress summary with completion status
            if ($failedCount -eq 0) {
                $script:ProgressLabel.Text = "COMPLETE - All $successCount job(s) completed successfully!"
                $script:ProgressLabel.ForeColor = [System.Drawing.Color]::DarkGreen
            }
            elseif ($successCount -eq 0) {
                $script:ProgressLabel.Text = "COMPLETE - All $failedCount job(s) failed"
                $script:ProgressLabel.ForeColor = [System.Drawing.Color]::DarkRed
            }
            else {
                $script:ProgressLabel.Text = "COMPLETE - Success: $successCount | Failed: $failedCount"
                $script:ProgressLabel.ForeColor = [System.Drawing.Color]::DarkOrange
            }
            
            Write-Host "`n=== PROCESSING COMPLETE ===" -ForegroundColor Cyan
            Write-Host "Success: $successCount | Failed: $failedCount" -ForegroundColor $(if ($failedCount -eq 0) { 'Green' } else { 'Yellow' })
        }
    })
    
    # Event Handlers
    $btnConnect.Add_Click({
        try {
            # Check if already connected - if so, disconnect
            $currentContext = Get-MgContext
            if ($currentContext -and $btnConnect.Text -eq "Disconnect") {
                Disconnect-MgGraph | Out-Null
                
                $btnConnect.BackColor = [System.Drawing.Color]::LightBlue
                $btnConnect.Text = "Connect to Graph"
                $btnConnect.Enabled = $true
                $btnSearchSource.Enabled = $false
                $btnSearchTarget.Enabled = $false
                
                $lblConnectionStatus.Text = "Disconnected"
                $lblConnectionStatus.ForeColor = [System.Drawing.Color]::Gray
                
                                return
            }
            
            # Connect
            $btnConnect.Enabled = $false
                        $form.Refresh()
            
            $connected = Connect-MgGraphForReplace -TenantId $null -AuthMethod 'Interactive'
            
            if ($connected) {
                $btnConnect.BackColor = [System.Drawing.Color]::LightCoral
                $btnConnect.Text = "Disconnect"
                $btnConnect.Enabled = $true
                $btnSearchSource.Enabled = $true
                $btnSearchTarget.Enabled = $true
                
                # Get and display connection info
                $context = Get-MgContext
                if ($context) {
                    $lblConnectionStatus.Text = "Connected | Tenant: $($context.TenantId) | Account: $($context.Account)"
                    $lblConnectionStatus.ForeColor = [System.Drawing.Color]::DarkGreen
                    
                                                                            } else {
                    $lblConnectionStatus.Text = "Connected"
                    $lblConnectionStatus.ForeColor = [System.Drawing.Color]::DarkGreen
                                    }
            }
            else {
                $btnConnect.Enabled = $true
                $btnConnect.BackColor = [System.Drawing.Color]::LightBlue
                $btnConnect.Text = "Connect to Graph"
                $lblConnectionStatus.Text = "Connection failed"
                $lblConnectionStatus.ForeColor = [System.Drawing.Color]::Red
                            }
        }
        catch {
            $btnConnect.Enabled = $true
            $btnConnect.BackColor = [System.Drawing.Color]::LightBlue
            $btnConnect.Text = "Connect to Graph"
            $lblConnectionStatus.Text = "Connection error"
            $lblConnectionStatus.ForeColor = [System.Drawing.Color]::Red
            Write-StatusLog "[ERROR] Connect button error: $_" -Color Red
        }
    })
    
    $btnSearchSource.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtSearchSource.Text)) { return }
        
        $btnSearchSource.Enabled = $false
        $lstGroupsSource.Items.Clear()
        $chkUsersSource.Items.Clear()
        $lstGroupsSource.Tag = $null
        
        Write-StatusLog "[Search] Searching source groups for: '$($txtSearchSource.Text)'" -Color Cyan
        
        try {
            $groups = Find-EntraIDGroups -SearchTerm $txtSearchSource.Text
            $groupArray = [System.Collections.ArrayList]::new()
            
            # Sort groups alphabetically by displayName (case-insensitive)
            $sortedGroups = $groups | Sort-Object -Property { $_.displayName.ToLower() }
            
            # TEST MODE: Uncomment the next line to simulate AD-synced groups for testing
            # if ($sortedGroups.Count -gt 0) { $sortedGroups[0].onPremisesSyncEnabled = $true }
            
            foreach ($group in $sortedGroups) {
                # Create a display string with warning for AD-synced or dynamic groups
                $adSyncWarning = if ($group.onPremisesSyncEnabled -eq $true) { " *!AD-Synced!*" } else { "" }
                $dynamicWarning = if ($group.groupTypes -and $group.groupTypes -contains 'DynamicMembership') { " *!Dynamic!*" } else { "" }
                $item = $group.displayName + $adSyncWarning + $dynamicWarning + " (" + $group.id + ")"
                $lstGroupsSource.Items.Add($item) | Out-Null
                # Store the actual group hashtable for retrieval later
                $groupArray.Add($group) | Out-Null
            }
            
            $lstGroupsSource.Tag = $groupArray.ToArray()
            Write-StatusLog "[Search] Found $($groups.Count) group(s) matching '$($txtSearchSource.Text)'" -Color Green
        }
        catch {
            Write-StatusLog "[Search] Error searching for '$($txtSearchSource.Text)': $_" -Color Red
        }
        finally {
            $btnSearchSource.Enabled = $true
        }
    })
    
    $btnSearchTarget.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtSearchTarget.Text)) { return }
        
        $btnSearchTarget.Enabled = $false
        $lstGroupsTarget.Items.Clear()
        $lstGroupsTarget.Tag = $null
        
        Write-StatusLog "[Search] Searching target groups for: '$($txtSearchTarget.Text)'" -Color Cyan
        
        try {
            $groups = Find-EntraIDGroups -SearchTerm $txtSearchTarget.Text
            $groupArray = [System.Collections.ArrayList]::new()
            
            # Sort groups alphabetically by displayName (case-insensitive)
            $sortedGroups = $groups | Sort-Object -Property { $_.displayName.ToLower() }
            
            # TEST MODE: Uncomment the next line to simulate AD-synced groups for testing
            # if ($sortedGroups.Count -gt 0) { $sortedGroups[0].onPremisesSyncEnabled = $true }
            
            foreach ($group in $sortedGroups) {
                # Create a display string with warning for AD-synced or dynamic groups
                $adSyncWarning = if ($group.onPremisesSyncEnabled -eq $true) { " *!AD-Synced!*" } else { "" }
                $dynamicWarning = if ($group.groupTypes -and $group.groupTypes -contains 'DynamicMembership') { " *!Dynamic!*" } else { "" }
                $item = $group.displayName + $adSyncWarning + $dynamicWarning + " (" + $group.id + ")"
                $lstGroupsTarget.Items.Add($item) | Out-Null
                # Store the actual group hashtable for retrieval later
                $groupArray.Add($group) | Out-Null
            }
            
            $lstGroupsTarget.Tag = $groupArray.ToArray()
            Write-StatusLog "[Search] Found $($groups.Count) group(s) matching '$($txtSearchTarget.Text)'" -Color Green
        }
        catch {
            Write-StatusLog "[Search] Error searching for '$($txtSearchTarget.Text)': $_" -Color Red
        }
        finally {
            $btnSearchTarget.Enabled = $true
        }
    })
    
    $lstGroupsSource.Add_SelectedIndexChanged({
        if ($lstGroupsSource.SelectedIndex -ge 0) {
            # Retrieve the actual group object from the Tag array
            $selectedGroup = $lstGroupsSource.Tag[$lstGroupsSource.SelectedIndex]
            
            # Validate group is not AD-synced
            if ($selectedGroup.onPremisesSyncEnabled -eq $true) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Cannot Use AD-Synced Group`n`n" +
                    "The selected group '$($selectedGroup.displayName)' is synchronized from Active Directory and cannot be modified via Microsoft Graph API.`n`n" +
                    "This tool requires Entra Cloud Groups where membership can be changed programmatically.`n`n" +
                    "SOLUTION: Create a cloud-only group in Microsoft Entra ID and assign your Cloud PC provisioning policy to it instead.`n`n" +
                    "See README for details on group requirements.",
                    "AD-Synced Group Not Supported",
                    "OK",
                    "Warning"
                )
                $lstGroupsSource.ClearSelected()
                $script:sourceGroupId = $null
                return
            }

            # Validate group is not a dynamic group
            if ($selectedGroup.groupTypes -and $selectedGroup.groupTypes -contains 'DynamicMembership') {
                [System.Windows.Forms.MessageBox]::Show(
                    "Cannot Use Dynamic Group`n`n" +
                    "The selected group '$($selectedGroup.displayName)' is a Dynamic Membership group. Membership is controlled by rules and cannot be modified manually via Microsoft Graph API.`n`n" +
                    "This tool requires Assigned (static) Entra Cloud Groups where membership can be changed programmatically.`n`n" +
                    "SOLUTION: Create an Assigned (static) cloud-only group in Microsoft Entra ID and assign your Cloud PC provisioning policy to it instead.`n`n" +
                    "See README for details on group requirements.",
                    "Dynamic Group Not Supported",
                    "OK",
                    "Warning"
                )
                $lstGroupsSource.ClearSelected()
                $script:sourceGroupId = $null
                return
            }
            
            $script:sourceGroupId = $selectedGroup.Id
            $chkUsersSource.Items.Clear()
            $txtSearchUsers.Clear()
            
            Write-StatusLog "[Group ] Loading members from '$($selectedGroup.DisplayName)'..." -Color Cyan
            Write-StatusLog "[API   ] GET /groups/$($script:sourceGroupId)/members" -Color Yellow
            
            try {
                $members = Get-GroupMembers -GroupId $script:sourceGroupId
                # Store full list for filtering and wrap each user in a display object
                $script:allSourceUsers = @()
                
                # Sort members alphabetically by displayName (case-insensitive)
                $sortedMembers = $members | Sort-Object -Property { $_.displayName.ToLower() }
                
                foreach ($member in $sortedMembers) {
                    # Create a wrapper object with ToString() override
                    $userObj = [PSCustomObject]@{
                        displayName = $member.displayName
                        userPrincipalName = $member.userPrincipalName
                        id = $member.id
                        mail = $member.mail
                    }
                    # Add script method for ToString
                    $userObj | Add-Member -MemberType ScriptMethod -Name ToString -Value {
                        return "$($this.displayName) ($($this.userPrincipalName))"
                    } -Force
                    $script:allSourceUsers += $userObj
                }
                
                foreach ($userObj in $script:allSourceUsers) {
                    $chkUsersSource.Items.Add($userObj) | Out-Null
                }
                $lblUsersSource.Text = "Users: $($script:allSourceUsers.Count) | Selected: 0"
                Write-StatusLog "[Group ] Loaded $($script:allSourceUsers.Count) user(s) from '$($selectedGroup.DisplayName)'" -Color Green
            }
            catch {
                Write-StatusLog "[Group ] Error loading members: $_" -Color Red
                $script:allSourceUsers = @()
            }
        }
    })
    
    $lstGroupsTarget.Add_SelectedIndexChanged({
        if ($lstGroupsTarget.SelectedIndex -ge 0) {
            # Retrieve the actual group object from the Tag array
            $selectedGroup = $lstGroupsTarget.Tag[$lstGroupsTarget.SelectedIndex]
            
            # Validate group is not AD-synced
            if ($selectedGroup.onPremisesSyncEnabled -eq $true) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Cannot Use AD-Synced Group`n`n" +
                    "The selected group '$($selectedGroup.displayName)' is synchronized from Active Directory and cannot be modified via Microsoft Graph API.`n`n" +
                    "This tool requires Entra Cloud Groups where membership can be changed programmatically.`n`n" +
                    "SOLUTION: Create a cloud-only group in Microsoft Entra ID and assign your Cloud PC provisioning policy to it instead.`n`n" +
                    "See README for details on group requirements.",
                    "AD-Synced Group Not Supported",
                    "OK",
                    "Warning"
                )
                $lstGroupsTarget.ClearSelected()
                $script:targetGroupId = $null
                return
            }

            # Validate group is not a dynamic group
            if ($selectedGroup.groupTypes -and $selectedGroup.groupTypes -contains 'DynamicMembership') {
                [System.Windows.Forms.MessageBox]::Show(
                    "Cannot Use Dynamic Group`n`n" +
                    "The selected group '$($selectedGroup.displayName)' is a Dynamic Membership group. Membership is controlled by rules and cannot be modified manually via Microsoft Graph API.`n`n" +
                    "This tool requires Assigned (static) Entra Cloud Groups where membership can be changed programmatically.`n`n" +
                    "SOLUTION: Create an Assigned (static) cloud-only group in Microsoft Entra ID and assign your Cloud PC provisioning policy to it instead.`n`n" +
                    "See README for details on group requirements.",
                    "Dynamic Group Not Supported",
                    "OK",
                    "Warning"
                )
                $lstGroupsTarget.ClearSelected()
                $script:targetGroupId = $null
                return
            }
            
            $script:targetGroupId = $selectedGroup.Id
            Write-StatusLog "[Group ] Target group selected: '$($selectedGroup.DisplayName)'" -Color Cyan
            
            # Enable Start button if we have source, target, and checked users
            if (-not $script:replaceRunning) {
                $checkedCount = 0
                for ($i = 0; $i -lt $chkUsersSource.Items.Count; $i++) {
                    if ($chkUsersSource.GetItemChecked($i)) { $checkedCount++ }
                }
                $btnAddToQueue.Enabled = ($checkedCount -gt 0 -and $script:sourceGroupId -and $script:targetGroupId)
            }
        }
    })
    
    # Checkbox change handler - supports dynamic queue management
    $chkUsersSource.Add_ItemCheck({
        param($sender, $e)
        
        Write-VerboseLog "[DEBUG] ItemCheck fired - Index: $($e.Index), OldValue: $($e.CurrentValue), NewValue: $($e.NewValue)" "Yellow"
        
        $index = $e.Index
        $newValue = $e.NewValue
        
        # Validate index
        if ($index -lt 0 -or $index -ge $chkUsersSource.Items.Count) {
            Write-VerboseLog "[DEBUG] Index out of bounds, returning" "Red"
            return
        }
        
        $upn = $chkUsersSource.Items[$index] -replace '.*\((.+)\).*', '$1'
        $isChecked = ($newValue -eq 'Checked')
        
        Write-VerboseLog "[DEBUG] UPN: $upn, IsChecked: $isChecked" "Cyan"
        
        # If replace is running, prevent unchecking jobs that are InProgress
        if ($script:replaceRunning) {
            if ($script:userStates.ContainsKey($upn)) {
                $state = $script:userStates[$upn]
                
                # Prevent unchecking InProgress or completed users
                if (-not $isChecked -and $state.Status -ne 'Queued') {
                    # Can't modify jobs that are already running or complete
                    $chkUsersSource.SetItemChecked($index, $true)
                }
            }
        }
        
        # Update count - account for the current item's pending state change
        Write-VerboseLog "[DEBUG] Calculating count..." "Cyan"
        $count = 0
        for ($i = 0; $i -lt $chkUsersSource.Items.Count; $i++) {
            $willBeChecked = if ($i -eq $index) { 
                $newValue -eq 'Checked' 
            } else { 
                $chkUsersSource.GetItemChecked($i) 
            }
            if ($willBeChecked) { $count++ }
        }
        Write-VerboseLog "[DEBUG] Count: $count" "Green"
        
        # Update label removed - was in Summary box that no longer exists
        
        
        # Enable/disable add to queue button (always enabled when users are selected)
        $btnAddToQueue.Enabled = ($count -gt 0 -and $script:sourceGroupId -and $script:targetGroupId)
    })
    
    $lblSelectAll.Add_Click({
        for ($i = 0; $i -lt $chkUsersSource.Items.Count; $i++) {
            # During replace, only check items that aren't InProgress or completed
            if ($script:replaceRunning) {
                $upn = $chkUsersSource.Items[$i] -replace '.*\((.+)\).*', '$1'
                if ($script:userStates.ContainsKey($upn)) {
                    $state = $script:userStates[$upn]
                    if ($state.Status -ne 'Queued') {
                        continue  # Skip InProgress and completed users
                    }
                }
            }
            $chkUsersSource.SetItemChecked($i, $true)
        }
    })
    
    $lblSelectNone.Add_Click({
        for ($i = 0; $i -lt $chkUsersSource.Items.Count; $i++) {
            # During replace, only uncheck items that are Queued
            if ($script:replaceRunning) {
                $upn = $chkUsersSource.Items[$i] -replace '.*\((.+)\).*', '$1'
                if ($script:userStates.ContainsKey($upn)) {
                    $state = $script:userStates[$upn]
                    if ($state.Status -ne 'Queued') {
                        continue  # Skip InProgress and completed users
                    }
                }
            }
            $chkUsersSource.SetItemChecked($i, $false)
        }
    })
    
    # Add to Queue button - adds selected users to queue
    $btnAddToQueue.Add_Click({
        if (-not $script:sourceGroupId -or -not $script:targetGroupId) {
            [System.Windows.Forms.MessageBox]::Show("Please select both source and target groups", "Missing Selection", "OK", "Warning")
            return
        }
        
        if ($script:sourceGroupId -eq $script:targetGroupId) {
            [System.Windows.Forms.MessageBox]::Show("Source and Target groups must be different", "Invalid Selection", "OK", "Warning")
            return
        }
        
        $selectedUsers = @()
        for ($i = 0; $i -lt $chkUsersSource.Items.Count; $i++) {
            if ($chkUsersSource.GetItemChecked($i)) {
                $upn = $chkUsersSource.Items[$i] -replace '.*\((.+)\).*', '$1'
                $selectedUsers += $upn
            }
        }
        
        if ($selectedUsers.Count -eq 0) { return }
        
        # Get group names for display
        $sourceGroupName = $lstGroupsSource.SelectedItem -replace ' \(.*\)', ''
        $targetGroupName = $lstGroupsTarget.SelectedItem -replace ' \(.*\)', ''
        
        Write-StatusLog "[QUEUE] Adding $($selectedUsers.Count) job(s) to queue" -Color Cyan
        Write-Host "  Source: $sourceGroupName" -ForegroundColor Gray
        Write-Host "  Target: $targetGroupName" -ForegroundColor Gray
        
        $addedCount = 0
        foreach ($upn in $selectedUsers) {
            # Check if already queued or in progress (allow re-queuing completed/failed jobs)
            $existingKey = $script:userStates.Keys | Where-Object { 
                $script:userStates[$_].UserPrincipalName -eq $upn -and 
                ($script:userStates[$_].Status -eq "Queued" -or $script:userStates[$_].Status -eq "InProgress")
            }
            if ($existingKey) {
                Write-Host "  [SKIP] $upn - already queued or in progress" -ForegroundColor Yellow
                continue
            }
            
            # Create unique key for this job
            $jobKey = "$upn-$(Get-Date -Format 'HHmmssfff')"
            
            # Create state object
            $state = [UserReplaceState]::new()
            $state.UserPrincipalName = $upn
            $state.SourceGroupId = $script:sourceGroupId
            $state.SourceGroupName = $sourceGroupName
            $state.TargetGroupId = $script:targetGroupId
            $state.TargetGroupName = $targetGroupName
            $state.Stage = "Queued"
            $state.Status = "Queued"
            $state.ProgressPercent = 0
            $state.StartTime = Get-Date
            
            Write-VerboseLog "[DEBUG] Job ${upn}: SourceGroupId='$($state.SourceGroupId)', TargetGroupId='$($state.TargetGroupId)'" "Magenta"
            
            # Assign queue order (max + 1 for new jobs)
            $maxOrder = ($script:userStates.Values | Measure-Object -Property QueueOrder -Maximum).Maximum
            $state.QueueOrder = if ($maxOrder -ne $null) { $maxOrder + 1 } else { 0 }
            
            # Add to grid
            $rowIndex = $gridStatus.Rows.Add()
            $gridStatus.Rows[$rowIndex].Cells["User"].Value = $upn
            $gridStatus.Rows[$rowIndex].Cells["Source"].Value = $sourceGroupName
            $gridStatus.Rows[$rowIndex].Cells["Target"].Value = $targetGroupName
            $gridStatus.Rows[$rowIndex].Cells["Stage"].Value = Get-StageDisplay "Queued"
            $gridStatus.Rows[$rowIndex].Cells["Status"].Value = "Queued"
            $gridStatus.Rows[$rowIndex].Cells["NextPoll"].Value = "-"
            $gridStatus.Rows[$rowIndex].Cells["Messages"].Value = ""
            $gridStatus.Rows[$rowIndex].DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGray
            
            $state.GridRowIndex = $rowIndex
            $script:userStates[$jobKey] = $state
            
            Write-Host "  [ADDED] $upn" -ForegroundColor Green
            $addedCount++
        }
        
        Write-StatusLog "[QUEUE] Added $addedCount job(s)" -Color Green
        
        # Uncheck all users
        for ($i = 0; $i -lt $chkUsersSource.Items.Count; $i++) {
            $chkUsersSource.SetItemChecked($i, $false)
        }
        
        # Enable Start Processing button if we have queued jobs
        $queuedCount = ($script:userStates.Values | Where-Object { $_.Status -eq "Queued" }).Count
        $btnStartProcessing.Enabled = ($queuedCount -gt 0)
        $btnClearQueue.Enabled = ($queuedCount -gt 0)
        
        # Select first queued job to enable queue management buttons
        if ($addedCount -gt 0 -and $gridStatus.Rows.Count -gt 0) {
            $gridStatus.ClearSelection()
            $gridStatus.Rows[0].Selected = $true
        }
    })
    
    # Start Processing button
    $btnStartProcessing.Add_Click({
        $queuedJobs = $script:userStates.Values | Where-Object { $_.Status -eq "Queued" }
        if ($queuedJobs.Count -eq 0) { return }
        
        Write-Host "`n=== STARTING PROCESSING ===" -ForegroundColor Cyan
        Write-Host "Queued jobs: $($queuedJobs.Count)" -ForegroundColor White
        Write-Host "Max concurrent: $($numConcurrency.Value)" -ForegroundColor White
        Write-Host ""
        
        Start-ReplaceProcessing -MaxConcurrent $numConcurrency.Value `
            -GridStatus $gridStatus -ProgressLabel $lblProgressSummary `
            -ProcessTimer $processTimer
            
        $btnStartProcessing.Enabled = $false
        $btnStopProcessing.Enabled = $true
    })
    
    # Stop Processing button
    $btnStopProcessing.Add_Click({
        $script:cancellationToken = $true
        $btnStopProcessing.Enabled = $false
        Write-StatusLog "[STOP] Stopping processing (active jobs will complete)..." -Color Yellow
    })
    
    # Clear Queue button
    $btnClearQueue.Add_Click({
        $queuedJobs = $script:userStates.Values | Where-Object { $_.Status -eq "Queued" }
        if ($queuedJobs.Count -eq 0) { return }
        
        $confirmation = [System.Windows.Forms.MessageBox]::Show(
            "Clear $($queuedJobs.Count) queued job(s)?",
            "Confirm",
            "YesNo",
            "Question"
        )
        
        if ($confirmation -eq "No") { return }
        
        # Collect jobs to remove (with their row indices) - use PSCustomObject so Sort-Object works
        $jobsToRemove = @()
        foreach ($key in $script:userStates.Keys) {
            if ($script:userStates[$key].Status -eq "Queued") {
                $jobsToRemove += [PSCustomObject]@{
                    Key = $key
                    RowIndex = $script:userStates[$key].GridRowIndex
                    UPN = $script:userStates[$key].UserPrincipalName
                }
            }
        }
        
        Write-DebugLog "[DEBUG] Clear Queue: Found $($jobsToRemove.Count) jobs to remove" "Cyan"
        Write-DebugLog "[DEBUG] Current grid row count: $($gridStatus.Rows.Count)" "Cyan"
        
        # Remove rows from bottom to top (highest index first) to avoid index shifting issues
        $jobsToRemove = $jobsToRemove | Sort-Object -Property RowIndex -Descending
        foreach ($job in $jobsToRemove) {
            Write-DebugLog "[DEBUG] Attempting to remove row $($job.RowIndex) for $($job.UPN)" "Cyan"
            if ($job.RowIndex -ge 0 -and $job.RowIndex -lt $gridStatus.Rows.Count) {
                $gridStatus.Rows.RemoveAt($job.RowIndex)
                $script:userStates.Remove($job.Key)
                Write-DebugLog "[DEBUG] Successfully removed row $($job.RowIndex)" "Green"
            } else {
                Write-StatusLog "[WARNING] Skipping invalid row index $($job.RowIndex) for $($job.UPN)" -Color Yellow
            }
        }
        
        # Reindex remaining rows
        foreach ($key in $script:userStates.Keys) {
            $upn = $script:userStates[$key].UserPrincipalName
            for ($i = 0; $i -lt $gridStatus.Rows.Count; $i++) {
                if ($gridStatus.Rows[$i].Cells["User"].Value -eq $upn) {
                    $script:userStates[$key].GridRowIndex = $i
                    break
                }
            }
        }
        
        Write-StatusLog "[QUEUE] Cleared queue" -Color Yellow
        
        $btnClearQueue.Enabled = $false
        $btnStartProcessing.Enabled = $false
    })
    
    # Export Summary button
    $btnExportSummary.Add_Click({
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        $saveDialog.Title = "Export Job Summary"
        $saveDialog.FileName = "CloudPC-Replace-Summary-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
        $saveDialog.InitialDirectory = $PSScriptRoot
        
        if ($saveDialog.ShowDialog() -eq "OK") {
            Export-JobSummary -FilePath $saveDialog.FileName
        }
    })
    
    # Remove Selected button
    $btnRemoveSelected.Add_Click({
        if ($gridStatus.SelectedRows.Count -eq 0) { return }
        
        $selectedRow = $gridStatus.SelectedRows[0]
        $upn = $selectedRow.Cells["User"].Value
        
        # Find the job key
        $jobKey = $script:userStates.Keys | Where-Object { $script:userStates[$_].UserPrincipalName -eq $upn }
        if (-not $jobKey) { return }
        
        $state = $script:userStates[$jobKey]
        if ($state.Status -ne "Queued") {
            [System.Windows.Forms.MessageBox]::Show("Can only remove queued jobs (not started yet)", "Cannot Remove", "OK", "Information")
            return
        }
        
        # Remove from grid
        $gridStatus.Rows.RemoveAt($state.GridRowIndex)
        $script:userStates.Remove($jobKey)
        
        # Reindex
        foreach ($key in $script:userStates.Keys) {
            $u = $script:userStates[$key].UserPrincipalName
            for ($i = 0; $i -lt $gridStatus.Rows.Count; $i++) {
                if ($gridStatus.Rows[$i].Cells["User"].Value -eq $u) {
                    $script:userStates[$key].GridRowIndex = $i
                    break
                }
            }
        }
        
        Write-StatusLog "[QUEUE] Removed $upn from queue" -Color Yellow
        
        $queuedCount = ($script:userStates.Values | Where-Object { $_.Status -eq "Queued" }).Count
        $btnClearQueue.Enabled = ($queuedCount -gt 0)
        $btnStartProcessing.Enabled = ($queuedCount -gt 0)
    })
    
    # Move Up button (swap with previous)
    $btnMoveUp.Add_Click({
        if ($gridStatus.SelectedRows.Count -eq 0) { return }
        
        $selectedRow = $gridStatus.SelectedRows[0]
        $upn = $selectedRow.Cells["User"].Value
        $jobKey = $script:userStates.Keys | Where-Object { $script:userStates[$_].UserPrincipalName -eq $upn }
        if (-not $jobKey) { return }
        
        $state = $script:userStates[$jobKey]
        if ($state.Status -ne "Queued") { return }
        
        # Find the queued job immediately before this one
        $queuedJobs = @($script:userStates.Values | Where-Object { $_.Status -eq "Queued" } | Sort-Object QueueOrder)
        
        # Find current position in sorted array
        $currentIndex = -1
        for ($i = 0; $i -lt $queuedJobs.Count; $i++) {
            if ($queuedJobs[$i].UserPrincipalName -eq $upn) {
                $currentIndex = $i
                break
            }
        }
        
        if ($currentIndex -gt 0) {
            $previousJob = $queuedJobs[$currentIndex - 1]
            # Swap queue orders
            $tempOrder = $state.QueueOrder
            $state.QueueOrder = $previousJob.QueueOrder
            $previousJob.QueueOrder = $tempOrder
            
            Write-StatusLog "[QUEUE] Moved $upn up in queue" -Color Cyan
            Update-QueueGridOrder -KeepSelection $upn
        }
    })
    
    # Move Down button (swap with next)
    $btnMoveDown.Add_Click({
        if ($gridStatus.SelectedRows.Count -eq 0) { return }
        
        $selectedRow = $gridStatus.SelectedRows[0]
        $upn = $selectedRow.Cells["User"].Value
        $jobKey = $script:userStates.Keys | Where-Object { $script:userStates[$_].UserPrincipalName -eq $upn }
        if (-not $jobKey) { return }
        
        $state = $script:userStates[$jobKey]
        if ($state.Status -ne "Queued") { return }
        
        # Find the queued job immediately after this one
        $queuedJobs = @($script:userStates.Values | Where-Object { $_.Status -eq "Queued" } | Sort-Object QueueOrder)
        
        # Find current position in sorted array
        $currentIndex = -1
        for ($i = 0; $i -lt $queuedJobs.Count; $i++) {
            if ($queuedJobs[$i].UserPrincipalName -eq $upn) {
                $currentIndex = $i
                break
            }
        }
        
        if ($currentIndex -ge 0 -and $currentIndex -lt ($queuedJobs.Count - 1)) {
            $nextJob = $queuedJobs[$currentIndex + 1]
            # Swap queue orders
            $tempOrder = $state.QueueOrder
            $state.QueueOrder = $nextJob.QueueOrder
            $nextJob.QueueOrder = $tempOrder
            
            Write-StatusLog "[QUEUE] Moved $upn down in queue" -Color Cyan
            Update-QueueGridOrder -KeepSelection $upn
        }
    })
    
    # Move to Top button
    $btnMoveTop.Add_Click({
        if ($gridStatus.SelectedRows.Count -eq 0) { return }
        
        $selectedRow = $gridStatus.SelectedRows[0]
        $upn = $selectedRow.Cells["User"].Value
        $jobKey = $script:userStates.Keys | Where-Object { $script:userStates[$_].UserPrincipalName -eq $upn }
        if (-not $jobKey) { return }
        
        $state = $script:userStates[$jobKey]
        if ($state.Status -ne "Queued") { return }
        
        # Get minimum queue order from all queued jobs
        $queuedJobs = $script:userStates.Values | Where-Object { $_.Status -eq "Queued" -and $_.UserPrincipalName -ne $upn }
        $minOrder = ($queuedJobs | Measure-Object -Property QueueOrder -Minimum).Minimum
        
        if ($minOrder -ne $null) {
            $state.QueueOrder = $minOrder - 1
        } else {
            $state.QueueOrder = 0
        }
        
        Write-StatusLog "[QUEUE] Moved $upn to top of queue (order: $($state.QueueOrder))" -Color Cyan
        
        # Re-sort grid and keep selection
        Update-QueueGridOrder -KeepSelection $upn
    })
    
    # Move to Bottom button
    $btnMoveBottom.Add_Click({
        if ($gridStatus.SelectedRows.Count -eq 0) { return }
        
        $selectedRow = $gridStatus.SelectedRows[0]
        $upn = $selectedRow.Cells["User"].Value
        $jobKey = $script:userStates.Keys | Where-Object { $script:userStates[$_].UserPrincipalName -eq $upn }
        if (-not $jobKey) { return }
        
        $state = $script:userStates[$jobKey]
        if ($state.Status -ne "Queued") { return }
        
        # Get maximum queue order from all queued jobs
        $queuedJobs = $script:userStates.Values | Where-Object { $_.Status -eq "Queued" -and $_.UserPrincipalName -ne $upn }
        $maxOrder = ($queuedJobs | Measure-Object -Property QueueOrder -Maximum).Maximum
        
        if ($maxOrder -ne $null) {
            $state.QueueOrder = $maxOrder + 1
        } else {
            $state.QueueOrder = 0
        }
        
        Write-StatusLog "[QUEUE] Moved $upn to bottom of queue (order: $($state.QueueOrder))" -Color Cyan
        
        # Re-sort grid and keep selection
        Update-QueueGridOrder -KeepSelection $upn
    })
    
    # Grid selection changed - enable/disable Remove and Move buttons
    $gridStatus.Add_SelectionChanged({
        try {
            if ($gridStatus.SelectedRows.Count -gt 0) {
                $selectedRow = $gridStatus.SelectedRows[0]
                $upn = $selectedRow.Cells["User"].Value
                
                # Guard against null or empty UPN
                if ([string]::IsNullOrEmpty($upn)) {
                    $btnRemoveSelected.Enabled = $false
                    $btnMoveUp.Enabled = $false
                    $btnMoveDown.Enabled = $false
                    $btnMoveTop.Enabled = $false
                    $btnMoveBottom.Enabled = $false
                    return
                }
                
                $jobKey = $script:userStates.Keys | Where-Object { $script:userStates[$_].UserPrincipalName -eq $upn }
                if ($jobKey) {
                    $state = $script:userStates[$jobKey]
                    $isQueued = ($state.Status -eq "Queued")
                    $btnRemoveSelected.Enabled = $isQueued
                    $btnMoveUp.Enabled = $isQueued
                    $btnMoveDown.Enabled = $isQueued
                    $btnMoveTop.Enabled = $isQueued
                    $btnMoveBottom.Enabled = $isQueued
                } else {
                    $btnRemoveSelected.Enabled = $false
                    $btnMoveUp.Enabled = $false
                    $btnMoveDown.Enabled = $false
                    $btnMoveTop.Enabled = $false
                    $btnMoveBottom.Enabled = $false
                }
            }
            else {
                $btnRemoveSelected.Enabled = $false
                $btnMoveUp.Enabled = $false
                $btnMoveDown.Enabled = $false
                $btnMoveTop.Enabled = $false
                $btnMoveBottom.Enabled = $false
            }
        } catch {
            # Silently ignore errors during selection changes (e.g., during form initialization)
            $btnRemoveSelected.Enabled = $false
            $btnMoveUp.Enabled = $false
            $btnMoveDown.Enabled = $false
            $btnMoveTop.Enabled = $false
            $btnMoveBottom.Enabled = $false
        }
    })
    
    # Handle clicking on empty space in grid to deselect rows
    $gridStatus.Add_MouseDown({
        param($sender, $e)
        $hitTest = $gridStatus.HitTest($e.X, $e.Y)
        # If clicked outside of any row (in empty space), clear selection
        if ($hitTest.RowIndex -eq -1 -and $hitTest.ColumnIndex -eq -1) {
            $gridStatus.ClearSelection()
        }
    })
    
    $script:LogFilePath = Initialize-Logging -LogPath (Join-Path $PSScriptRoot "CloudPCReplace_Logs")
    
    # Handle window resize to adjust Source and Target boxes proportionally
    $form.Add_Resize({
        $formWidth = $form.ClientSize.Width
        $gap = 10
        $middleGap = 10
        
        # Each column gets half the available width (minus margins and gap)
        $halfWidth = [int](($formWidth - ($gap * 2) - $middleGap) / 2)
        
        # Left column: Source Group (width only, Target stays below it)
        $sourceGroup.Width = $halfWidth
        $targetGroup.Width = $halfWidth
        
        # Right column: Users Group (position and width)
        $usersGroup.Left = $gap + $halfWidth + $middleGap
        $usersGroup.Width = $halfWidth
    })
    
    # Ensure window shows on top and activates
    $form.Add_Shown({
        $form.Activate()
        $form.BringToFront()
        $form.TopMost = $true
        $form.TopMost = $false
    })
    
    $form.ShowDialog()
}

function Start-ReplaceProcessing {
    param($MaxConcurrent, $GridStatus, $ProgressLabel, $ProcessTimer)
    
    $script:cancellationToken = $false
    $script:replaceRunning = $true
    $script:maxConcurrent = $MaxConcurrent
    
    # Store UI controls in script scope so timer can access them
    $script:GridStatus = $GridStatus
    $script:ProgressLabel = $ProgressLabel
    $script:ProcessTimer = $ProcessTimer
    
    Write-VerboseLog "[DEBUG] Starting processing with $($script:userStates.Count) jobs in queue" "Magenta"
    
    # $script:userStates already populated by Add to Queue
    # Each job already has SourceGroupId/TargetGroupId set
    
    # Start timer
    $ProcessTimer.Start()
}

function Process-UserState {
    param($State, $GridStatus)
    
    Write-VerboseLog "[DEBUG] $($State.UserPrincipalName): Process-UserState called - SourceGroupId='$($State.SourceGroupId)', TargetGroupId='$($State.TargetGroupId)', Stage='$($State.Stage)'" "DarkCyan"
    
    try {
        switch ($State.Stage) {
            "Getting User Info" {
                Write-StatusLog "[STAGE] $($State.UserPrincipalName): Getting user info..." -Color Cyan
                
                # Get user using Graph API directly
                $uri = "https://graph.microsoft.com/v1.0/users?`$filter=userPrincipalName eq '$($State.UserPrincipalName)'&$select=id,userPrincipalName,displayName"
                $userResponse = Invoke-MgGraphRequest -Uri $uri -Method GET
                $user = $userResponse.value | Select-Object -First 1
                
                if (-not $user) { throw "User not found" }
                $State.UserId = $user.id
                Write-StatusLog "[STAGE] $($State.UserPrincipalName): User found (ID: $($State.UserId))" -Color Green
                
                $State.Stage = "Getting Current Cloud PC"
                $State.ProgressPercent = 10
            }
            
            "Getting Current Cloud PC" {
                Write-StatusLog "[STAGE] $($State.UserPrincipalName): Checking for existing Cloud PC..." -Color Cyan
                
                # Get provisioning policies for THIS job's source group (if not cached)
                if (-not $State.SourcePolicyIds) {
                    Write-DebugLog "[DEBUG] Getting provisioning policies for source group: $($State.SourceGroupName)" "Gray"
                    $State.SourcePolicyIds = @(Get-ProvisioningPoliciesForGroup -GroupId $State.SourceGroupId)
                    if ($State.SourcePolicyIds.Count -gt 0) {
                        Write-DebugLog "[DEBUG] Source group uses $($State.SourcePolicyIds.Count) provisioning policy(ies)" "Gray"
                    }
                }
                
                $cloudPCs = Get-CloudPCForUser -UserId $State.UserPrincipalName
                
                # Ensure it's an array for consistent counting
                $cloudPCsArray = @($cloudPCs)
                
                if ($cloudPCsArray.Count -gt 0) {
                    Write-DebugLog "[DEBUG] $($State.UserPrincipalName): Found $($cloudPCsArray.Count) Cloud PC(s) for user" "Yellow"
                    
                    foreach ($cpc in $cloudPCsArray) {
                        $cpcName = if ($cpc.managedDeviceName) { $cpc.managedDeviceName } else { $cpc.displayName }
                        Write-DebugLog "[DEBUG]   $($State.UserPrincipalName): CPC: $cpcName | Status: $($cpc.status) | ID: $($cpc.id)" "Yellow"
                        Write-DebugLog "[DEBUG]     $($State.UserPrincipalName): Policy: $($cpc.provisioningPolicyId)" "Yellow"
                    }
                    
                    # Find ALL Cloud PCs that match the SOURCE provisioning policy
                    $matchingCPCs = @()
                    if ($State.SourcePolicyIds -and $State.SourcePolicyIds.Count -gt 0) {
                        Write-DebugLog "[DEBUG] $($State.UserPrincipalName): Looking for CPCs matching source policy..." "Cyan"
                        $matchingCPCs = @($cloudPCsArray | Where-Object { $State.SourcePolicyIds -contains $_.provisioningPolicyId })
                    }
                    
                    if ($matchingCPCs.Count -gt 0) {
                        # Store old CPC details (ID, Name, ServicePlan) for matching later
                        $State.OldCPCs = @($matchingCPCs | ForEach-Object {
                            @{
                                Id = $_.id
                                Name = if ($_.managedDeviceName) { $_.managedDeviceName } else { $_.displayName }
                                ServicePlan = $_.servicePlanName
                            }
                        })
                        
                        # Also store IDs separately for backward compat with deprovision logic
                        $State.CloudPCIds = @($matchingCPCs | ForEach-Object { $_.id })
                        
                        # Capture old CPC name(s) for summary
                        $oldCPCNames = $matchingCPCs | ForEach-Object {
                            if ($_.managedDeviceName) { $_.managedDeviceName } else { $_.displayName }
                        }
                        $State.OldCPCName = $oldCPCNames -join ", "
                        
                        Write-StatusLog "[STAGE] $($State.UserPrincipalName): Found $($matchingCPCs.Count) CPC(s) from SOURCE policy" -Color Green
                        foreach ($cpc in $matchingCPCs) {
                            $cpcName = if ($cpc.managedDeviceName) { $cpc.managedDeviceName } else { $cpc.displayName }
                            Write-InfoLog "[INFO]   $($State.UserPrincipalName): Will deprovision: $cpcName (ID: $($cpc.id), Plan: $($cpc.servicePlanName))" "Cyan"
                        }
                        
                        if ($matchingCPCs.Count -gt 1) {
                            Write-StatusLog "[Info] $($State.UserPrincipalName): User has MULTIPLE CPCs from same policy (likely multiple licenses)" -Color Cyan
                            Write-InfoLog "[INFO] $($State.UserPrincipalName): All will be deprovisioned before provisioning new CPC(s)" "Cyan"
                        }
                    }
                    else {
                        Write-StatusLog "[WARNING] $($State.UserPrincipalName): Has $($cloudPCs.Count) CPC(s) but NONE match source policy!" -Color Yellow
                        Write-InfoLog "[INFO] Source policy IDs: $($State.SourcePolicyIds -join ', ')" "Yellow"
                        $State.OldCPCs = @()
                        $State.CloudPCIds = @()
                    }
                }
                else {
                    Write-StatusLog "[STAGE] $($State.UserPrincipalName): No existing Cloud PC found" -Color Yellow
                    $State.OldCPCs = @()
                    $State.CloudPCIds = @()
                }
                
                # Show current group memberships
                Write-DebugLog "[DEBUG] Current group memberships:" "Magenta"
                $currentGroups = Get-UserGroupMemberships -UserId $State.UserId
                foreach ($grp in $currentGroups) {
                    Write-DebugLog "[DEBUG]   - Member of: $($grp.displayName) (ID: $($grp.id))" "Magenta"
                }
                
                $State.Stage = "Removing from Source"
                $State.ProgressPercent = 20
                $State.StageStartTime = Get-Date
            }
            
            "Removing from Source" {
                # VALIDATE FIRST - only proceed if we have at least one CPC to track
                if (-not $State.CloudPCIds -or $State.CloudPCIds.Count -eq 0) {
                    Write-StatusLog "[ERROR] $($State.UserPrincipalName): User has NO Cloud PC matching source policy!" -Color Red
                    Write-StatusLog "[ERROR] Cannot replace - admin must investigate manually" -Color Red
                    Write-StatusLog "[ERROR] User remains in source group (no changes made)" -Color Red
                    throw "No Cloud PC found matching source provisioning policy. User needs manual intervention."
                }
                
                # Safe to proceed - remove from source group
                Write-StatusLog "[STAGE] $($State.UserPrincipalName): Removing from source group ($($State.SourceGroupName))..." -Color Cyan
                Remove-UserFromGroup -UserId $State.UserId -GroupId $State.SourceGroupId -UserPrincipalName $State.UserPrincipalName | Out-Null
                Write-StatusLog "[STAGE] $($State.UserPrincipalName): Removed from source group" -Color Green
                Write-InfoLog "[INFO] $($State.UserPrincipalName): Tracking $($State.CloudPCIds.Count) CPC(s) - waiting for grace period..." "Cyan"
                
                $State.Stage = "Waiting for Grace Period"
                $State.ProgressPercent = 30
                $State.StageStartTime = Get-Date
            }
            
            "Waiting for Grace Period" {
                # Only proceed if we're tracking specific CPCs
                if (-not $State.CloudPCIds -or $State.CloudPCIds.Count -eq 0) {
                    Write-StatusLog "[ERROR] No CPC IDs to track - should not be in grace period stage!" -Color Red
                    throw "Logic error: No CPC IDs set"
                }
                
                $cloudPCs = Get-CloudPCForUser -UserId $State.UserPrincipalName
                
                # Check status of ALL tracked CPCs
                $trackedCPCs = @($cloudPCs | Where-Object { $State.CloudPCIds -contains $_.id })
                
                if ($trackedCPCs.Count -eq 0) {
                    Write-StatusLog "[STATE CHANGE] $($State.UserPrincipalName): All tracked CPCs disappeared (already deprovisioned)" -Color Green
                    Write-Log "$($State.UserPrincipalName): State Change - All tracked CPCs already gone" -Level Info
                    $State.Stage = "Adding to Target"
                    $State.ProgressPercent = 60
                    $State.StageStartTime = Get-Date
                }
                else {
                    # Check if any CPCs are already deprovisioning or notProvisioned - skip to deprovision wait
                    $alreadyDeprovisioning = $trackedCPCs | Where-Object { $_.status -in @('deprovisioning', 'notProvisioned') }
                    
                    if ($alreadyDeprovisioning) {
                        Write-StatusLog "[STATE CHANGE] $($State.UserPrincipalName): CPCs already deprovisioning - skipping grace period stages!" -Color Yellow
                        foreach ($cpc in $alreadyDeprovisioning) {
                            $cpcName = if ($cpc.managedDeviceName) { $cpc.managedDeviceName } else { $cpc.displayName }
                            Write-StatusLog "[STATE CHANGE]   - $cpcName (status: $($cpc.status))" -Color Yellow
                        }
                        $State.Stage = "Waiting for Deprovision"
                        $State.ProgressPercent = 55
                        $State.StageStartTime = Get-Date
                    }
                    else {
                        # Check if ALL tracked CPCs are in grace period
                        $allInGrace = $true
                        $statusSummary = @{}
                        
                        foreach ($cpc in $trackedCPCs) {
                            $cpcName = if ($cpc.managedDeviceName) { $cpc.managedDeviceName } else { $cpc.displayName }
                            
                            if ($cpc.status -ne 'inGracePeriod') {
                                $allInGrace = $false
                            }
                            
                            if (-not $statusSummary.ContainsKey($cpc.status)) {
                                $statusSummary[$cpc.status] = @()
                            }
                            $statusSummary[$cpc.status] += $cpcName
                        }
                        
                        if ($allInGrace) {
                            Write-StatusLog "[STATE CHANGE] $($State.UserPrincipalName): ALL $($trackedCPCs.Count) tracked CPCs entered grace period!" -Color Green
                            foreach ($cpc in $trackedCPCs) {
                                $cpcName = if ($cpc.managedDeviceName) { $cpc.managedDeviceName } else { $cpc.displayName }
                                Write-StatusLog "[STATE CHANGE]   - $cpcName (status: $($cpc.status))" -Color Green
                            }
                            Write-Log "$($State.UserPrincipalName): State Change - All Cloud PCs entered grace period" -Level Info
                            
                            $State.Stage = "Ending Grace Period"
                            $State.ProgressPercent = 40
                            $State.StageStartTime = Get-Date
                        }
                        else {
                            Write-PollingLog "[POLLING] $($State.UserPrincipalName): Waiting for all CPCs to enter grace..." "Gray"
                            foreach ($status in $statusSummary.Keys) {
                                $names = $statusSummary[$status] -join ', '
                                Write-DebugLog "[DEBUG]   Status '$status': $names" "Gray"
                            }
                            
                            if ($elapsed.TotalMinutes -gt $script:gracePeriodTimeoutMinutes) {
                                throw "Timeout waiting for grace period"
                            }
                        }
                    }
                }
            }
            
            "Ending Grace Period" {
                # First time in this stage - send the End Grace API calls
                if (-not $State.GraceEndedCPCIds -or $State.GraceEndedCPCIds.Count -eq 0) {
                    Write-StatusLog "[STAGE] $($State.UserPrincipalName): Ending grace period on $($State.CloudPCIds.Count) tracked CPC(s)..." -Color Cyan
                    
                    # Track which CPCs we've ended grace on
                    $State.GraceEndedCPCIds = @()
                    
                    # End grace period on ALL tracked CPCs
                    foreach ($cpcId in $State.CloudPCIds) {
                        Write-InfoLog "[INFO] Ending grace for CPC ID: $cpcId" "Cyan"
                        Stop-CloudPCGracePeriod -CloudPCId $cpcId -UserPrincipalName $State.UserPrincipalName | Out-Null
                        $State.GraceEndedCPCIds += $cpcId
                    }
                    
                    Write-StatusLog "[STAGE] $($State.UserPrincipalName): End grace API called - waiting for deprovisioning to start..." -Color Green
                }
                
                # Now wait until we see deprovisioning status
                $cloudPCs = Get-CloudPCForUser -UserId $State.UserPrincipalName
                $trackedCPCs = @($cloudPCs | Where-Object { $State.CloudPCIds -contains $_.id })
                
                # Check if any tracked CPCs have disappeared (already done)
                if ($trackedCPCs.Count -eq 0) {
                    Write-StatusLog "[STATE CHANGE] $($State.UserPrincipalName): All tracked CPCs disappeared (already deprovisioned)" -Color Green
                    $State.Stage = "Adding to Target"
                    $State.ProgressPercent = 60
                    $State.StageStartTime = Get-Date
                }
                else {
                    # Check if ANY CPC is deprovisioning or notProvisioned
                    $anyDeprovisioning = $trackedCPCs | Where-Object { $_.status -in @('deprovisioning', 'notProvisioned') }
                    
                    if ($anyDeprovisioning) {
                        # At least one CPC confirmed deprovisioning - move to next stage
                        Write-StatusLog "[STATE CHANGE] $($State.UserPrincipalName): Deprovisioning confirmed - backend processing started!" -Color Green
                        foreach ($cpc in $anyDeprovisioning) {
                            $cpcName = if ($cpc.managedDeviceName) { $cpc.managedDeviceName } else { $cpc.displayName }
                            Write-StatusLog "[STATE CHANGE]   - $cpcName (status: $($cpc.status))" -Color Green
                        }
                        
                        $State.Stage = "Waiting for Deprovision"
                        $State.ProgressPercent = 55
                        $State.StageStartTime = Get-Date
                        $State.LastPollTime = Get-Date
                    }
                    else {
                        # Still showing inGracePeriod - wait for API to catch up
                        $elapsed = (Get-Date) - $State.StageStartTime
                        Write-PollingLog "[POLLING] $($State.UserPrincipalName): Waiting for deprovisioning to start (elapsed: $([math]::Round($elapsed.TotalMinutes, 1))min)..." "Gray"
                        
                        foreach ($cpc in $trackedCPCs) {
                            $cpcName = if ($cpc.managedDeviceName) { $cpc.managedDeviceName } else { $cpc.displayName }
                            Write-DebugLog "[DEBUG]   - $cpcName (status: $($cpc.status))" "Gray"
                        }
                        
                        # Timeout check
                        if ($elapsed.TotalMinutes -gt $script:endingGracePeriodTimeoutMinutes) {
                            throw "Timeout waiting for deprovisioning to start after ending grace period"
                        }
                    }
                }
            }
            
            "Waiting for Deprovision" {
                $cloudPCs = Get-CloudPCForUser -UserId $State.UserPrincipalName
                
                # Check if ALL tracked CPCs are gone OR notProvisioned (license freed)
                $trackedCPCs = @($cloudPCs | Where-Object { $State.CloudPCIds -contains $_.id })
                
                # Context-aware filtering: Consider what actions we've already taken
                # - notProvisioned = license freed, ready to provision new CPC
                # - inGracePeriod (after we've ended grace) = API lag or flip-flopping
                # - deprovisioning = actively deprovisioning
                # - Status can flip-flop between deprovisioning and inGracePeriod due to API inconsistency
                
                $activeCPCs = @()
                foreach ($cpc in $trackedCPCs) {
                    # Already freed = not blocking
                    if ($cpc.status -eq 'notProvisioned') {
                        continue
                    }
                    
                    # Track if this CPC has ever reached 'deprovisioning' status
                    if ($cpc.status -eq 'deprovisioning' -and -not $State.DeprovisioningSeenCPCs.ContainsKey($cpc.id)) {
                        $State.DeprovisioningSeenCPCs[$cpc.id] = $true
                        Write-StatusLog "[STATE] CPC $($cpc.id) reached 'deprovisioning' status - backend is processing" -Color Green
                    }
                    
                    # We've ended grace on this CPC and it's showing inGracePeriod
                    if ($cpc.status -eq 'inGracePeriod' -and $State.GraceEndedCPCIds -contains $cpc.id) {
                        $timeSinceGraceEnded = ((Get-Date) - $State.StageStartTime).TotalMinutes
                        
                        # Check statusDetails for more info
                        $statusInfo = ""
                        if ($cpc.statusDetails) {
                            $statusInfo = " | statusDetails: $($cpc.statusDetails.code) - $($cpc.statusDetails.message)"
                        }
                        if ($cpc.lastModifiedDateTime) {
                            # Graph API returns UTC timestamps (e.g., "2024-02-13T14:05:34Z")
                            # Use SpecifyKind to tell PowerShell this is UTC, don't convert it
                            $lastModified = [DateTime]::SpecifyKind([DateTime]::Parse($cpc.lastModifiedDateTime), [DateTimeKind]::Utc)
                            $minutesSinceModified = ((Get-Date).ToUniversalTime() - $lastModified).TotalMinutes
                            $statusInfo += " | Last modified: $([math]::Round($minutesSinceModified, 1))min ago"
                        }
                        
                        # If we've seen this CPC reach 'deprovisioning' before, treat inGracePeriod as flip-flop
                        if ($State.DeprovisioningSeenCPCs.ContainsKey($cpc.id)) {
                            Write-DebugLog "[DEBUG] CPC $($cpc.id) flip-flopped back to inGracePeriod (seen deprovisioning before) - API inconsistency$statusInfo" "Yellow"
                            $script:GridStatus.Rows[$State.GridRowIndex].Cells["Messages"].Value = "Status flip-flop detected (API lag)"
                            # Don't block on this - we know it's actually deprovisioning
                            $activeCPCs += $cpc  # Still track it, but we won't fail as quickly
                        }
                        else {
                            # We ended grace on this CPC but backend hasn't switched to deprovisioning yet - this is normal
                            Write-DebugLog "[DEBUG] CPC $($cpc.id) still shows inGracePeriod $([math]::Round($timeSinceGraceEnded, 1))min after ending grace - waiting for backend$statusInfo" "Gray"
                            $activeCPCs += $cpc  # Keep tracking - waiting for backend to update status
                        }
                    }
                    # Currently deprovisioning = expected, track and continue
                    elseif ($cpc.status -eq 'deprovisioning') {
                        $statusInfo = ""
                        if ($cpc.lastModifiedDateTime) {
                            # Graph API returns UTC timestamps (e.g., "2024-02-13T14:05:34Z")
                            # Use SpecifyKind to tell PowerShell this is UTC, don't convert it
                            $lastModified = [DateTime]::SpecifyKind([DateTime]::Parse($cpc.lastModifiedDateTime), [DateTimeKind]::Utc)
                            $minutesSinceModified = ((Get-Date).ToUniversalTime() - $lastModified).TotalMinutes
                            $statusInfo = " (last updated $([math]::Round($minutesSinceModified, 1))min ago)"
                        }
                        Write-DebugLog "[DEBUG] CPC $($cpc.id) is actively deprovisioning$statusInfo - waiting..." "Gray"
                        $activeCPCs += $cpc
                    }
                    # Any other status = treat as active
                    else {
                        $activeCPCs += $cpc
                    }
                }
                
                if ($activeCPCs.Count -eq 0) {
                    Write-StatusLog "[STATE CHANGE] $($State.UserPrincipalName): ALL tracked CPCs deprovisioned (licenses freed)!" -Color Green
                    
                    # Show notProvisioned CPCs if any remain as "ghosts"
                    $notProvisionedCPCs = @($trackedCPCs | Where-Object { $_.status -eq 'notProvisioned' })
                    if ($notProvisionedCPCs.Count -gt 0) {
                        Write-InfoLog "[INFO] $($State.UserPrincipalName): $($notProvisionedCPCs.Count) CPC(s) in 'notProvisioned' state (licenses freed)" "Cyan"
                    }
                    
                    Write-Log "$($State.UserPrincipalName): State Change - All tracked CPCs deprovisioned" -Level Info
                    
                    if ($cloudPCs -and $cloudPCs.Count -gt 0) {
                        $otherCPCs = @($cloudPCs | Where-Object { $State.CloudPCIds -notcontains $_.id })
                        if ($otherCPCs.Count -gt 0) {
                            Write-InfoLog "[INFO] $($State.UserPrincipalName): User has $($otherCPCs.Count) other Cloud PC(s) (e.g., Frontline)" "Cyan"
                            foreach ($cpc in $otherCPCs) {
                                $cpcName = if ($cpc.managedDeviceName) { $cpc.managedDeviceName } else { $cpc.displayName }
                                Write-DebugLog "[DEBUG]   $($State.UserPrincipalName): $cpcName | Status: $($cpc.status) | Policy: $($cpc.provisioningPolicyId)" "Cyan"
                            }
                        }
                    }
                    
                    $State.Stage = "Adding to Target"
                    $State.ProgressPercent = 60
                    $State.StageStartTime = Get-Date
                }
                else {
                    Write-PollingLog "[POLLING] $($State.UserPrincipalName): Still has $($activeCPCs.Count) active tracked CPC(s) deprovisioning..." "Gray"
                    foreach ($cpc in $activeCPCs) {
                        $cpcName = if ($cpc.managedDeviceName) { $cpc.managedDeviceName } else { $cpc.displayName }
                        Write-DebugLog "[DEBUG]   $($State.UserPrincipalName): $cpcName | Status: $($cpc.status)" "Gray"
                    }
                    
                    if ($elapsed.TotalMinutes -gt $script:deprovisionTimeoutMinutes) {
                        throw "Timeout waiting for deprovision - still has $($activeCPCs.Count) active CPC(s)"
                    }
                }
            }
            
            "Adding to Target" {
                Write-StatusLog "[STAGE] $($State.UserPrincipalName): Adding to target group ($($State.TargetGroupName))..." -Color Cyan
                Add-UserToGroup -UserId $State.UserId -GroupId $State.TargetGroupId -UserPrincipalName $State.UserPrincipalName | Out-Null
                Write-StatusLog "[STAGE] $($State.UserPrincipalName): Added to target group, new Cloud PC will provision" -Color Green
                Write-InfoLog "[INFO] $($State.UserPrincipalName): Admin work complete - monitoring provisioning (does not block queue)" "Cyan"
                $State.Stage = "Waiting for Provisioning"
                $State.ProgressPercent = 70
                $State.StageStartTime = Get-Date
                
                # Clear flip-flop messages but preserve multiple CPC warnings
                $currentMessage = $script:GridStatus.Rows[$State.GridRowIndex].Cells["Messages"].Value
                if ($currentMessage -notlike "Multiple CPCs detected*") {
                    $script:GridStatus.Rows[$State.GridRowIndex].Cells["Messages"].Value = "Waiting for new CPC to provision..."
                }
                
                # Update grid to show "Monitoring" status (not blocking other jobs)
                $script:GridStatus.Rows[$State.GridRowIndex].Cells["Status"].Value = "Monitoring"
                $script:GridStatus.Rows[$State.GridRowIndex].DefaultCellStyle.BackColor = [System.Drawing.Color]::LightCyan
            }
            
            "Waiting for Provisioning" {
                # Known Cloud PC status values from Microsoft Graph API:
                # - provisioned: Successfully provisioned and ready
                # - provisionedWithWarnings: Provisioned but with non-blocking issues
                # - provisioning: Currently being provisioned
                # - failed: Provisioning failed
                # - notProvisioned: Not yet provisioned
                # - inGracePeriod: In grace period (being deprovisioned)
                # - deprovisioning: Being deprovisioned
                # - resizing, restoring, movingRegion: Other temporary states
                
                $cloudPCs = @(Get-CloudPCForUser -UserId $State.UserPrincipalName)
                
                # Ensure it's an array for proper counting
                if ($cloudPCs -isnot [array]) {
                    $cloudPCs = @($cloudPCs)
                }
                
                # Get target policy IDs if not already cached
                if (-not $State.TargetPolicyIds) {
                    Write-DebugLog "[DEBUG] Getting provisioning policies for target group: $($State.TargetGroupName)" "Gray"
                    $State.TargetPolicyIds = @(Get-ProvisioningPoliciesForGroup -GroupId $State.TargetGroupId)
                    if ($State.TargetPolicyIds.Count -gt 0) {
                        Write-DebugLog "[DEBUG] Target group uses $($State.TargetPolicyIds.Count) provisioning policy(ies)" "Gray"
                    }
                }
                
                if ($cloudPCs.Count -gt 1) {
                    Write-StatusLog "[Info] $($State.UserPrincipalName): Multiple Cloud PCs detected ($($cloudPCs.Count))" -Color Cyan
                    $script:GridStatus.Rows[$State.GridRowIndex].Cells["Messages"].Value = "Multiple CPCs detected ($($cloudPCs.Count))"
                    foreach ($cpc in $cloudPCs) {
                        $cpcName = if ($cpc.managedDeviceName) { $cpc.managedDeviceName } else { $cpc.displayName }
                        Write-DebugLog "[DEBUG]   $($State.UserPrincipalName): $cpcName | Status: $($cpc.status) | Policy: $($cpc.provisioningPolicyId) | Plan: $($cpc.servicePlanName)" "Yellow"
                    }
                }
                
                # Find NEW CPCs from TARGET policy
                $newProvisionedCPCs = @($cloudPCs | Where-Object { 
                    $_.status -in @('provisioned', 'provisionedWithWarnings') -and 
                    $State.TargetPolicyIds -contains $_.provisioningPolicyId 
                })
                
                Write-DebugLog "[DEBUG] Found $($newProvisionedCPCs.Count) provisioned CPC(s) from TARGET policy (expected: $($State.OldCPCs.Count))" "Magenta"
                
                # Match by service plan - ensure we have the same number of each plan type
                if ($newProvisionedCPCs.Count -gt 0 -and $State.OldCPCs.Count -gt 0) {
                    # Group old and new CPCs by service plan
                    $oldByPlan = $State.OldCPCs | Group-Object ServicePlan
                    $newByPlan = $newProvisionedCPCs | Group-Object servicePlanName
                    
                    Write-DebugLog "[DEBUG] $($State.UserPrincipalName): Old CPCs by plan: $($oldByPlan.Count) plan type(s)" "Cyan"
                    foreach ($group in $oldByPlan) {
                        Write-DebugLog "[DEBUG]   $($State.UserPrincipalName): $($group.Name): $($group.Count) CPC(s)" "Cyan"
                    }
                    
                    Write-DebugLog "[DEBUG] $($State.UserPrincipalName): New CPCs by plan: $($newByPlan.Count) plan type(s)" "Cyan"
                    foreach ($group in $newByPlan) {
                        Write-DebugLog "[DEBUG]   $($State.UserPrincipalName): $($group.Name): $($group.Count) CPC(s)" "Cyan"
                    }
                    
                    # Check if counts match for each plan type
                    $allMatched = $true
                    foreach ($oldGroup in $oldByPlan) {
                        $newGroup = $newByPlan | Where-Object { $_.Name -eq $oldGroup.Name }
                        if (-not $newGroup -or $newGroup.Count -lt $oldGroup.Count) {
                            Write-DebugLog "[DEBUG] Plan '$($oldGroup.Name)': Need $($oldGroup.Count), have $($newGroup.Count)" "Yellow"
                            $allMatched = $false
                            break
                        }
                    }
                    
                    if ($allMatched) {
                        # SUCCESS! All service plans have matching counts
                        Write-StatusLog "[STATE CHANGE] $($State.UserPrincipalName): All $($State.OldCPCs.Count) Cloud PC(s) provisioned!" -Color Green
                        Write-Log "$($State.UserPrincipalName): State Change - All new Cloud PCs provisioned" -Level Success
                        
                        # Store new CPC details for export
                        $State.NewCPCs = @($newProvisionedCPCs | ForEach-Object {
                            @{
                                Id = $_.id
                                Name = if ($_.managedDeviceName) { $_.managedDeviceName } else { $_.displayName }
                                ServicePlan = $_.servicePlanName
                                Status = $_.status
                            }
                        })
                        
                        # Build summary messages
                        $hasWarnings = $newProvisionedCPCs | Where-Object { $_.status -eq 'provisionedWithWarnings' }
                        $State.Stage = "Complete"
                        $State.Status = if ($hasWarnings) { "Success (Warnings)" } else { "Success" }
                        $State.ProgressPercent = 100
                        $State.EndTime = Get-Date
                        
                        # Calculate duration
                        $duration = $State.EndTime - $State.StartTime
                        $durationText = if ($duration.TotalMinutes -lt 60) {
                            "{0:N0} minutes" -f $duration.TotalMinutes
                        } else {
                            "{0:N1} hours" -f $duration.TotalHours
                        }
                        
                        $State.FinalMessage = if ($hasWarnings) {
                            "Completed $($State.NewCPCs.Count) CPC(s) in $durationText (with warnings)"
                        } else {
                            "Completed $($State.NewCPCs.Count) CPC(s) in $durationText"
                        }
                        $State.NextPollDisplay = "-"
                        
                        $script:GridStatus.Rows[$State.GridRowIndex].Cells["Status"].Value = $State.Status
                        $script:GridStatus.Rows[$State.GridRowIndex].Cells["Stage"].Value = Get-StageDisplay "Complete"
                        $script:GridStatus.Rows[$State.GridRowIndex].Cells["NextPoll"].Value = "-"
                        $script:GridStatus.Rows[$State.GridRowIndex].Cells["Messages"].Value = $State.FinalMessage
                        $script:GridStatus.Rows[$State.GridRowIndex].DefaultCellStyle.BackColor = if ($hasWarnings) { 
                            [System.Drawing.Color]::LightYellow 
                        } else { 
                            [System.Drawing.Color]::LightGreen 
                        }
                        
                        # Auto-save job summary
                        Save-JobSummaryAuto -State $State
                        
                        return
                    }
                }
                
                # Not yet complete - still waiting for CPCs to provision
                Write-DebugLog "[DEBUG] Waiting for provisioning... (have $($newProvisionedCPCs.Count) of $($State.OldCPCs.Count) expected)" "Yellow"
                
                # Check for failed provisioning
                $failedCPC = $cloudPCs | Where-Object { $_.status -eq 'failed' } | Select-Object -First 1
                if ($failedCPC) {
                    Write-StatusLog "[STATE CHANGE] $($State.UserPrincipalName): Provisioning FAILED!" -Color Red
                    Write-DebugLog "[DEBUG] Provisioning Policy: $($failedCPC.provisioningPolicyId)" "Magenta"
                    Write-Log "$($State.UserPrincipalName): State Change - Provisioning failed" -Level Error
                    throw "Provisioning failed"
                }
                
                # Still provisioning
                if ($cloudPCs.Count -gt 0) {
                    $provisioningCPC = $cloudPCs | Where-Object { $_.status -eq 'provisioning' } | Select-Object -First 1
                    if ($provisioningCPC) {
                        $cpcName = if ($provisioningCPC.managedDeviceName) { $provisioningCPC.managedDeviceName } else { $provisioningCPC.displayName }
                        Write-PollingLog "[POLLING] $($State.UserPrincipalName): Still provisioning (status: $($provisioningCPC.status), Name: $cpcName)" "Gray"
                        Write-DebugLog "[DEBUG] Provisioning Policy: $($provisioningCPC.provisioningPolicyId)" "Magenta"
                    }
                    else {
                        # Log any unexpected statuses for visibility
                        $otherStatuses = $cloudPCs.status | Select-Object -Unique
                        Write-PollingLog "[POLLING] $($State.UserPrincipalName): Has $($cloudPCs.Count) CPC(s) with status: $($otherStatuses -join ', ')" "Gray"
                    }
                    
                    if ($elapsed.TotalMinutes -gt $script:provisioningTimeoutMinutes) {
                        # Don't fail - just stop monitoring with warning
                        $State.Status = "Warning"
                        $State.Stage = "Monitoring Stopped"
                        $State.EndTime = Get-Date
                        $State.FinalMessage = "Provisioning exceeded $script:provisioningTimeoutMinutes min - monitoring stopped but provisioning may still complete"
`n                        $State.NextPollDisplay = "-"
                        
                        Write-StatusLog "[WARNING] $($State.UserPrincipalName): Provisioning exceeded timeout ($script:provisioningTimeoutMinutes min) - stopping monitoring" -Color Yellow
                        Write-InfoLog "[INFO] Provisioning may still complete in the background" "Cyan"
                        
                        $script:GridStatus.Rows[$State.GridRowIndex].Cells["Status"].Value = "Warning"
                        $script:GridStatus.Rows[$State.GridRowIndex].Cells["Stage"].Value = "Stopped"
                        $script:GridStatus.Rows[$State.GridRowIndex].Cells["NextPoll"].Value = "-"
                        $script:GridStatus.Rows[$State.GridRowIndex].Cells["Messages"].Value = $State.FinalMessage
                        $script:GridStatus.Rows[$State.GridRowIndex].DefaultCellStyle.BackColor = [System.Drawing.Color]::Yellow
                        
                        Save-JobSummaryAuto -State $State
                        return
                    }
                }
                else {
                    Write-PollingLog "[POLLING] $($State.UserPrincipalName): No Cloud PC found yet, still waiting..." "Gray"
                    
                    if ($elapsed.TotalMinutes -gt $script:provisioningTimeoutMinutes) {
                        # Don't fail - just stop monitoring with warning
                        $State.Status = "Warning"
                        $State.Stage = "Monitoring Stopped"
                        $State.EndTime = Get-Date
                        $State.FinalMessage = "Provisioning exceeded $script:provisioningTimeoutMinutes min - monitoring stopped but provisioning may still complete"
`n                        $State.NextPollDisplay = "-"
                        
                        Write-StatusLog "[WARNING] $($State.UserPrincipalName): Provisioning exceeded timeout ($script:provisioningTimeoutMinutes min) - stopping monitoring" -Color Yellow
                        Write-InfoLog "[INFO] Provisioning may still complete in the background" "Cyan"
                        
                        $script:GridStatus.Rows[$State.GridRowIndex].Cells["Status"].Value = "Warning"
                        $script:GridStatus.Rows[$State.GridRowIndex].Cells["Stage"].Value = "Stopped"
                        $script:GridStatus.Rows[$State.GridRowIndex].Cells["NextPoll"].Value = "-"
                        $script:GridStatus.Rows[$State.GridRowIndex].Cells["Messages"].Value = $State.FinalMessage
                        $script:GridStatus.Rows[$State.GridRowIndex].DefaultCellStyle.BackColor = [System.Drawing.Color]::Yellow
                        
                        Save-JobSummaryAuto -State $State
                        return
                    }
                }
            }
        }
        
        # Update grid for ongoing stages
        $script:GridStatus.Rows[$State.GridRowIndex].Cells["Stage"].Value = Get-StageDisplay $State.Stage
    }
    catch {
        $State.Status = "Failed"
        $State.ErrorMessage = $_.Exception.Message
        $State.EndTime = Get-Date
        $State.FinalMessage = $State.ErrorMessage
`n        $State.NextPollDisplay = "-"
        
        Write-Host "`n[ERROR] $($State.UserPrincipalName) FAILED" -ForegroundColor Red
        Write-Host "  Stage: $($State.Stage)" -ForegroundColor Yellow
        Write-Host "  Error: $($State.ErrorMessage)" -ForegroundColor Red
        Write-Host "  Stack: $($_.ScriptStackTrace)" -ForegroundColor Gray
        
        $script:GridStatus.Rows[$State.GridRowIndex].Cells["Status"].Value = "Failed"
        $script:GridStatus.Rows[$State.GridRowIndex].Cells["Messages"].Value = $State.ErrorMessage
        $script:GridStatus.Rows[$State.GridRowIndex].Cells["NextPoll"].Value = "-"
        $script:GridStatus.Rows[$State.GridRowIndex].DefaultCellStyle.BackColor = [System.Drawing.Color]::LightCoral
        
        # Auto-save job summary for failed job
        Save-JobSummaryAuto -State $State
    }
}

# Show disclaimer and get user acceptance
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$disclaimerForm = New-Object System.Windows.Forms.Form
$disclaimerForm.Text = "Cloud PC Replace Tool v$script:ToolVersion - Important Notice"
$disclaimerForm.Size = New-Object System.Drawing.Size(600, 400)
$disclaimerForm.StartPosition = "CenterScreen"
$disclaimerForm.FormBorderStyle = "FixedDialog"
$disclaimerForm.MaximizeBox = $false
$disclaimerForm.MinimizeBox = $false

$disclaimerText = New-Object System.Windows.Forms.TextBox
$disclaimerText.Multiline = $true
$disclaimerText.ReadOnly = $true
$disclaimerText.ScrollBars = "Vertical"
$disclaimerText.Location = New-Object System.Drawing.Point(20, 20)
$disclaimerText.Size = New-Object System.Drawing.Size(540, 260)
$disclaimerText.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$disclaimerText.Text = @"
IMPORTANT NOTICE - PLEASE READ CAREFULLY

This Cloud PC Replace Tool is provided AS-IS without warranty or support.

By using this tool, you acknowledge and agree that:

- This is an UNOFFICIAL, COMMUNITY-DEVELOPED tool
- It is NOT supported, endorsed, or warrantied by Microsoft
- You use this tool entirely AT YOUR OWN RISK
- You are responsible for understanding what this tool does
- You have had the opportunity to review the source code
- You should TEST in a non-production environment first
- You accept full responsibility for any consequences
- Changes made by this tool affect production Cloud PCs
- There is no built-in rollback mechanism
- You have proper backups and disaster recovery plans

This tool automates Cloud PC Replace operations using Microsoft Graph API. It will:
- Remove users from groups
- End grace periods on Cloud PCs
- Deprovision existing Cloud PCs
- Add users to new groups
- Trigger new Cloud PC provisioning

These are PERMANENT operations that affect user productivity.

Recommended: Review the code in Start-CloudPCReplaceGUI.ps1 and CloudPCReplace.psm1 before proceeding.
"@
$disclaimerText.Select(0, 0)  # Clear selection

$disclaimerForm.Controls.Add($disclaimerText)

$checkboxAccept = New-Object System.Windows.Forms.CheckBox
$checkboxAccept.Location = New-Object System.Drawing.Point(20, 290)
$checkboxAccept.Size = New-Object System.Drawing.Size(540, 25)
$checkboxAccept.Text = "I have read and understand the above. I accept all risks and responsibilities."
$checkboxAccept.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$disclaimerForm.Controls.Add($checkboxAccept)

$btnAccept = New-Object System.Windows.Forms.Button
$btnAccept.Location = New-Object System.Drawing.Point(370, 320)
$btnAccept.Size = New-Object System.Drawing.Size(90, 30)
$btnAccept.Text = "Accept"
$btnAccept.Enabled = $false
$btnAccept.DialogResult = [System.Windows.Forms.DialogResult]::OK
$disclaimerForm.Controls.Add($btnAccept)

$btnDecline = New-Object System.Windows.Forms.Button
$btnDecline.Location = New-Object System.Drawing.Point(470, 320)
$btnDecline.Size = New-Object System.Drawing.Size(90, 30)
$btnDecline.Text = "Decline"
$btnDecline.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$disclaimerForm.Controls.Add($btnDecline)

$checkboxAccept.Add_CheckedChanged({
    $btnAccept.Enabled = $checkboxAccept.Checked
})

$disclaimerForm.AcceptButton = $btnAccept
$disclaimerForm.CancelButton = $btnDecline

$result = $disclaimerForm.ShowDialog()

if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "User declined terms. Exiting." -ForegroundColor Yellow
    exit
}

Write-Host "User accepted terms. Starting application..." -ForegroundColor Green

# Check for Microsoft Graph module
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Microsoft Graph SDK required. Install?", "Missing Dependency", "YesNo", "Question"
    )
    if ($result -eq "Yes") {
        Write-Host "Installing Microsoft Graph PowerShell SDK..." -ForegroundColor Yellow
        Install-Module Microsoft.Graph -Scope CurrentUser -Force
        Write-Host "Installation complete!" -ForegroundColor Green
    }
    else {
        exit
    }
}

# Import required modules
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

Show-ReplaceGUI





