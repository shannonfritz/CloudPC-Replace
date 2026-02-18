#Requires -Version 5.1

<#
.SYNOPSIS
    Cloud PC Replace Module - Automates replacing between provisioning profiles

.DESCRIPTION
    This module provides functions to automate replacing users' Cloud PCs from one
    provisioning profile to another while reusing Windows 365 Enterprise licenses.
    
    The process works by:
    1. Removing user from source group (deprovisions current Cloud PC)
    2. Waiting for grace period and forcing deprovision (frees license)
    3. Adding user to target group (provisions new Cloud PC with different profile)
    4. Monitoring new Cloud PC provisioning
    
    Use cases: Different join type, network, region, image, size, security policies, etc.

.NOTES
    Author: Cloud PC Replace Tool
    Requires: Microsoft.Graph PowerShell SDK
#>

#region Logging Helpers
# These will be controlled by the GUI's verbose logging setting
$script:verboseLogging = $false

function Write-DebugLog {
    param([string]$Message, [string]$Color = "DarkGray")
    if ($script:verboseLogging) {
        $timestamp = Get-Date -Format "HH:mm:ss"
        Write-Host "[$timestamp] $Message" -ForegroundColor $Color
    }
}

function Write-InfoLog {
    param([string]$Message, [string]$Color = "Cyan")
    if ($script:verboseLogging) {
        $timestamp = Get-Date -Format "HH:mm:ss"
        Write-Host "[$timestamp] $Message" -ForegroundColor $Color
    }
}

function Write-ApiLog {
    param([string]$Message, [string]$Color = "Yellow")
    $timestamp = Get-Date -Format "HH:mm:ss"
    Write-Host "[$timestamp] $Message" -ForegroundColor $Color
}
#endregion

#region Configuration
class ReplaceStatus {
    [string]$UserId
    [string]$UserPrincipalName
    [string]$Stage
    [string]$Status
    [datetime]$StartTime
    [datetime]$LastUpdate
    [string]$ErrorMessage
    [string]$CloudPCId  # Legacy - keeping for backwards compat
    [array]$CloudPCIds  # Array of CPC IDs to track
    [int]$ProgressPercent
}
#endregion

#region Graph Authentication
function Connect-MgGraphForReplace {
    <#
    .SYNOPSIS
        Connects to Microsoft Graph with required permissions
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$TenantId,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Interactive', 'DeviceCode', 'ClientSecret')]
        [string]$AuthMethod = 'Interactive'
    )
    
    $requiredScopes = @(
        'CloudPC.ReadWrite.All',
        'Group.ReadWrite.All',
        'User.Read.All'
    )
    
    try {
        Write-Log "Connecting to Microsoft Graph..." -Level Info
        
        $connectParams = @{
            Scopes = $requiredScopes
            NoWelcome = $true
        }
        
        if ($TenantId) {
            $connectParams['TenantId'] = $TenantId
        }
        
        if ($AuthMethod -eq 'DeviceCode') {
            $connectParams['UseDeviceAuthentication'] = $true
        }
        
        Connect-MgGraph @connectParams -ErrorAction Stop
        
        $context = Get-MgContext
        if (-not $context) {
            throw "Failed to get Graph context after connection"
        }
        
        $tenantInfo = if ($context.TenantId) { $context.TenantId } else { "Unknown" }
        $accountInfo = if ($context.Account) { $context.Account } else { "Unknown" }
        
        Write-Log "Connected to tenant: $tenantInfo (Account: $accountInfo)" -Level Success
        
        return $true
    }
    catch {
        Write-Log "Failed to connect to Microsoft Graph: $_" -Level Error
        return $false
    }
}

function Test-GraphConnection {
    <#
    .SYNOPSIS
        Tests if connected to Microsoft Graph
    #>
    try {
        $context = Get-MgContext
        return ($null -ne $context -and $null -ne $context.TenantId)
    }
    catch {
        return $false
    }
}
#endregion

#region Cloud PC Operations
function Get-CloudPCForUser {
    <#
    .SYNOPSIS
        Gets ALL Cloud PC information for a specific user
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId
    )
    
    try {
        $uri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/cloudPCs?`$filter=userPrincipalName eq '$UserId'"
        $response = Invoke-MgGraphRequest -Uri $uri -Method GET
        
        if ($response.value) {
            # Ensure we return an array
            if ($response.value -is [array]) {
                $cpcs = @($response.value)
            }
            else {
                # Single object - wrap in array
                $cpcs = @($response.value)
            }
            Write-DebugLog "[DEBUG] ${UserId}: Get-CloudPCForUser: response.value is array=$($response.value -is [array]), count=$($cpcs.Count)" "DarkGray"
            
            # Log each CPC found
            foreach ($cpc in $cpcs) {
                Write-DebugLog "[DEBUG]   ${UserId}: Found CPC: $($cpc.managedDeviceName), ID: $($cpc.id)" "DarkGray"
                Write-DebugLog "[DEBUG]     ${UserId}: ServicePlanName: $($cpc.servicePlanName), ServicePlanType: $($cpc.servicePlanType)" "DarkGray"
            }
            
            return $cpcs
        }
        return @()
    }
    catch {
        Write-Log "Error getting Cloud PC for user $UserId : $_" -Level Error
        throw
    }
}

function Get-UserGroupMemberships {
    <#
    .SYNOPSIS
        Gets all group memberships for a user with pagination support
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId
    )
    
    try {
        $uri = "https://graph.microsoft.com/v1.0/users/$UserId/memberOf?`$top=100"
        $allGroups = @()
        
        do {
            $response = Invoke-MgGraphRequest -Uri $uri -Method GET
            
            if ($response.value -and $response.value.Count -gt 0) {
                $groups = $response.value | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.group' }
                $allGroups += @($groups)
            }
            
            $uri = $response.'@odata.nextLink'
        } while ($uri)
        
        Write-DebugLog "[DEBUG] Get-UserGroupMemberships: Found $($allGroups.Count) group(s)" "DarkGray"
        
        return $allGroups
    }
    catch {
        Write-Log "Error getting group memberships for user $UserId : $_" -Level Error
        return @()
    }
}

function Get-ProvisioningPoliciesForGroup {
    <#
    .SYNOPSIS
        Gets Cloud PC provisioning policies assigned to a group
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupId
    )
    
    try {
        # Get all provisioning policies with assignments expanded
        $uri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies?`$expand=assignments"
        Write-DebugLog "[DEBUG] Querying provisioning policies with assignments..." "DarkGray"
        $response = Invoke-MgGraphRequest -Uri $uri -Method GET
        
        Write-DebugLog "[DEBUG] Found $(@($response.value).Count) total provisioning policies" "DarkGray"
        
        $policies = @()
        foreach ($policy in $response.value) {
            Write-DebugLog "[DEBUG] Checking policy: $($policy.displayName) (ID: $($policy.id))" "DarkGray"
            
            # Check if this group is assigned to this policy
            if ($policy.assignments) {
                Write-DebugLog "[DEBUG]   - Policy has $($policy.assignments.Count) assignment(s)" "DarkGray"
                foreach ($assignment in $policy.assignments) {
                    $odataType = $assignment.target.'@odata.type'
                    $targetGroupId = $assignment.target.groupId
                    Write-DebugLog "[DEBUG]   - Assignment target: $odataType, groupId: $targetGroupId" "DarkGray"
                    if ($assignment.target.groupId -eq $GroupId) {
                        $policies += $policy.id
                        Write-InfoLog "[INFO] Group IS assigned to policy: $($policy.displayName) (ID: $($policy.id))" "Green"
                    }
                }
            }
            else {
                Write-DebugLog "[DEBUG] Policy has NO assignments" "DarkGray"
            }
        }
        
        if ($policies.Count -eq 0) {
            $timestamp = Get-Date -Format "HH:mm:ss"
            Write-Host "[$timestamp] [WARNING] No provisioning policies found for group $GroupId" -ForegroundColor Yellow
        }
        
        return $policies
    }
    catch {
        $timestamp = Get-Date -Format "HH:mm:ss"
        Write-Host "[$timestamp] [ERROR] Error getting provisioning policies: $_" -ForegroundColor Red
        Write-Log "Error getting provisioning policies for group $GroupId : $_" -Level Error
        return @()
    }
}

function Stop-CloudPCGracePeriod {
    <#
    .SYNOPSIS
        Ends the grace period for a Cloud PC to deprovision it
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CloudPCId,
        
        [Parameter(Mandatory = $false)]
        [string]$UserPrincipalName = ""
    )
    
    try {
        $userInfo = if ($UserPrincipalName) { " for user $UserPrincipalName" } else { "" }
        Write-Log "API Call: Ending grace period$userInfo (Cloud PC: $CloudPCId)" -Level Info
        Write-ApiLog "[API] POST /cloudPCs/$CloudPCId/endGracePeriod"
        
        $uri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/cloudPCs/$CloudPCId/endGracePeriod"
        Invoke-MgGraphRequest -Uri $uri -Method POST
        
        Write-Log "Successfully ended grace period$userInfo" -Level Success
        return $true
    }
    catch {
        Write-Log "Error ending grace period$userInfo : $_" -Level Error
        return $false
    }
}

function Wait-CloudPCGracePeriod {
    <#
    .SYNOPSIS
        Waits for Cloud PC to enter grace period
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,
        
        [Parameter(Mandatory = $false)]
        [int]$TimeoutMinutes = 30,
        
        [Parameter(Mandatory = $false)]
        [int]$PollingIntervalSeconds = 60,
        
        [Parameter(Mandatory = $false)]
        [ref]$CancellationToken
    )
    
    $startTime = Get-Date
    $timeout = $startTime.AddMinutes($TimeoutMinutes)
    
    Write-Log "Waiting for Cloud PC to enter grace period for user: $UserId" -Level Info
    
    while ((Get-Date) -lt $timeout) {
        if ($CancellationToken -and $CancellationToken.Value) {
            Write-Log "Operation cancelled by user" -Level Warning
            return $null
        }
        
        $cloudPC = Get-CloudPCForUser -UserId $UserId
        
        if (-not $cloudPC) {
            Write-Log "Cloud PC not found for user (may be deprovisioned)" -Level Info
            return $null
        }
        
        if ($cloudPC.status -eq 'inGracePeriod') {
            Write-Log "Cloud PC entered grace period" -Level Success
            return $cloudPC
        }
        
        Write-Log "Current status: $($cloudPC.status). Waiting..." -Level Info
        Start-Sleep -Seconds $PollingIntervalSeconds
    }
    
    Write-Log "Timeout waiting for grace period" -Level Error
    return $null
}

function Wait-CloudPCProvisioning {
    <#
    .SYNOPSIS
        Waits for new Cloud PC to provision
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,
        
        [Parameter(Mandatory = $false)]
        [int]$TimeoutMinutes = 60,
        
        [Parameter(Mandatory = $false)]
        [int]$PollingIntervalSeconds = 60,
        
        [Parameter(Mandatory = $false)]
        [ref]$CancellationToken
    )
    
    $startTime = Get-Date
    $timeout = $startTime.AddMinutes($TimeoutMinutes)
    
    Write-Log "Waiting for new Cloud PC to provision for user: $UserId" -Level Info
    
    while ((Get-Date) -lt $timeout) {
        if ($CancellationToken -and $CancellationToken.Value) {
            Write-Log "Operation cancelled by user" -Level Warning
            return $null
        }
        
        $cloudPC = Get-CloudPCForUser -UserId $UserId
        
        if ($cloudPC) {
            if ($cloudPC.status -eq 'provisioned') {
                Write-Log "Cloud PC successfully provisioned" -Level Success
                return $cloudPC
            }
            elseif ($cloudPC.status -eq 'failed') {
                Write-Log "Cloud PC provisioning failed" -Level Error
                return $cloudPC
            }
            else {
                Write-Log "Current provisioning status: $($cloudPC.status)" -Level Info
            }
        }
        else {
            Write-Log "No Cloud PC found yet, waiting for provisioning to start..." -Level Info
        }
        
        Start-Sleep -Seconds $PollingIntervalSeconds
    }
    
    Write-Log "Timeout waiting for provisioning" -Level Error
    return $null
}
#endregion

#region Group Operations
function Find-EntraIDGroups {
    <#
    .SYNOPSIS
        Searches for Microsoft Entra ID groups by display name
    .DESCRIPTION
        Searches for groups in Microsoft Entra ID that match the provided search term.
        Returns group display name and object ID.
    .PARAMETER SearchTerm
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SearchTerm
    )
    
    try {
        # Use $search parameter which supports partial matching (requires ConsistencyLevel header)
        # Search format: "displayName:searchterm" matches groups containing the term
        $searchQuery = "displayName:$SearchTerm"
        $uri = 'https://graph.microsoft.com/v1.0/groups?$search=' + [System.Web.HttpUtility]::UrlEncode('"' + $searchQuery + '"') + '&$select=id,displayName,description,onPremisesSyncEnabled&$top=100&$count=true'
        
        $allGroups = @()
        
        do {
            # $search requires ConsistencyLevel: eventual header
            $response = Invoke-MgGraphRequest -Uri $uri -Method GET -Headers @{ ConsistencyLevel = 'eventual' }
            if ($response.value -and $response.value.Count -gt 0) {
                $allGroups += @($response.value)
            }
            $uri = $response.'@odata.nextLink'
        } while ($uri)
        
        Write-DebugLog "[DEBUG] Returning $($allGroups.Count) groups" "Cyan"
        return ,$allGroups
    }
    catch {
        Write-Log "Error searching for groups: $_" -Level Error
        throw
    }
}

function Get-GroupMembers {
    <#
    .SYNOPSIS
        Gets members of a specified group with pagination support
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupId
    )
    
    try {
        $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/members?`$select=id,userPrincipalName,displayName,mail&`$top=100"
        $allMembers = @()
        
        do {
            $response = Invoke-MgGraphRequest -Uri $uri -Method GET
            
            if ($response.value -and $response.value.Count -gt 0) {
                # Filter to only users (exclude devices, groups, etc.)
                $users = $response.value | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.user' }
                $allMembers += @($users)
            }
            
            $uri = $response.'@odata.nextLink'
        } while ($uri)
        
        Write-DebugLog "[DEBUG] Get-GroupMembers: Retrieved $($allMembers.Count) users from group $GroupId" "Cyan"
        return $allMembers
    }
    catch {
        Write-Log "Error getting group members: $_" -Level Error
        throw
    }
}

function Remove-UserFromGroup {
    <#
    .SYNOPSIS
        Removes a user from a specified group
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,
        
        [Parameter(Mandatory = $true)]
        [string]$GroupId,
        
        [Parameter(Mandatory = $false)]
        [string]$UserPrincipalName = ""
    )
    
    try {
        $upnPrefix = if ($UserPrincipalName) { "${UserPrincipalName}: " } else { "" }
        Write-Log "${upnPrefix}Removing from group $GroupId" -Level Info
        Write-ApiLog "[API] ${upnPrefix}DELETE /groups/$GroupId/members/$UserId/`$ref"
        
        $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/members/$UserId/`$ref"
        Invoke-MgGraphRequest -Uri $uri -Method DELETE
        
        Write-Log "${upnPrefix}Removed from group" -Level Success
        return $true
    }
    catch {
        if ($_.Exception.Message -like "*does not exist*") {
            $upnPrefix = if ($UserPrincipalName) { "${UserPrincipalName}: " } else { "" }
            Write-Log "${upnPrefix}Already not in group" -Level Info
            return $true
        }
        $upnPrefix = if ($UserPrincipalName) { "${UserPrincipalName}: " } else { "" }
        Write-Log "${upnPrefix}Error removing from group: $_" -Level Error
        return $false
    }
}

function Add-UserToGroup {
    <#
    .SYNOPSIS
        Adds a user to a specified group
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserId,
        
        [Parameter(Mandatory = $true)]
        [string]$GroupId,
        
        [Parameter(Mandatory = $false)]
        [string]$UserPrincipalName = ""
    )
    
    try {
        $upnPrefix = if ($UserPrincipalName) { "${UserPrincipalName}: " } else { "" }
        Write-Log "${upnPrefix}Adding to group $GroupId" -Level Info
        Write-ApiLog "[API] ${upnPrefix}POST /groups/$GroupId/members/`$ref"
        
        $body = @{
            "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$UserId"
        } | ConvertTo-Json
        
        $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/members/`$ref"
        Invoke-MgGraphRequest -Uri $uri -Method POST -Body $body -ContentType "application/json"
        
        Write-Log "${upnPrefix}Added to group" -Level Success
        return $true
    }
    catch {
        if ($_.Exception.Message -like "*already exist*") {
            $upnPrefix = if ($UserPrincipalName) { "${UserPrincipalName}: " } else { "" }
            Write-Log "${upnPrefix}Already in group" -Level Info
            return $true
        }
        $upnPrefix = if ($UserPrincipalName) { "${UserPrincipalName}: " } else { "" }
        Write-Log "${upnPrefix}Error adding to group: $_" -Level Error
        return $false
    }
}
#endregion

#region Replace Logic
function Start-CloudPCReplace {
    <#
    .SYNOPSIS
        Replaces a user's Cloud PC from one provisioning profile to another
    .DESCRIPTION
        Automates the process of replacing a Cloud PC by managing group memberships.
        Removes user from source group, waits for deprovision, adds to target group,
        and monitors new Cloud PC provisioning.
    .PARAMETER UserPrincipalName
        The UPN of the user whose Cloud PC should be replaced
    .PARAMETER SourceGroupId
        Microsoft Entra ID Object ID of the group with the current provisioning policy
    .PARAMETER TargetGroupId
        Microsoft Entra ID Object ID of the group with the new provisioning policy
    .PARAMETER PollingIntervalSeconds
        How often to poll for status changes (default: 60)
    .PARAMETER GracePeriodTimeoutMinutes
        Maximum time to wait for grace period/deprovision (default: 60)
    .PARAMETER ProvisioningTimeoutMinutes
        Maximum time to wait for new Cloud PC to provision (default: 90)
    .PARAMETER CancellationToken
        Optional cancellation token for stopping the operation
    .PARAMETER ProgressCallback
        Optional callback for progress updates
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName,
        
        [Parameter(Mandatory = $true)]
        [string]$SourceGroupId,
        
        [Parameter(Mandatory = $true)]
        [string]$TargetGroupId,
        
        [Parameter(Mandatory = $false)]
        [int]$PollingIntervalSeconds = 60,
        
        [Parameter(Mandatory = $false)]
        [int]$GracePeriodTimeoutMinutes = 60,
        
        [Parameter(Mandatory = $false)]
        [int]$ProvisioningTimeoutMinutes = 90,
        
        [Parameter(Mandatory = $false)]
        [ref]$CancellationToken,
        
        [Parameter(Mandatory = $false)]
        [scriptblock]$ProgressCallback
    )
    
    $status = [ReplaceStatus]@{
        UserPrincipalName = $UserPrincipalName
        StartTime = Get-Date
        LastUpdate = Get-Date
        Stage = 'Initializing'
        Status = 'InProgress'
        ProgressPercent = 0
    }
    
    try {
        # Get user ID
        if ($ProgressCallback) { & $ProgressCallback $status }
        $status.Stage = 'Getting User Info'
        $status.LastUpdate = Get-Date
        $status.ProgressPercent = 5
        
        # Get user using Graph API directly
        $uri = 'https://graph.microsoft.com/v1.0/users?$filter=userPrincipalName eq ' + "'$UserPrincipalName'" + '&$select=id,userPrincipalName,displayName'
        $userResponse = Invoke-MgGraphRequest -Uri $uri -Method GET
        $user = $userResponse.value | Select-Object -First 1
        
        if (-not $user) {
            throw "User not found: $UserPrincipalName"
        }
        $status.UserId = $user.id
        
        # Get current Cloud PC
        if ($ProgressCallback) { & $ProgressCallback $status }
        $status.Stage = 'Getting Current Cloud PC'
        $status.LastUpdate = Get-Date
        $status.ProgressPercent = 10
        
        $cloudPC = Get-CloudPCForUser -UserId $UserPrincipalName
        if ($cloudPC) {
            $status.CloudPCId = $cloudPC.id
            Write-Log "Found existing Cloud PC: $($cloudPC.id) - Status: $($cloudPC.status)" -Level Info
        }
        
        # Step 1: Remove from source group
        if ($ProgressCallback) { & $ProgressCallback $status }
        $status.Stage = 'Removing from Source Group'
        $status.LastUpdate = Get-Date
        $status.ProgressPercent = 20
        
        $removed = Remove-UserFromGroup -UserId $user.Id -GroupId $SourceGroupId
        if (-not $removed) {
            throw "Failed to remove user from source group"
        }
        
        # Step 2: Wait for grace period
        if ($cloudPC) {
            if ($ProgressCallback) { & $ProgressCallback $status }
            $status.Stage = 'Waiting for Grace Period'
            $status.LastUpdate = Get-Date
            $status.ProgressPercent = 30
            
            $cloudPC = Wait-CloudPCGracePeriod -UserId $UserPrincipalName `
                -TimeoutMinutes $GracePeriodTimeoutMinutes `
                -PollingIntervalSeconds $PollingIntervalSeconds `
                -CancellationToken $CancellationToken
            
            if ($CancellationToken -and $CancellationToken.Value) {
                $status.Status = 'Cancelled'
                return $status
            }
            
            # Step 3: End grace period
            if ($cloudPC) {
                if ($ProgressCallback) { & $ProgressCallback $status }
                $status.Stage = 'Ending Grace Period'
                $status.LastUpdate = Get-Date
                $status.ProgressPercent = 50
                
                $ended = Stop-CloudPCGracePeriod -CloudPCId $cloudPC.id
                if (-not $ended) {
                    Write-Log "Failed to end grace period, continuing anyway..." -Level Warning
                }
            }
        }
        
        # Step 4: Add to target group
        if ($ProgressCallback) { & $ProgressCallback $status }
        $status.Stage = 'Adding to Target Group'
        $status.LastUpdate = Get-Date
        $status.ProgressPercent = 60
        
        $added = Add-UserToGroup -UserId $user.Id -GroupId $TargetGroupId
        if (-not $added) {
            throw "Failed to add user to target group"
        }
        
        # Step 5: Wait for new Cloud PC
        if ($ProgressCallback) { & $ProgressCallback $status }
        $status.Stage = 'Waiting for New Cloud PC'
        $status.LastUpdate = Get-Date
        $status.ProgressPercent = 70
        
        $newCloudPC = Wait-CloudPCProvisioning -UserId $UserPrincipalName `
            -TimeoutMinutes $ProvisioningTimeoutMinutes `
            -PollingIntervalSeconds $PollingIntervalSeconds `
            -CancellationToken $CancellationToken
        
        if ($CancellationToken -and $CancellationToken.Value) {
            $status.Status = 'Cancelled'
            return $status
        }
        
        if ($newCloudPC -and $newCloudPC.status -eq 'provisioned') {
            $status.Stage = 'Completed'
            $status.Status = 'Success'
            $status.CloudPCId = $newCloudPC.id
            $status.ProgressPercent = 100
            Write-Log "Replace completed successfully for $UserPrincipalName" -Level Success
        }
        else {
            $status.Stage = 'Failed'
            $status.Status = 'Failed'
            $status.ProgressPercent = 70
            $status.ErrorMessage = "New Cloud PC did not provision successfully"
            Write-Log $status.ErrorMessage -Level Error
        }
    }
    catch {
        $status.Stage = 'Failed'
        $status.Status = 'Failed'
        $status.ProgressPercent = 0
        $status.ErrorMessage = $_.Exception.Message
        Write-Log "Replace failed for $UserPrincipalName : $($status.ErrorMessage)" -Level Error
    }
    
    $status.LastUpdate = Get-Date
    if ($ProgressCallback) { & $ProgressCallback $status }
    return $status
}
#endregion

#region Logging
function Write-Log {
    <#
    .SYNOPSIS
        Writes log messages to file and console
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Success', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Console output with colors
    switch ($Level) {
        'Success' { Write-Host $logMessage -ForegroundColor Green }
        'Warning' { Write-Host $logMessage -ForegroundColor Yellow }
        'Error' { Write-Host $logMessage -ForegroundColor Red }
        default { Write-Host $logMessage }
    }
    
    # File output
    if ($script:LogFilePath) {
        Add-Content -Path $script:LogFilePath -Value $logMessage
    }
}

function Initialize-Logging {
    <#
    .SYNOPSIS
        Initializes logging for the session
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$LogPath = ".\CloudPCReplace_Logs"
    )
    
    if (-not (Test-Path $LogPath)) {
        New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
    }
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $script:LogFilePath = Join-Path $LogPath "Replace_$timestamp.log"
    
    Write-Log "=== Cloud PC replace Session Started ===" -Level Info
    return $script:LogFilePath
}
#endregion

#region Helper Functions
function Test-GraphConnection {
    <#
    .SYNOPSIS
        Tests if connected to Microsoft Graph
    #>
    [CmdletBinding()]
    param()
    
    try {
        $context = Get-MgContext
        return ($null -ne $context)
    }
    catch {
        return $false
    }
}
#endregion

# Export module members
Export-ModuleMember -Function @(
    'Connect-MgGraphForReplace',
    'Start-CloudPCReplace',
    'Get-CloudPCForUser',
    'Get-UserGroupMemberships',
    'Get-ProvisioningPoliciesForGroup',
    'Stop-CloudPCGracePeriod',
    'Remove-UserFromGroup',
    'Add-UserToGroup',
    'Find-EntraIDGroups',
    'Get-GroupMembers',
    'Initialize-Logging',
    'Write-Log',
    'Test-GraphConnection'
)
