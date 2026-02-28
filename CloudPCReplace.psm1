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
    Write-Host "[$timestamp] [API   ] $Message" -ForegroundColor $Color
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

class UserReplaceState {
    [string]$UserPrincipalName
    [string]$UserId
    [string]$SourceGroupId        # Each job has its own source
    [string]$SourceGroupName      # For display in grid
    [string]$TargetGroupId        # Each job has its own target
    [string]$TargetGroupName      # For display in grid
    [array]$SourcePolicyIds       # Cached provisioning policy IDs for this job's source group
    [array]$TargetPolicyIds       # Cached provisioning policy IDs for this job's target group
    [string]$Stage
    [string]$Status  # Queued, InProgress, Success, Failed
    [int]$ProgressPercent
    [datetime]$StartTime
    [datetime]$StageStartTime
    [datetime]$LastPollTime
    [datetime]$EndTime
    [string]$ErrorMessage
    [array]$OldCPCs = @()         # Array of hashtables: @{Id, Name, ServicePlan}
    [array]$NewCPCs = @()         # Array of hashtables: @{Id, Name, ServicePlan}
    [array]$CloudPCIds     # DEPRECATED - kept for backward compat, use OldCPCs instead
    [array]$GraceEndedCPCIds  # Track which CPCs we've already ended grace on
    [hashtable]$DeprovisioningSeenCPCs = @{}  # Track which CPCs have reached 'deprovisioning' status
    [int]$GridRowIndex
    [int]$QueueOrder = 0  # Lower number = higher priority
    [string]$NextPollDisplay = "-"  # Pre-calculated display value for NextPoll column
    
    # Summary fields for CSV export
    [string]$OldCPCName = ""
    [string]$NewCPCName = ""
    [string]$FinalMessage = ""
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
        Write-Log "Connecting to Microsoft Graph..." -Level 'Info  '
        
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
        
        Write-Log "Connected to tenant: $tenantInfo (Account: $accountInfo)" -Level 'OK    '
        
        return $true
    }
    catch {
        Write-Log "Failed to connect to Microsoft Graph: $_" -Level 'FAIL  '
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
        Write-Log "Error getting Cloud PC for user $UserId : $_" -Level 'FAIL  '
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
        Write-Log "Error getting group memberships for user $UserId : $_" -Level 'FAIL  '
        return @()
    }
}

function Get-UserInfo {
    <#
    .SYNOPSIS
        Gets a user's basic info (id, userPrincipalName, displayName) by UPN
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName
    )
    $uri = "https://graph.microsoft.com/v1.0/users?`$filter=userPrincipalName eq '$UserPrincipalName'&`$select=id,userPrincipalName,displayName"
    $response = Invoke-MgGraphRequest -Uri $uri -Method GET
    return $response.value | Select-Object -First 1
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
        Write-Log "Error getting provisioning policies for group $GroupId : $_" -Level 'FAIL  '
        return @()
    }
}

function Get-EnterprisePolicyGroups {
    <#
    .SYNOPSIS
        Gets all Entra ID groups assigned to Enterprise Cloud PC provisioning policies
    #>
    try {
        Write-Log "API Call: Fetching Enterprise provisioning policies with assignments" -Level 'Info  '
        $uri      = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/provisioningPolicies?`$expand=assignments"
        $response = Invoke-MgGraphRequest -Uri $uri -Method GET

        $groupIds    = [System.Collections.Generic.HashSet[string]]::new()
        $groupPolicyMap = @{}   # groupId → policy displayName(s)
        $policyCount = 0
        foreach ($policy in @($response.value)) {
            # Skip Frontline Worker policies (cloudPcType = 'frontline')
            if ($policy.cloudPcType -eq 'frontline') {
                Write-DebugLog "[DEBUG] Skipping Frontline policy: $($policy.displayName)" "DarkGray"
                continue
            }
            $policyCount++
            foreach ($assignment in @($policy.assignments)) {
                $gid = $assignment.target.groupId
                if ($gid) {
                    $groupIds.Add($gid) | Out-Null
                    if (-not $groupPolicyMap.ContainsKey($gid)) { $groupPolicyMap[$gid] = @() }
                    $groupPolicyMap[$gid] += $policy.displayName
                }
            }
        }
        Write-Log "Found $policyCount Enterprise policy/policies covering $($groupIds.Count) unique assigned group(s)" -Level 'Info  '
        if ($groupIds.Count -eq 0) { return }

        # Bulk-resolve group details via getByIds (max 1000 per call)
        $allGroups = @()
        $idList    = @($groupIds)
        for ($i = 0; $i -lt $idList.Count; $i += 1000) {
            $batch    = $idList[$i..([math]::Min($i + 999, $idList.Count - 1))]
            $body     = @{ ids = $batch; types = @('group') } | ConvertTo-Json
            $resolved = Invoke-MgGraphRequest `
                -Uri         "https://graph.microsoft.com/v1.0/directoryObjects/getByIds?`$select=id,displayName,onPremisesSyncEnabled,groupTypes" `
                -Method      POST `
                -Body        $body `
                -ContentType 'application/json'
            $allGroups += @($resolved.value | ForEach-Object {
                [PSCustomObject]@{
                    id                    = $_.id
                    displayName           = $_.displayName
                    onPremisesSyncEnabled = $_.onPremisesSyncEnabled
                    groupTypes            = $_.groupTypes
                    policyName            = if ($groupPolicyMap.ContainsKey($_.id)) { $groupPolicyMap[$_.id] -join ', ' } else { '' }
                }
            })
        }
        return $allGroups
    }
    catch {
        Write-Log "Error fetching Enterprise policy groups: $_" -Level 'FAIL  '
        throw
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
        Write-Log "API Call: Ending grace period$userInfo (Cloud PC: $CloudPCId)" -Level 'Info  '
        Write-ApiLog "POST /cloudPCs/$CloudPCId/endGracePeriod"
        
        $uri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/cloudPCs/$CloudPCId/endGracePeriod"
        Invoke-MgGraphRequest -Uri $uri -Method POST
        
        Write-Log "Successfully ended grace period$userInfo" -Level 'OK    '
        return $true
    }
    catch {
        Write-Log "Error ending grace period$userInfo : $_" -Level 'FAIL  '
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
    
    Write-Log "Waiting for Cloud PC to enter grace period for user: $UserId" -Level 'Info  '
    
    while ((Get-Date) -lt $timeout) {
        if ($CancellationToken -and $CancellationToken.Value) {
            Write-Log "Operation cancelled by user" -Level 'WARN  '
            return $null
        }
        
        $cloudPC = Get-CloudPCForUser -UserId $UserId
        
        if (-not $cloudPC) {
            Write-Log "Cloud PC not found for user (may be deprovisioned)" -Level 'Info  '
            return $null
        }
        
        if ($cloudPC.status -eq 'inGracePeriod') {
            Write-Log "Cloud PC entered grace period" -Level 'OK    '
            return $cloudPC
        }
        
        Write-Log "Current status: $($cloudPC.status). Waiting..." -Level 'Info  '
        Start-Sleep -Seconds $PollingIntervalSeconds
    }
    
    Write-Log "Timeout waiting for grace period" -Level 'FAIL  '
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
    
    Write-Log "Waiting for new Cloud PC to provision for user: $UserId" -Level 'Info  '
    
    while ((Get-Date) -lt $timeout) {
        if ($CancellationToken -and $CancellationToken.Value) {
            Write-Log "Operation cancelled by user" -Level 'WARN  '
            return $null
        }
        
        $cloudPC = Get-CloudPCForUser -UserId $UserId
        
        if ($cloudPC) {
            if ($cloudPC.status -eq 'provisioned') {
                Write-Log "Cloud PC successfully provisioned" -Level 'OK    '
                return $cloudPC
            }
            elseif ($cloudPC.status -eq 'failed') {
                Write-Log "Cloud PC provisioning failed" -Level 'FAIL  '
                return $cloudPC
            }
            else {
                Write-Log "Current provisioning status: $($cloudPC.status)" -Level 'Info  '
            }
        }
        else {
            Write-Log "No Cloud PC found yet, waiting for provisioning to start..." -Level 'Info  '
        }
        
        Start-Sleep -Seconds $PollingIntervalSeconds
    }
    
    Write-Log "Timeout waiting for provisioning" -Level 'FAIL  '
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
        $uri = 'https://graph.microsoft.com/v1.0/groups?$search=' + [System.Web.HttpUtility]::UrlEncode('"' + $searchQuery + '"') + '&$select=id,displayName,description,onPremisesSyncEnabled,groupTypes&$top=100&$count=true'
        
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
        Write-Log "Error searching for groups: $_" -Level 'FAIL  '
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
        Write-Log "Error getting group members: $_" -Level 'FAIL  '
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
        Write-Log "${upnPrefix}Removing from group $GroupId" -Level 'Info  '
        Write-ApiLog "${upnPrefix}DELETE /groups/$GroupId/members/$UserId/`$ref"
        
        $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/members/$UserId/`$ref"
        Invoke-MgGraphRequest -Uri $uri -Method DELETE
        
        Write-Log "${upnPrefix}Removed from group" -Level 'OK    '
        return $true
    }
    catch {
        if ($_.Exception.Message -like "*does not exist*") {
            $upnPrefix = if ($UserPrincipalName) { "${UserPrincipalName}: " } else { "" }
            Write-Log "${upnPrefix}Already not in group" -Level 'Info  '
            return $true
        }
        $upnPrefix = if ($UserPrincipalName) { "${UserPrincipalName}: " } else { "" }
        Write-Log "${upnPrefix}Error removing from group: $_" -Level 'FAIL  '
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
        Write-Log "${upnPrefix}Adding to group $GroupId" -Level 'Info  '
        Write-ApiLog "${upnPrefix}POST /groups/$GroupId/members/`$ref"
        
        $body = @{
            "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$UserId"
        } | ConvertTo-Json
        
        $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/members/`$ref"
        Invoke-MgGraphRequest -Uri $uri -Method POST -Body $body -ContentType "application/json"
        
        Write-Log "${upnPrefix}Added to group" -Level 'OK    '
        return $true
    }
    catch {
        if ($_.Exception.Message -like "*already exist*") {
            $upnPrefix = if ($UserPrincipalName) { "${UserPrincipalName}: " } else { "" }
            Write-Log "${upnPrefix}Already in group" -Level 'Info  '
            return $true
        }
        $upnPrefix = if ($UserPrincipalName) { "${UserPrincipalName}: " } else { "" }
        Write-Log "${upnPrefix}Error adding to group: $_" -Level 'FAIL  '
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
            Write-Log "Found existing Cloud PC: $($cloudPC.id) - Status: $($cloudPC.status)" -Level 'Info  '
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
                    Write-Log "Failed to end grace period, continuing anyway..." -Level 'WARN  '
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
            Write-Log "Replace completed successfully for $UserPrincipalName" -Level 'OK    '
        }
        else {
            $status.Stage = 'Failed'
            $status.Status = 'Failed'
            $status.ProgressPercent = 70
            $status.ErrorMessage = "New Cloud PC did not provision successfully"
            Write-Log $status.ErrorMessage -Level 'FAIL  '
        }
    }
    catch {
        $status.Stage = 'Failed'
        $status.Status = 'Failed'
        $status.ProgressPercent = 0
        $status.ErrorMessage = $_.Exception.Message
        Write-Log "Replace failed for $UserPrincipalName : $($status.ErrorMessage)" -Level 'FAIL  '
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
        [ValidateSet('Info  ', 'OK    ', 'WARN  ', 'FAIL  ')]
        [string]$Level = 'Info  '
    )
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Console output with colors
    switch ($Level) {
        'OK    ' { Write-Host $logMessage -ForegroundColor Green }
        'WARN  ' { Write-Host $logMessage -ForegroundColor Yellow }
        'FAIL  ' { Write-Host $logMessage -ForegroundColor Red }
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
    
    Write-Log "=== Cloud PC replace Session Started ===" -Level 'Info  '
    return $script:LogFilePath
}
#endregion

#region Helper Functions
function Get-StageDisplay {
    <#
    .SYNOPSIS
        Formats stage name with step number for display
    .PARAMETER Stage
        The stage name to format
    .EXAMPLE
        Get-StageDisplay -Stage "Waiting for Grace Period"
        Returns: "5/9 - Waiting for Grace Period"
    #>
    [CmdletBinding()]
    param([string]$Stage)
    
    $stageMap = @{
        'Queued' = '1/9'
        'Getting User Info' = '2/9'
        'Getting Current Cloud PC' = '3/9'
        'Removing from Source' = '4/9'
        'Waiting for Grace Period' = '5/9'
        'Ending Grace Period' = '6/9'
        'Waiting for Deprovision' = '7/9'
        'Adding to Target' = '8/9'
        'Waiting for Provisioning' = '9/9'
        'Complete' = 'Done'
    }
    
    $prefix = $stageMap[$Stage]
    if ($prefix) {
        return $prefix + ' - ' + $Stage
    }
    return $Stage
}

function Invoke-CloudPCReplaceStep {
    <#
    .SYNOPSIS
        Processes one step of the Cloud PC replacement state machine
    .DESCRIPTION
        Examines the current state and performs the next action in the replacement process.
        This function is called repeatedly by the UI on a timer to advance through stages.
        It performs ONE action per call and updates the state accordingly.
    .PARAMETER State
        The UserReplaceState object containing the current job state
    .PARAMETER Timeouts
        Hashtable containing timeout values in minutes:
        - GracePeriodTimeout
        - EndingGracePeriodTimeout
        - DeprovisionTimeout
        - ProvisioningTimeout
    .PARAMETER OnLog
        Scriptblock called for logging: { param($Message, $Level, $Color) }
        Levels: "Status", "Info", "Debug", "Polling", "Verbose"
    .PARAMETER OnGridUpdate
        Optional scriptblock called for grid updates: { param($State, $ColumnName, $Value) }
    .EXAMPLE
        Invoke-CloudPCReplaceStep -State $jobState -Timeouts $timeouts `
            -OnLog { param($msg,$lvl,$clr) Write-Host $msg -ForegroundColor $clr }
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [UserReplaceState]$State,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$Timeouts,
        
        [Parameter(Mandatory = $true)]
        [scriptblock]$OnLog,
        
        [Parameter(Mandatory = $false)]
        [scriptblock]$OnGridUpdate
    )
    
    # Helper function to log messages
    function Write-StepLog {
        param([string]$Message, [string]$Level = "Info", [string]$Color = "White")
        & $OnLog $Message $Level $Color
    }
    
    try {
        switch ($State.Stage) {
            "Getting User Info" {
                # Skip Graph lookup if UserId already set (e.g. pre-populated from group member list)
                if (-not [string]::IsNullOrEmpty($State.UserId)) {
                    Write-StepLog "[Action] $($State.UserPrincipalName): User info pre-populated (ID: $($State.UserId))" "Status" "Green"
                    $State.Stage = "Getting Current Cloud PC"
                    $State.ProgressPercent = 10
                    break
                }
                Write-StepLog "[Action] $($State.UserPrincipalName): Getting user info..." "Status" "Cyan"
                
                $user = Get-UserInfo -UserPrincipalName $State.UserPrincipalName
                
                if (-not $user) { throw "User not found" }
                $State.UserId = $user.id
                Write-StepLog "[Action] $($State.UserPrincipalName): User found (ID: $($State.UserId))" "Status" "Green"
                
                $State.Stage = "Getting Current Cloud PC"
                $State.ProgressPercent = 10
            }
            
            "Getting Current Cloud PC" {
                Write-StepLog "[Action] $($State.UserPrincipalName): Checking for existing Cloud PC..." "Status" "Cyan"
                
                if (-not $State.SourcePolicyIds) {
                    Write-StepLog "[Debug ] Getting provisioning policies for source group: $($State.SourceGroupName)" "Debug" "Gray"
                    $State.SourcePolicyIds = @(Get-ProvisioningPoliciesForGroup -GroupId $State.SourceGroupId)
                    if ($State.SourcePolicyIds.Count -gt 0) {
                        Write-StepLog "[Debug ] Source group uses $($State.SourcePolicyIds.Count) provisioning policy(ies)" "Debug" "Gray"
                    }
                }
                
                $cloudPCs = Get-CloudPCForUser -UserId $State.UserPrincipalName
                $cloudPCsArray = @($cloudPCs)
                
                if ($cloudPCsArray.Count -gt 0) {
                    Write-StepLog "[Debug ] $($State.UserPrincipalName): Found $($cloudPCsArray.Count) Cloud PC(s) for user" "Debug" "Yellow"
                    
                    foreach ($cpc in $cloudPCsArray) {
                        $cpcName = if ($cpc.managedDeviceName) { $cpc.managedDeviceName } else { $cpc.displayName }
                        Write-StepLog "[Debug ]   $($State.UserPrincipalName): CPC: $cpcName | Status: $($cpc.status) | ID: $($cpc.id)" "Debug" "Yellow"
                        Write-StepLog "[Debug ]     $($State.UserPrincipalName): Policy: $($cpc.provisioningPolicyId)" "Debug" "Yellow"
                    }
                    
                    $matchingCPCs = @()
                    if ($State.SourcePolicyIds -and $State.SourcePolicyIds.Count -gt 0) {
                        Write-StepLog "[Debug ] $($State.UserPrincipalName): Looking for CPCs matching source policy..." "Debug" "Cyan"
                        $matchingCPCs = @($cloudPCsArray | Where-Object { $State.SourcePolicyIds -contains $_.provisioningPolicyId })
                    }
                    
                    if ($matchingCPCs.Count -gt 0) {
                        $State.OldCPCs = @($matchingCPCs | ForEach-Object {
                            @{
                                Id = $_.id
                                Name = if ($_.managedDeviceName) { $_.managedDeviceName } else { $_.displayName }
                                ServicePlan = $_.servicePlanName
                            }
                        })
                        
                        $State.CloudPCIds = @($matchingCPCs | ForEach-Object { $_.id })
                        
                        Write-StepLog "[Debug ] Captured $($State.CloudPCIds.Count) old CPC ID(s) for tracking: $($State.CloudPCIds -join ', ')" "Debug" "Magenta"
                        
                        $oldCPCNames = $matchingCPCs | ForEach-Object {
                            if ($_.managedDeviceName) { $_.managedDeviceName } else { $_.displayName }
                        }
                        $State.OldCPCName = $oldCPCNames -join ", "
                        
                        Write-StepLog "[Action] $($State.UserPrincipalName): Found $($matchingCPCs.Count) CPC(s) from SOURCE policy" "Status" "Green"
                        foreach ($cpc in $matchingCPCs) {
                            $cpcName = if ($cpc.managedDeviceName) { $cpc.managedDeviceName } else { $cpc.displayName }
                            Write-StepLog "[Action]   OLD: $cpcName | ID: $($cpc.id) | Plan: $($cpc.servicePlanName)" "Status" "Cyan"
                        }
                        
                        if ($matchingCPCs.Count -gt 1) {
                            Write-StepLog "[Info  ] $($State.UserPrincipalName): User has MULTIPLE CPCs from same policy (likely multiple licenses)" "Status" "Cyan"
                            Write-StepLog "[Info  ] $($State.UserPrincipalName): All will be deprovisioned before provisioning new CPC(s)" "Info" "Cyan"
                        }
                    }
                    else {
                        Write-StepLog "[WARN  ] $($State.UserPrincipalName): Has $($cloudPCs.Count) CPC(s) but NONE match source policy!" "Status" "Yellow"
                        Write-StepLog "[Info  ] Source policy IDs: $($State.SourcePolicyIds -join ', ')" "Info" "Yellow"
                        $State.OldCPCs = @()
                        $State.CloudPCIds = @()
                    }
                }
                else {
                    Write-StepLog "[Action] $($State.UserPrincipalName): No existing Cloud PC found" "Status" "Yellow"
                    $State.OldCPCs = @()
                    $State.CloudPCIds = @()
                }
                
                Write-StepLog "[Debug ] Current group memberships:" "Debug" "Magenta"
                $currentGroups = Get-UserGroupMemberships -UserId $State.UserId
                foreach ($grp in $currentGroups) {
                    Write-StepLog "[Debug ]   - Member of: $($grp.displayName) (ID: $($grp.id))" "Debug" "Magenta"
                }
                
                $State.Stage = "Removing from Source"
                $State.ProgressPercent = 20
                $State.StageStartTime = Get-Date
            }
            
            "Removing from Source" {
                if (-not $State.CloudPCIds -or $State.CloudPCIds.Count -eq 0) {
                    Write-StepLog "[FAIL  ] $($State.UserPrincipalName): User has NO Cloud PC matching source policy!" "Status" "Red"
                    Write-StepLog "[FAIL  ] Cannot replace - admin must investigate manually" "Status" "Red"
                    Write-StepLog "[FAIL  ] User remains in source group (no changes made)" "Status" "Red"
                    throw "No Cloud PC found matching source provisioning policy. User needs manual intervention."
                }
                
                Write-StepLog "[Action] $($State.UserPrincipalName): Removing from source group ($($State.SourceGroupName))..." "Status" "Cyan"
                Remove-UserFromGroup -UserId $State.UserId -GroupId $State.SourceGroupId -UserPrincipalName $State.UserPrincipalName | Out-Null
                Write-StepLog "[Action] $($State.UserPrincipalName): Removed from source group" "Status" "Green"
                Write-StepLog "[Info  ] $($State.UserPrincipalName): Tracking $($State.CloudPCIds.Count) CPC(s) - waiting for grace period..." "Info" "Cyan"
                
                $State.Stage = "Waiting for Grace Period"
                $State.ProgressPercent = 30
                $State.StageStartTime = Get-Date
            }
            
            "Waiting for Grace Period" {
                if (-not $State.CloudPCIds -or $State.CloudPCIds.Count -eq 0) {
                    Write-StepLog "[FAIL  ] No CPC IDs to track - should not be in grace period stage!" "Status" "Red"
                    throw "Logic error: No CPC IDs set"
                }
                
                Write-StepLog "[Poll  ] $($State.UserPrincipalName): Polling Graph API - checking CPC grace period status..." "Polling" "Gray"
                $cloudPCs = Get-CloudPCForUser -UserId $State.UserPrincipalName
                $trackedCPCs = @($cloudPCs | Where-Object { $State.CloudPCIds -contains $_.id })
                
                if ($trackedCPCs.Count -eq 0) {
                    Write-StepLog "[Detect] $($State.UserPrincipalName): All tracked CPCs disappeared (already deprovisioned)" "Status" "Green"
                    # Log the old CPC details before clearing
                    foreach ($oldCPC in $State.OldCPCs) {
                        Write-StepLog "[Detect]   OLD: $($oldCPC.Name) | ID: $($oldCPC.Id) - deprovisioned" "Status" "Gray"
                    }
                    # Clear CloudPCIds so any CPC found later is considered "new"
                    # NOTE: Azure appears to reuse the same cloudPcId GUID when deprovisioning/reprovisioning
                    # for the same user. The device name changes but the Cloud PC object ID may persist.
                    # Clearing this array ensures we detect the reprovisioned CPC as "new".
                    Write-StepLog "[Debug ] [Grace] Clearing old CPC ID tracking array (was: $($State.CloudPCIds -join ', '))" "Debug" "Magenta"
                    $State.CloudPCIds = @()
                    Write-StepLog "[Debug ] After clear, CloudPCIds.Count = $($State.CloudPCIds.Count)" "Debug" "Magenta"
                    $State.Stage = "Adding to Target"
                    $State.ProgressPercent = 60
                    $State.StageStartTime = Get-Date
                }
                else {
                    $anyDeprovisioning = $trackedCPCs | Where-Object { $_.status -in @('deprovisioning', 'notProvisioned') }
                    
                    if ($anyDeprovisioning) {
                        Write-StepLog "[Detect] $($State.UserPrincipalName): CPCs already deprovisioning - skipping grace period stages!" "Status" "Yellow"
                        foreach ($cpc in $anyDeprovisioning) {
                            $cpcName = if ($cpc.managedDeviceName) { $cpc.managedDeviceName } else { $cpc.displayName }
                            Write-StepLog "[Detect]   - $cpcName (status: $($cpc.status))" "Status" "Yellow"
                        }
                        # Log the old CPC details if we're skipping ahead
                        foreach ($oldCPC in $State.OldCPCs) {
                            Write-StepLog "[Info  ]   OLD: $($oldCPC.Name) | ID: $($oldCPC.Id)" "Info" "Gray"
                        }
                        $State.Stage = "Waiting for Deprovision"
                        $State.ProgressPercent = 55
                        $State.StageStartTime = Get-Date
                    }
                    else {
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
                            Write-StepLog "[Detect] $($State.UserPrincipalName): ALL $($trackedCPCs.Count) tracked CPCs entered grace period!" "Status" "Green"
                            foreach ($cpc in $trackedCPCs) {
                                $cpcName = if ($cpc.managedDeviceName) { $cpc.managedDeviceName } else { $cpc.displayName }
                                Write-StepLog "[Detect]   - $cpcName (status: $($cpc.status))" "Status" "Green"
                            }
                            
                            $State.Stage = "Ending Grace Period"
                            $State.ProgressPercent = 40
                            $State.StageStartTime = Get-Date
                        }
                        else {
                            Write-StepLog "[Poll  ] $($State.UserPrincipalName): Waiting for all CPCs to enter grace..." "Polling" "Gray"
                            foreach ($status in $statusSummary.Keys) {
                                $names = $statusSummary[$status] -join ', '
                                Write-StepLog "[Debug ]   Status '$status': $names" "Debug" "Gray"
                            }
                            
                            $elapsed = (Get-Date) - $State.StageStartTime
                            if ($elapsed.TotalMinutes -gt $Timeouts.GracePeriodTimeout) {
                                throw "Timeout waiting for grace period"
                            }
                        }
                    }
                }
            }
            
            "Ending Grace Period" {
                if (-not $State.GraceEndedCPCIds -or $State.GraceEndedCPCIds.Count -eq 0) {
                    Write-StepLog "[Action] $($State.UserPrincipalName): Ending grace period on $($State.CloudPCIds.Count) tracked CPC(s)..." "Status" "Cyan"
                    
                    $State.GraceEndedCPCIds = @()
                    
                    foreach ($cpcId in $State.CloudPCIds) {
                        Write-StepLog "[Debug ] Ending grace for CPC ID: $cpcId" "Debug" "Gray"
                        Stop-CloudPCGracePeriod -CloudPCId $cpcId -UserPrincipalName $State.UserPrincipalName | Out-Null
                        $State.GraceEndedCPCIds += $cpcId
                    }
                    
                    Write-StepLog "[Action] $($State.UserPrincipalName): End grace API called - waiting for deprovisioning to start..." "Status" "Green"
                }
                
                $cloudPCs = Get-CloudPCForUser -UserId $State.UserPrincipalName
                $trackedCPCs = @($cloudPCs | Where-Object { $State.CloudPCIds -contains $_.id })
                
                if ($trackedCPCs.Count -eq 0) {
                    Write-StepLog "[Detect] $($State.UserPrincipalName): All tracked CPCs disappeared (already deprovisioned)" "Status" "Green"
                    # Log the old CPC details before clearing
                    foreach ($oldCPC in $State.OldCPCs) {
                        Write-StepLog "[Detect]   OLD: $($oldCPC.Name) | ID: $($oldCPC.Id) - deprovisioned" "Status" "Gray"
                    }
                    # Clear CloudPCIds so any CPC found later is considered "new"
                    # NOTE: Azure appears to reuse the same cloudPcId GUID when deprovisioning/reprovisioning
                    # for the same user. The device name changes but the Cloud PC object ID may persist.
                    # Clearing this array ensures we detect the reprovisioned CPC as "new".
                    Write-StepLog "[Debug ] [EndGrace] Clearing old CPC ID tracking array (was: $($State.CloudPCIds -join ', '))" "Debug" "Magenta"
                    $State.CloudPCIds = @()
                    Write-StepLog "[Debug ] After clear, CloudPCIds.Count = $($State.CloudPCIds.Count)" "Debug" "Magenta"
                    $State.Stage = "Adding to Target"
                    $State.ProgressPercent = 60
                    $State.StageStartTime = Get-Date
                }
                else {
                    $anyDeprovisioning = $trackedCPCs | Where-Object { $_.status -in @('deprovisioning', 'notProvisioned') }
                    
                    if ($anyDeprovisioning) {
                        Write-StepLog "[Detect] $($State.UserPrincipalName): Deprovisioning confirmed - backend processing started!" "Status" "Green"
                        foreach ($cpc in $anyDeprovisioning) {
                            $cpcName = if ($cpc.managedDeviceName) { $cpc.managedDeviceName } else { $cpc.displayName }
                            Write-StepLog "[Detect]   - $cpcName (status: $($cpc.status))" "Status" "Green"
                        }
                        
                        $State.Stage = "Waiting for Deprovision"
                        $State.ProgressPercent = 55
                        $State.StageStartTime = Get-Date
                        $State.LastPollTime = Get-Date
                    }
                    else {
                        $elapsed = (Get-Date) - $State.StageStartTime
                        Write-StepLog "[Poll  ] $($State.UserPrincipalName): Waiting for deprovisioning to start (elapsed: $([math]::Round($elapsed.TotalMinutes, 1))min)..." "Polling" "Gray"
                        
                        foreach ($cpc in $trackedCPCs) {
                            $cpcName = if ($cpc.managedDeviceName) { $cpc.managedDeviceName } else { $cpc.displayName }
                            Write-StepLog "[Debug ]   - $cpcName (status: $($cpc.status))" "Debug" "Gray"
                        }
                        
                        if ($elapsed.TotalMinutes -gt $Timeouts.EndingGracePeriodTimeout) {
                            throw "Timeout waiting for deprovisioning to start after ending grace period"
                        }
                    }
                }
            }
            
            "Waiting for Deprovision" {
                Write-StepLog "[Poll  ] $($State.UserPrincipalName): Polling Graph API - checking CPC deprovision status..." "Polling" "Gray"
                $cloudPCs = Get-CloudPCForUser -UserId $State.UserPrincipalName
                $trackedCPCs = @($cloudPCs | Where-Object { $State.CloudPCIds -contains $_.id })
                
                $activeCPCs = @()
                foreach ($cpc in $trackedCPCs) {
                    if ($cpc.status -eq 'notProvisioned') {
                        continue
                    }
                    
                    if ($cpc.status -eq 'deprovisioning' -and -not $State.DeprovisioningSeenCPCs.ContainsKey($cpc.id)) {
                        $State.DeprovisioningSeenCPCs[$cpc.id] = $true
                        Write-StepLog "[Detect] CPC $($cpc.id) reached 'deprovisioning' status - backend is processing" "Status" "Green"
                    }
                    
                    if ($cpc.status -eq 'inGracePeriod' -and $State.GraceEndedCPCIds -contains $cpc.id) {
                        $timeSinceGraceEnded = ((Get-Date) - $State.StageStartTime).TotalMinutes
                        
                        $statusInfo = ""
                        if ($cpc.statusDetails) {
                            $statusInfo = " | statusDetails: $($cpc.statusDetails.code) - $($cpc.statusDetails.message)"
                        }
                        if ($cpc.lastModifiedDateTime) {
                            $lastModified = [DateTime]::SpecifyKind([DateTime]::Parse($cpc.lastModifiedDateTime), [DateTimeKind]::Utc)
                            $minutesSinceModified = ((Get-Date).ToUniversalTime() - $lastModified).TotalMinutes
                            $statusInfo += " | Last modified: $([math]::Round($minutesSinceModified, 1))min ago"
                        }
                        
                        if ($State.DeprovisioningSeenCPCs.ContainsKey($cpc.id)) {
                            Write-StepLog "[Debug ] CPC $($cpc.id) flip-flopped back to inGracePeriod (seen deprovisioning before) - API inconsistency$statusInfo" "Debug" "Yellow"
                            if ($OnGridUpdate) {
                                & $OnGridUpdate $State "Messages" "Status flip-flop detected (API lag)"
                            }
                            $activeCPCs += $cpc
                        }
                        else {
                            Write-StepLog "[Debug ] CPC $($cpc.id) still shows inGracePeriod $([math]::Round($timeSinceGraceEnded, 1))min after ending grace - waiting for backend$statusInfo" "Debug" "Gray"
                            $activeCPCs += $cpc
                        }
                    }
                    elseif ($cpc.status -eq 'deprovisioning') {
                        $activeCPCs += $cpc
                    }
                    else {
                        Write-StepLog "[WARN  ] CPC $($cpc.id) in unexpected status: $($cpc.status)" "Status" "Yellow"
                        $activeCPCs += $cpc
                    }
                }
                
                if ($activeCPCs.Count -eq 0) {
                    Write-StepLog "[Detect] $($State.UserPrincipalName): All tracked CPCs deprovisioned!" "Status" "Green"
                    # Log the old CPC details before clearing
                    foreach ($oldCPC in $State.OldCPCs) {
                        Write-StepLog "[Detect]   OLD: $($oldCPC.Name) | ID: $($oldCPC.Id) - deprovisioned" "Status" "Gray"
                    }
                    # Clear CloudPCIds so any CPC found later is considered "new"
                    # NOTE: Azure appears to reuse the same cloudPcId GUID when deprovisioning/reprovisioning
                    # for the same user. The device name changes but the Cloud PC object ID may persist.
                    # Clearing this array ensures we detect the reprovisioned CPC as "new".
                    Write-StepLog "[Debug ] Clearing old CPC ID tracking array (was: $($State.CloudPCIds -join ', '))" "Debug" "Magenta"
                    $State.CloudPCIds = @()
                    Write-StepLog "[Debug ] After clear, CloudPCIds.Count = $($State.CloudPCIds.Count)" "Debug" "Magenta"
                    $State.Stage = "Adding to Target"
                    $State.ProgressPercent = 60
                    $State.StageStartTime = Get-Date
                }
                else {
                    $elapsed = (Get-Date) - $State.StageStartTime
                    Write-StepLog "[Poll  ] $($State.UserPrincipalName): Waiting for $($activeCPCs.Count) CPC(s) to deprovision (elapsed: $([math]::Round($elapsed.TotalMinutes, 1))min)..." "Polling" "Gray"
                    
                    foreach ($cpc in $activeCPCs) {
                        $cpcName = if ($cpc.managedDeviceName) { $cpc.managedDeviceName } else { $cpc.displayName }
                        Write-StepLog "[Debug ]   - $cpcName (status: $($cpc.status))" "Debug" "Gray"
                    }
                    
                    if ($elapsed.TotalMinutes -gt $Timeouts.DeprovisionTimeout) {
                        throw "Timeout waiting for Cloud PCs to deprovision"
                    }
                }
            }
            
            "Adding to Target" {
                Write-StepLog "[Action] $($State.UserPrincipalName): Adding to target group ($($State.TargetGroupName))..." "Status" "Cyan"
                Add-UserToGroup -UserId $State.UserId -GroupId $State.TargetGroupId -UserPrincipalName $State.UserPrincipalName | Out-Null
                Write-StepLog "[Action] $($State.UserPrincipalName): Added to target group" "Status" "Green"
                Write-StepLog "[Info  ] $($State.UserPrincipalName): Waiting for new Cloud PC to provision..." "Info" "Cyan"
                
                $State.Stage = "Waiting for Provisioning"
                $State.ProgressPercent = 70
                $State.StageStartTime = Get-Date
                $State.LastPollTime = Get-Date
            }
            
            "Waiting for Provisioning" {
                Write-StepLog "[Poll  ] $($State.UserPrincipalName): Polling Graph API - checking new CPC provisioning status..." "Polling" "Gray"
                $cloudPCs = Get-CloudPCForUser -UserId $State.UserPrincipalName
                $cloudPCsArray = @($cloudPCs)
                
                # Log what we're tracking - show DETAILED old CPC info
                Write-StepLog "[Debug ] Tracked old CPC IDs ($($State.CloudPCIds.Count)): $($State.CloudPCIds -join ', ')" "Debug" "Cyan"
                if ($State.OldCPCs -and $State.OldCPCs.Count -gt 0) {
                    Write-StepLog "[Debug ] Old CPC details from initial capture:" "Debug" "Cyan"
                    foreach ($oldCPC in $State.OldCPCs) {
                        Write-StepLog "[Debug ]   - Name: $($oldCPC.Name), ID: $($oldCPC.Id)" "Debug" "Cyan"
                    }
                }
                
                if ($cloudPCsArray.Count -gt 0) {
                    # Log all CPCs found for this user with FULL details
                    Write-StepLog "[Debug ] Found $($cloudPCsArray.Count) total CPC(s) from API:" "Debug" "Cyan"
                    foreach ($cpc in $cloudPCsArray) {
                        $cpcName = if ($cpc.managedDeviceName) { $cpc.managedDeviceName } else { $cpc.displayName }
                        $isOld = if ($State.CloudPCIds -contains $cpc.id) { "OLD" } else { "NEW" }
                        Write-StepLog "[Debug ]   [$isOld] Name: $cpcName | ID: $($cpc.id) | Status: $($cpc.status)" "Debug" "Cyan"
                        Write-StepLog "[Debug ]        managedDeviceName: '$($cpc.managedDeviceName)' | displayName: '$($cpc.displayName)'" "Debug" "Gray"
                    }
                    
                    $newCPCs = @()
                    foreach ($cpc in $cloudPCsArray) {
                        if ($State.CloudPCIds -notcontains $cpc.id) {
                            $newCPCs += $cpc
                        }
                    }
                    
                    Write-StepLog "[Debug ] Identified $($newCPCs.Count) NEW CPC(s) (not in old ID list)" "Debug" "Yellow"
                    
                    if ($newCPCs.Count -gt 0) {
                        $provisionedCPCs = @($newCPCs | Where-Object { $_.status -in @('provisioned', 'provisionedWithWarnings') })
                        
                        if ($provisionedCPCs.Count -gt 0) {
                            $hasWarnings = @($provisionedCPCs | Where-Object { $_.status -eq 'provisionedWithWarnings' })
                            $State.NewCPCs = @($provisionedCPCs | ForEach-Object {
                                @{
                                    Id = $_.id
                                    Name = if ($_.managedDeviceName) { $_.managedDeviceName } else { $_.displayName }
                                    ServicePlan = $_.servicePlanName
                                }
                            })
                            
                            $newCPCNames = $provisionedCPCs | ForEach-Object {
                                if ($_.managedDeviceName) { $_.managedDeviceName } else { $_.displayName }
                            }
                            $State.NewCPCName = $newCPCNames -join ", "
                            
                            Write-StepLog "[Detect] $($State.UserPrincipalName): New Cloud PC(s) provisioned$(if ($hasWarnings.Count) {' (with warnings)'})!" "Status" "Green"
                            foreach ($cpc in $provisionedCPCs) {
                                $cpcName = if ($cpc.managedDeviceName) { $cpc.managedDeviceName } else { $cpc.displayName }
                                Write-StepLog "[Detect]   NEW: $cpcName | ID: $($cpc.id) | Plan: $($cpc.servicePlanName) | Status: $($cpc.status)" "Status" "Green"
                            }
                            
                            $State.Stage          = "Complete"
                            $State.ProgressPercent = 100
                            $State.Status         = if ($hasWarnings.Count -gt 0) { "Success (Warnings)" } else { "Success" }
                            $State.EndTime        = Get-Date
                            $State.FinalMessage   = if ($hasWarnings.Count -gt 0) {
                                "Provisioned with warnings - review Cloud PC health in Intune"
                            } else {
                                "Successfully replaced Cloud PC(s)"
                            }
                        }
                        else {
                            $elapsed = (Get-Date) - $State.StageStartTime
                            Write-StepLog "[Poll  ] $($State.UserPrincipalName): New CPC(s) found but not provisioned yet (elapsed: $([math]::Round($elapsed.TotalMinutes, 1))min)..." "Polling" "Gray"
                            
                            foreach ($cpc in $newCPCs) {
                                $cpcName = if ($cpc.managedDeviceName) { $cpc.managedDeviceName } else { $cpc.displayName }
                                Write-StepLog "[Debug ]   - $cpcName (status: $($cpc.status))" "Debug" "Gray"
                                
                                if ($cpc.status -eq 'failed') {
                                    throw "New Cloud PC provisioning failed: $($cpc.statusDetails.message)"
                                }
                            }
                            
                            if ($elapsed.TotalMinutes -gt $Timeouts.ProvisioningTimeout) {
                                throw "Timeout waiting for new Cloud PC(s) to provision"
                            }
                        }
                    }
                    else {
                        $elapsed = (Get-Date) - $State.StageStartTime
                        Write-StepLog "[Poll  ] $($State.UserPrincipalName): No new Cloud PC found yet (elapsed: $([math]::Round($elapsed.TotalMinutes, 1))min)..." "Polling" "Gray"
                        
                        if ($elapsed.TotalMinutes -gt $Timeouts.ProvisioningTimeout) {
                            throw "Timeout waiting for new Cloud PC(s) to appear"
                        }
                    }
                }
                else {
                    $elapsed = (Get-Date) - $State.StageStartTime
                    Write-StepLog "[Poll  ] $($State.UserPrincipalName): No Cloud PCs found yet (elapsed: $([math]::Round($elapsed.TotalMinutes, 1))min)..." "Polling" "Gray"
                    
                    if ($elapsed.TotalMinutes -gt $Timeouts.ProvisioningTimeout) {
                        throw "Timeout waiting for new Cloud PC(s) to provision"
                    }
                }
            }
            
            default {
                Write-StepLog "[WARN  ] Unknown stage: $($State.Stage)" "Status" "Yellow"
            }
        }
    }
    catch {
        $State.Status = "Failed"
        $State.ErrorMessage = $_.Exception.Message
        $State.EndTime = Get-Date
        Write-StepLog "[FAIL  ] $($State.UserPrincipalName): $($_.Exception.Message)" "Status" "Red"
    }
}

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
    'Get-UserInfo',
    'Get-ProvisioningPoliciesForGroup',
    'Stop-CloudPCGracePeriod',
    'Remove-UserFromGroup',
    'Add-UserToGroup',
    'Find-EntraIDGroups',
    'Get-EnterprisePolicyGroups',
    'Get-GroupMembers',
    'Initialize-Logging',
    'Write-Log',
    'Test-GraphConnection',
    'Get-StageDisplay',
    'Invoke-CloudPCReplaceStep'
)




