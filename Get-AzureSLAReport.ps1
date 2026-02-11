<#
.SYNOPSIS
    Azure SLA & Service Health Report Generator

.DESCRIPTION
    Generates an Excel report with:
      - Tab 1 (SLA Overview): Resource availability aggregated by region,
        service category (Compute, SQL DB, Web Apps, Storage), and month for the past 12 months.
      - Tab 2 (Incidents & Alerts): Service Health incidents and alerts reported in your environment
        for the past month.

    Prerequisites:
      - Az PowerShell module (Az.Accounts, Az.ResourceGraph, Az.Monitor, Az.Resources)
      - ImportExcel module
      - An active Azure subscription with Reader access

    Subscription scope:
      - By default, queries ALL subscriptions accessible to the authenticated account.
      - Use -SubscriptionIds to limit to specific subscriptions.

.NOTES
    Author  : Guil Lima (Microsoft)
    Date    : 2026-02-11
    Version : 1.1
#>

[CmdletBinding()]
param(
    # Region scope: leave empty for ALL regions, or specify specific ones
    [string[]]$Regions = @(),
    [int]$MonthsBack = 12,
    [string]$OutputPath = (Join-Path $PSScriptRoot "AzureSLA_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"),

    # Subscription scope: pass one or more subscription IDs, or leave empty for ALL subscriptions
    [string[]]$SubscriptionIds = @()
)

#region ── 0. HELPER: COLOUR / STYLE CONSTANTS ──────────────────────────────────
$HeaderBg       = [System.Drawing.Color]::FromArgb(0, 120, 215)   # Azure blue
$HeaderFg       = [System.Drawing.Color]::White
$GreenBg        = [System.Drawing.Color]::FromArgb(198, 239, 206)
$YellowBg       = [System.Drawing.Color]::FromArgb(255, 235, 156)
$RedBg          = [System.Drawing.Color]::FromArgb(255, 199, 206)
#endregion

#region ── 1. TROUBLESHOOTING & PREREQUISITES ────────────────────────────────────
function Test-Prerequisites {
    Write-Host "`n╔══════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "║   Azure SLA & Service Health Report Generator    ║" -ForegroundColor Cyan
    Write-Host "╚══════════════════════════════════════════════════╝`n" -ForegroundColor Cyan

    # ── Check required modules ──────────────────────────────────────────────
    $requiredModules = @(
        @{ Name = 'Az.Accounts';       MinVersion = '2.0.0' },
        @{ Name = 'Az.ResourceGraph';   MinVersion = '0.11.0' },
        @{ Name = 'Az.Monitor';         MinVersion = '3.0.0' },
        @{ Name = 'Az.Resources';       MinVersion = '5.0.0' },
        @{ Name = 'ImportExcel';        MinVersion = '7.0.0' }
    )

    foreach ($mod in $requiredModules) {
        $installed = Get-Module -ListAvailable -Name $mod.Name | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $installed) {
            Write-Host "[MISSING] Module '$($mod.Name)' is not installed." -ForegroundColor Red
            Write-Host "          Run:  Install-Module -Name $($mod.Name) -Scope CurrentUser -Force" -ForegroundColor Yellow
            $missingModules = $true
        } else {
            Write-Host "[  OK  ] $($mod.Name) v$($installed.Version)" -ForegroundColor Green
        }
    }
    if ($missingModules) {
        Write-Host "`n[ACTION] Install missing modules before proceeding. Example:" -ForegroundColor Red
        Write-Host "         Install-Module Az -Scope CurrentUser -Force" -ForegroundColor Yellow
        Write-Host "         Install-Module ImportExcel -Scope CurrentUser -Force`n" -ForegroundColor Yellow
        throw "Missing required PowerShell modules. See messages above."
    }

    # ── Import modules ──────────────────────────────────────────────────────
    Import-Module Az.Accounts, Az.ResourceGraph, Az.Monitor, Az.Resources, ImportExcel -ErrorAction Stop

    # ── Check Azure connection ──────────────────────────────────────────────
    Write-Host "`n── Checking Azure connection ──" -ForegroundColor Cyan
    $ctx = Get-AzContext -ErrorAction SilentlyContinue
    if (-not $ctx -or -not $ctx.Account) {
        Write-Host "[WARN ] Not connected to Azure. Attempting interactive login..." -ForegroundColor Yellow
        try {
            Connect-AzAccount -ErrorAction Stop | Out-Null
            $ctx = Get-AzContext
        } catch {
            Write-Host "`n[ERROR] Failed to authenticate to Azure." -ForegroundColor Red
            Write-Host @"

  ╔═══ TROUBLESHOOTING STEPS ═══════════════════════════════════════════╗
  ║                                                                     ║
  ║  1. Run 'Connect-AzAccount' manually and sign in.                   ║
  ║  2. If MFA is required, use:                                        ║
  ║       Connect-AzAccount -TenantId <your-tenant-id>                  ║
  ║  3. If using a service principal:                                    ║
  ║       Connect-AzAccount -ServicePrincipal -ApplicationId <id> `     ║
  ║         -CertificateThumbprint <thumbprint> -TenantId <tenant>      ║
  ║  4. Verify your network can reach https://login.microsoftonline.com  ║
  ║  5. If behind a proxy, configure:                                   ║
  ║       [System.Net.WebRequest]::DefaultWebProxy.Credentials =        ║
  ║         [System.Net.CredentialCache]::DefaultCredentials             ║
  ║  6. Update the Az module: Update-Module Az -Force                   ║
  ║  7. Clear cached tokens: Clear-AzContext -Force                     ║
  ║     then retry Connect-AzAccount.                                   ║
  ║                                                                     ║
  ╚═════════════════════════════════════════════════════════════════════╝
"@ -ForegroundColor Yellow
            throw "Azure authentication failed. See troubleshooting steps above."
        }
    }

    Write-Host "[  OK  ] Connected as: $($ctx.Account.Id)" -ForegroundColor Green
    Write-Host "[  OK  ] Default subscription: $($ctx.Subscription.Name) ($($ctx.Subscription.Id))" -ForegroundColor Green

    # ── Resolve subscription scope ─────────────────────────────────────────
    Write-Host "`n── Resolving subscription scope ──" -ForegroundColor Cyan
    if ($SubscriptionIds -and $SubscriptionIds.Count -gt 0) {
        # User specified explicit subscription IDs
        $targetSubs = @()
        foreach ($sid in $SubscriptionIds) {
            try {
                $s = Get-AzSubscription -SubscriptionId $sid -ErrorAction Stop
                $targetSubs += $s
                Write-Host "[  OK  ] $($s.Name) ($($s.Id)) — $($s.State)" -ForegroundColor Green
            } catch {
                Write-Host "[WARN ] Subscription '$sid' not accessible — skipping" -ForegroundColor Yellow
            }
        }
        if ($targetSubs.Count -eq 0) {
            throw "None of the specified subscriptions are accessible."
        }
    } else {
        # Default: ALL subscriptions the account can access
        $targetSubs = Get-AzSubscription -ErrorAction Stop | Where-Object { $_.State -eq 'Enabled' }
        if ($targetSubs.Count -eq 0) {
            throw "No enabled subscriptions found for this account."
        }
        Write-Host "[  OK  ] Found $($targetSubs.Count) enabled subscription(s):" -ForegroundColor Green
        foreach ($s in $targetSubs) {
            Write-Host "         • $($s.Name) ($($s.Id))" -ForegroundColor Gray
        }
    }

    # Store resolved subscription IDs in script scope for other functions
    $script:ResolvedSubscriptionIds = $targetSubs | ForEach-Object { $_.Id }

    # ── Verify Resource Graph provider (on current context subscription) ──
    $rgProvider = Get-AzResourceProvider -ProviderNamespace 'Microsoft.ResourceHealth' -ErrorAction SilentlyContinue
    if (-not $rgProvider -or $rgProvider[0].RegistrationState -ne 'Registered') {
        Write-Host "[WARN ] Microsoft.ResourceHealth provider not registered. Attempting registration..." -ForegroundColor Yellow
        Register-AzResourceProvider -ProviderNamespace 'Microsoft.ResourceHealth' -ErrorAction SilentlyContinue | Out-Null
        Write-Host "[INFO ] Registration initiated. It may take a few minutes to propagate." -ForegroundColor Yellow
    } else {
        Write-Host "[  OK  ] Microsoft.ResourceHealth provider registered" -ForegroundColor Green
    }

    Write-Host ""
    return $ctx
}
#endregion

#region ── 2. REGION RESOLUTION ──────────────────────────────────────────────────
function Resolve-Regions {
    <#
    .SYNOPSIS
        Resolves the target regions. If none specified, discovers all regions that
        contain resources in the target subscriptions. Builds a display-name lookup.
    #>
    [CmdletBinding()]
    param(
        [string[]]$RequestedRegions
    )

    Write-Host "`n── Resolving target regions ──" -ForegroundColor Cyan

    # Build a full Azure location lookup (internal name → display name)
    $allLocations = Get-AzLocation -ErrorAction SilentlyContinue
    $script:RegionDisplayNames = @{}
    foreach ($loc in $allLocations) {
        $script:RegionDisplayNames[$loc.Location] = $loc.DisplayName
    }

    if ($RequestedRegions -and $RequestedRegions.Count -gt 0) {
        # User specified explicit regions
        $resolved = $RequestedRegions | ForEach-Object { $_.ToLower() }
        Write-Host "[  OK  ] Using $($resolved.Count) specified region(s):" -ForegroundColor Green
        foreach ($r in $resolved) {
            $display = if ($script:RegionDisplayNames[$r]) { $script:RegionDisplayNames[$r] } else { $r }
            Write-Host "         • $display ($r)" -ForegroundColor Gray
        }
        return $resolved
    }

    # Default: discover all regions that have relevant resources
    Write-Host "[INFO ] No regions specified — discovering regions with resources..." -ForegroundColor Yellow
    $query = @"
Resources
| where type in~ (
    'microsoft.compute/virtualmachines',
    'microsoft.compute/virtualmachinescalesets',
    'microsoft.sql/servers/databases',
    'microsoft.sql/servers',
    'microsoft.sql/managedinstances',
    'microsoft.web/sites',
    'microsoft.web/serverfarms',
    'microsoft.storage/storageaccounts'
  )
| distinct location
| order by location asc
"@

    try {
        $regionResults = Search-AzGraph -Query $query -First 1000 -Subscription $script:ResolvedSubscriptionIds -ErrorAction Stop
        $resolved = $regionResults | ForEach-Object { $_.location.ToLower() }

        if ($resolved.Count -eq 0) {
            Write-Host "[WARN ] No resources found in any region. Falling back to all Azure regions." -ForegroundColor Yellow
            $resolved = $allLocations | Where-Object { $_.RegionType -eq 'Physical' } | ForEach-Object { $_.Location }
        }

        Write-Host "[  OK  ] Found resources in $($resolved.Count) region(s):" -ForegroundColor Green
        foreach ($r in $resolved) {
            $display = if ($script:RegionDisplayNames[$r]) { $script:RegionDisplayNames[$r] } else { $r }
            Write-Host "         • $display" -ForegroundColor Gray
        }
        return $resolved
    } catch {
        Write-Host "[WARN ] Region discovery failed: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "[INFO ] Falling back to all physical Azure regions." -ForegroundColor Yellow
        $resolved = $allLocations | Where-Object { $_.RegionType -eq 'Physical' } | ForEach-Object { $_.Location }
        return $resolved
    }
}

# Initialize as empty — will be populated by Resolve-Regions
$script:RegionDisplayNames = @{}
#endregion

#region ── 3. DATA COLLECTION FUNCTIONS ──────────────────────────────────────────

function Get-ResourceHealthEvents {
    <#
    .SYNOPSIS
        Retrieves Resource Health availability events using Azure Resource Graph
        for the specified regions and date range.
    #>
    [CmdletBinding()]
    param(
        [string[]]$TargetRegions,
        [datetime]$StartDate,
        [datetime]$EndDate
    )

    Write-Host "── Querying Resource Health events via Resource Graph ──" -ForegroundColor Cyan

    # Query resource health availability status changes
    $query = @"
ServiceHealthResources
| where type == "microsoft.resourcehealth/events"
| where properties.EventType == "ServiceIssue"
| extend impactStartTime = todatetime(properties.ImpactStartTime)
| extend impactEndTime   = todatetime(properties.ImpactMitigationTime)
| extend status          = tostring(properties.Status)
| extend title           = tostring(properties.Title)
| extend summary         = tostring(properties.Summary)
| extend eventLevel      = tostring(properties.EventLevel)
| extend impactedServices = properties.Impact
| where impactStartTime >= datetime('$($StartDate.ToString("yyyy-MM-dd"))') and impactStartTime <= datetime('$($EndDate.ToString("yyyy-MM-dd"))')
| project id, name, impactStartTime, impactEndTime, status, title, summary, eventLevel, impactedServices
| order by impactStartTime desc
"@

    try {
        $results = Search-AzGraph -Query $query -First 1000 -Subscription $script:ResolvedSubscriptionIds -ErrorAction Stop
        Write-Host "[  OK  ] Retrieved $($results.Count) service health events" -ForegroundColor Green
        return $results
    } catch {
        Write-Host "[WARN ] Resource Graph query failed: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "[INFO ] Falling back to Activity Log method..." -ForegroundColor Yellow
        return @()
    }
}

function Get-ServiceHealthAlerts {
    <#
    .SYNOPSIS
        Retrieves Service Health alerts from Activity Log for the past month.
    #>
    [CmdletBinding()]
    param(
        [datetime]$StartDate,
        [datetime]$EndDate
    )

    Write-Host "── Querying Service Health alerts from Activity Log ──" -ForegroundColor Cyan

    $alerts = @()
    try {
        # Save current context to restore later
        $originalContext = Get-AzContext

        foreach ($subId in $script:ResolvedSubscriptionIds) {
            # Switch context to each subscription for Activity Log queries
            Set-AzContext -SubscriptionId $subId -ErrorAction SilentlyContinue | Out-Null
            $subName = (Get-AzContext).Subscription.Name
            Write-Host "  Scanning subscription: $subName" -ForegroundColor Gray

            # Get Resource Health events from Activity Log
            $logs = Get-AzActivityLog -StartTime $StartDate -EndTime $EndDate `
                -ResourceProvider "Microsoft.ResourceHealth" -MaxRecord 1000 -ErrorAction SilentlyContinue

            if ($logs) {
                foreach ($log in $logs) {
                    $alerts += [PSCustomObject]@{
                        Timestamp       = $log.EventTimestamp
                        Category        = $log.Category.Value
                        Level           = $log.Level
                        OperationName   = $log.OperationName.Value
                        Status          = $log.Status.Value
                        Description     = $log.Description
                        ResourceId      = $log.ResourceId
                        CorrelationId   = $log.CorrelationId
                        Subscription    = $subName
                    }
                }
            }

            # Also get ServiceHealth category events
            $shLogs = Get-AzActivityLog -StartTime $StartDate -EndTime $EndDate `
                -MaxRecord 1000 -ErrorAction SilentlyContinue |
                Where-Object { $_.Category.Value -eq 'ServiceHealth' }

            if ($shLogs) {
                foreach ($log in $shLogs) {
                    $alerts += [PSCustomObject]@{
                        Timestamp       = $log.EventTimestamp
                        Category        = 'ServiceHealth'
                        Level           = $log.Level
                        OperationName   = $log.OperationName.Value
                        Status          = $log.Status.Value
                        Description     = $log.Description
                        ResourceId      = $log.ResourceId
                        CorrelationId   = $log.CorrelationId
                        Subscription    = $subName
                    }
                }
            }
        }

        # Restore original context
        Set-AzContext -SubscriptionId $originalContext.Subscription.Id -ErrorAction SilentlyContinue | Out-Null

        Write-Host "[  OK  ] Retrieved $($alerts.Count) health alerts from Activity Log across $($script:ResolvedSubscriptionIds.Count) subscription(s)" -ForegroundColor Green
    } catch {
        Write-Host "[WARN ] Activity Log query failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    return $alerts
}

function Get-ResourceInventory {
    <#
    .SYNOPSIS
        Queries the current resource inventory in the target regions using Resource Graph,
        grouped by service category.
    #>
    [CmdletBinding()]
    param(
        [string[]]$TargetRegions
    )

    Write-Host "── Querying resource inventory in target regions ──" -ForegroundColor Cyan

    $regionFilter = ($TargetRegions | ForEach-Object { "'$_'" }) -join ', '

    $query = @"
Resources
| where location in~ ($regionFilter)
| where type in~ (
    'microsoft.compute/virtualmachines',
    'microsoft.compute/virtualmachinescalesets',
    'microsoft.sql/servers/databases',
    'microsoft.sql/servers',
    'microsoft.sql/managedinstances',
    'microsoft.web/sites',
    'microsoft.web/serverfarms',
    'microsoft.storage/storageaccounts'
  )
| extend ServiceCategory = case(
    type =~ 'microsoft.compute/virtualmachines'          , 'Compute',
    type =~ 'microsoft.compute/virtualmachinescalesets'  , 'Compute',
    type =~ 'microsoft.sql/servers/databases'            , 'SQL DB',
    type =~ 'microsoft.sql/servers'                      , 'SQL DB',
    type =~ 'microsoft.sql/managedinstances'             , 'SQL DB',
    type =~ 'microsoft.web/sites'                        , 'Web Apps',
    type =~ 'microsoft.web/serverfarms'                  , 'Web Apps',
    type =~ 'microsoft.storage/storageaccounts'          , 'Storage',
    'Other'
  )
| project name, type, location, resourceGroup, ServiceCategory, subscriptionId, id
| order by ServiceCategory asc, location asc, name asc
"@

    try {
        $resources = Search-AzGraph -Query $query -First 1000 -Subscription $script:ResolvedSubscriptionIds -ErrorAction Stop
        Write-Host "[  OK  ] Found $($resources.Count) resources across target regions" -ForegroundColor Green
        foreach ($region in $TargetRegions) {
            $displayName = if ($RegionDisplayNames[$region]) { $RegionDisplayNames[$region] } else { $region }
            $count = ($resources | Where-Object { $_.location -eq $region }).Count
            Write-Host "         $displayName : $count resources" -ForegroundColor Gray
        }
        return $resources
    } catch {
        Write-Host "[WARN ] Resource Graph query failed: $($_.Exception.Message)" -ForegroundColor Yellow
        return @()
    }
}

function Get-ResourceAvailability {
    <#
    .SYNOPSIS
        Queries resource health availability for each resource to calculate SLA metrics.
        Uses Resource Health availability status changes over the past 12 months.
    #>
    [CmdletBinding()]
    param(
        [string[]]$TargetRegions,
        [datetime]$StartDate,
        [datetime]$EndDate
    )

    Write-Host "── Querying resource availability data ──" -ForegroundColor Cyan

    $regionFilter = ($TargetRegions | ForEach-Object { "'$_'" }) -join ', '

    # Query for resource health availability changes
    $query = @"
HealthResources
| where type == "microsoft.resourcehealth/availabilitystatuses"
| extend resourceId = tolower(tostring(properties.targetResourceId))
| extend availabilityState = tostring(properties.availabilityState)
| extend occurredTime = todatetime(properties.occurredTime)
| extend reasonType = tostring(properties.reasonType)
| extend resourceType = tostring(properties.targetResourceType)
| extend location = tostring(properties.location)
| where location in~ ($regionFilter)
| where resourceType in~ (
    'microsoft.compute/virtualmachines',
    'microsoft.compute/virtualmachinescalesets',
    'microsoft.sql/servers/databases',
    'microsoft.sql/servers',
    'microsoft.sql/managedinstances',
    'microsoft.web/sites',
    'microsoft.web/serverfarms',
    'microsoft.storage/storageaccounts'
  )
| extend ServiceCategory = case(
    resourceType =~ 'microsoft.compute/virtualmachines'        , 'Compute',
    resourceType =~ 'microsoft.compute/virtualmachinescalesets', 'Compute',
    resourceType =~ 'microsoft.sql/servers/databases'          , 'SQL DB',
    resourceType =~ 'microsoft.sql/servers'                    , 'SQL DB',
    resourceType =~ 'microsoft.sql/managedinstances'           , 'SQL DB',
    resourceType =~ 'microsoft.web/sites'                      , 'Web Apps',
    resourceType =~ 'microsoft.web/serverfarms'                , 'Web Apps',
    resourceType =~ 'microsoft.storage/storageaccounts'        , 'Storage',
    'Other'
  )
| project resourceId, availabilityState, occurredTime, reasonType, resourceType, location, ServiceCategory
| order by location asc, ServiceCategory asc, occurredTime desc
"@

    try {
        $healthData = Search-AzGraph -Query $query -First 5000 -Subscription $script:ResolvedSubscriptionIds -ErrorAction Stop
        Write-Host "[  OK  ] Retrieved $($healthData.Count) availability records" -ForegroundColor Green
        return $healthData
    } catch {
        Write-Host "[WARN ] Health resources query failed: $($_.Exception.Message)" -ForegroundColor Yellow
        return @()
    }
}

function Get-ServiceHealthIncidents {
    <#
    .SYNOPSIS
        Queries detailed service health incidents impacting the target regions
        with service-level breakdown.
    #>
    [CmdletBinding()]
    param(
        [string[]]$TargetRegions,
        [datetime]$StartDate,
        [datetime]$EndDate
    )

    Write-Host "── Querying Service Health incidents (detailed) ──" -ForegroundColor Cyan

    $query = @"
ServiceHealthResources
| where type =~ "microsoft.resourcehealth/events"
| extend eventType        = tostring(properties.EventType)
| extend status           = tostring(properties.Status)
| extend title            = tostring(properties.Title)
| extend summary          = tostring(properties.Summary)
| extend impactStartTime  = todatetime(properties.ImpactStartTime)
| extend impactEndTime    = todatetime(properties.ImpactMitigationTime)
| extend lastUpdateTime   = todatetime(properties.LastUpdateTime)
| extend level            = tostring(properties.EventLevel)
| extend impactedServices = properties.Impact
| where impactStartTime >= datetime('$($StartDate.ToString("yyyy-MM-dd"))') 
    and impactStartTime <= datetime('$($EndDate.ToString("yyyy-MM-dd"))')
| project name, eventType, status, title, summary, impactStartTime, impactEndTime,
          lastUpdateTime, level, impactedServices
| order by impactStartTime desc
"@

    try {
        $incidents = Search-AzGraph -Query $query -First 1000 -Subscription $script:ResolvedSubscriptionIds -ErrorAction Stop
        Write-Host "[  OK  ] Retrieved $($incidents.Count) service health incidents" -ForegroundColor Green
        return $incidents
    } catch {
        Write-Host "[WARN ] Service health incidents query failed: $($_.Exception.Message)" -ForegroundColor Yellow
        return @()
    }
}
#endregion

#region ── 4. DATA PROCESSING ────────────────────────────────────────────────────

function Build-SLAMatrix {
    <#
    .SYNOPSIS
        Builds the month-by-month SLA matrix for each region and service category.
        Calculates availability % based on resource health data and incidents.
    #>
    [CmdletBinding()]
    param(
        [array]$HealthData,
        [array]$Incidents,
        [array]$Resources,
        [string[]]$TargetRegions,
        [datetime]$StartDate,
        [datetime]$EndDate
    )

    Write-Host "`n── Building SLA matrix ──" -ForegroundColor Cyan

    $serviceCategories = @('Compute', 'SQL DB', 'Web Apps', 'Storage')
    $slaRows = @()

    foreach ($region in $TargetRegions) {
        $regionDisplay = if ($RegionDisplayNames[$region]) { $RegionDisplayNames[$region] } else { $region }

        foreach ($category in $serviceCategories) {
            $row = [ordered]@{
                'Region'   = $regionDisplay
                'Service'  = $category
            }

            # Count resources in this region/category
            $resourceCount = ($Resources | Where-Object {
                $_.location -eq $region -and $_.ServiceCategory -eq $category
            }).Count

            $row['Resource Count'] = $resourceCount

            # Build month columns
            for ($i = $MonthsBack - 1; $i -ge 0; $i--) {
                $monthStart = (Get-Date).AddMonths(-$i).Date
                $monthStart = Get-Date -Year $monthStart.Year -Month $monthStart.Month -Day 1
                $monthEnd   = $monthStart.AddMonths(1).AddSeconds(-1)
                $monthLabel = $monthStart.ToString("MMM yyyy")

                # Calculate availability for this month
                $availability = Calculate-MonthlyAvailability `
                    -HealthData $HealthData `
                    -Incidents $Incidents `
                    -Region $region `
                    -ServiceCategory $category `
                    -MonthStart $monthStart `
                    -MonthEnd $monthEnd `
                    -ResourceCount $resourceCount

                $row[$monthLabel] = $availability
            }

            $slaRows += [PSCustomObject]$row
        }
    }

    Write-Host "[  OK  ] SLA matrix built: $($slaRows.Count) rows" -ForegroundColor Green
    return $slaRows
}

function Calculate-MonthlyAvailability {
    <#
    .SYNOPSIS
        Calculates the availability percentage for a given region, service category,
        and month based on health data and incident records.
    #>
    [CmdletBinding()]
    param(
        [array]$HealthData,
        [array]$Incidents,
        [string]$Region,
        [string]$ServiceCategory,
        [datetime]$MonthStart,
        [datetime]$MonthEnd,
        [int]$ResourceCount
    )

    # If no resources exist for this category, return N/A
    if ($ResourceCount -eq 0) {
        return "N/A"
    }

    $totalMinutesInMonth = ($MonthEnd - $MonthStart).TotalMinutes
    $downtimeMinutes = 0

    # ── Check for unavailable health records in this period ──
    $unhealthyRecords = $HealthData | Where-Object {
        $_.location -eq $Region -and
        $_.ServiceCategory -eq $ServiceCategory -and
        $_.availabilityState -ne 'Available' -and
        $_.occurredTime -ge $MonthStart -and
        $_.occurredTime -le $MonthEnd
    }

    if ($unhealthyRecords -and $unhealthyRecords.Count -gt 0) {
        # Estimate downtime based on number of unhealthy events (each event ~= some downtime window)
        $downtimeMinutes += ($unhealthyRecords.Count * 30)  # conservative 30-min estimate per event
    }

    # ── Check for incidents impacting this region/service ──
    $serviceTypeMap = @{
        'Compute'  = @('Virtual Machines', 'Compute', 'Virtual Machine Scale Sets')
        'SQL DB'   = @('SQL Database', 'SQL Managed Instance', 'Azure SQL', 'SQL')
        'Web Apps' = @('App Service', 'Web Apps', 'App Service (Web Apps)')
        'Storage'  = @('Storage', 'Storage Accounts')
    }

    $relevantServiceNames = $serviceTypeMap[$ServiceCategory]

    foreach ($incident in $Incidents) {
        if ($null -eq $incident.impactedServices) { continue }

        $impactedServicesArray = if ($incident.impactedServices -is [array]) {
            $incident.impactedServices
        } else {
            @($incident.impactedServices)
        }

        foreach ($impact in $impactedServicesArray) {
            $serviceName = if ($impact.ImpactedService) { $impact.ImpactedService } else { $impact.ServiceName }
            $impactedRegions = if ($impact.ImpactedRegions) { $impact.ImpactedRegions } else { @() }

            $regionMatch = $false
            foreach ($ir in $impactedRegions) {
                $irName = if ($ir.ImpactedRegion) { $ir.ImpactedRegion } else { $ir }
                if ($irName -like "*$(if ($RegionDisplayNames[$Region]) { $RegionDisplayNames[$Region] } else { $Region })*" -or $irName -eq $Region) {
                    $regionMatch = $true
                    break
                }
            }

            if (-not $regionMatch) { continue }

            $serviceMatch = $false
            foreach ($svcName in $relevantServiceNames) {
                if ($serviceName -like "*$svcName*") {
                    $serviceMatch = $true
                    break
                }
            }

            if (-not $serviceMatch) { continue }

            # Calculate actual downtime from incident window
            $incStart = [datetime]$incident.impactStartTime
            $incEnd   = if ($incident.impactEndTime) { [datetime]$incident.impactEndTime } else { $MonthEnd }

            # Clamp to month boundaries
            $effectiveStart = [datetime]([Math]::Max($incStart.Ticks, $MonthStart.Ticks))
            $effectiveEnd   = [datetime]([Math]::Min($incEnd.Ticks, $MonthEnd.Ticks))

            if ($effectiveEnd -gt $effectiveStart) {
                $downtimeMinutes += ($effectiveEnd - $effectiveStart).TotalMinutes
            }
        }
    }

    # Cap downtime to total minutes in month
    $downtimeMinutes = [Math]::Min($downtimeMinutes, $totalMinutesInMonth)

    # Calculate availability percentage
    $availability = (($totalMinutesInMonth - $downtimeMinutes) / $totalMinutesInMonth) * 100
    return [Math]::Round($availability, 4)
}

function Build-IncidentsTable {
    <#
    .SYNOPSIS
        Builds a flat table of incidents and alerts for Tab 2 of the report.
    #>
    [CmdletBinding()]
    param(
        [array]$Incidents,
        [array]$Alerts,
        [string[]]$TargetRegions
    )

    Write-Host "`n── Building incidents & alerts table ──" -ForegroundColor Cyan

    $rows = @()

    # ── Process Service Health incidents ──
    foreach ($inc in $Incidents) {
        $regionsAffected = @()
        $servicesAffected = @()

        if ($null -ne $inc.impactedServices) {
            $impactArray = if ($inc.impactedServices -is [array]) { $inc.impactedServices } else { @($inc.impactedServices) }
            foreach ($impact in $impactArray) {
                $svcName = if ($impact.ImpactedService) { $impact.ImpactedService } else { $impact.ServiceName }
                if ($svcName) { $servicesAffected += $svcName }

                $impRegions = if ($impact.ImpactedRegions) { $impact.ImpactedRegions } else { @() }
                foreach ($ir in $impRegions) {
                    $rName = if ($ir.ImpactedRegion) { $ir.ImpactedRegion } else { $ir }
                    if ($rName) { $regionsAffected += $rName }
                }
            }
        }

        # Filter: only include if it impacts our target regions (or if no region info available)
        $regionRelevant = $false
        if ($regionsAffected.Count -eq 0) {
            $regionRelevant = $true  # No region info, include for safety
        } else {
            foreach ($region in $TargetRegions) {
                $displayName = if ($RegionDisplayNames[$region]) { $RegionDisplayNames[$region] } else { $region }
                foreach ($ra in $regionsAffected) {
                    if ($ra -like "*$displayName*" -or $ra -eq $region) {
                        $regionRelevant = $true
                        break
                    }
                }
                if ($regionRelevant) { break }
            }
        }

        if (-not $regionRelevant) { continue }

        $durationHours = if ($inc.impactStartTime -and $inc.impactEndTime) {
            [Math]::Round(([datetime]$inc.impactEndTime - [datetime]$inc.impactStartTime).TotalHours, 2)
        } else { "Ongoing" }

        $rows += [PSCustomObject][ordered]@{
            'Source'             = 'Service Health'
            'Type'               = $inc.eventType
            'Status'             = $inc.status
            'Title'              = $inc.title
            'Impact Start (UTC)' = if ($inc.impactStartTime) { ([datetime]$inc.impactStartTime).ToString("yyyy-MM-dd HH:mm") } else { "" }
            'Impact End (UTC)'   = if ($inc.impactEndTime) { ([datetime]$inc.impactEndTime).ToString("yyyy-MM-dd HH:mm") } else { "Ongoing" }
            'Duration (Hours)'   = $durationHours
            'Level'              = $inc.level
            'Affected Services'  = ($servicesAffected | Select-Object -Unique) -join '; '
            'Affected Regions'   = ($regionsAffected | Select-Object -Unique) -join '; '
            'Summary'            = if ($inc.summary) { ($inc.summary -replace '<[^>]+>', '' ).Substring(0, [Math]::Min(500, ($inc.summary -replace '<[^>]+>', '').Length)) } else { "" }
            'Tracking ID'        = $inc.name
        }
    }

    # ── Process Activity Log alerts ──
    foreach ($alert in $Alerts) {
        $rows += [PSCustomObject][ordered]@{
            'Source'             = 'Activity Log'
            'Type'               = $alert.Category
            'Status'             = $alert.Status
            'Title'              = $alert.OperationName
            'Impact Start (UTC)' = if ($alert.Timestamp) { $alert.Timestamp.ToString("yyyy-MM-dd HH:mm") } else { "" }
            'Impact End (UTC)'   = ""
            'Duration (Hours)'   = ""
            'Level'              = $alert.Level
            'Affected Services'  = ""
            'Affected Regions'   = ""
            'Subscription'       = if ($alert.Subscription) { $alert.Subscription } else { "" }
            'Summary'            = if ($alert.Description) { $alert.Description.Substring(0, [Math]::Min(500, $alert.Description.Length)) } else { "" }
            'Tracking ID'        = $alert.CorrelationId
        }
    }

    Write-Host "[  OK  ] Incidents table: $($rows.Count) entries for target regions" -ForegroundColor Green
    return $rows
}

function Build-ServiceHealthTimeline {
    <#
    .SYNOPSIS
        Builds a month-by-month timeline of all service health events
        over the full reporting period (12 months by default).
    #>
    [CmdletBinding()]
    param(
        [array]$Incidents,
        [string[]]$TargetRegions,
        [datetime]$StartDate,
        [datetime]$EndDate
    )

    Write-Host "`n── Building service health timeline ──" -ForegroundColor Cyan

    $rows = @()

    foreach ($inc in $Incidents) {
        if ($null -eq $inc.impactStartTime) { continue }

        $incStart = [datetime]$inc.impactStartTime

        # Extract affected regions and services from the incident
        $regionsAffected  = @()
        $servicesAffected = @()

        if ($null -ne $inc.impactedServices) {
            $impactArray = if ($inc.impactedServices -is [array]) { $inc.impactedServices } else { @($inc.impactedServices) }
            foreach ($impact in $impactArray) {
                $svcName = if ($impact.ImpactedService) { $impact.ImpactedService } else { $impact.ServiceName }
                if ($svcName) { $servicesAffected += $svcName }

                $impRegions = if ($impact.ImpactedRegions) { $impact.ImpactedRegions } else { @() }
                foreach ($ir in $impRegions) {
                    $rName = if ($ir.ImpactedRegion) { $ir.ImpactedRegion } else { $ir }
                    if ($rName) { $regionsAffected += $rName }
                }
            }
        }

        # Filter to target regions
        $regionRelevant = $false
        if ($regionsAffected.Count -eq 0) {
            $regionRelevant = $true
        } else {
            foreach ($region in $TargetRegions) {
                $displayName = if ($RegionDisplayNames[$region]) { $RegionDisplayNames[$region] } else { $region }
                foreach ($ra in $regionsAffected) {
                    if ($ra -like "*$displayName*" -or $ra -eq $region) {
                        $regionRelevant = $true
                        break
                    }
                }
                if ($regionRelevant) { break }
            }
        }
        if (-not $regionRelevant) { continue }

        # Calculate duration
        $incEnd = if ($inc.impactEndTime) { [datetime]$inc.impactEndTime } else { $null }
        $durationHours = if ($incStart -and $incEnd) {
            [Math]::Round(($incEnd - $incStart).TotalHours, 2)
        } else { "Ongoing" }

        $rows += [PSCustomObject][ordered]@{
            'Month'              = $incStart.ToString("yyyy-MM")
            'Month Name'         = $incStart.ToString("MMM yyyy")
            'Event Type'         = $inc.eventType
            'Status'             = $inc.status
            'Title'              = $inc.title
            'Impact Start (UTC)' = $incStart.ToString("yyyy-MM-dd HH:mm")
            'Impact End (UTC)'   = if ($incEnd) { $incEnd.ToString("yyyy-MM-dd HH:mm") } else { "Ongoing" }
            'Duration (Hours)'   = $durationHours
            'Level'              = $inc.level
            'Affected Services'  = ($servicesAffected | Select-Object -Unique) -join '; '
            'Affected Regions'   = ($regionsAffected | Select-Object -Unique) -join '; '
            'Summary'            = if ($inc.summary) { ($inc.summary -replace '<[^>]+>', '').Substring(0, [Math]::Min(500, ($inc.summary -replace '<[^>]+>', '').Length)) } else { "" }
            'Tracking ID'        = $inc.name
        }
    }

    # Sort by month descending, then by start time descending
    $rows = $rows | Sort-Object { $_.'Month' }, { $_.'Impact Start (UTC)' } -Descending

    Write-Host "[  OK  ] Service health timeline: $($rows.Count) events across reporting period" -ForegroundColor Green
    return $rows
}
#endregion

#region ── 5. EXCEL EXPORT ───────────────────────────────────────────────────────

function Export-SLAReport {
    <#
    .SYNOPSIS
        Exports the SLA matrix, incidents table, and service health timeline
        to a formatted Excel workbook.
    #>
    [CmdletBinding()]
    param(
        [array]$SLAMatrix,
        [array]$IncidentsTable,
        [array]$HealthTimeline,
        [string]$OutputFile
    )

    Write-Host "`n── Exporting Excel report ──" -ForegroundColor Cyan

    # Remove existing file if present
    if (Test-Path $OutputFile) { Remove-Item $OutputFile -Force }

    # ══════════════════════════════════════════════════════════════════════
    # TAB 1: SLA Overview
    # ══════════════════════════════════════════════════════════════════════
    # Build a dynamic title showing the regions covered
    $regionDisplayList = ($Regions | ForEach-Object {
        if ($script:RegionDisplayNames[$_]) { $script:RegionDisplayNames[$_] } else { $_ }
    })
    if ($regionDisplayList.Count -le 5) {
        $titleRegions = $regionDisplayList -join ', '
    } else {
        $titleRegions = "$($regionDisplayList.Count) regions"
    }

    $tab1Name = "SLA Overview"

    $excelPkg = $SLAMatrix | Export-Excel -Path $OutputFile -WorksheetName $tab1Name `
        -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
        -Title "Azure SLA Report — $titleRegions" `
        -TitleBold -TitleSize 14 `
        -PassThru

    $ws1 = $excelPkg.Workbook.Worksheets[$tab1Name]

    # Style header row (row 2, since row 1 is the title)
    $headerRow = 2
    $lastCol = $ws1.Dimension.End.Column
    $lastRow = $ws1.Dimension.End.Row

    for ($col = 1; $col -le $lastCol; $col++) {
        $cell = $ws1.Cells[$headerRow, $col]
        $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $cell.Style.Fill.BackgroundColor.SetColor($HeaderBg)
        $cell.Style.Font.Color.SetColor($HeaderFg)
        $cell.Style.Font.Bold = $true
        $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    }

    # Conditional formatting for SLA percentages (columns 4 onwards = month columns)
    $monthColStart = 4  # Column D (after Region, Service, Resource Count)
    for ($col = $monthColStart; $col -le $lastCol; $col++) {
        for ($row = $headerRow + 1; $row -le $lastRow; $row++) {
            $cell = $ws1.Cells[$row, $col]
            $val = $cell.Value

            if ($val -is [double] -or $val -is [decimal] -or $val -is [float] -or $val -is [int]) {
                $cell.Style.Numberformat.Format = "0.00\%"
                $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center

                if ($val -ge 99.99) {
                    $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $cell.Style.Fill.BackgroundColor.SetColor($GreenBg)
                } elseif ($val -ge 99.9) {
                    $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $cell.Style.Fill.BackgroundColor.SetColor($YellowBg)
                } elseif ($val -ne 0) {
                    $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $cell.Style.Fill.BackgroundColor.SetColor($RedBg)
                }
            } elseif ($val -eq "N/A") {
                $cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
                $cell.Style.Font.Italic = $true
                $cell.Style.Font.Color.SetColor([System.Drawing.Color]::Gray)
            }
        }
    }

    # Add alternating row shading for region grouping
    $currentRegion = ""
    $shadeToggle = $false
    $shadeBg = [System.Drawing.Color]::FromArgb(242, 242, 242)
    for ($row = $headerRow + 1; $row -le $lastRow; $row++) {
        $regionVal = $ws1.Cells[$row, 1].Value
        if ($regionVal -ne $currentRegion) {
            $currentRegion = $regionVal
            $shadeToggle = -not $shadeToggle
        }
        if ($shadeToggle) {
            for ($col = 1; $col -le 3; $col++) {
                $c = $ws1.Cells[$row, $col]
                if ($c.Style.Fill.PatternType -ne [OfficeOpenXml.Style.ExcelFillStyle]::Solid) {
                    $c.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $c.Style.Fill.BackgroundColor.SetColor($shadeBg)
                }
            }
        }
    }

    # Set column widths
    $ws1.Column(1).Width = 18   # Region
    $ws1.Column(2).Width = 14   # Service
    $ws1.Column(3).Width = 16   # Resource Count
    for ($col = $monthColStart; $col -le $lastCol; $col++) {
        $ws1.Column($col).Width = 14
    }

    # ══════════════════════════════════════════════════════════════════════
    # TAB 2: Incidents & Alerts
    # ══════════════════════════════════════════════════════════════════════
    $tab2Name = "Incidents & Alerts"

    if ($IncidentsTable.Count -gt 0) {
        $excelPkg = $IncidentsTable | Export-Excel -ExcelPackage $excelPkg -WorksheetName $tab2Name `
            -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
            -Title "Service Health Incidents & Alerts — Past Month" `
            -TitleBold -TitleSize 14 `
            -PassThru

        $ws2 = $excelPkg.Workbook.Worksheets[$tab2Name]
        $headerRow2 = 2
        $lastCol2 = $ws2.Dimension.End.Column

        for ($col = 1; $col -le $lastCol2; $col++) {
            $cell = $ws2.Cells[$headerRow2, $col]
            $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $cell.Style.Fill.BackgroundColor.SetColor($HeaderBg)
            $cell.Style.Font.Color.SetColor($HeaderFg)
            $cell.Style.Font.Bold = $true
        }

        # Set reasonable column widths for tab 2
        $ws2.Column(1).Width  = 16  # Source
        $ws2.Column(2).Width  = 16  # Type
        $ws2.Column(3).Width  = 14  # Status
        $ws2.Column(4).Width  = 50  # Title
        $ws2.Column(5).Width  = 18  # Impact Start
        $ws2.Column(6).Width  = 18  # Impact End
        $ws2.Column(7).Width  = 16  # Duration
        $ws2.Column(8).Width  = 12  # Level
        $ws2.Column(9).Width  = 30  # Affected Services
        $ws2.Column(10).Width = 30  # Affected Regions
        $ws2.Column(11).Width = 60  # Summary
        $ws2.Column(12).Width = 36  # Tracking ID

        # Wrap text for Summary column
        $lastRow2 = $ws2.Dimension.End.Row
        for ($row = $headerRow2 + 1; $row -le $lastRow2; $row++) {
            $ws2.Cells[$row, 11].Style.WrapText = $true
            $ws2.Row($row).Height = 45
        }
    } else {
        # Create empty tab with a message
        $emptyData = @([PSCustomObject]@{ Message = "No incidents or alerts found for the target regions in the past month." })
        $excelPkg = $emptyData | Export-Excel -ExcelPackage $excelPkg -WorksheetName $tab2Name `
            -AutoSize -PassThru
    }

    # ══════════════════════════════════════════════════════════════════════
    # TAB 3: Service Health Timeline (month by month)
    # ══════════════════════════════════════════════════════════════════════
    $tab3Name = "Health Timeline"

    if ($HealthTimeline -and $HealthTimeline.Count -gt 0) {
        $excelPkg = $HealthTimeline | Export-Excel -ExcelPackage $excelPkg -WorksheetName $tab3Name `
            -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
            -Title "Service Health Events — Month by Month ($MonthsBack months)" `
            -TitleBold -TitleSize 14 `
            -PassThru

        $ws3 = $excelPkg.Workbook.Worksheets[$tab3Name]
        $headerRow3 = 2
        $lastCol3 = $ws3.Dimension.End.Column
        $lastRow3 = $ws3.Dimension.End.Row

        # Style header row
        for ($col = 1; $col -le $lastCol3; $col++) {
            $cell = $ws3.Cells[$headerRow3, $col]
            $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $cell.Style.Fill.BackgroundColor.SetColor($HeaderBg)
            $cell.Style.Font.Color.SetColor($HeaderFg)
            $cell.Style.Font.Bold = $true
        }

        # Set column widths
        $ws3.Column(1).Width  = 12  # Month
        $ws3.Column(2).Width  = 14  # Month Name
        $ws3.Column(3).Width  = 16  # Event Type
        $ws3.Column(4).Width  = 14  # Status
        $ws3.Column(5).Width  = 50  # Title
        $ws3.Column(6).Width  = 18  # Impact Start
        $ws3.Column(7).Width  = 18  # Impact End
        $ws3.Column(8).Width  = 16  # Duration
        $ws3.Column(9).Width  = 12  # Level
        $ws3.Column(10).Width = 30  # Affected Services
        $ws3.Column(11).Width = 30  # Affected Regions
        $ws3.Column(12).Width = 60  # Summary
        $ws3.Column(13).Width = 36  # Tracking ID

        # Alternating row shading by month for visual grouping
        $currentMonth = ""
        $monthShadeToggle = $false
        $monthShadeBg = [System.Drawing.Color]::FromArgb(230, 240, 250)  # light blue
        for ($row = $headerRow3 + 1; $row -le $lastRow3; $row++) {
            $monthVal = $ws3.Cells[$row, 1].Value
            if ($monthVal -ne $currentMonth) {
                $currentMonth = $monthVal
                $monthShadeToggle = -not $monthShadeToggle
            }
            if ($monthShadeToggle) {
                for ($col = 1; $col -le $lastCol3; $col++) {
                    $c = $ws3.Cells[$row, $col]
                    if ($c.Style.Fill.PatternType -ne [OfficeOpenXml.Style.ExcelFillStyle]::Solid) {
                        $c.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                        $c.Style.Fill.BackgroundColor.SetColor($monthShadeBg)
                    }
                }
            }
            # Wrap summary text
            $ws3.Cells[$row, 12].Style.WrapText = $true
            $ws3.Row($row).Height = 40
        }
    } else {
        $emptyData3 = @([PSCustomObject]@{ Message = "No service health events found for the target regions in the reporting period." })
        $excelPkg = $emptyData3 | Export-Excel -ExcelPackage $excelPkg -WorksheetName $tab3Name `
            -AutoSize -PassThru
    }

    # Save and close
    Close-ExcelPackage $excelPkg
    Write-Host "[  OK  ] Report saved to: $OutputFile" -ForegroundColor Green
}
#endregion

#region ── 6. MAIN EXECUTION ─────────────────────────────────────────────────────

try {
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    # Step 1: Prerequisites & connection
    $context = Test-Prerequisites

    # Step 2: Resolve regions
    $Regions = Resolve-Regions -RequestedRegions $Regions

    # Step 3: Define date ranges
    $now          = Get-Date
    $startDate12m = (Get-Date -Year $now.Year -Month $now.Month -Day 1).AddMonths(-($MonthsBack - 1))
    $endDate      = $now
    $startDate1m  = $now.AddMonths(-1)

    Write-Host "`n── Date ranges ──" -ForegroundColor Cyan
    Write-Host "  SLA period : $($startDate12m.ToString('yyyy-MM-dd')) to $($endDate.ToString('yyyy-MM-dd')) ($MonthsBack months)" -ForegroundColor Gray
    Write-Host "  Incidents  : $($startDate1m.ToString('yyyy-MM-dd')) to $($endDate.ToString('yyyy-MM-dd')) (past month)" -ForegroundColor Gray
    Write-Host "  Regions    : $($Regions.Count) region(s)" -ForegroundColor Gray
    Write-Host "  Subscriptions: $($script:ResolvedSubscriptionIds.Count) subscription(s)`n" -ForegroundColor Gray

    # Step 3: Collect data
    $resources      = Get-ResourceInventory -TargetRegions $Regions
    $healthData     = Get-ResourceAvailability -TargetRegions $Regions -StartDate $startDate12m -EndDate $endDate
    $incidents12m   = Get-ServiceHealthIncidents -TargetRegions $Regions -StartDate $startDate12m -EndDate $endDate
    $incidents1m    = Get-ServiceHealthIncidents -TargetRegions $Regions -StartDate $startDate1m -EndDate $endDate
    $alerts1m       = Get-ServiceHealthAlerts -StartDate $startDate1m -EndDate $endDate

    # Also get Resource Health events for additional context
    $healthEvents   = Get-ResourceHealthEvents -TargetRegions $Regions -StartDate $startDate1m -EndDate $endDate

    # Step 4: Build report data
    $slaMatrix = Build-SLAMatrix `
        -HealthData $healthData `
        -Incidents $incidents12m `
        -Resources $resources `
        -TargetRegions $Regions `
        -StartDate $startDate12m `
        -EndDate $endDate

    $incidentsTable = Build-IncidentsTable `
        -Incidents ($incidents1m + $healthEvents) `
        -Alerts $alerts1m `
        -TargetRegions $Regions

    $healthTimeline = Build-ServiceHealthTimeline `
        -Incidents ($incidents12m + $healthEvents) `
        -TargetRegions $Regions `
        -StartDate $startDate12m `
        -EndDate $endDate

    # Step 5: Export to Excel
    Export-SLAReport -SLAMatrix $slaMatrix -IncidentsTable $incidentsTable `
        -HealthTimeline $healthTimeline -OutputFile $OutputPath

    $stopwatch.Stop()

    # ── Summary ─────────────────────────────────────────────────────────────
    Write-Host "`n╔══════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║           Report Generated Successfully          ║" -ForegroundColor Green
    Write-Host "╚══════════════════════════════════════════════════╝" -ForegroundColor Green
    Write-Host "  File     : $OutputPath" -ForegroundColor White
    Write-Host "  Duration : $([Math]::Round($stopwatch.Elapsed.TotalSeconds, 1)) seconds" -ForegroundColor White
    Write-Host "  Subs     : $($script:ResolvedSubscriptionIds.Count) subscription(s)" -ForegroundColor White
    Write-Host "  Resources: $($resources.Count) across $($Regions.Count) regions" -ForegroundColor White
    Write-Host "  Incidents: $($incidentsTable.Count) in past month" -ForegroundColor White
    Write-Host ""

    # Open the file
    if ($OutputPath -and (Test-Path $OutputPath)) {
        Write-Host "Opening report..." -ForegroundColor Cyan
        Invoke-Item $OutputPath
    }

} catch {
    Write-Host "`n[FATAL] $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Line: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
    Write-Host "  Stack: $($_.ScriptStackTrace)" -ForegroundColor DarkRed

    if ($_.Exception.Message -like "*Login*" -or $_.Exception.Message -like "*auth*" -or $_.Exception.Message -like "*token*") {
        Write-Host @"

  ╔═══ AUTHENTICATION TROUBLESHOOTING ═════════════════════════════════╗
  ║                                                                     ║
  ║  1. Clear cached credentials:                                       ║
  ║       Disconnect-AzAccount                                          ║
  ║       Clear-AzContext -Force                                        ║
  ║  2. Re-authenticate:                                                ║
  ║       Connect-AzAccount                                             ║
  ║  3. Verify subscription:                                            ║
  ║       Get-AzSubscription | Format-Table                             ║
  ║  4. Set correct subscription:                                       ║
  ║       Set-AzContext -SubscriptionId <your-id>                       ║
  ║                                                                     ║
  ╚═════════════════════════════════════════════════════════════════════╝
"@ -ForegroundColor Yellow
    }

    exit 1
}
#endregion
