<#
MIT License

Copyright (c) 2024 Guil Lima

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

DISCLAIMER:
  This report is an estimation tool, NOT an official SLA measurement.
  The availability percentages shown are approximations based on Azure Resource
  Health signals and Service Health incident tracking data. They are NOT the same
  as Microsoft's contractual SLA metrics. Official Azure SLAs are defined in the
  Service Level Agreements for Online Services:
  https://www.microsoft.com/licensing/docs/view/Service-Level-Agreements-SLA-for-Online-Services

  Always review and validate the results against your own monitoring data
  (Azure Monitor, Application Insights, third-party tools, etc.) before
  presenting them in reports or making decisions based on them.

  This tool is intended as a supplementary data source for operational reviews,
  governance reporting, and trend analysis — not as the single source of truth
  for availability.

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
    Date    : 2026-04-07
    Version : 2.0.0
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

# ── Resolve OutputPath: if it's a directory, append the default filename ────
$defaultFileName = "AzureSLA_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
if ($OutputPath -and (Test-Path $OutputPath -PathType Container)) {
    $OutputPath = Join-Path $OutputPath $defaultFileName
} elseif ($OutputPath -and (-not [System.IO.Path]::HasExtension($OutputPath))) {
    # Path doesn't exist yet but has no extension — treat as directory
    $null = New-Item -Path $OutputPath -ItemType Directory -Force -ErrorAction SilentlyContinue
    $OutputPath = Join-Path $OutputPath $defaultFileName
}

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
            $troubleshootMsg = @"

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
"@
            Write-Host $troubleshootMsg -ForegroundColor Yellow
            throw "Azure authentication failed. See troubleshooting steps above."
        }
    }

    Write-Host "[  OK  ] Connected as: $($ctx.Account.Id)" -ForegroundColor Green
    Write-Host "[  OK  ] Default subscription: $($ctx.Subscription.Name) ($($ctx.Subscription.Id))" -ForegroundColor Green

    # ── Resolve subscription scope ─────────────────────────────────────────
    Write-Host "`n── Resolving subscription scope ──" -ForegroundColor Cyan
    if ($SubscriptionIds -and $SubscriptionIds.Count -gt 0) {
        # User specified explicit subscription IDs
        $targetSubs = [System.Collections.Generic.List[object]]::new()
        foreach ($sid in $SubscriptionIds) {
            try {
                $s = Get-AzSubscription -SubscriptionId $sid -ErrorAction Stop
                $targetSubs.Add($s)
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
        $regionResults = Invoke-PaginatedGraphQuery -Query $query -Label 'region records'
        $resolved = @($regionResults | ForEach-Object { $_.location.ToLower() } | Select-Object -Unique)

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

function Convert-TicksToDateTime {
    <#
    .SYNOPSIS
        Converts .NET ticks (Int64) to a PowerShell DateTime (UTC).
        ServiceHealthResources stores ImpactStartTime / ImpactMitigationTime as
        ticks, not ISO‑8601 strings.
    #>
    param([object]$Ticks)
    if ($null -eq $Ticks) { return $null }
    try {
        $val = [long]$Ticks
        # Ticks of 0 or near-zero → DateTime.MinValue (0001-01-01) — treat as null
        if ($val -le 0) { return $null }
        $dt = [datetime]::new($val, [System.DateTimeKind]::Utc)
        # Sanity check: dates before 2000 are clearly invalid for Azure events
        if ($dt.Year -lt 2000) { return $null }
        return $dt
    } catch {
        return $null
    }
}

function Invoke-PaginatedGraphQuery {
    <#
    .SYNOPSIS
        Executes an Azure Resource Graph query with automatic pagination.
        First runs a count query to log the total, then fetches all rows in
        batches of 1000 using -First / -Skip (up to 5 000) and $SkipToken
        (beyond 5 000).

        Automatically batches subscriptions into groups of 200 to stay within
        the Search-AzGraph -Subscription limit.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Query,
        [string]$Label = 'records',
        [int]$BatchSize = 1000
    )

    # ── 0. Split subscriptions into batches of 200 (ARG limit) ──────────
    $maxSubsPerQuery = 200
    $allSubIds = @($script:ResolvedSubscriptionIds)
    $subBatches = [System.Collections.Generic.List[string[]]]::new()

    for ($i = 0; $i -lt $allSubIds.Count; $i += $maxSubsPerQuery) {
        $end = [Math]::Min($i + $maxSubsPerQuery, $allSubIds.Count) - 1
        $subBatches.Add([string[]]@($allSubIds[$i..$end]))
    }

    if ($subBatches.Count -gt 1) {
        Write-Host "[INFO ] Splitting $($allSubIds.Count) subscriptions into $($subBatches.Count) batches of max $maxSubsPerQuery for $Label" -ForegroundColor Cyan
    }

    $allResults = [System.Collections.Generic.List[object]]::new()
    $subBatchIndex = 0

    foreach ($subBatch in $subBatches) {
        $subBatchIndex++
        $subBatchPrefix = if ($subBatches.Count -gt 1) { "Sub-batch $subBatchIndex/$($subBatches.Count): " } else { "" }

        # ── 1. Count query (informational only — never skip based on count) ─
        $countQuery = @"
$Query
| count
"@
        $totalCount = -1
        try {
            $countResult = Search-AzGraph -Query $countQuery -First 1 `
                -Subscription $subBatch -ErrorAction Stop

            # Try multiple property name patterns (varies by Az.ResourceGraph version)
            $totalCount = -1
            foreach ($propName in @('Count_', 'count_', 'Count')) {
                $val = $countResult | Select-Object -ExpandProperty $propName -ErrorAction SilentlyContinue
                if ($null -ne $val -and $val -is [int] -or $val -is [long]) {
                    $totalCount = [int]$val
                    break
                }
            }
            # Fallback: try direct property access
            if ($totalCount -lt 0) {
                if ($null -ne $countResult.Count_)  { $totalCount = [int]$countResult.Count_ }
                elseif ($null -ne $countResult.count_) { $totalCount = [int]$countResult.count_ }
            }

            if ($totalCount -ge 0) {
                Write-Host "[INFO ] ${subBatchPrefix}Total $Label to retrieve: $totalCount" -ForegroundColor Cyan
            } else {
                Write-Host "[WARN ] ${subBatchPrefix}Could not parse count — will paginate until exhausted." -ForegroundColor Yellow
            }
        } catch {
            Write-Host "[WARN ] ${subBatchPrefix}Count query failed — will paginate until exhausted." -ForegroundColor Yellow
        }

        # ── 2. Paginated fetch ──────────────────────────────────────────
        $skip      = 0
        $skipToken = $null
        $batchNum  = 0
        $batchResults = 0

        while ($true) {
            $batchNum++
            $params = @{
                Query        = $Query
                First        = $BatchSize
                Subscription = $subBatch
                ErrorAction  = 'Stop'
            }

            if ($skipToken) {
                $params['SkipToken'] = $skipToken
            } elseif ($skip -gt 0) {
                $params['Skip'] = $skip
            }

            try {
                $batch = Search-AzGraph @params
            } catch {
                Write-Host "[WARN ] ${subBatchPrefix}Batch $batchNum failed: $($_.Exception.Message)" -ForegroundColor Yellow
                break
            }

            if (-not $batch -or $batch.Count -eq 0) { break }

            $allResults.AddRange(@($batch))
            $batchResults += $batch.Count
            Write-Host "[INFO ] ${subBatchPrefix}Batch $batchNum — fetched $($batch.Count) $Label (sub-batch total: $batchResults)" -ForegroundColor Gray

            # Determine next page strategy
            if ($batch.SkipToken) {
                $skipToken = $batch.SkipToken
            } else {
                $skipToken = $null
                $skip += $batch.Count
            }

            # Stop when we've collected everything for this sub-batch
            if ($totalCount -ge 0 -and $batchResults -ge $totalCount) { break }
            if ($batch.Count -lt $BatchSize) { break }
        }
    }

    Write-Host "[  OK  ] Retrieved $($allResults.Count) $Label (paginated across $($subBatches.Count) subscription batch(es))" -ForegroundColor Green
    return , $allResults.ToArray()
}

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

    # Pull all service-issue events; date filtering done in PowerShell because
    # ImpactStartTime is stored as .NET ticks, not a Kusto-native datetime.
    $query = @"
servicehealthresources
| where type =~ 'microsoft.resourcehealth/events'
| where tostring(properties.EventType) =~ 'ServiceIssue'
| extend status          = tostring(properties.Status)
| extend summary         = tostring(properties.Summary)
| extend eventLevel      = tostring(properties.EventLevel)
| extend impactedServices = properties.Impact
| project id, name, properties, status, Title = properties.Title, summary, eventLevel, impactedServices
"@

    try {
        $raw = Invoke-PaginatedGraphQuery -Query $query -Label 'health events'

        # Deduplicate: servicehealthresources is tenant-scoped, so subscription batching returns duplicates
        if ($raw.Count -gt 0) {
            $rawBefore = $raw.Count
            $raw = @($raw | Sort-Object -Property name -Unique)
            if ($raw.Count -lt $rawBefore) {
                Write-Host "[INFO ] Deduplicated health events (from $rawBefore to $($raw.Count)) by tracking ID" -ForegroundColor Gray
            }
        }

        # Convert ticks → DateTime and filter by date range in PowerShell
        $results = foreach ($r in $raw) {
            $start = Convert-TicksToDateTime $r.properties.ImpactStartTime
            $end   = Convert-TicksToDateTime $r.properties.ImpactMitigationTime
            if ($null -eq $start) { continue }
            if ($start -lt $StartDate -or $start -gt $EndDate) { continue }

            $r | Add-Member -NotePropertyName 'impactStartTime' -NotePropertyValue $start -Force -PassThru |
                 Add-Member -NotePropertyName 'impactEndTime'   -NotePropertyValue $end   -Force -PassThru
        }

        Write-Host "[  OK  ] Retrieved $(@($results).Count) service health events (filtered to date range)" -ForegroundColor Green
        return @($results)
    } catch {
        Write-Host "[WARN ] Resource Graph query failed: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "[INFO ] Falling back to Activity Log method..." -ForegroundColor Yellow
        return @()
    }
}

function Get-ActivityLogViaApi {
    <#
    .SYNOPSIS
        Queries Activity Log events for a single subscription using the REST API.
        Returns parsed event objects. Handles pagination via nextLink.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SubscriptionId,
        [Parameter(Mandatory)][string]$Filter,
        [string]$SubscriptionName = $SubscriptionId
    )

    $apiVersion = '2015-04-01'
    $path = "/subscriptions/$SubscriptionId/providers/Microsoft.Insights/eventtypes/management/values?api-version=$apiVersion&`$filter=$Filter"
    $events = [System.Collections.Generic.List[object]]::new()

    while ($path) {
        try {
            $response = Invoke-AzRestMethod -Path $path -ErrorAction Stop
            if ($response.StatusCode -ne 200) {
                Write-Host "  [WARN ] API returned $($response.StatusCode) for subscription $SubscriptionName" -ForegroundColor Yellow
                break
            }
            $body = $response.Content | ConvertFrom-Json
            if ($body.value) {
                $events.AddRange(@($body.value))
            }
            # Follow pagination
            if ($body.nextLink) {
                # nextLink is a full URL; extract the path+query portion for Invoke-AzRestMethod
                $uri = [System.Uri]$body.nextLink
                $path = $uri.PathAndQuery
            } else {
                $path = $null
            }
        } catch {
            Write-Host "  [WARN ] API call failed for subscription $SubscriptionName : $($_.Exception.Message)" -ForegroundColor Yellow
            break
        }
    }
    return , $events.ToArray()
}

function Get-ServiceHealthAlerts {
    <#
    .SYNOPSIS
        Retrieves Service Health alerts from Activity Log for the past month.
        Uses direct REST API calls via Invoke-AzRestMethod to avoid the overhead
        of Set-AzContext per subscription (much faster for large environments).
        Parallelises requests on PowerShell 7+ using ForEach-Object -Parallel.
    #>
    [CmdletBinding()]
    param(
        [datetime]$StartDate,
        [datetime]$EndDate
    )

    Write-Host "── Querying Service Health alerts from Activity Log (REST API) ──" -ForegroundColor Cyan

    # ── Build subscription-name lookup upfront ──────────────────────────
    $subNameMap = @{}
    foreach ($sub in (Get-AzSubscription -ErrorAction SilentlyContinue)) {
        $subNameMap[$sub.Id] = $sub.Name
    }

    # OData filter strings for the Activity Log API
    $startIso = $StartDate.ToUniversalTime().ToString('o')
    $endIso   = $EndDate.ToUniversalTime().ToString('o')

    $filterResourceHealth = "eventTimestamp ge '$startIso' and eventTimestamp le '$endIso' and resourceProvider eq 'Microsoft.ResourceHealth'"
    $filterServiceHealth  = "eventTimestamp ge '$startIso' and eventTimestamp le '$endIso' and eventChannels eq 'Admin, Operation, Service'"

    $alerts = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
    $totalSubs = $script:ResolvedSubscriptionIds.Count

    # ── Helper script block (used by both parallel and sequential paths) ─
    $processSubscription = {
        param($subId, $subName, $filterRH, $filterSH)

        # 1. Resource Health events
        $rhEvents = Get-ActivityLogViaApi -SubscriptionId $subId -Filter $filterRH -SubscriptionName $subName
        foreach ($ev in $rhEvents) {
            [PSCustomObject]@{
                Timestamp     = $ev.eventTimestamp
                Category      = $ev.category.value
                Level         = $ev.level
                OperationName = $ev.operationName.value
                Status        = $ev.status.value
                Description   = $ev.description
                ResourceId    = $ev.resourceId
                CorrelationId = $ev.correlationId
                Subscription  = $subName
            }
        }

        # 2. ServiceHealth events
        $shEvents = Get-ActivityLogViaApi -SubscriptionId $subId -Filter $filterSH -SubscriptionName $subName
        foreach ($ev in $shEvents) {
            if ($ev.category.value -ne 'ServiceHealth') { continue }
            [PSCustomObject]@{
                Timestamp     = $ev.eventTimestamp
                Category      = 'ServiceHealth'
                Level         = $ev.level
                OperationName = $ev.operationName.value
                Status        = $ev.status.value
                Description   = $ev.description
                ResourceId    = $ev.resourceId
                CorrelationId = $ev.correlationId
                Subscription  = $subName
            }
        }
    }

    try {
        $isPwsh7 = $PSVersionTable.PSVersion.Major -ge 7

        if ($isPwsh7) {
            # ── Parallel execution (PowerShell 7+) ──────────────────────
            Write-Host "[INFO ] Using parallel API calls ($totalSubs subscriptions, throttle 10)" -ForegroundColor Cyan

            # Thread-safe counter for progress tracking
            $progressCounter = [System.Collections.Concurrent.ConcurrentDictionary[string,string]]::new()
            $completedCount  = [ref] 0

            # Run the parallel work as a PowerShell job so we can poll progress
            $parallelJob = {
                param($subIds, $subNameMap, $alerts, $filterRH, $filterSH, $progressCounter, $completedCount)
                $subIds | ForEach-Object -ThrottleLimit 10 -Parallel {
                    $subId    = $_
                    $nameMap  = $using:subNameMap
                    $subName  = if ($nameMap[$subId]) { $nameMap[$subId] } else { $subId }
                    $bag      = $using:alerts
                    $progress = $using:progressCounter
                    $done     = $using:completedCount

                    $progress[$subId] = $subName  # mark as in-progress

                    function Get-ActivityLogViaApi {
                        param([string]$SubscriptionId, [string]$Filter, [string]$SubscriptionName)
                        $path = "/subscriptions/$SubscriptionId/providers/Microsoft.Insights/eventtypes/management/values?api-version=2015-04-01&`$filter=$Filter"
                        $events = [System.Collections.Generic.List[object]]::new()
                        while ($path) {
                            try {
                                $response = Invoke-AzRestMethod -Path $path -ErrorAction Stop
                                if ($response.StatusCode -ne 200) { break }
                                $body = $response.Content | ConvertFrom-Json
                                if ($body.value) { $events.AddRange(@($body.value)) }
                                if ($body.nextLink) { $path = ([System.Uri]$body.nextLink).PathAndQuery } else { $path = $null }
                            } catch { break }
                        }
                        return , $events.ToArray()
                    }

                    # Resource Health events
                    $rhEvents = Get-ActivityLogViaApi -SubscriptionId $subId -Filter $using:filterRH -SubscriptionName $subName
                    foreach ($ev in $rhEvents) {
                        $bag.Add([PSCustomObject]@{
                            Timestamp     = $ev.eventTimestamp
                            Category      = $ev.category.value
                            Level         = $ev.level
                            OperationName = $ev.operationName.value
                            Status        = $ev.status.value
                            Description   = $ev.description
                            ResourceId    = $ev.resourceId
                            CorrelationId = $ev.correlationId
                            Subscription  = $subName
                        })
                    }

                    # ServiceHealth events
                    $shEvents = Get-ActivityLogViaApi -SubscriptionId $subId -Filter $using:filterSH -SubscriptionName $subName
                    foreach ($ev in $shEvents) {
                        if ($ev.category.value -ne 'ServiceHealth') { continue }
                        $bag.Add([PSCustomObject]@{
                            Timestamp     = $ev.eventTimestamp
                            Category      = 'ServiceHealth'
                            Level         = $ev.level
                            OperationName = $ev.operationName.value
                            Status        = $ev.status.value
                            Description   = $ev.description
                            ResourceId    = $ev.resourceId
                            CorrelationId = $ev.correlationId
                            Subscription  = $subName
                        })
                    }

                    # Mark completed
                    [System.Threading.Interlocked]::Increment($done) | Out-Null
                    $removed = $null
                    $progress.TryRemove($subId, [ref]$removed) | Out-Null
                }
            }

            # Start the parallel work in a background thread
            $ps = [powershell]::Create()
            $ps.AddScript($parallelJob).AddArgument($script:ResolvedSubscriptionIds).AddArgument($subNameMap).AddArgument($alerts).AddArgument($filterResourceHealth).AddArgument($filterServiceHealth).AddArgument($progressCounter).AddArgument($completedCount) | Out-Null
            $asyncResult = $ps.BeginInvoke()

            # ── Poll progress bar while parallel work runs ──────────────
            $progressId = 1
            while (-not $asyncResult.IsCompleted) {
                $done    = $completedCount.Value
                $pct     = if ($totalSubs -gt 0) { [int][Math]::Min(100, ($done / $totalSubs) * 100) } else { 0 }
                $running = @($progressCounter.Values)
                $statusMsg = if ($running.Count -gt 0) {
                    "Active: $($running[0..([Math]::Min(2, $running.Count - 1))] -join ', ')" +
                    $(if ($running.Count -gt 3) { " +$($running.Count - 3) more" })
                } else { "Starting..." }

                Write-Progress -Id $progressId `
                    -Activity "Querying Activity Logs ($done/$totalSubs subscriptions)" `
                    -Status $statusMsg `
                    -PercentComplete $pct

                Start-Sleep -Milliseconds 500
            }

            # Final update and cleanup
            $ps.EndInvoke($asyncResult)
            $ps.Dispose()
            Write-Progress -Id $progressId -Activity "Querying Activity Logs" -Completed
        } else {
            # ── Sequential execution (Windows PowerShell 5.1) ───────────
            Write-Host "[INFO ] Using sequential API calls ($totalSubs subscriptions)" -ForegroundColor Cyan
            $counter = 0
            foreach ($subId in $script:ResolvedSubscriptionIds) {
                $counter++
                $subName = if ($subNameMap[$subId]) { $subNameMap[$subId] } else { $subId }
                Write-Host "  [$counter/$totalSubs] $subName" -ForegroundColor Gray

                $results = & $processSubscription $subId $subName $filterResourceHealth $filterServiceHealth
                foreach ($r in $results) { $alerts.Add($r) }
            }
        }

        $alertList = @($alerts.ToArray())
        Write-Host "[  OK  ] Retrieved $($alertList.Count) health alerts from Activity Log across $totalSubs subscription(s)" -ForegroundColor Green
    } catch {
        Write-Host "[WARN ] Activity Log query failed: $($_.Exception.Message)" -ForegroundColor Yellow
        $alertList = @()
    }

    return $alertList
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
        $resources = Invoke-PaginatedGraphQuery -Query $query -Label 'resources'
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
        $healthData = Invoke-PaginatedGraphQuery -Query $query -Label 'availability records'

        # Deduplicate: HealthResources is tenant-scoped, so subscription batching returns duplicates
        if ($healthData.Count -gt 0) {
            $seen = @{}
            $unique = [System.Collections.Generic.List[object]]::new()
            foreach ($h in $healthData) {
                $key = "$($h.resourceId)|$($h.availabilityState)|$($h.occurredTime)"
                if (-not $seen.ContainsKey($key)) {
                    $seen[$key] = $true
                    $unique.Add($h)
                }
            }
            $dupeCount = $healthData.Count - $unique.Count
            if ($dupeCount -gt 0) {
                Write-Host "[INFO ] Deduplicated $dupeCount availability records (from $($healthData.Count) to $($unique.Count))" -ForegroundColor Gray
            }
            $healthData = $unique.ToArray()
        }

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

    # Pull all events; date filtering done in PowerShell because
    # ImpactStartTime is stored as .NET ticks, not a Kusto-native datetime.
    $query = @"
servicehealthresources
| where type =~ 'microsoft.resourcehealth/events'
| extend eventType        = tostring(properties.EventType)
| extend status           = tostring(properties.Status)
| extend summary          = tostring(properties.Summary)
| extend level            = tostring(properties.EventLevel)
| extend impactedServices = properties.Impact
| project name, eventType, status, Title = properties.Title, summary, level, impactedServices, properties
| order by name desc
"@

    try {
        $raw = Invoke-PaginatedGraphQuery -Query $query -Label 'service health incidents'

        # Deduplicate: servicehealthresources is tenant-scoped, so subscription batching returns duplicates
        if ($raw.Count -gt 0) {
            $rawBefore = $raw.Count
            $raw = @($raw | Sort-Object -Property name -Unique)
            if ($raw.Count -lt $rawBefore) {
                Write-Host "[INFO ] Deduplicated service health incidents (from $rawBefore to $($raw.Count)) by tracking ID" -ForegroundColor Gray
            }
        }

        # Convert ticks → DateTime and filter by date range in PowerShell
        $incidents = foreach ($r in $raw) {
            $start  = Convert-TicksToDateTime $r.properties.ImpactStartTime
            $end    = Convert-TicksToDateTime $r.properties.ImpactMitigationTime
            $update = Convert-TicksToDateTime $r.properties.LastUpdateTime
            if ($null -eq $start) { continue }
            if ($start -lt $StartDate -or $start -gt $EndDate) { continue }

            $r | Add-Member -NotePropertyName 'impactStartTime' -NotePropertyValue $start  -Force -PassThru |
                 Add-Member -NotePropertyName 'impactEndTime'   -NotePropertyValue $end    -Force -PassThru |
                 Add-Member -NotePropertyName 'lastUpdateTime'  -NotePropertyValue $update -Force -PassThru
        }

        Write-Host "[  OK  ] Retrieved $(@($incidents).Count) service health incidents (filtered to date range)" -ForegroundColor Green
        return @($incidents)
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
    $slaRows = [System.Collections.Generic.List[object]]::new()

    # ── Pre-index health data by region|category|month for O(1) lookup ──
    Write-Host "[INFO ] Pre-indexing health data..." -ForegroundColor Gray

    # Pre-compute month boundaries once (used throughout)
    $monthBoundaries = [System.Collections.Generic.List[object]]::new()
    for ($i = $MonthsBack - 1; $i -ge 0; $i--) {
        $ms = (Get-Date).AddMonths(-$i).Date
        $ms = Get-Date -Year $ms.Year -Month $ms.Month -Day 1
        $me = $ms.AddMonths(1).AddSeconds(-1)
        $monthBoundaries.Add(@{
            Label = $ms.ToString("MMM yyyy")
            Start = $ms
            End   = $me
            Key   = $ms.ToString("yyyy-MM")
        })
    }

    # healthIndex: key = "region|category|yyyy-MM" → count of unhealthy events
    $healthIndex = @{}
    $healthTotal = 0
    foreach ($h in $HealthData) {
        if ($h.availabilityState -eq 'Available') { continue }
        if ($null -eq $h.occurredTime) { continue }
        $healthTotal++

        $rcBase = "$($h.location)|$($h.ServiceCategory)"
        $monthKey = ([datetime]$h.occurredTime).ToString("yyyy-MM")
        $fullKey  = "$rcBase|$monthKey"

        if ($healthIndex.ContainsKey($fullKey)) {
            $healthIndex[$fullKey]++
        } else {
            $healthIndex[$fullKey] = 1
        }
    }
    Write-Host "[INFO ] Health index: $healthTotal unhealthy events binned into $($healthIndex.Count) region|category|month buckets (from $($HealthData.Count) total records)" -ForegroundColor Gray

    # ── Pre-index resource counts by region|category ────────────────────
    $resourceCountIndex = @{}
    foreach ($r in $Resources) {
        $key = "$($r.location)|$($r.ServiceCategory)"
        if ($resourceCountIndex.ContainsKey($key)) {
            $resourceCountIndex[$key]++
        } else {
            $resourceCountIndex[$key] = 1
        }
    }

    # ── Pre-process incidents: build a lookup of (region, service) → incident windows ──
    Write-Host "[INFO ] Pre-indexing incidents..." -ForegroundColor Gray
    $serviceTypeMap = @{
        'Compute'  = @('Virtual Machines', 'Compute', 'Virtual Machine Scale Sets')
        'SQL DB'   = @('SQL Database', 'SQL Managed Instance', 'Azure SQL', 'SQL')
        'Web Apps' = @('App Service', 'Web Apps', 'App Service (Web Apps)')
        'Storage'  = @('Storage', 'Storage Accounts')
    }

    # incidentIndex: key = "region|category|yyyy-MM" → list of @{ Start; End }
    $incidentIndex = @{}
    $skippedNonServiceIssue = 0
    $skippedActive          = 0
    $indexedIncidents       = 0

    foreach ($incident in $Incidents) {
        if ($null -eq $incident.impactedServices) { continue }

        # Only count actual service outages for SLA — not maintenance, advisories, or security notices
        if ($incident.eventType -and $incident.eventType -ne 'ServiceIssue') {
            $skippedNonServiceIssue++
            continue
        }

        # Determine the effective end time
        $incStart = [datetime]$incident.impactStartTime
        $incEnd   = $null
        if ($incident.impactEndTime) {
            $incEnd = [datetime]$incident.impactEndTime
        } elseif ($incident.status -eq 'Active') {
            # Active/ongoing incidents: use lastUpdateTime as a conservative proxy end,
            # or skip entirely — they don't have a confirmed outage window yet.
            if ($incident.lastUpdateTime) {
                $incEnd = [datetime]$incident.lastUpdateTime
            } else {
                # No end time and no update time — skip to avoid inflating downtime
                $skippedActive++
                continue
            }
        }
        # If resolved but no end time (rare), use start + 1 hour as conservative estimate
        if ($null -eq $incEnd) {
            $incEnd = $incStart.AddHours(1)
        }

        $impactedServicesArray = if ($incident.impactedServices -is [array]) {
            $incident.impactedServices
        } else {
            @($incident.impactedServices)
        }

        foreach ($impact in $impactedServicesArray) {
            $serviceName = if ($impact.ImpactedService) { $impact.ImpactedService } else { $impact.ServiceName }
            $impactedRegions = if ($impact.ImpactedRegions) { $impact.ImpactedRegions } else { @() }

            # Determine which of our categories this service matches
            $matchedCategories = @()
            foreach ($cat in $serviceCategories) {
                foreach ($svcName in $serviceTypeMap[$cat]) {
                    if ($serviceName -like "*$svcName*") {
                        $matchedCategories += $cat
                        break
                    }
                }
            }
            if ($matchedCategories.Count -eq 0) { continue }

            # Determine which of our target regions this impact matches
            $matchedRegions = @()
            foreach ($region in $TargetRegions) {
                $regionDisplay = if ($RegionDisplayNames[$region]) { $RegionDisplayNames[$region] } else { $region }
                foreach ($ir in $impactedRegions) {
                    $irName = if ($ir.ImpactedRegion) { $ir.ImpactedRegion } else { $ir }
                    if ($irName -like "*$regionDisplay*" -or $irName -eq $region) {
                        $matchedRegions += $region
                        break
                    }
                }
            }
            if ($matchedRegions.Count -eq 0) { continue }

            # Add the incident window to each matching region|category|month
            $indexedIncidents++
            foreach ($mr in $matchedRegions) {
                foreach ($mc in $matchedCategories) {
                    $rcBase = "$mr|$mc"
                    # Bin into each month this incident overlaps
                    foreach ($mb in $monthBoundaries) {
                        if ($incStart -gt $mb.End -or $incEnd -lt $mb.Start) { continue }
                        $fullKey = "$rcBase|$($mb.Key)"
                        if (-not $incidentIndex.ContainsKey($fullKey)) {
                            $incidentIndex[$fullKey] = [System.Collections.Generic.List[object]]::new()
                        }
                        $incidentIndex[$fullKey].Add(@{ Start = $incStart; End = $incEnd })
                    }
                }
            }
        }
    }
    Write-Host "[INFO ] Incident index: $($incidentIndex.Count) region|category|month buckets ($indexedIncidents service issues indexed)" -ForegroundColor Gray
    if ($skippedNonServiceIssue -gt 0) {
        Write-Host "[INFO ] Skipped $skippedNonServiceIssue non-ServiceIssue events (maintenance/advisories) — not counted as downtime" -ForegroundColor Gray
    }
    if ($skippedActive -gt 0) {
        Write-Host "[WARN ] Skipped $skippedActive active incidents without end/update time — excluded from SLA calculation" -ForegroundColor Yellow
    }

    # ── Build SLA rows with progress tracking ───────────────────────────
    $totalCells   = $TargetRegions.Count * $serviceCategories.Count
    $cellsDone    = 0
    $progressId   = 2
    $swMatrix     = [System.Diagnostics.Stopwatch]::StartNew()

    foreach ($region in $TargetRegions) {
        $regionDisplay = if ($RegionDisplayNames[$region]) { $RegionDisplayNames[$region] } else { $region }

        foreach ($category in $serviceCategories) {
            $cellsDone++
            $pct = [int][Math]::Min(100, ($cellsDone / $totalCells) * 100)
            $elapsed = $swMatrix.Elapsed
            Write-Progress -Id $progressId `
                -Activity "Building SLA matrix ($cellsDone/$totalCells) — elapsed $([Math]::Round($elapsed.TotalSeconds,0))s" `
                -Status "$regionDisplay — $category" `
                -PercentComplete $pct

            $row = [ordered]@{
                'Region'   = $regionDisplay
                'Service'  = $category
            }

            # Count resources in this region/category (pre-indexed)
            $rcKey = "$region|$category"
            $resourceCount = if ($resourceCountIndex.ContainsKey($rcKey)) { $resourceCountIndex[$rcKey] } else { 0 }

            $row['Resource Count'] = $resourceCount

            # Build month columns using pre-computed boundaries and pre-binned data
            foreach ($mb in $monthBoundaries) {
                if ($resourceCount -eq 0) {
                    $row[$mb.Label] = "N/A"
                    continue
                }

                $fullKey = "$rcKey|$($mb.Key)"
                $unhealthyCount   = if ($healthIndex.ContainsKey($fullKey))   { $healthIndex[$fullKey] }   else { 0 }
                $monthIncidents   = if ($incidentIndex.ContainsKey($fullKey)) { $incidentIndex[$fullKey] } else { @() }

                $totalMinutes    = ($mb.End - $mb.Start).TotalMinutes
                $downtimeMinutes = 0
                $merged          = $null

                # Health data: each event is per-resource, not per-service.
                # Calculate as weighted average: (unhealthy resources / total resources) × 30 min
                if ($unhealthyCount -gt 0) {
                    $affectedFraction = [Math]::Min(1.0, $unhealthyCount / $resourceCount)
                    $downtimeMinutes += $affectedFraction * 30
                }
                $healthDowntime = $downtimeMinutes

                # Service health incidents: merge overlapping windows, then cap each
                # merged window's contribution. Incident tracking windows represent the
                # investigation period, NOT continuous downtime. The actual outage is
                # typically a fraction of the window. We cap each merged window at 4 hours
                # (240 min) — the real per-resource impact is already captured by HealthResources.
                $maxDowntimePerIncidentMinutes = 240  # 4 hours cap per merged window
                $incidentDowntime = 0

                if ($monthIncidents.Count -gt 0) {
                    # Clamp all windows to month boundaries and collect
                    $clampedWindows = [System.Collections.Generic.List[object]]::new()
                    foreach ($iw in $monthIncidents) {
                        $iwEnd = if ($iw.End) { $iw.End } else { $mb.End }
                        $s = [datetime]([Math]::Max($iw.Start.Ticks, $mb.Start.Ticks))
                        $e = [datetime]([Math]::Min($iwEnd.Ticks,    $mb.End.Ticks))
                        if ($e -gt $s) {
                            $clampedWindows.Add(@{ S = $s; E = $e })
                        }
                    }

                    # Sort by start time, then merge overlapping intervals
                    if ($clampedWindows.Count -gt 0) {
                        $sorted = $clampedWindows | Sort-Object { $_.S }
                        $merged = [System.Collections.Generic.List[object]]::new()
                        $cur = $sorted[0]
                        for ($wi = 1; $wi -lt $sorted.Count; $wi++) {
                            $nxt = $sorted[$wi]
                            if ($nxt.S -le $cur.E) {
                                # Overlapping or adjacent — extend current window
                                if ($nxt.E -gt $cur.E) { $cur = @{ S = $cur.S; E = $nxt.E } }
                            } else {
                                $merged.Add($cur)
                                $cur = $nxt
                            }
                        }
                        $merged.Add($cur)

                        foreach ($mw in $merged) {
                            $windowMinutes = ($mw.E - $mw.S).TotalMinutes
                            # Cap each merged window — tracking period ≠ actual downtime
                            $incidentDowntime += [Math]::Min($windowMinutes, $maxDowntimePerIncidentMinutes)
                        }
                        $downtimeMinutes += $incidentDowntime
                    }
                }

                $downtimeMinutes = [Math]::Min($downtimeMinutes, $totalMinutes)
                $slaValue = [Math]::Round((($totalMinutes - $downtimeMinutes) / $totalMinutes) * 100, 4)

                # ── Diagnostic: log cells with very low SLA ──
                if ($slaValue -le 50 -and $resourceCount -gt 0) {
                    $incCount    = $monthIncidents.Count
                    $mergedCount = if ($merged) { $merged.Count } else { 0 }
                    Write-Host "[DIAG ] LOW SLA $slaValue% — $regionDisplay | $category | $($mb.Label) — resources: $resourceCount, healthEvents: $unhealthyCount (${healthDowntime}min), incidents: $incCount→${mergedCount} merged (capped ${incidentDowntime}min), total downtime: $([Math]::Round($downtimeMinutes, 1))/$([Math]::Round($totalMinutes, 0))min" -ForegroundColor Yellow
                }

                $row[$mb.Label] = $slaValue
            }

            $slaRows.Add([PSCustomObject]$row)
        }
    }

    Write-Progress -Id $progressId -Activity "Building SLA matrix" -Completed
    $swMatrix.Stop()
    Write-Host "[  OK  ] SLA matrix built: $($slaRows.Count) rows in $([Math]::Round($swMatrix.Elapsed.TotalSeconds, 1))s" -ForegroundColor Green
    return $slaRows
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

    $rows = [System.Collections.Generic.List[object]]::new()

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

        $rows.Add([PSCustomObject][ordered]@{
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
        })
    }

    # ── Process Activity Log alerts ──
    foreach ($alert in $Alerts) {
        $rows.Add([PSCustomObject][ordered]@{
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
        })
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

    $rows = [System.Collections.Generic.List[object]]::new()

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

        $rows.Add([PSCustomObject][ordered]@{
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
        })
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
        -Incidents (@($incidents1m) + @($healthEvents)) `
        -Alerts $alerts1m `
        -TargetRegions $Regions

    $healthTimeline = Build-ServiceHealthTimeline `
        -Incidents (@($incidents12m) + @($healthEvents)) `
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
        $authTroubleshootMsg = @"

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
"@
        Write-Host $authTroubleshootMsg -ForegroundColor Yellow
    }

    exit 1
}
#endregion
