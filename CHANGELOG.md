# Changelog

All notable changes to the Azure SLA & Service Health Report Generator are documented in this file.

## [2.1.0] — 2026-04-08

### Added

- **Azure Blob Storage upload** — New `-BlobContainerUrl` parameter to upload the generated report to an Azure Storage blob container via azcopy. Works both locally and from Azure Cloud Shell.
  - Supports plain URLs (with Azure CLI / `azcopy login` auth) and SAS token URLs
  - Validates azcopy is installed before attempting upload (only when parameter is used)
  - Pre-upload access validation against the blob container with detailed troubleshooting guidance
  - Required role: `Storage Blob Data Contributor` on the storage account, or a SAS token with write + create permissions
- **`.gitignore`** — Prevents `.xlsx` report files from being committed to the repository.
- **Azure Cloud Shell support** — `Invoke-Item` is skipped when running in Cloud Shell; download instructions shown instead.
- **MIT license and disclaimer in script header** — Visible at the top of the `.ps1` file for anyone running the code.
- **Detailed SLA Calculation section in README** — Technical deep-dive into the availability formula, data sources, and edge cases.

## [2.0.0] — 2026-04-07

### Added

- **REST API for Activity Log** — `Get-ServiceHealthAlerts` now uses `Invoke-AzRestMethod` instead of `Set-AzContext`/`Get-AzActivityLog` loops, dramatically reducing query time for large environments.
- **Parallel API calls (PowerShell 7+)** — Activity Log queries run in parallel across subscriptions using `ForEach-Object -Parallel` with `-ThrottleLimit 10`. Falls back to sequential on PowerShell 5.1.
- **Progress bars** — Real-time `Write-Progress` bars for both parallel API calls and SLA matrix building, using `[powershell]::Create().BeginInvoke()` async pattern with `ConcurrentDictionary` tracking.
- **Low-SLA diagnostics** — `[DIAG]` logging emitted to console for any cell with SLA ≤ 50%, showing the breakdown of health events, incidents, and merged windows.
- **Interval merge algorithm** — Overlapping incident time windows are now merged before calculating downtime, preventing the same time period from being counted twice.
- **Comprehensive README** — Full rewrite with sections for non-technical users: metric calculation methodology, results interpretation guide, analysis guidance, data accuracy disclaimer, and step-by-step requirements.

### Changed

- **Subscription batching** — `Invoke-PaginatedGraphQuery` now batches subscriptions in groups of 200 (Azure Resource Graph limit) instead of passing all at once.
- **O(n²) array elimination** — All `+=` array concatenation patterns replaced with `[System.Collections.Generic.List[object]]` and `.Add()` for O(1) appends.
- **Pre-indexed health data** — Health events and incidents pre-binned into hashtables keyed by `region|category|yyyy-MM` for O(1) lookups during SLA matrix building, replacing per-cell iteration over all events.
- **Month boundaries pre-computed** — Start/end dates and total minutes for each month calculated once upfront instead of per-cell.
- **Weighted health downtime** — Changed from flat 30 minutes per unhealthy event to weighted fraction: `min(1, unhealthy/total) × 30 minutes`, correctly scaling impact by fleet size.
- **Incident window cap** — Each merged incident window is capped at 4 hours (240 minutes) since tracking windows represent investigation periods, not continuous outage.
- **ServiceIssue-only filtering** — Only `ServiceIssue` events count toward SLA downtime. `PlannedMaintenance`, `HealthAdvisory`, and `SecurityAdvisory` events are excluded from availability calculations (still shown in Tabs 2 and 3).
- **Active incident handling** — Incidents with `Active` status and no end time now use `lastUpdateTime` as a proxy end, instead of spanning to the end of each month.

### Fixed

- **N/A for all SLAs** — Count query returning `$null` was cast to `[int]0`, triggering `if ($totalCount -eq 0) { continue }` which skipped all data fetching. Removed the premature skip; count is now informational only.
- **0% SLA from overlapping incidents** — Multiple overlapping incident windows were double-counted. Fixed with interval merge algorithm (sort by start, extend overlapping windows).
- **0% SLA from non-ServiceIssue events** — Planned maintenance and health advisories were incorrectly counted as downtime. Now filtered to `ServiceIssue` only.
- **0% SLA from Active incidents** — Active incidents with no end time used month-end as fallback, creating full-month downtime across multiple months.
- **0% SLA from excessive health events** — Flat `unhealthyCount × 30min` with thousands of events exceeded total month minutes. Fixed with weighted fraction.
- **0% SLA from long tracking windows** — Single incidents tracked across multiple months (e.g., Jan→Apr) produced full-month downtime in each month. Fixed by the 4-hour cap per merged window.
- **Duplicate rows** — Tenant-scoped queries (`servicehealthresources`, `HealthResources`) returned identical results in each subscription batch. Added deduplication: health data by composite key hashtable, incidents/events by `Sort-Object -Property name -Unique`, regions by `Select-Object -Unique`.
- **Negative duration (-17752462 hours)** — `ImpactMitigationTime` with ticks = 0 converted to `DateTime 0001-01-01`. `Convert-TicksToDateTime` now returns `$null` for ticks ≤ 0 or dates before year 2000.
- **Robust count parsing** — `Invoke-PaginatedGraphQuery` no longer fails silently when the count query returns unexpected formats.

## [1.3.1] — 2026-02-11

### Fixed

- **OutputPath directory handling** — Auto-appends the default filename when `-OutputPath` is a directory instead of a file path.
- **PowerShell 5.1 here-string syntax** — Fixed here-string formatting that caused parse errors on Windows PowerShell 5.1.

### Added

- **Azure Workbook template** — `AzureSLA.workbook.json` added (in testing).

## [1.3.0] — 2026-02-11

### Added

- **Paginated Resource Graph queries** — New `Invoke-PaginatedGraphQuery` helper function fetches results in batches of 1 000, supporting environments with 250 000+ resources. Uses `-Skip` for the first 5 000 rows and `$SkipToken` beyond that. A count query runs first to log the total before fetching begins.
- All five `Search-AzGraph` call sites (`Resolve-Regions`, `Get-ResourceHealthEvents`, `Get-ResourceInventory`, `Get-ResourceAvailability`, `Get-ServiceHealthIncidents`) now route through the paginated helper.

## [1.2.0] — 2026-02-11

### Fixed

- **KQL `title` reserved word** — `title` conflicts with a built-in column in Azure Resource Graph; replaced with `Title = properties.Title` in the `project` statement.
- **`-First 5000` exceeds Search-AzGraph limit** — `Get-ResourceAvailability` passed `-First 5000`, but the maximum is 1000. Reduced to `-First 1000`.
- **`ImpactStartTime` stored as .NET ticks** — `ServiceHealthResources` stores timestamps as .NET ticks, not ISO-8601 datetimes. Added `Convert-TicksToDateTime` helper and moved date filtering to PowerShell.
- **Array concatenation failure** — `($incidents1m + $healthEvents)` threw `op_Addition` error when one side was an empty `PSObject`. Wrapped both sides in `@()` to guarantee array types.

## [1.1.0] — 2026-02-11

### Added

- **Tab 3 — Health Timeline** — New worksheet showing all service health events month by month across the full reporting period, with alternating row shading per month for readability.
- **README updated** with Tab 3 documentation and output examples.

## [1.0.2] — 2026-02-11

### Changed

- **Default to all regions** — When `-Regions` is omitted, the script now auto-discovers every region that contains tracked resources via Resource Graph instead of defaulting to Canada Central/East.
- Dynamic `RegionDisplayNames` lookup built from `Get-AzLocation` at runtime.

## [1.0.1] — 2026-02-11

### Added

- **Multi-subscription support** — New `-SubscriptionIds` parameter. Defaults to all enabled subscriptions; pass one or more IDs to narrow scope.
- All `Search-AzGraph` calls now pass `-Subscription $script:ResolvedSubscriptionIds`.
- Activity Log queries loop per subscription with `Set-AzContext`.

## [1.0.0] — 2026-02-11

### Added

- Initial release.
- **Tab 1 — SLA Overview**: Resource availability aggregated by region, service category (Compute, SQL DB, Web Apps, Storage), and month for the past 12 months.
- **Tab 2 — Incidents & Alerts**: Service Health incidents and Activity Log alerts for the past month.
- Prerequisites check with troubleshooting boxes for authentication and missing modules.
- Conditional formatting: green (≥ 99.99 %), yellow (≥ 99.9 %), red (< 99.9 %).
- README with full documentation, usage examples, and implementation guide.
