# Changelog

All notable changes to the Azure SLA & Service Health Report Generator are documented in this file.

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
