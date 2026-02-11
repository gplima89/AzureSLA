# Azure SLA & Service Health Report Generator

Automated PowerShell script that generates an Excel report tracking Azure service availability and health incidents across your Azure environment.

> **Note:** This report summarizes **observed availability and incidents** for your resources using Resource Health / Service Health signals. It does not replace [Microsoft's official Service Level Agreements](https://www.microsoft.com/licensing/docs/view/Service-Level-Agreements-SLA-for-Online-Services).

---

## Overview

This tool queries Azure Resource Health, Service Health, and Activity Log APIs to produce a three-tab Excel workbook:

| Tab | Content |
|-----|-------|
| **SLA Overview** | Resource availability % aggregated by region, service category, and month for the past 12 months |
| **Incidents & Alerts** | Service Health incidents and Activity Log alerts from the past month affecting your environment |
| **Health Timeline** | All service health events listed month by month over the full reporting period (12 months by default) |
> **Multi-subscription**: By default the script queries **all enabled subscriptions** your account has access to. Use `-SubscriptionIds` to narrow the scope to specific subscriptions.
### Service Categories Tracked

| Category | Azure Resource Types |
|----------|---------------------|
| **Compute** | Virtual Machines, VM Scale Sets |
| **SQL DB** | SQL Databases, SQL Servers, Managed Instances |
| **Web Apps** | App Services, App Service Plans |
| **Storage** | Storage Accounts |

### SLA Color Coding

| Color | Availability |
|-------|-------------|
| 🟢 Green | ≥ 99.99% |
| 🟡 Yellow | ≥ 99.90% |
| 🔴 Red | < 99.90% |

---

## Prerequisites

### PowerShell Modules

| Module | Purpose |
|--------|---------|
| `Az.Accounts` | Azure authentication |
| `Az.ResourceGraph` | Resource & health queries |
| `Az.Monitor` | Activity Log access |
| `Az.Resources` | Provider registration |
| `ImportExcel` | Excel file generation (no Excel installation required) |

### Install Modules

```powershell
# Install the Az module (includes Az.Accounts, Az.ResourceGraph, Az.Monitor, Az.Resources)
Install-Module Az -Scope CurrentUser -Force

# Install the ImportExcel module
Install-Module ImportExcel -Scope CurrentUser -Force
```

### Azure Permissions

- **Minimum role**: `Reader` on the target subscription(s)
- The `Microsoft.ResourceHealth` provider must be registered (the script will attempt auto-registration)
- For multi-subscription, the account must have Reader on each subscription to be queried

---

## How to Run

### 1. Connect to Azure

```powershell
Connect-AzAccount
```

> The script will attempt interactive login automatically if you're not connected.

### 2. Run with Defaults

```powershell
.\Get-AzureSLAReport.ps1
```

This will:
- Query **all subscriptions** your account can access
- Auto-discover **all regions** that contain tracked resources (Compute, SQL DB, Web Apps, Storage)
- Cover the past **12 months** for SLA data
- Cover the past **1 month** for incidents
- Save the report to the script directory as `AzureSLA_Report_<timestamp>.xlsx`

### 3. Run with Custom Parameters

```powershell
# Specific regions and time range
.\Get-AzureSLAReport.ps1 `
    -Regions @("canadacentral", "canadaeast") `
    -MonthsBack 6 `
    -OutputPath "C:\Reports\MySLAReport.xlsx"

# Run for specific subscriptions only
.\Get-AzureSLAReport.ps1 `
    -SubscriptionIds @("aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee", "11111111-2222-3333-4444-555555555555")

# Combine: specific regions + specific subscriptions
.\Get-AzureSLAReport.ps1 `
    -Regions @("canadacentral", "eastus") `
    -SubscriptionIds @("aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee")
```

### Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Regions` | `string[]` | `@()` (all regions with resources) | Azure region identifiers. Leave empty to auto-discover all regions containing tracked resources |
| `-MonthsBack` | `int` | `12` | Number of months of historical SLA data |
| `-OutputPath` | `string` | `.\AzureSLA_Report_<timestamp>.xlsx` | Full path for the output Excel file |
| `-SubscriptionIds` | `string[]` | `@()` (all subscriptions) | Specific subscription IDs to query. Leave empty to query ALL enabled subscriptions |

---

## Implementation Guide

### Filtering by Region

By default the script auto-discovers all regions that contain tracked resources. To limit to specific regions:

```powershell
.\Get-AzureSLAReport.ps1 -Regions @("canadacentral", "canadaeast", "eastus", "westus2")
```

Region display names are resolved automatically from Azure — no manual mapping needed.

### Adding New Service Categories

1. In `Get-ResourceInventory`, add the resource type to the query's `where type in~(...)` clause and the `case` statement.
2. In `Get-ResourceAvailability`, mirror the same changes.
3. In `Calculate-MonthlyAvailability`, add the service name mapping to `$serviceTypeMap`.

### Scheduling with Task Scheduler

```powershell
# Create a scheduled task to run monthly
$action  = New-ScheduledTaskAction -Execute "pwsh.exe" -Argument "-NoProfile -File `"C:\Scripts\Get-AzureSLAReport.ps1`""
$trigger = New-ScheduledTaskTrigger -Monthly -At "08:00" -DaysOfMonth 1
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -RunLevel Highest

Register-ScheduledTask -TaskName "Azure SLA Report" -Action $action -Trigger $trigger -Principal $principal
```

> **Note**: For unattended execution, authenticate with a service principal:
> ```powershell
> Connect-AzAccount -ServicePrincipal -ApplicationId <AppId> -CertificateThumbprint <Thumbprint> -TenantId <TenantId>
> ```

### Scheduling with Azure Automation

1. Create an **Azure Automation Account**
2. Import the `Az.Accounts`, `Az.ResourceGraph`, `Az.Monitor`, `Az.Resources`, and `ImportExcel` modules
3. Create a **Runbook** with the script content
4. Configure a **Managed Identity** with Reader access
5. Set a **Schedule** (e.g., first day of each month)

---

## Troubleshooting

The script includes built-in diagnostics. Common issues:

| Issue | Resolution |
|-------|-----------|
| **Not connected to Azure** | Run `Connect-AzAccount` manually |
| **MFA required** | Use `Connect-AzAccount -TenantId <tenant-id>` |
| **Service principal auth** | `Connect-AzAccount -ServicePrincipal -ApplicationId <id> -CertificateThumbprint <thumb> -TenantId <tenant>` |
| **Behind a proxy** | Set `[System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials` |
| **Stale tokens** | Run `Clear-AzContext -Force` then reconnect |
| **Missing modules** | `Install-Module Az -Scope CurrentUser -Force` and `Install-Module ImportExcel -Scope CurrentUser -Force` |
| **ResourceHealth not registered** | `Register-AzResourceProvider -ProviderNamespace Microsoft.ResourceHealth` |
| **No data returned** | Verify you have resources in the target regions and subscription access |

---

## Output Example

### Tab 1 — SLA Overview

| Region | Service | Resource Count | Jan 2026 | Feb 2026 | ... |
|--------|---------|---------------|----------|----------|-----|
| Canada Central | Compute | 12 | 99.9987% | 100.0000% | ... |
| Canada Central | SQL DB | 5 | 99.9900% | 99.9956% | ... |
| Canada Central | Web Apps | 8 | 100.0000% | 100.0000% | ... |
| Canada Central | Storage | 3 | 100.0000% | 100.0000% | ... |
| Canada East | Compute | 4 | 100.0000% | 99.9500% | ... |
| ... | ... | ... | ... | ... | ... |

### Tab 2 — Incidents & Alerts

| Source | Type | Status | Title | Impact Start | Impact End | Duration | Affected Services | Affected Regions | ... |
|--------|------|--------|-------|-------------|-----------|----------|-------------------|-----------------|-----|
| Service Health | ServiceIssue | Resolved | VM connectivity issue | 2026-01-15 08:30 | 2026-01-15 10:45 | 2.25h | Virtual Machines | Canada Central | ... |

### Tab 3 — Health Timeline

Month-by-month breakdown of all service health events over the full reporting period, with alternating row shading per month for visual grouping.

| Month | Month Name | Event Type | Status | Title | Impact Start | Impact End | Duration | Level | Affected Services | Affected Regions | Summary | Tracking ID |
|-------|-----------|------------|--------|-------|-------------|-----------|----------|-------|-------------------|-----------------|---------|-------------|
| 2026-02 | Feb 2026 | ServiceIssue | Resolved | Storage latency in Canada | 2026-02-03 14:00 | 2026-02-03 16:30 | 2.50 | Warning | Storage Accounts | Canada Central | Intermittent latency... | XXXX-YYYY |
| 2026-01 | Jan 2026 | ServiceIssue | Resolved | VM connectivity issue | 2026-01-15 08:30 | 2026-01-15 10:45 | 2.25 | Error | Virtual Machines | Canada Central; Canada East | VMs experienced... | AAAA-BBBB |
| 2025-12 | Dec 2025 | HealthAdvisory | Active | SQL maintenance window | 2025-12-20 02:00 | 2025-12-20 06:00 | 4.00 | Informational | SQL Database | Canada East | Planned maintenance... | CCCC-DDDD |

---

## License

MIT
