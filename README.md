# Azure SLA & Service Health Report Generator

A PowerShell script that generates an Excel report showing how available your Azure resources have been over the past 12 months, along with any service health incidents that affected your environment.

**In simple terms**: this tool looks at your Azure resources (virtual machines, databases, web apps, and storage accounts) and tells you how often they were "up" or "down" — broken down by region, service type, and month.

---

## ⚠️ Important Disclaimer

> **This report is an estimation tool, not an official SLA measurement.**
>
> - The availability percentages shown are **approximations** based on Azure Resource Health signals and Service Health incident tracking data. They are **not** the same as Microsoft's contractual SLA metrics.
> - Official Azure SLAs are defined in the [Service Level Agreements for Online Services](https://www.microsoft.com/licensing/docs/view/Service-Level-Agreements-SLA-for-Online-Services). To file an SLA credit claim, use the Azure portal's official process.
> - **Always review and validate** the results against your own monitoring data (Azure Monitor, Application Insights, third-party tools, etc.) before presenting them in reports or making decisions based on them.
> - This tool is intended as a **supplementary data source** for operational reviews, governance reporting, and trend analysis — not as the single source of truth for availability.
> - The author(s) and Microsoft are not responsible for decisions made based on this report's output.

---

## What Does This Tool Do?

If you manage Azure resources — even if you're not deeply technical — here's what this tool does in plain language:

1. **Connects to your Azure account** (with your permission) and looks at all the subscriptions you have access to.
2. **Finds your resources** — it looks for Virtual Machines, SQL Databases, Web Apps, and Storage Accounts across all your Azure regions.
3. **Checks their health history** — Azure keeps track of when resources were "Available", "Unavailable", or "Degraded". The script reads this history.
4. **Checks for known incidents** — Azure publicly tracks service outages ("Service Issues"). The script finds which ones affected your regions and services.
5. **Calculates an estimated availability percentage** for each region, service type, and month — then produces a colour-coded Excel spreadsheet you can share with your team or management.

The script is **read-only** — it does not create, modify, or delete any Azure resources. It only reads data.

---

## What's in the Report?

The output is an Excel file (`.xlsx`) with three tabs:

### Tab 1 — SLA Overview

A table showing estimated availability for each combination of **region** and **service category**, broken down by month.

| Region | Service | Resource Count | May 2025 | Jun 2025 | ... | Apr 2026 |
|--------|---------|---------------|----------|----------|-----|----------|
| Canada Central | Compute | 12 | 99.9987% | 100.0000% | ... | 99.9500% |
| Canada Central | SQL DB | 5 | 99.9900% | 99.9956% | ... | 100.0000% |
| East US | Web Apps | 42 | 100.0000% | 100.0000% | ... | 99.9800% |
| East US | Storage | 18 | 100.0000% | 100.0000% | ... | 100.0000% |

**Colour coding:**
| Colour | Meaning |
|--------|---------|
| 🟢 Green | ≥ 99.99% — Excellent availability |
| 🟡 Yellow | ≥ 99.90% — Minor degradation detected |
| 🔴 Red | < 99.90% — Significant availability impact |
| *N/A (grey)* | No resources of that type exist in that region |

### Tab 2 — Incidents & Alerts

A list of service health incidents and activity log alerts from the **past month** that affected your environment.

| Column | What It Means |
|--------|--------------|
| Source | Where the alert came from (Service Health or Activity Log) |
| Type | The kind of event (ServiceIssue, HealthAdvisory, etc.) |
| Status | Whether the incident is Active, Resolved, etc. |
| Title | A short description of what happened |
| Impact Start/End | When the incident window began and ended (UTC) |
| Duration | How long the tracking window lasted |
| Affected Services | Which Azure services were impacted |
| Affected Regions | Which Azure regions were impacted |
| Summary | A brief description of the issue |

### Tab 3 — Health Timeline

The same type of information as Tab 2, but covering the **full reporting period** (12 months by default) and organized by month. Useful for spotting patterns or recurring issues.

---

## How Are the Metrics Calculated?

Understanding how the numbers are produced helps you interpret them correctly.

### Data Sources

The script pulls from three Azure data sources:

| Source | What It Provides | Scope |
|--------|-----------------|-------|
| **HealthResources** (Resource Graph) | Per-resource availability status changes (Available, Unavailable, Degraded) | Current status per resource |
| **ServiceHealthResources** (Resource Graph) | Service-level incidents — outages reported by Azure | Tenant-wide events filtered to your regions |
| **Activity Log** (REST API) | Resource Health and Service Health events from the activity log | Per-subscription |

### Availability Calculation

For each **region + service category + month** cell, the script calculates:

```
Availability % = ((Total minutes in month − Estimated downtime minutes) / Total minutes in month) × 100
```

**Estimated downtime** comes from two components:

1. **Resource health events** — If some of your resources in that region+category reported "Unavailable" or "Degraded" during that month, the script estimates the impact as:
   ```
   (number of unhealthy events / total resource count) × 30 minutes
   ```
   This treats each event as a point-in-time status, weighted by what fraction of your fleet was affected.

2. **Service health incidents** — If Azure reported a `ServiceIssue` affecting that region+service during that month, each distinct outage window contributes up to **4 hours** of downtime (capped). This cap exists because incident tracking windows represent the investigation period, not continuous outage — the actual downtime is typically much shorter.

   When multiple incidents overlap in the same month, their time windows are **merged** (not stacked), so the same time period is never counted twice.

### What "N/A" Means

If a cell shows **N/A**, it means you have **zero resources** of that service type in that region. No availability calculation is possible.

### What to Watch For

- **100.0000%** — No health events or incidents were recorded. This is the most common result and typically means the service was fully available.
- **99.90%–99.99%** — Minor degradation was detected. Review the Incidents tab for details.
- **Below 99.90%** — Potentially significant. Cross-reference with:
  - The **Incidents & Alerts** tab for specific incidents
  - Your own monitoring data (Azure Monitor, Application Insights)
  - Your team's incident reports for that period
- **Numbers that seem too low** — The script may have picked up planned maintenance or health advisories. Only `ServiceIssue` events count toward the SLA calculation, but review the Health Timeline tab if something looks off.

---

## Requirements

### What You Need on Your Computer

| Requirement | Details |
|------------|---------|
| **PowerShell 7+** | Recommended. Download from [https://aka.ms/powershell](https://aka.ms/powershell). Windows PowerShell 5.1 also works but parallel API calls (faster for large environments) require PowerShell 7+. |
| **Internet access** | The script connects to Azure APIs — your machine must be able to reach `https://management.azure.com` and `https://login.microsoftonline.com`. |
| **No Excel needed** | The report is generated using the `ImportExcel` module, which creates `.xlsx` files without requiring Microsoft Excel to be installed. You'll need Excel (or a compatible viewer) only to **open** the report. |

### PowerShell Modules

These are installed once and reused across runs:

| Module | What It Does | Install Command |
|--------|-------------|-----------------|
| `Az.Accounts` | Handles Azure authentication (login) | `Install-Module Az -Scope CurrentUser -Force` |
| `Az.ResourceGraph` | Queries Azure resources and health data | *(included in the Az module)* |
| `Az.Monitor` | Reads Azure Activity Logs | *(included in the Az module)* |
| `Az.Resources` | Checks provider registration | *(included in the Az module)* |
| `ImportExcel` | Creates the Excel report file | `Install-Module ImportExcel -Scope CurrentUser -Force` |

**Quick install (copy-paste this once):**
```powershell
Install-Module Az -Scope CurrentUser -Force
Install-Module ImportExcel -Scope CurrentUser -Force
```

### Azure Access Requirements

| Requirement | Details |
|------------|---------|
| **Azure account** | You need an Azure account (work, school, or personal) that can sign into the Azure portal. |
| **Minimum role** | **Reader** on the subscription(s) you want to report on. This is the lowest-privilege role — it can only view resources, not change them. Ask your Azure administrator to assign this if you don't have it. |
| **Resource Health provider** | The `Microsoft.ResourceHealth` provider must be registered on at least one subscription. The script tries to register it automatically, but if it fails, ask your admin to run: `Register-AzResourceProvider -ProviderNamespace Microsoft.ResourceHealth` |

**How to check your access**: Open the [Azure Portal](https://portal.azure.com), go to **Subscriptions**, and verify you can see the subscriptions you want to report on. If you can see them, you likely have at least Reader access.

---

## How to Run

### Step 1 — Install Modules (One Time Only)

Open PowerShell and run:

```powershell
Install-Module Az -Scope CurrentUser -Force
Install-Module ImportExcel -Scope CurrentUser -Force
```

This downloads the required modules. You only need to do this once per computer.

### Step 2 — Connect to Azure

```powershell
Connect-AzAccount
```

A browser window will open asking you to sign in with your Azure credentials. After signing in, return to PowerShell.

> **Tip**: If your organization uses MFA (multi-factor authentication), you may need to specify your tenant:
> ```powershell
> Connect-AzAccount -TenantId "your-tenant-id"
> ```

### Step 3 — Run the Script

Navigate to the folder where the script is saved, then run:

```powershell
.\Get-AzureSLAReport.ps1
```

The script will:
1. Check that all required modules are installed
2. Verify your Azure connection
3. Discover all your subscriptions and regions
4. Query resource health and incident data
5. Build the SLA matrix
6. Generate and open the Excel report

**Typical run time**: 2–10 minutes for small environments (1–10 subscriptions). Larger environments (100+ subscriptions) may take 15–30 minutes.

### Step 4 — Review the Report

The Excel file opens automatically. Start with the **SLA Overview** tab to see the big picture, then drill into **Incidents & Alerts** for details on any months with reduced availability.

### Advanced: Custom Parameters

```powershell
# Report on specific regions only
.\Get-AzureSLAReport.ps1 -Regions @("canadacentral", "eastus", "westus2")

# Report on specific subscriptions only
.\Get-AzureSLAReport.ps1 -SubscriptionIds @("aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee")

# Change the lookback period to 6 months
.\Get-AzureSLAReport.ps1 -MonthsBack 6

# Save the report to a specific location
.\Get-AzureSLAReport.ps1 -OutputPath "C:\Reports\MySLAReport.xlsx"

# Combine parameters
.\Get-AzureSLAReport.ps1 `
    -Regions @("canadacentral", "canadaeast") `
    -SubscriptionIds @("aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee") `
    -MonthsBack 6 `
    -OutputPath "C:\Reports\MySLAReport.xlsx"
```

### Parameters Reference

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `-Regions` | `string[]` | All regions with resources | Azure region names (e.g., `canadacentral`, `eastus`). Leave empty to auto-discover. |
| `-MonthsBack` | `int` | `12` | How many months of history to include |
| `-OutputPath` | `string` | Script folder | Full path for the output `.xlsx` file |
| `-SubscriptionIds` | `string[]` | All enabled subscriptions | Specific subscription IDs to include |

---

## How to Read and Analyse the Results

### Starting Point: SLA Overview Tab

1. **Look for red/yellow cells** — these indicate months where availability dropped below 99.99%.
2. **Check the Resource Count column** — if a region has very few resources (e.g., 1–2), a single health event has a larger proportional impact.
3. **Compare across months** — a single bad month surrounded by green months likely indicates a one-time incident. A pattern of yellow/red months suggests a recurring issue worth investigating.

### Digging Deeper: Incidents & Alerts Tab

1. **Filter by region or service** — use Excel's filter dropdowns to focus on a specific region or service that showed reduced availability.
2. **Check the Duration column** — long durations may indicate extended outages, but remember that "Duration" reflects the incident tracking window, not necessarily continuous downtime.
3. **Read the Summary** — this gives context on what happened. Many incidents are brief and limited in scope.

### Common Patterns

| Pattern | Likely Cause | Action |
|---------|-------------|--------|
| All months green (100%) | No issues detected | No action needed |
| One month yellow, rest green | Isolated incident | Review the incident in Tab 2 for details |
| Multiple months slightly below 100% | Recurring minor health events | May warrant investigation — check resource health in the Azure Portal |
| N/A across all months | No resources of that type in the region | Expected — you can ignore these rows |

### Sharing the Report

The Excel file is self-contained and can be shared via email or file shares. The colour coding and formatting are preserved. No Azure access is needed to view the report.

---

## Service Categories Tracked

| Category | Azure Resource Types Included |
|----------|------------------------------|
| **Compute** | Virtual Machines, Virtual Machine Scale Sets |
| **SQL DB** | SQL Databases, SQL Servers, SQL Managed Instances |
| **Web Apps** | App Services (Web Apps), App Service Plans |
| **Storage** | Storage Accounts |

Other resource types (e.g., Kubernetes, Cosmos DB, Redis Cache) are not currently tracked. See the "Extending" section below if you need to add more.

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| **"Not connected to Azure"** | Run `Connect-AzAccount` before running the script |
| **MFA prompt not appearing** | Use `Connect-AzAccount -TenantId "your-tenant-id"` |
| **"No enabled subscriptions found"** | Your account doesn't have Reader access to any subscription. Contact your Azure admin. |
| **"Module not installed"** | Run `Install-Module Az -Scope CurrentUser -Force` and `Install-Module ImportExcel -Scope CurrentUser -Force` |
| **Script is slow (>30 min)** | Use PowerShell 7+ for parallel API calls. Consider narrowing scope with `-Regions` or `-SubscriptionIds`. |
| **`[DIAG] LOW SLA` messages** | These are informational — they show the breakdown of what contributed to low availability for a specific cell. Use them to understand the data, not as an error. |
| **Report shows N/A everywhere** | You may not have any of the tracked resource types, or the script couldn't access your subscriptions. Check the console output for warnings. |
| **Behind a corporate proxy** | Add this before running: `[System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials` |
| **Stale authentication tokens** | Run `Clear-AzContext -Force` then `Connect-AzAccount` again |

---

## Extending the Script

### Adding New Service Categories

To track additional Azure resource types:

1. In `Get-ResourceInventory`, add the resource type to the `where type in~(...)` clause and the `case` statement.
2. In `Get-ResourceAvailability`, mirror the same changes.
3. In `Build-SLAMatrix`, add the service name mapping to `$serviceTypeMap`.

### Scheduling Automated Runs

**Windows Task Scheduler:**
```powershell
$action  = New-ScheduledTaskAction -Execute "pwsh.exe" -Argument "-NoProfile -File `"C:\Scripts\Get-AzureSLAReport.ps1`""
$trigger = New-ScheduledTaskTrigger -Monthly -At "08:00" -DaysOfMonth 1
Register-ScheduledTask -TaskName "Azure SLA Report" -Action $action -Trigger $trigger
```

> For unattended execution, authenticate with a service principal:
> ```powershell
> Connect-AzAccount -ServicePrincipal -ApplicationId <AppId> -CertificateThumbprint <Thumbprint> -TenantId <TenantId>
> ```

---

## License

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
