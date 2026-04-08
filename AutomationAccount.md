# Running the Azure SLA Report in an Azure Automation Account

This guide walks you through setting up an Azure Automation Account to run the SLA report on a schedule, with the output uploaded to Azure Blob Storage automatically.

---

## Architecture Overview

```
┌─────────────────────┐     ┌───────────────────────┐     ┌──────────────────────┐
│  Azure Automation   │     │   Azure Managed       │     │  Azure Storage       │
│  Account            │────▶│   Identity            │────▶│  Account             │
│                     │     │                       │     │                      │
│  • Runbook (PS 7.2) │     │  Permissions:         │     │  • Blob container    │
│  • Schedule         │     │  • Reader (subs)      │     │    "sla-reports"     │
│  • Modules          │     │  • Blob Data Contrib  │     │  • Report .xlsx      │
└─────────────────────┘     └───────────────────────┘     └──────────────────────┘
```

---

## Prerequisites

- An Azure subscription with **Contributor** access (to create resources)
- Azure CLI or Azure Portal access
- The `Get-AzureSLAReport.ps1` script from this repository

---

## Step 1 — Create a Storage Account

The Storage Account will hold the generated SLA reports.

### Azure Portal

1. Go to **Storage accounts** → **+ Create**
2. Fill in:
   - **Resource group**: Create new or use existing (e.g., `rg-sla-reports`)
   - **Storage account name**: e.g., `stslareports` (must be globally unique)
   - **Region**: Choose your preferred region
   - **Performance**: Standard
   - **Redundancy**: LRS (locally redundant) is sufficient
3. Click **Review + Create** → **Create**
4. After deployment, go to the storage account → **Containers** → **+ Container**
   - **Name**: `sla-reports`
   - **Public access level**: Private (no anonymous access)
   - Click **Create**

### Azure CLI

```bash
# Variables — adjust these
RESOURCE_GROUP="rg-sla-reports"
LOCATION="canadacentral"
STORAGE_ACCOUNT="stslareports"       # must be globally unique, lowercase, no dashes
CONTAINER_NAME="sla-reports"

# Create resource group
az group create --name $RESOURCE_GROUP --location $LOCATION

# Create storage account
az storage account create \
    --name $STORAGE_ACCOUNT \
    --resource-group $RESOURCE_GROUP \
    --location $LOCATION \
    --sku Standard_LRS \
    --kind StorageV2

# Create blob container
az storage container create \
    --name $CONTAINER_NAME \
    --account-name $STORAGE_ACCOUNT \
    --auth-mode login
```

> **Note the container URL** — you'll need it later:
> `https://<storage-account-name>.blob.core.windows.net/sla-reports`

---

## Step 2 — Create an Automation Account

### Azure Portal

1. Go to **Automation Accounts** → **+ Create**
2. Fill in:
   - **Resource group**: Same as above (e.g., `rg-sla-reports`)
   - **Name**: e.g., `aa-sla-reports`
   - **Region**: Same region as your storage account
3. Under **Identity**: Ensure **System assigned managed identity** is **On**
4. Click **Review + Create** → **Create**

### Azure CLI

```bash
AUTOMATION_ACCOUNT="aa-sla-reports"

# Create automation account
az automation account create \
    --name $AUTOMATION_ACCOUNT \
    --resource-group $RESOURCE_GROUP \
    --location $LOCATION

# Enable system-assigned managed identity
az automation account identity assign \
    --name $AUTOMATION_ACCOUNT \
    --resource-group $RESOURCE_GROUP \
    --identity-type SystemAssigned
```

**Save the Managed Identity Object ID** — you'll need it in the next step:

```bash
# Get the managed identity principal ID
az automation account show \
    --name $AUTOMATION_ACCOUNT \
    --resource-group $RESOURCE_GROUP \
    --query identity.principalId \
    --output tsv
```

---

## Step 3 — Assign Permissions to the Managed Identity

The Automation Account's managed identity needs two sets of permissions:

| Permission | Scope | Purpose |
|-----------|-------|---------|
| **Reader** | Each subscription to report on | Read resources, health data, service health, activity logs |
| **Storage Blob Data Contributor** | Storage account | Upload the `.xlsx` report to blob storage |

### Option A: All subscriptions (recommended for full coverage)

Assign **Reader** at the **Management Group** level (root or a specific group) to cover all current and future subscriptions.

### Option B: Specific subscriptions

Assign **Reader** on each subscription individually.

### Azure Portal

#### Assign Reader on subscriptions:

1. Go to **Subscriptions** → Select a subscription → **Access control (IAM)**
2. Click **+ Add** → **Add role assignment**
3. **Role**: Reader
4. **Assign access to**: Managed identity
5. **Members**: Select the Automation Account's managed identity (`aa-sla-reports`)
6. Click **Review + assign**
7. Repeat for each subscription you want to include in the report

#### Assign Storage Blob Data Contributor:

1. Go to **Storage accounts** → Select your storage account → **Access control (IAM)**
2. Click **+ Add** → **Add role assignment**
3. **Role**: Storage Blob Data Contributor
4. **Assign access to**: Managed identity
5. **Members**: Select the Automation Account's managed identity (`aa-sla-reports`)
6. Click **Review + assign**

### Azure CLI

```bash
PRINCIPAL_ID=$(az automation account show \
    --name $AUTOMATION_ACCOUNT \
    --resource-group $RESOURCE_GROUP \
    --query identity.principalId \
    --output tsv)

# Assign Reader on a subscription
SUBSCRIPTION_ID="your-subscription-id-here"
az role assignment create \
    --assignee-object-id $PRINCIPAL_ID \
    --assignee-principal-type ServicePrincipal \
    --role "Reader" \
    --scope "/subscriptions/$SUBSCRIPTION_ID"

# Assign Storage Blob Data Contributor on the storage account
STORAGE_ACCOUNT_ID=$(az storage account show \
    --name $STORAGE_ACCOUNT \
    --resource-group $RESOURCE_GROUP \
    --query id --output tsv)

az role assignment create \
    --assignee-object-id $PRINCIPAL_ID \
    --assignee-principal-type ServicePrincipal \
    --role "Storage Blob Data Contributor" \
    --scope "$STORAGE_ACCOUNT_ID"
```

> **Wait 5–10 minutes** for role assignments to propagate before running the runbook.

---

## Step 4 — Import PowerShell Modules (Runtime 7.2)

The runbook runs on **PowerShell 7.2** runtime. You need to import the required modules into the Automation Account.

### Required Modules

| Module | Recommended Version | Notes |
|--------|---------------------|-------|
| `Az.Accounts` | **Use the built-in global** (2.15.0) | **Do NOT import a newer version** — see warning below |
| `Az.ResourceGraph` | **0.13.0** | Import as custom module (Runtime 7.2) |
| `Az.Monitor` | **Use the built-in global** (5.0.0) | Pre-installed in the Automation sandbox |
| `Az.Resources` | **Use the built-in global** (6.13.0) | Pre-installed in the Automation sandbox |
| `ImportExcel` | **7.8.10** (or latest) | Import as custom module (Runtime 7.2) |

> **⚠️ Critical: Do NOT import Az.Accounts 3.0.0 or newer as a custom module.**
>
> Newer versions of `Az.Accounts` (≥ 3.0.0) use an assembly load context (`AzAssemblyLoadContextInitializer`) that is **incompatible with the Azure Automation sandbox**, causing the runbook to fail immediately with:
> ```
> Unable to find type [Microsoft.Azure.PowerShell.AuthenticationAssemblyLoadContext.AzAssemblyLoadContextInitializer]
> ```
> The built-in global `Az.Accounts 2.15.0` works correctly. Similarly, keep `Az.Monitor` and `Az.Resources` at their built-in global versions.
>
> **Az.ResourceGraph** must be imported as a custom module because it’s not pre-installed. Use version **0.13.0** which requires `Az.Accounts ≥ 2.9.1` (compatible with the built-in 2.15.0). Do NOT use v1.0.0+ which requires `Az.Accounts ≥ 4.2.0`.

> **Only two custom modules are needed**: `Az.ResourceGraph` (0.13.0) and `ImportExcel` (7.8.10).

### Azure Portal

1. Go to your Automation Account → **Modules** (under Shared Resources)
2. Click **+ Add a module**
3. **Module source**: Browse gallery
4. Search for `Az.ResourceGraph`
5. Select it → Choose **Runtime version**: `7.2`
6. Click **Import**
7. **Wait for it to finish** (status changes from "Importing" to "Available") — this can take 5–15 minutes
8. Repeat for `ImportExcel`

> **That’s it.** You only need to import `Az.ResourceGraph` and `ImportExcel`. The other Az modules (`Az.Accounts`, `Az.Monitor`, `Az.Resources`) are already available as built-in globals.

### PowerShell (Az Module)

```powershell
$automationAccount = "aa-sla-reports"
$resourceGroup     = "rg-sla-reports"

# Import Az.ResourceGraph 0.13.0 (compatible with built-in Az.Accounts 2.15.0)
New-AzAutomationModule -AutomationAccountName $automationAccount `
    -ResourceGroupName $resourceGroup `
    -Name "Az.ResourceGraph" `
    -ContentLinkUri "https://www.powershellgallery.com/api/v2/package/Az.ResourceGraph/0.13.0" `
    -RuntimeVersion "7.2"

# Import ImportExcel
New-AzAutomationModule -AutomationAccountName $automationAccount `
    -ResourceGroupName $resourceGroup `
    -Name "ImportExcel" `
    -ContentLinkUri "https://www.powershellgallery.com/api/v2/package/ImportExcel" `
    -RuntimeVersion "7.2"

Write-Host "Modules queued for import. Check the portal for status."
```

---

## Step 5 — Create the Runbook

### Azure Portal

1. Go to your Automation Account → **Runbooks** → **+ Create a runbook**
2. Fill in:
   - **Name**: `Get-AzureSLAReport`
   - **Runbook type**: PowerShell
   - **Runtime version**: 7.2
   - **Description**: Generates monthly Azure SLA & Service Health Excel report
3. Click **Create**
4. In the editor, paste the **entire contents** of `Get-AzureSLAReport.ps1`
5. Click **Save** → **Publish**

### Azure CLI

```bash
# Upload and create the runbook from the local script file
az automation runbook create \
    --automation-account-name $AUTOMATION_ACCOUNT \
    --resource-group $RESOURCE_GROUP \
    --name "Get-AzureSLAReport" \
    --type "PowerShell" \
    --runtime-version "7.2" \
    --description "Generates monthly Azure SLA & Service Health Excel report"

# Upload the script content
az automation runbook replace-content \
    --automation-account-name $AUTOMATION_ACCOUNT \
    --resource-group $RESOURCE_GROUP \
    --name "Get-AzureSLAReport" \
    --content @Get-AzureSLAReport.ps1

# Publish the runbook
az automation runbook publish \
    --automation-account-name $AUTOMATION_ACCOUNT \
    --resource-group $RESOURCE_GROUP \
    --name "Get-AzureSLAReport"
```

---

## Step 6 — Configure Runbook Parameters

When starting the runbook (manually or via schedule), pass these parameters:

| Parameter | Value | Required? |
|-----------|-------|-----------|
| `BlobContainerUrl` | `https://stslareports.blob.core.windows.net/sla-reports` | **Yes** (for automation — otherwise the file is lost when the sandbox exits) |
| `Regions` | Leave empty for all regions, or e.g., `canadacentral,eastus` | No |
| `MonthsBack` | `12` (default) | No |
| `SubscriptionIds` | Leave empty for all, or comma-separated IDs | No |
| `OutputPath` | Leave default (the file is created in the sandbox temp dir and uploaded to blob) | No |

> **Critical**: Always use `-BlobContainerUrl` in automation. The Automation sandbox is ephemeral — the report file is deleted when the job completes. Without blob upload, the report is lost.

### Authentication in Automation

The script automatically detects it's running with a Managed Identity because `Connect-AzAccount` is called if not already authenticated. In an Automation Account, the managed identity handles authentication automatically.

However, you may need to add this to the **beginning** of the runbook (before the script runs) to explicitly use the managed identity:

```powershell
# Connect using the Automation Account's managed identity
Connect-AzAccount -Identity | Out-Null
```

**To add this**: Edit the runbook → add the line **above** the `[CmdletBinding()]` block but **below** the comment header → Save → Publish.

Alternatively, you can configure an **Automation pre-script** or modify the script's `Test-Prerequisites` function to call `Connect-AzAccount -Identity` when it detects it's running in an Automation Account.

---

## Step 7 — Test the Runbook

Before scheduling, run it manually to verify everything works.

### Azure Portal

1. Go to your Automation Account → **Runbooks** → `Get-AzureSLAReport`
2. Click **Start**
3. Fill in parameters:
   - **BlobContainerUrl**: `https://stslareports.blob.core.windows.net/sla-reports`
   - Leave other parameters at defaults
4. Click **OK**
5. Monitor the job:
   - **Output** tab — shows the script's Write-Host messages
   - **Errors** tab — shows any exceptions
   - **All Logs** tab — shows everything

### Expected Output

```
[AUTOMATION] Azure Automation Account detected — configuring output streams...
[START] Azure SLA Report Generator v2.2.2 — 2026-04-08 15:00:00 UTC
[STEP ] Step 1/6: Checking prerequisites...
[  OK  ] Az.Accounts v2.15.0
[  OK  ] Az.ResourceGraph v0.13.0
[  OK  ] Az.Monitor v5.0.0
[  OK  ] Az.Resources v6.13.0
[  OK  ] ImportExcel v7.8.10
[AUTH ] Authenticating with Managed Identity...
[AUTH ] Managed Identity authentication successful
[  OK  ] Found X enabled subscription(s)
...
[STEP ] Step 5/6: Exporting Excel report...
[STEP ] Report saved to: /tmp/AzureSLA_Report_20260408_150000.xlsx
[UPLOAD] Starting blob upload...
[UPLOAD] azcopy not available, using REST API
[UPLOAD] Uploading via REST API with bearer token...
[UPLOAD] SUCCESS via REST API: https://stslareports.blob.core.windows.net/sla-reports/AzureSLA_Report_20260408_150000.xlsx
[DONE ] Report uploaded successfully to blob storage.
[DONE ] Report generated in 60s — X resources, Y incidents, Z subs, N regions
```

### Common Issues

| Issue | Solution |
|-------|----------|
| **Module not found** | Ensure `Az.ResourceGraph` and `ImportExcel` are imported with runtime version 7.2 and status is "Available". Other Az modules use the built-in globals. |
| **Authentication failed** | Verify the managed identity is enabled and has Reader on the target subscriptions |
| **Blob upload failed** | Verify `Storage Blob Data Contributor` is assigned on the storage account (not subscription). The script uses the REST API with a bearer token — no `azcopy` needed. |
| **AzAssemblyLoadContextInitializer error** | You imported a newer Az.Accounts (≥ 3.0.0) as a custom module. Remove it and use the built-in global (2.15.0). |
| **Az.ResourceGraph requires newer Az.Accounts** | You imported Az.ResourceGraph 1.0.0+. Use version 0.13.0 instead (compatible with Az.Accounts 2.15.0). |
| **Timeout** | Default job timeout is 3 hours. For very large environments (700+ subs), consider using `-Regions` or `-SubscriptionIds` to narrow scope |

---

## Step 8 — Schedule the Runbook

### Azure Portal

1. Go to your Automation Account → **Runbooks** → `Get-AzureSLAReport`
2. Click **Link to schedule** → **Schedule** → **+ Add a schedule**
3. Fill in:
   - **Name**: `Monthly SLA Report`
   - **Starts**: First day of next month, e.g., `2026-05-01 08:00`
   - **Recurrence**: Recurring → every **1 Month**
   - **Timezone**: Your preferred timezone
   - **Set expiration**: No (or set a far-future date)
4. Click **Create**
5. Back on the "Link to schedule" page, click **Parameters and run settings**
6. Enter:
   - **BlobContainerUrl**: `https://stslareports.blob.core.windows.net/sla-reports`
7. Click **OK** → **OK**

### Azure CLI

```bash
# Create a schedule (monthly on the 1st at 08:00 UTC)
az automation schedule create \
    --automation-account-name $AUTOMATION_ACCOUNT \
    --resource-group $RESOURCE_GROUP \
    --name "Monthly-SLA-Report" \
    --frequency "Month" \
    --interval 1 \
    --start-time "2026-05-01T08:00:00Z" \
    --time-zone "UTC" \
    --description "Generates SLA report on the 1st of each month"
```

> **Note**: Linking the schedule to the runbook with parameters must be done in the portal or via the `New-AzAutomationScheduledRunbook` PowerShell cmdlet — Azure CLI doesn't support parameter binding for schedule-runbook links.

### PowerShell (link schedule with parameters)

```powershell
$params = @{
    BlobContainerUrl = "https://stslareports.blob.core.windows.net/sla-reports"
}

Register-AzAutomationScheduledRunbook `
    -AutomationAccountName $automationAccount `
    -ResourceGroupName $resourceGroup `
    -RunbookName "Get-AzureSLAReport" `
    -ScheduleName "Monthly-SLA-Report" `
    -Parameters $params `
    -RunOn ""   # empty = run on Azure
```

---

## Summary Checklist

| Step | What to Do | Status |
|------|-----------|--------|
| 1 | Create Storage Account + `sla-reports` container | ☐ |
| 2 | Create Automation Account with System Managed Identity | ☐ |
| 3a | Assign **Reader** to Managed Identity on target subscription(s) | ☐ |
| 3b | Assign **Storage Blob Data Contributor** to Managed Identity on Storage Account | ☐ |
| 4 | Import custom modules: Az.ResourceGraph (0.13.0), ImportExcel (runtime 7.2). Other Az modules use built-in globals. | ☐ |
| 5 | Create Runbook (PowerShell 7.2), paste script, publish | ☐ |
| 6 | Test the runbook manually with `-BlobContainerUrl` | ☐ |
| 7 | Create a monthly schedule and link it to the runbook | ☐ |

---

## Appendix: Upload Method

The script automatically handles blob upload without `azcopy`:

1. **If `azcopy` is available** — uses `azcopy copy` with MSI auth (`AZCOPY_AUTO_LOGIN_TYPE=MSI`)
2. **If `azcopy` is not available** (typical in Automation sandbox) — falls back to the **Azure Storage REST API** using a bearer token from `Get-AzAccessToken -ResourceUrl 'https://storage.azure.com/'`
3. **If a SAS token URL is provided** — uploads via REST API directly with the SAS token in the URL

No additional modules (like `Az.Storage`) are needed. The REST API fallback uses only `Invoke-RestMethod` and `Get-AzAccessToken` (from `Az.Accounts`).

This uses the managed identity's RBAC permissions (same `Storage Blob Data Contributor` role) without needing any extra tools.

---

## Appendix: Cost Estimate

| Resource | Approximate Monthly Cost |
|----------|------------------------|
| Automation Account (free tier) | **$0** — 500 minutes/month included |
| Storage Account (LRS, minimal) | **~$0.02/month** — one small .xlsx file |
| **Total** | **< $0.05/month** |

The free tier of Azure Automation includes 500 job minutes per month. A single SLA report run typically takes 5–30 minutes depending on environment size, well within the free tier.
