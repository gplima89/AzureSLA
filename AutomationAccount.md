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

| Module | Version | Notes |
|--------|---------|-------|
| `Az.Accounts` | Latest | Core authentication — **import first** |
| `Az.ResourceGraph` | Latest | Resource Graph queries |
| `Az.Monitor` | Latest | Activity Log queries |
| `Az.Resources` | Latest | Provider registration |
| `ImportExcel` | Latest | Excel file generation |

> **Important**: `Az.Accounts` must be imported **before** the other `Az.*` modules.

### Azure Portal

1. Go to your Automation Account → **Modules** (under Shared Resources)
2. Click **+ Add a module**
3. **Module source**: Browse gallery
4. Search for `Az.Accounts`
5. Select it → Choose **Runtime version**: `7.2`
6. Click **Import**
7. **Wait for it to finish** (status changes from "Importing" to "Available") — this can take 5–15 minutes
8. Repeat for each remaining module in order:
   - `Az.ResourceGraph`
   - `Az.Monitor`
   - `Az.Resources`
   - `ImportExcel`

### PowerShell (Az Module)

```powershell
$automationAccount = "aa-sla-reports"
$resourceGroup     = "rg-sla-reports"

# Import Az.Accounts first (dependency for others)
New-AzAutomationModule -AutomationAccountName $automationAccount `
    -ResourceGroupName $resourceGroup `
    -Name "Az.Accounts" `
    -ContentLinkUri "https://www.powershellgallery.com/api/v2/package/Az.Accounts" `
    -RuntimeVersion "7.2"

# Wait for Az.Accounts to finish importing before continuing
Write-Host "Waiting for Az.Accounts to import..."
do {
    Start-Sleep -Seconds 30
    $mod = Get-AzAutomationModule -AutomationAccountName $automationAccount `
        -ResourceGroupName $resourceGroup -Name "Az.Accounts" `
        -RuntimeVersion "7.2" -ErrorAction SilentlyContinue
    Write-Host "  Status: $($mod.ProvisioningState)"
} while ($mod.ProvisioningState -ne "Succeeded")

# Import remaining modules
$modules = @("Az.ResourceGraph", "Az.Monitor", "Az.Resources", "ImportExcel")
foreach ($modName in $modules) {
    Write-Host "Importing $modName..."
    New-AzAutomationModule -AutomationAccountName $automationAccount `
        -ResourceGroupName $resourceGroup `
        -Name $modName `
        -ContentLinkUri "https://www.powershellgallery.com/api/v2/package/$modName" `
        -RuntimeVersion "7.2"
    Start-Sleep -Seconds 10
}

Write-Host "All modules queued for import. Check the portal for status."
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
[  OK  ] Az.Accounts v...
[  OK  ] Az.ResourceGraph v...
[  OK  ] Az.Monitor v...
[  OK  ] Az.Resources v...
[  OK  ] ImportExcel v...
[  OK  ] Connected as: (managed identity)
[  OK  ] Found X enabled subscription(s)
...
[  OK  ] Report saved to: /tmp/AzureSLA_Report_20260408_080000.xlsx
[  OK  ] azcopy found: /usr/bin/azcopy
[  OK  ] Blob container accessible
[  OK  ] Report uploaded to: https://stslareports.blob.core.windows.net/sla-reports/AzureSLA_Report_20260408_080000.xlsx
```

### Common Issues

| Issue | Solution |
|-------|----------|
| **Module not found** | Ensure all modules are imported with runtime version 7.2 and status is "Available" |
| **Authentication failed** | Verify the managed identity is enabled and has Reader on the target subscriptions |
| **Blob upload failed** | Verify `Storage Blob Data Contributor` is assigned on the storage account (not subscription) |
| **azcopy not found** | azcopy is pre-installed in the Automation sandbox. If missing, use the Az.Storage module as an alternative (see Appendix below) |
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
| 4 | Import modules: Az.Accounts → Az.ResourceGraph → Az.Monitor → Az.Resources → ImportExcel (all runtime 7.2) | ☐ |
| 5 | Create Runbook (PowerShell 7.2), paste script, publish | ☐ |
| 6 | Test the runbook manually with `-BlobContainerUrl` | ☐ |
| 7 | Create a monthly schedule and link it to the runbook | ☐ |

---

## Appendix: Alternative Blob Upload Without azcopy

If azcopy is not available in the Automation sandbox, you can upload using the `Az.Storage` module instead. Add this module to Step 4, then replace the `-BlobContainerUrl` parameter with a manual upload after the script runs:

```powershell
# Alternative: upload using Az.Storage (add after Export-SLAReport in the runbook)
$storageAccount = Get-AzStorageAccount -ResourceGroupName "rg-sla-reports" -Name "stslareports"
$ctx = $storageAccount.Context
Set-AzStorageBlobContent -File $OutputPath `
    -Container "sla-reports" `
    -Blob (Split-Path $OutputPath -Leaf) `
    -Context $ctx `
    -Force
```

This uses the managed identity's RBAC permissions (same `Storage Blob Data Contributor` role) without needing azcopy.

---

## Appendix: Cost Estimate

| Resource | Approximate Monthly Cost |
|----------|------------------------|
| Automation Account (free tier) | **$0** — 500 minutes/month included |
| Storage Account (LRS, minimal) | **~$0.02/month** — one small .xlsx file |
| **Total** | **< $0.05/month** |

The free tier of Azure Automation includes 500 job minutes per month. A single SLA report run typically takes 5–30 minutes depending on environment size, well within the free tier.
