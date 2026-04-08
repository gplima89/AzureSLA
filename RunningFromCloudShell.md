# Running the Azure SLA Report from Azure Cloud Shell

This guide walks you through running the SLA report directly from [Azure Cloud Shell](https://shell.azure.com) — no local installation required.

---

## Why Cloud Shell?

Azure Cloud Shell is a browser-based terminal that comes **pre-authenticated** with your Azure account and **pre-installed** with PowerShell 7, the Az modules, `azcopy`, and `git`. You only need to install one extra module (`ImportExcel`) and clone the repository.

---

## Prerequisites

| Requirement | Status in Cloud Shell |
|------------|----------------------|
| PowerShell 7+ | ✅ Pre-installed |
| Az.Accounts | ✅ Pre-installed |
| Az.ResourceGraph | ✅ Pre-installed |
| Az.Monitor | ✅ Pre-installed |
| Az.Resources | ✅ Pre-installed |
| ImportExcel | ❌ Must be installed (one-time) |
| Git | ✅ Pre-installed |
| Azure authentication | ✅ Automatic — you're already signed in |
| azcopy | ✅ Pre-installed (for blob upload) |

### Azure Access Requirements

- **Reader** role on the subscription(s) you want to report on
- **Storage Blob Data Contributor** on the storage account *(only if using `-BlobContainerUrl`)*

---

## Step 1 — Open Azure Cloud Shell

1. Go to [https://shell.azure.com](https://shell.azure.com) or click the **Cloud Shell** icon (` >_ `) in the Azure Portal toolbar
2. If prompted, select **PowerShell** as the shell type
3. If this is your first time, Cloud Shell will ask you to create a storage account for your home directory — follow the prompts

---

## Step 2 — Install the ImportExcel Module

This only needs to be done **once** — the module persists in your Cloud Shell home directory across sessions.

```powershell
Install-Module ImportExcel -Force
```

---

## Step 3 — Clone the Repository

```powershell
git clone https://github.com/gplima89/AzureSLA.git
```

Then navigate into the folder:

```powershell
cd AzureSLA
```

---

## Step 4 — Run the Script

### Basic run (all subscriptions, all regions, 12 months)

```powershell
./Get-AzureSLAReport.ps1
```

### Common options

```powershell
# Specific regions only
./Get-AzureSLAReport.ps1 -Regions @("canadacentral", "eastus")

# Specific subscriptions only
./Get-AzureSLAReport.ps1 -SubscriptionIds @("aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee")

# Shorter lookback period
./Get-AzureSLAReport.ps1 -MonthsBack 6

# Upload report to Azure Blob Storage (plain URL — Cloud Shell is already authenticated)
./Get-AzureSLAReport.ps1 -BlobContainerUrl "https://mystorageaccount.blob.core.windows.net/sla-reports"

# Upload with SAS token
./Get-AzureSLAReport.ps1 -BlobContainerUrl "https://mystorageaccount.blob.core.windows.net/sla-reports?sv=2022-11-02&ss=b&srt=o&sp=wc&se=2026-12-31T00:00:00Z&sig=..."

# Combine parameters
./Get-AzureSLAReport.ps1 `
    -Regions @("canadacentral", "eastus") `
    -MonthsBack 6 `
    -BlobContainerUrl "https://mystorageaccount.blob.core.windows.net/sla-reports"
```

---

## Step 5 — Download the Report

Cloud Shell cannot open Excel files directly. You have two options:

### Option A: Upload to Blob Storage (recommended)

Use the `-BlobContainerUrl` parameter as shown above. The report is uploaded automatically and you can download it from the Azure Portal (Storage account → Containers → your container).

### Option B: Download from Cloud Shell

1. After the script completes, note the output path (e.g., `/home/youruser/AzureSLA/AzureSLA_Report_20260408_150000.xlsx`)
2. Click the **Upload/Download** button in the Cloud Shell toolbar (↑↓ icon)
3. Select **Download**
4. Enter the file path and click **Download**

> **Note**: Cloud Shell's download feature has a **1 GB file size limit**. SLA reports are typically small (< 1 MB), so this should not be an issue.

---

## Updating the Script

To pull the latest version from GitHub:

```powershell
cd ~/AzureSLA
git pull
```

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| **"Module ImportExcel not found"** | Run `Install-Module ImportExcel -Force` |
| **"Not connected to Azure"** | Cloud Shell should auto-authenticate. If not, run `Connect-AzAccount` |
| **Script is slow** | Cloud Shell uses PowerShell 7+ with parallel API calls. For very large environments (700+ subs), narrow scope with `-Regions` or `-SubscriptionIds`. |
| **"Permission denied" running the script** | Run `chmod +x ./Get-AzureSLAReport.ps1` or use `pwsh ./Get-AzureSLAReport.ps1` |
| **Cloud Shell session timed out** | Sessions time out after 20 minutes of inactivity. Re-open Cloud Shell and re-run. The Az modules and ImportExcel persist across sessions. |
| **File not found after session restart** | Files in `~/` (your home directory) persist. Files in `/tmp/` do not. The cloned repo stays in `~/AzureSLA/`. |
| **Blob upload fails** | Ensure your account has **Storage Blob Data Contributor** on the storage account. Cloud Shell's `azcopy` uses your Azure CLI credentials automatically. |
