# patch-status

Reports the Azure Update Manager patch status of every Linux VM in one or more Azure subscriptions, and flags the ones that need attention. The report is an Excel workbook (default) or a CSV file, plus a short summary in the console.

All Azure data comes from Azure Resource Graph, so a whole subscription takes a handful of queries rather than one API call per VM.

## What it checks

Each VM gets a `status` of **OK** or **Attention**, and an `issues` column that says why.

| Check | Flagged when |
|---|---|
| Power state | The VM isn't running (a deallocated VM can't be assessed or patched) |
| Patch orchestration | `patchMode` isn't `AutomaticByPlatform`, or `bypassPlatformSafetyChecksOnUserSchedule` isn't set (without both, a maintenance schedule doesn't apply) |
| Periodic assessment | `assessmentMode` isn't `AutomaticByPlatform` |
| Schedule | No maintenance configuration assigned, it isn't weekly, or it isn't the one named with `--schedule` |
| Assessment | Older than `--stale-days` on a running VM, never run, or failed |
| Patch run | Last run older than `--max-run-age`, none in the last 30 days, failed, completed with warnings, or some patches failed to install |
| Pending work | A reboot is pending, or security or critical patches are waiting |
| OS support | The OS release is past end of life, or reaches it within `--eol-warn-days` (dates from [endoflife.date](https://endoflife.date)) |

## Requirements

- [uv](https://docs.astral.sh/uv/). It installs Python 3.14 and the dependencies (`azure-identity`, `azure-mgmt-resourcegraph`, `xlsxwriter`) itself.
- An Azure identity with **Reader** on each subscription or resource group being scanned. To show schedule details, it also needs Reader wherever the maintenance configurations live, if that's a different subscription.
- Outbound HTTPS to `endoflife.date` for the OS support check. This is optional: see `--no-eol-check`.

Authentication uses `DefaultAzureCredential`, so any of these work:

- `az login` locally
- The service connection of an `AzureCLI@2` task in Azure DevOps
- `AZURE_TENANT_ID`, `AZURE_CLIENT_ID` and `AZURE_CLIENT_SECRET` environment variables (or in `.env`)

## Quick start

```bash
uv sync
cp patch-status.example.toml patch-status.toml   # then fill in your subscription IDs
az login
uv run patch_status.py
```

That writes `patch-status-YYYYMMDD.xlsx` in the current directory and prints the VMs that need attention.

## Configuration

### What to scan

The script uses the first of these that is set:

1. **Command line:** `-s/--subscription` (repeatable), optionally narrowed with `-g/--resource-group` (repeatable).
2. **Config file:** `[[subscriptions]]` in `patch-status.toml`. The file is found from `--config`, then `$PATCH_STATUS_CONFIG`, then `./patch-status.toml`.
3. **Environment:** `$AZURE_SUBSCRIPTION_ID`, comma-separated for several subscriptions.

The config file wins over the environment on purpose. A stray `AZURE_SUBSCRIPTION_ID` left in a shell can't silently change what an explicit config file says to scan.

Thresholds (`schedule`, `stale_days`, `max_run_age`, `eol_warn_days`, `format`) use the command line first, then the config file, then the built-in defaults.

### Config file (`patch-status.toml`)

```toml
# Defaults for every subscription
schedule = "mc-linux-weekly"  # omit to accept any weekly schedule
format = "xlsx"               # or "csv"
stale_days = 2
max_run_age = 8
eol_warn_days = 180

[[subscriptions]]
name = "Production"           # optional label (default: the subscription's Azure name)
id = "00000000-0000-0000-0000-000000000000"
# resource_groups = ["rg-app-prod", "rg-data-prod"]  # omit to scan the whole subscription
# schedule = "mc-linux-weekly-prod"                   # overrides the top-level schedule

[[subscriptions]]
name = "Non-production"
id = "11111111-1111-1111-1111-111111111111"
```

Unknown keys are an error, so a typo such as `shedule` stops the run instead of being ignored. Subscription IDs aren't secrets, so the file can be committed. `patch-status.example.toml` is a commented template.

### `.env` file

If `./.env` exists (or a file is given with `--env-file`), it's loaded before anything else. Variables already set in the real environment take priority, so pipeline variables override a local `.env`.

```bash
AZURE_SUBSCRIPTION_ID=00000000-0000-0000-0000-000000000000,11111111-1111-1111-1111-111111111111
# PATCH_STATUS_CONFIG=config/prod.toml
# AZURE_TENANT_ID=...
# AZURE_CLIENT_ID=...
# AZURE_CLIENT_SECRET=...
```

`.env` is in `.gitignore`. `.env.example` is the template.

## Usage

```bash
uv run patch_status.py --help
```

| Option | Default | Description |
|---|---|---|
| `-s`, `--subscription ID` | from config or env | Subscription to scan. Repeat for several. |
| `-g`, `--resource-group NAME` | whole subscription | Only scan this resource group. Repeat for several. Applies to every subscription being scanned. |
| `-c`, `--config FILE` | `$PATCH_STATUS_CONFIG`, then `./patch-status.toml` | TOML config file. |
| `--env-file FILE` | `./.env` if it exists | Env file to load first. |
| `--schedule NAME` | any weekly schedule | Maintenance configuration every VM should be assigned to. |
| `--stale-days DAYS` | `2` | Flag running VMs whose last assessment is older than this. |
| `--max-run-age DAYS` | `8` | Flag VMs whose last patch run is older than this. |
| `--eol-warn-days DAYS` | `180` | Flag OS releases reaching end of life within this many days. |
| `--no-eol-check` | off | Skip the endoflife.date lookup. |
| `--format {xlsx,csv}` | `xlsx` | Report format. |
| `-o`, `--output FILE` | `patch-status-YYYYMMDD.<format>` | Where to write the report. |
| `--fail-on-issues` | off | Exit 1 if any VM needs attention. |

Examples:

```bash
# Subscriptions from patch-status.toml or .env
uv run patch_status.py

# Two whole subscriptions, CSV output
uv run patch_status.py -s <sub-id> -s <sub-id> --format csv

# Two resource groups in one subscription
uv run patch_status.py -s <sub-id> -g rg-app-prod -g rg-data-prod

# Require a specific schedule and fail the run if anything is flagged
uv run patch_status.py --schedule mc-linux-weekly --fail-on-issues
```

### Exit codes

| Code | Meaning |
|---|---|
| `0` | Report written. VMs may still need attention unless `--fail-on-issues` is set. |
| `1` | No Linux VMs found, or `--fail-on-issues` is set and at least one VM needs attention. |
| `2` | Bad arguments or config file. |

## The report

**Excel** (default):

- **Patch status** sheet: one filterable table with a frozen header row and name column, and a red or green highlight on `status`.
- **About** sheet: when the report ran, the thresholds used, and a table for each subscription showing its scope, expected schedule, VM count and how many need attention.

**CSV** has the same columns as the main sheet.

Times are shown in `Australia/Sydney`. To change that, edit `LOCAL_TZ` in the script.

| Column | Contents |
|---|---|
| `name`, `subscription`, `resourceGroup` | Where the VM is |
| `status`, `issues` | OK or Attention, and why |
| `powerState` | e.g. `running`, `deallocated` |
| `os` | OS name and version reported by the VM agent, e.g. `ubuntu 24.04`. If the agent isn't reporting it reads `unknown (built from image: ...)`, because the image name goes stale after in-place upgrades. |
| `osSupport` | e.g. `Supported until 2029-05-31`, or `End of life 2025-05-29 (extended support to 2030-04-23)` |
| `patchMode`, `assessmentMode`, `bypassSafetyChecks` | Patch orchestration settings |
| `maintenanceConfig`, `schedule` | Assigned maintenance configuration(s) and their recurrence, start time, time zone and duration |
| `lastAssessment`, `assessmentStatus` | Latest assessment |
| `securityPending`, `criticalPending`, `otherPending`, `rebootPending` | What the latest assessment found waiting |
| `lastPatchRun`, `patchRunStatus`, `patchesInstalled`, `patchesFailed`, `patchRunStartedBy` | Latest patch run |
| `location`, `resourceId` | Azure region and full resource ID |

## Running in Azure DevOps

Use an `AzureCLI@2` task so `DefaultAzureCredential` picks up the service connection. Adjust the paths to suit the repo.

```yaml
trigger: none

schedules:
  - cron: "0 22 * * 1"          # Tuesday 09:00 AEDT (cron is in UTC)
    displayName: Weekly patch status report
    branches:
      include: [main]
    always: true

pool:
  vmImage: ubuntu-latest

steps:
  - script: curl -LsSf https://astral.sh/uv/install.sh | sh
    displayName: Install uv

  - task: AzureCLI@2
    displayName: Run patch status report
    inputs:
      azureSubscription: <service-connection-name>
      scriptType: bash
      scriptLocation: inlineScript
      workingDirectory: scripts/patch-status
      inlineScript: |
        export PATH="$HOME/.local/bin:$PATH"
        uv run --locked patch_status.py -o "$(Build.ArtifactStagingDirectory)/patch-status.xlsx"

  - publish: $(Build.ArtifactStagingDirectory)
    artifact: patch-status
    condition: always()
```

Add `--fail-on-issues` if a flagged VM should fail the run. Keep `condition: always()` on the publish step so the report is still saved when it does.

## Limitations

- **Dynamic scopes:** schedules are found through direct maintenance configuration assignments. VMs that are only in a schedule through a dynamic scope will show "no maintenance schedule assigned".
- **Patch run history:** Resource Graph only keeps about 30 days of patch runs, so older runs show as "no patch run in the last 30 days".
- **Assessment freshness:** results are only as fresh as the last assessment. Periodic assessment (checked above) keeps this current.
- **Extended support:** an OS past its standard end of life is flagged even if it has extended support, such as Ubuntu Pro ESM, because the script can't tell whether that's active on the VM. The `osSupport` column shows the extended support date where one exists.
- **OS support lookup:** if endoflife.date can't be reached, `osSupport` reads `lookup failed` and nothing is flagged for OS support. The rest of the report still runs.

## Development

```bash
uvx ruff format .
uvx ruff check .         # select = ["ALL"], see pyproject.toml
uvx ty check patch_status.py
```

Commit `uv.lock` so pipeline runs with `uv run --locked` get the same dependency versions.
