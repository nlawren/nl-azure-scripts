"""Report Azure Update Manager patch status for Linux VMs across one or more subscriptions.

Usage:
    uv run patch_status.py                          # subscriptions from patch-status.toml or .env
    uv run patch_status.py --config prod.toml --format csv
    uv run patch_status.py -s <sub-id> -s <sub-id>  # whole subscriptions
    uv run patch_status.py -s <sub-id> -g rg-app-prod -g rg-data-prod
    uv run patch_status.py --fail-on-issues         # exit 1 if anything needs attention

What to scan comes from the first of these that is set:
    1. -s/--subscription on the command line (repeatable; -g narrows to resource groups)
    2. [[subscriptions]] in the config file: --config, $PATCH_STATUS_CONFIG, or
       ./patch-status.toml if it exists (see patch-status.example.toml)
    3. $AZURE_SUBSCRIPTION_ID, comma-separated for several subscriptions
Thresholds (schedule, stale_days, ...) use command line, then config file, then defaults.

A .env file in the current directory (or --env-file) is loaded first. Variables already set
in the environment win, so pipeline variables override a local .env.

All data comes from Azure Resource Graph, one query per table, merged by VM ID:
    resources                   VM inventory, power state and patch orchestration settings
    patchassessmentresources    latest assessment and pending patch counts
    patchinstallationresources  latest patch run (Resource Graph keeps about 30 days of runs)
    maintenanceresources        maintenance configuration assignments (the schedule)
    resources                   maintenance configurations (recurrence, start time, time zone)

OS support status comes from the endoflife.date API (https://endoflife.date), looked up
once per distro per run. If it can't be reached the column says so and nothing is flagged;
--no-eol-check skips the lookup entirely.

Auth uses DefaultAzureCredential: `az login` locally, the AzureCLI@2 task's service
connection in a pipeline, or AZURE_CLIENT_ID/SECRET/TENANT_ID env vars.

Permissions: Reader on each subscription (or resource group) scanned, plus Reader wherever
the maintenance configurations live if that is elsewhere, so their schedules can be shown.
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sys
import tomllib
import urllib.request
from dataclasses import dataclass
from datetime import UTC, date, datetime, timedelta
from pathlib import Path
from typing import Any, cast
from urllib.error import URLError
from zoneinfo import ZoneInfo

import xlsxwriter
from azure.identity import DefaultAzureCredential
from azure.mgmt.resourcegraph import ResourceGraphClient
from azure.mgmt.resourcegraph.models import (
    QueryRequest,
    QueryRequestOptions,
    ResultFormat,
    ResultTruncated,
)

LOCAL_TZ = ZoneInfo("Australia/Sydney")
PAGE_SIZE = 1000  # Resource Graph maximum rows per page
RG_NAME_RE = re.compile(r"^[\w.()-]{1,90}$")
GUID_RE = re.compile(r"^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$", re.IGNORECASE)
DEFAULT_CONFIG = Path("patch-status.toml")
DEFAULT_ENV_FILE = Path(".env")
DEFAULTS: dict[str, Any] = {
    "schedule": None,
    "format": "xlsx",
    "stale_days": 2,
    "max_run_age": 8,
    "eol_warn_days": 180,
}
SUBSCRIPTION_KEYS = {"id", "name", "resource_groups", "schedule"}
REQUIRED_MODE = "AutomaticByPlatform"
PATCH_SCOPE = "InGuestPatch"
MAX_COL_WIDTH = 60
EXIT_ISSUES = 1
EXIT_CONFIG = 2  # same as argparse errors

EOL_API = "https://endoflife.date/api/v1/products/{product}"
EOL_TIMEOUT = 15
# Substring of the guest agent's osName -> endoflife.date product slug (first match wins).
EOL_PRODUCTS = [
    ("ubuntu", "ubuntu"),
    ("centos", "centos"),
    ("almalinux", "almalinux"),
    ("rocky", "rocky-linux"),
    ("redhat", "rhel"),
    ("rhel", "rhel"),
    ("debian", "debian"),
    ("sles", "sles"),
    ("suse", "sles"),
    ("oracle", "oracle-linux"),
]

COLUMNS = [
    "name",
    "subscription",
    "resourceGroup",
    "status",
    "issues",
    "powerState",
    "os",
    "osSupport",
    "patchMode",
    "assessmentMode",
    "bypassSafetyChecks",
    "maintenanceConfig",
    "schedule",
    "lastAssessment",
    "assessmentStatus",
    "securityPending",
    "criticalPending",
    "otherPending",
    "rebootPending",
    "lastPatchRun",
    "patchRunStatus",
    "patchesInstalled",
    "patchesFailed",
    "patchRunStartedBy",
    "location",
    "resourceId",
]
DATE_COLUMNS = {"lastAssessment", "lastPatchRun"}

# --- Resource Graph queries --------------------------------------------------------------
# {rg_filter} is empty (whole subscription) or a resource group filter built by rg_filter()
# from names validated against RG_NAME_RE.
# Every query projects the resource's own `id`: without it Resource Graph returns no skip
# token, so results stop at the first page.

VM_QUERY = """
resources
| where type =~ 'microsoft.compute/virtualmachines'
{rg_filter}
| where properties.storageProfile.osDisk.osType =~ 'Linux'
| extend ps = properties.osProfile.linuxConfiguration.patchSettings
| project
    id,
    vmId = tolower(id),
    name,
    subscriptionId,
    resourceGroup,
    location,
    powerState = tostring(properties.extended.instanceView.powerState.code),
    osName = tostring(properties.extended.instanceView.osName),
    osVersion = tostring(properties.extended.instanceView.osVersion),
    imageOffer = tostring(properties.storageProfile.imageReference.offer),
    imageSku = tostring(properties.storageProfile.imageReference.sku),
    patchMode = tostring(ps.patchMode),
    assessmentMode = tostring(ps.assessmentMode),
    bypass = tobool(ps.automaticByPlatformSettings.bypassPlatformSafetyChecksOnUserSchedule)
"""

ASSESSMENT_QUERY = """
patchassessmentresources
| where type =~ 'microsoft.compute/virtualmachines/patchassessmentresults'
{rg_filter}
| extend vmId = tolower(substring(id, 0, indexof(tolower(id), '/patchassessmentresults')))
| project
    id,
    vmId,
    assessmentStatus = tostring(properties.status),
    lastAssessment = tostring(properties.lastModifiedDateTime),
    rebootPending = tobool(properties.rebootPending),
    securityPending = toint(properties.availablePatchCountByClassification.security),
    criticalPending = toint(properties.availablePatchCountByClassification.critical),
    otherPending = toint(properties.availablePatchCountByClassification.other)
"""

INSTALLATION_QUERY = """
patchinstallationresources
| where type =~ 'microsoft.compute/virtualmachines/patchinstallationresults'
{rg_filter}
| extend vmId = tolower(substring(id, 0, indexof(tolower(id), '/patchinstallationresults')))
| extend runTime = todatetime(properties.lastModifiedDateTime)
| summarize arg_max(runTime, id, properties) by vmId
| project
    id,
    vmId,
    lastPatchRun = tostring(runTime),
    patchRunStatus = tostring(properties.status),
    patchesInstalled = toint(properties.installedPatchCount),
    patchesFailed = toint(properties.failedPatchCount),
    patchRunStartedBy = tostring(properties.startedBy)
"""

ASSIGNMENT_QUERY = """
maintenanceresources
| where type =~ 'microsoft.maintenance/configurationassignments'
| extend vmId = tolower(tostring(properties.resourceId))
| extend vmId = iff(isempty(vmId),
    tolower(substring(id, 0, indexof(tolower(id), '/providers/microsoft.maintenance/'))), vmId)
| where vmId has '/providers/microsoft.compute/virtualmachines/'
| extend resourceGroup = tostring(split(vmId, '/')[4])
{rg_filter}
| project id, vmId, configId = tolower(tostring(properties.maintenanceConfigurationId))
"""

# Run without a subscription list, so configs held in a central subscription are found too.
CONFIG_QUERY = """
resources
| where type =~ 'microsoft.maintenance/maintenanceconfigurations'
| project
    id,
    configId = tolower(id),
    configName = name,
    scope = tostring(properties.maintenanceScope),
    recurEvery = tostring(properties.maintenanceWindow.recurEvery),
    startDateTime = tostring(properties.maintenanceWindow.startDateTime),
    duration = tostring(properties.maintenanceWindow.duration),
    timeZone = tostring(properties.maintenanceWindow.timeZone)
"""

# Subscriptions and resource groups the identity can read. Reader on just a resource group
# shows that group but not its subscription.
CONTAINER_QUERY = """
resourcecontainers
| where type in~ ('microsoft.resources/subscriptions',
    'microsoft.resources/subscriptions/resourcegroups')
| project
    id,
    isSubscription = type =~ 'microsoft.resources/subscriptions',
    subscriptionId = tolower(subscriptionId),
    resourceGroup = tolower(resourceGroup),
    name
"""


@dataclass(frozen=True)
class Target:
    """One subscription to scan, optionally narrowed to some resource groups."""

    id: str
    name: str | None = None
    resource_groups: tuple[str, ...] = ()
    schedule: str | None = None

    @property
    def scope(self) -> str:
        """Describe what is scanned, for the console and the About sheet."""
        return ", ".join(self.resource_groups) if self.resource_groups else "all resource groups"


@dataclass(frozen=True)
class Settings:
    """Resolved run settings (command line over config file over defaults)."""

    targets: tuple[Target, ...]
    output_format: str
    output: Path
    stale_days: int
    max_run_age: int
    eol_warn_days: int
    eol_check: bool
    fail_on_issues: bool
    config_source: str


def rg_filter(resource_groups: tuple[str, ...]) -> str:
    """Build a KQL resource group filter; names are pre-validated so quoting is safe."""
    if not resource_groups:
        return ""
    names = ", ".join(f"'{rg.lower()}'" for rg in resource_groups)
    return f"| where tolower(resourceGroup) in ({names})"


def run_query(
    client: ResourceGraphClient, subscriptions: list[str] | None, query: str
) -> list[dict[str, Any]]:
    """Run a Resource Graph query, following skip tokens until every page is read.

    With no subscriptions the query covers everything the identity can read.
    """
    rows: list[dict[str, Any]] = []
    skip_token: str | None = None
    while True:
        options = QueryRequestOptions(
            result_format=ResultFormat.OBJECT_ARRAY, top=PAGE_SIZE, skip_token=skip_token
        )
        response = client.resources(
            QueryRequest(subscriptions=subscriptions, query=query, options=options)
        )
        rows.extend(cast("list[dict[str, Any]]", response.data))
        skip_token = response.skip_token
        if not skip_token:
            if response.result_truncated == ResultTruncated.TRUE:
                msg = f"Resource Graph truncated results at {len(rows)} rows with no skip token"
                raise RuntimeError(msg)
            return rows


class EolLookup:
    """Look up distro release support dates from endoflife.date, caching per product."""

    def __init__(self, warn_days: int) -> None:
        """Set how many days before end of life to start warning."""
        self.warn_days = warn_days
        self.today = datetime.now(LOCAL_TZ).date()
        self._cache: dict[str, list[dict[str, Any]] | None] = {}

    def _releases(self, product: str) -> list[dict[str, Any]] | None:
        if product not in self._cache:
            request = urllib.request.Request(  # noqa: S310 - fixed https URL
                EOL_API.format(product=product), headers={"User-Agent": "patch-status/0.1"}
            )
            try:
                with urllib.request.urlopen(request, timeout=EOL_TIMEOUT) as resp:  # noqa: S310
                    self._cache[product] = json.load(resp)["result"]["releases"]
            except (URLError, TimeoutError, ValueError, KeyError) as exc:
                print(
                    f"warning: endoflife.date lookup for {product} failed: {exc}", file=sys.stderr
                )
                self._cache[product] = None
        return self._cache[product]

    def check(self, os_name: str, os_version: str) -> tuple[str, str | None]:
        """Return (support label, issue or None) for an agent-reported OS name and version."""
        name = os_name.lower()
        product = next((slug for key, slug in EOL_PRODUCTS if key in name), None)
        if not product or not os_version:
            return "unknown", None
        releases = self._releases(product)
        if releases is None:
            return "lookup failed", None
        by_name = {str(r.get("name")): r for r in releases}
        # Agent versions are often more precise than release names: 7.9.2009 -> 7.9 -> 7.
        parts = os_version.split(".")
        release = next(
            (by_name[v] for i in range(len(parts), 0, -1) if (v := ".".join(parts[:i])) in by_name),
            None,
        )
        if release is None:
            return f"unknown ({product} {os_version} not listed)", None
        return self._assess(release)

    def _assess(self, release: dict[str, Any]) -> tuple[str, str | None]:
        eol = _to_date(release.get("eolFrom"))
        extended = _to_date(release.get("eoesFrom"))
        if eol is None:
            return (
                ("End of life", "OS end of life") if release.get("isEol") else ("Supported", None)
            )
        if eol <= self.today:
            label = f"End of life {eol}"
            if extended and extended > self.today:
                label += f" (extended support to {extended})"
            return label, f"OS end of life since {eol}"
        if eol <= self.today + timedelta(days=self.warn_days):
            return f"Supported until {eol}", f"OS end of life on {eol}"
        return f"Supported until {eol}", None


def _to_date(value: object) -> date | None:
    """Parse an endoflife.date field, which may be a date string, a boolean or null."""
    if isinstance(value, str):
        try:
            return date.fromisoformat(value)
        except ValueError:
            return None
    return None


def parse_time(value: str | None) -> datetime | None:
    """Parse a Resource Graph timestamp (UTC ISO 8601) into an aware datetime."""
    if not value:
        return None
    try:
        parsed = datetime.fromisoformat(value)
    except ValueError:
        return None
    return parsed if parsed.tzinfo else parsed.replace(tzinfo=UTC)


def describe_schedule(config: dict[str, Any]) -> str:
    """Turn a maintenance window into e.g. '1Week Sunday at 02:00 AUS Eastern ... (03:55)'."""
    start = config.get("startDateTime") or ""
    start_time = start.split(" ")[-1] if " " in start else start
    parts = [config.get("recurEvery") or "?"]
    if start_time:
        parts.append(f"at {start_time}")
    if config.get("timeZone"):
        parts.append(config["timeZone"])
    if config.get("duration"):
        parts.append(f"({config['duration']})")
    return " ".join(parts)


def config_issues(row: dict[str, Any]) -> list[str]:
    """Check the VM's patch orchestration settings."""
    issues: list[str] = []
    if row["powerState"] != "running":
        issues.append(f"VM {row['powerState'] or 'state unknown'}")
    if row["patchMode"] != REQUIRED_MODE:
        issues.append(f"patch mode is {row['patchMode'] or 'unset'} (needs {REQUIRED_MODE})")
    elif not row["bypassSafetyChecks"]:
        issues.append("bypassPlatformSafetyChecksOnUserSchedule not set")
    if row["assessmentMode"] != REQUIRED_MODE:
        issues.append("periodic assessment off")
    return issues


def schedule_issues(configs: list[dict[str, Any]], expected: str | None) -> list[str]:
    """Check the VM is assigned to a weekly patch maintenance configuration (and the right one).

    Only InGuestPatch configs patch the guest OS; Host or OSImage configs don't count. A config
    the identity can't read has no scope, so it's given the benefit of the doubt.
    """
    if not configs:
        return ["no maintenance schedule assigned"]
    candidates = [c for c in configs if c.get("scope", PATCH_SCOPE).lower() == PATCH_SCOPE.lower()]
    if not candidates:
        scopes = ", ".join(sorted({c["scope"] or "unknown" for c in configs}))
        return [f"no patch schedule assigned (only {scopes} maintenance)"]
    issues: list[str] = []
    names = {c["configName"].lower() for c in candidates}
    if expected and expected.lower() not in names:
        issues.append(f"not on schedule {expected}")
    visible = [c for c in candidates if "recurEvery" in c]
    if not visible:
        issues.append("schedule details not visible (needs Reader where the config lives)")
    elif not any("week" in (c.get("recurEvery") or "").lower() for c in visible):
        issues.append("schedule is not weekly")
    return issues


def assessment_issues(row: dict[str, Any], stale_days: int, now: datetime) -> list[str]:
    """Check the latest assessment is recent and clean."""
    assessed: datetime | None = row["lastAssessment"]
    if assessed is None:
        return ["never assessed"]
    issues: list[str] = []
    age = (now - assessed).days
    if row["powerState"] == "running" and age > stale_days:
        issues.append(f"assessment {age} days old")
    if row["assessmentStatus"] == "Failed":
        issues.append("last assessment failed")
    if row["rebootPending"]:
        issues.append("reboot pending")
    urgent = (row["securityPending"] or 0) + (row["criticalPending"] or 0)
    if urgent:
        issues.append(f"{urgent} security/critical patches pending")
    return issues


def patch_run_issues(row: dict[str, Any], max_run_age: int, now: datetime) -> list[str]:
    """Check the latest patch run happened on time and succeeded."""
    patched: datetime | None = row["lastPatchRun"]
    if patched is None:
        return ["no patch run in the last 30 days"]
    issues: list[str] = []
    age = (now - patched).days
    if age > max_run_age:
        issues.append(f"last patch run {age} days ago")
    status = row["patchRunStatus"]
    if status == "Failed":
        issues.append("last patch run failed")
    elif status == "CompletedWithWarnings":
        issues.append("last patch run completed with warnings")
    if row["patchesFailed"]:
        issues.append(f"{row['patchesFailed']} patches failed to install")
    return issues


def find_issues(
    row: dict[str, Any],
    configs: list[dict[str, Any]],
    schedule: str | None,
    settings: Settings,
    now: datetime,
) -> list[str]:
    """Return the problems worth flagging for one VM; an empty list means it looks healthy."""
    return [
        *config_issues(row),
        *schedule_issues(configs, schedule),
        *assessment_issues(row, settings.stale_days, now),
        *patch_run_issues(row, settings.max_run_age, now),
    ]


def collect(settings: Settings) -> list[dict[str, Any]]:
    """Query Resource Graph and build one merged report row per Linux VM, for every target."""
    client = ResourceGraphClient(DefaultAzureCredential())
    sub_ids = [t.id for t in settings.targets]
    containers = run_query(client, sub_ids, CONTAINER_QUERY)
    sub_names = {r["subscriptionId"]: r["name"] for r in containers if r["isSubscription"]}
    check_readable(settings.targets, sub_names, containers)
    configs = {r["configId"]: r for r in run_query(client, None, CONFIG_QUERY)}
    eol = EolLookup(settings.eol_warn_days) if settings.eol_check else None
    now = datetime.now(UTC)

    rows: list[dict[str, Any]] = []
    for target in settings.targets:
        label = target.name or sub_names.get(target.id) or target.id
        found = collect_target(
            client, target, label, configs=configs, eol=eol, settings=settings, now=now
        )
        print(
            f"{label} ({target.scope}): {len(found)} Linux VM{'' if len(found) == 1 else 's'}",
            file=sys.stderr,
        )
        rows.extend(found)
    rows.sort(key=lambda r: (r["subscription"].lower(), r["resourceGroup"], r["name"].lower()))
    return rows


def check_readable(
    targets: tuple[Target, ...], sub_names: dict[str, str], containers: list[dict[str, Any]]
) -> None:
    """Fail if a target can't be read, rather than reporting it as having no Linux VMs.

    Resource Graph silently returns nothing for scopes the identity can't read. A target with
    resource groups needs each group to be visible; a whole-subscription target needs Reader on
    the subscription itself, or the scan would only cover the groups that happen to be readable.
    """
    groups = {(r["subscriptionId"], r["resourceGroup"]) for r in containers}
    missing: list[str] = []
    for t in targets:
        if t.resource_groups:
            missing += [
                f"{t.id}/{rg}" for rg in t.resource_groups if (t.id, rg.lower()) not in groups
            ]
        elif t.id not in sub_names:
            hint = " (Reader on only some resource groups? list them with -g)"
            missing.append(t.id + (hint if any(s == t.id for s, _ in groups) else ""))
    if missing:
        msg = f"not found or not readable: {', '.join(missing)}"
        raise LookupError(msg)


def collect_target(  # noqa: PLR0913 - plumbing for one subscription
    client: ResourceGraphClient,
    target: Target,
    label: str,
    *,
    configs: dict[str, dict[str, Any]],
    eol: EolLookup | None,
    settings: Settings,
    now: datetime,
) -> list[dict[str, Any]]:
    """Build report rows for the Linux VMs in one subscription (or its resource groups)."""
    flt = rg_filter(target.resource_groups)

    def query(template: str) -> list[dict[str, Any]]:
        return run_query(client, [target.id], template.format(rg_filter=flt))

    vms = query(VM_QUERY)
    if not vms:
        return []
    assessments = {r["vmId"]: r for r in query(ASSESSMENT_QUERY)}
    installations = {r["vmId"]: r for r in query(INSTALLATION_QUERY)}
    assigned: dict[str, list[dict[str, Any]]] = {}
    for a in query(ASSIGNMENT_QUERY):
        # A config the identity can't read won't be in `configs`; fall back to its name.
        config = configs.get(a["configId"]) or {"configName": a["configId"].rsplit("/", 1)[-1]}
        assigned.setdefault(a["vmId"], []).append(config)

    rows: list[dict[str, Any]] = []
    for vm in vms:
        vm_id = vm["vmId"]
        assess = assessments.get(vm_id, {})
        install = installations.get(vm_id, {})
        vm_configs = assigned.get(vm_id, [])
        # Prefer what the guest agent reports. The image reference is only what the VM was
        # built from, so it goes stale after in-place upgrades; label it so it isn't trusted.
        os_label = " ".join(filter(None, [vm["osName"], vm["osVersion"]]))
        if not os_label:
            image = " ".join(filter(None, [vm["imageOffer"], vm["imageSku"]]))
            os_label = f"unknown (built from image: {image})" if image else "unknown"
        support, os_issue = (
            eol.check(vm["osName"], vm["osVersion"]) if eol else ("not checked", None)
        )
        row: dict[str, Any] = {
            "name": vm["name"],
            "subscription": label,
            "resourceGroup": vm["resourceGroup"].lower(),
            "powerState": vm["powerState"].removeprefix("PowerState/"),
            "os": os_label,
            "osSupport": support,
            "patchMode": vm["patchMode"],
            "assessmentMode": vm["assessmentMode"],
            "bypassSafetyChecks": bool(vm["bypass"]),
            "maintenanceConfig": ", ".join(c["configName"] for c in vm_configs),
            "schedule": "; ".join(describe_schedule(c) for c in vm_configs if "recurEvery" in c),
            "lastAssessment": parse_time(assess.get("lastAssessment")),
            "assessmentStatus": assess.get("assessmentStatus"),
            "securityPending": assess.get("securityPending"),
            "criticalPending": assess.get("criticalPending"),
            "otherPending": assess.get("otherPending"),
            "rebootPending": assess.get("rebootPending"),
            "lastPatchRun": parse_time(install.get("lastPatchRun")),
            "patchRunStatus": install.get("patchRunStatus"),
            "patchesInstalled": install.get("patchesInstalled"),
            "patchesFailed": install.get("patchesFailed"),
            "patchRunStartedBy": install.get("patchRunStartedBy"),
            "location": vm["location"],
            "resourceId": vm_id,
        }
        issues = [os_issue] if os_issue else []
        issues += find_issues(row, vm_configs, target.schedule, settings, now)
        row["status"] = "Attention" if issues else "OK"
        row["issues"] = "; ".join(issues)
        rows.append(row)
    return rows


def to_local(value: datetime | None) -> datetime | None:
    """Convert to local wall-clock time without tzinfo (Excel has no time zones)."""
    return value.astimezone(LOCAL_TZ).replace(tzinfo=None) if value else None


def write_csv(rows: list[dict[str, Any]], path: Path) -> None:
    """Write the report as CSV with local timestamps."""
    with path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=COLUMNS)
        writer.writeheader()
        for row in rows:
            out = dict(row)
            for col in DATE_COLUMNS:
                local = to_local(row[col])
                out[col] = local.strftime("%Y-%m-%d %H:%M") if local else ""
            writer.writerow(out)


def write_xlsx(rows: list[dict[str, Any]], path: Path, settings: Settings) -> None:
    """Write the report as a filterable Excel table with a frozen header and an About sheet."""
    with xlsxwriter.Workbook(path) as wb:
        date_fmt = wb.add_format({"num_format": "yyyy-mm-dd hh:mm"})
        bad_fmt = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})
        good_fmt = wb.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"})

        ws = wb.add_worksheet("Patch status")
        data = [[to_local(r[c]) if c in DATE_COLUMNS else r[c] for c in COLUMNS] for r in rows]
        last_row = max(len(data), 1)
        ws.add_table(
            0,
            0,
            last_row,
            len(COLUMNS) - 1,
            {
                "name": "PatchStatus",
                "style": "Table Style Medium 2",
                "data": data,
                "columns": [
                    {"header": c, "format": date_fmt} if c in DATE_COLUMNS else {"header": c}
                    for c in COLUMNS
                ],
            },
        )
        status_col = COLUMNS.index("status")
        for value, fmt in (("Attention", bad_fmt), ("OK", good_fmt)):
            ws.conditional_format(
                1,
                status_col,
                last_row,
                status_col,
                {"type": "cell", "criteria": "==", "value": f'"{value}"', "format": fmt},
            )
        for idx, col in enumerate(COLUMNS):
            longest = max(
                [len(col)] + [len(str(r[col])) for r in rows if r[col] is not None], default=0
            )
            width = 17 if col in DATE_COLUMNS else min(longest + 2, MAX_COL_WIDTH)
            ws.set_column(idx, idx, width)
        ws.freeze_panes(1, 1)

        about = wb.add_worksheet("About")
        attention = sum(r["status"] == "Attention" for r in rows)
        info: list[tuple[str, Any]] = [
            ("Generated", datetime.now(LOCAL_TZ).strftime("%Y-%m-%d %H:%M %Z")),
            ("Subscriptions from", settings.config_source),
            ("Linux VMs", len(rows)),
            ("Need attention", attention),
            ("Assessment stale after (days)", settings.stale_days),
            ("Patch run overdue after (days)", settings.max_run_age),
            ("OS support data", "endoflife.date" if settings.eol_check else "not checked"),
            ("OS end-of-life warning (days)", settings.eol_warn_days),
            ("Times shown in", str(LOCAL_TZ)),
        ]
        for i, (key, value) in enumerate(info):
            about.write(i, 0, key)
            about.write(i, 1, value)

        bold = wb.add_format({"bold": True})
        start = len(info) + 1
        headers = [
            "Subscription",
            "Subscription ID",
            "Scope",
            "Expected schedule",
            "VMs",
            "Attention",
        ]
        about.write_row(start, 0, headers, bold)
        for i, target in enumerate(settings.targets, start=start + 1):
            mine = [r for r in rows if r["resourceId"].startswith(f"/subscriptions/{target.id}/")]
            label = mine[0]["subscription"] if mine else (target.name or target.id)
            about.write_row(
                i,
                0,
                [
                    label,
                    target.id,
                    target.scope,
                    target.schedule or "(any weekly schedule)",
                    len(mine),
                    sum(r["status"] == "Attention" for r in mine),
                ],
            )
        about.set_column(0, 0, 30)
        about.set_column(1, 1, 40)
        about.set_column(2, 3, 30)


def print_summary(rows: list[dict[str, Any]]) -> None:
    """Print a short console summary of VMs that need attention."""
    attention = [r for r in rows if r["status"] == "Attention"]
    print(
        f"{len(rows)} Linux VM{'' if len(rows) == 1 else 's'}, "
        f"{len(rows) - len(attention)} OK, {len(attention)} need attention"
    )
    if attention:
        width = max(len(r["name"]) for r in attention)
        current = None
        for r in attention:
            if r["subscription"] != current:
                current = r["subscription"]
                print(f"\n{current}")
            print(f"  {r['name']:<{width}}  {r['issues']}  [{r['resourceGroup']}]")
        print()


def load_env_file(path: Path) -> None:
    """Load KEY=VALUE lines into os.environ without overriding variables already set."""
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.removeprefix("export ").partition("=")
        key, value = key.strip(), value.strip()
        if len(value) > 1 and value[0] == value[-1] and value[0] in "\"'":
            value = value[1:-1]
        os.environ.setdefault(key, value)


def load_config(path: Path) -> dict[str, Any]:
    """Read the TOML config, rejecting unknown keys so typos don't pass silently."""
    with path.open("rb") as fh:
        config = tomllib.load(fh)
    unknown = set(config) - set(DEFAULTS) - {"subscriptions"}
    if unknown:
        msg = f"{path}: unknown setting(s): {', '.join(sorted(unknown))}"
        raise ValueError(msg)
    for i, sub in enumerate(config.get("subscriptions", []), start=1):
        bad = set(sub) - SUBSCRIPTION_KEYS
        if bad:
            msg = f"{path}: subscription {i}: unknown key(s): {', '.join(sorted(bad))}"
            raise ValueError(msg)
        if not isinstance(sub.get("id"), str):
            msg = f"{path}: subscription {i} needs an id string"
            raise TypeError(msg)
        groups = sub.get("resource_groups", [])
        if not isinstance(groups, list) or not all(isinstance(g, str) for g in groups):
            msg = f'{path}: subscription {i}: resource_groups must be a list, e.g. ["rg-app"]'
            raise TypeError(msg)
    return config


def build_targets(
    args: argparse.Namespace, config: dict[str, Any], config_path: Path | None
) -> tuple[list[Target], str]:
    """Work out which subscriptions to scan, and say where that came from."""
    schedule = args.schedule or config.get("schedule")
    groups = tuple(args.resource_group or ())
    if args.subscription:
        subs = [Target(s, None, groups, schedule) for s in args.subscription]
        return subs, "command line"
    if config.get("subscriptions"):
        subs = [
            Target(
                id=sub["id"],
                name=sub.get("name"),
                resource_groups=groups or tuple(sub.get("resource_groups", ())),
                schedule=args.schedule or sub.get("schedule") or schedule,
            )
            for sub in config["subscriptions"]
        ]
        return subs, str(config_path)
    env = os.environ.get("AZURE_SUBSCRIPTION_ID", "")
    ids = [s.strip() for s in env.split(",") if s.strip()]
    return [Target(s, None, groups, schedule) for s in ids], "AZURE_SUBSCRIPTION_ID"


HELP_DESCRIPTION = """\
Report Azure Update Manager patch status for Linux VMs across one or more subscriptions.

Reads Azure Resource Graph and flags VMs that need attention: wrong patch settings, not on
a weekly maintenance schedule, stale or failed assessments, missed or failed patch runs,
pending security patches or reboots, and OS releases at or near end of life.

What to scan comes from the first of these that is set:
  1. -s/--subscription on the command line
  2. [[subscriptions]] in the config file (--config, $PATCH_STATUS_CONFIG,
     or ./patch-status.toml if it exists)
  3. $AZURE_SUBSCRIPTION_ID, comma-separated for several subscriptions

Thresholds use the command line first, then the config file, then the defaults shown below.
A .env file is loaded before any of this; variables already in the environment win."""

HELP_EPILOG = """\
examples:
  uv run patch_status.py
      scan the subscriptions listed in ./patch-status.toml (or .env)
  uv run patch_status.py -s <sub-id> -s <sub-id> --format csv
      scan two whole subscriptions, write CSV
  uv run patch_status.py -s <sub-id> -g rg-app-prod -g rg-data-prod
      scan two resource groups in one subscription
  uv run patch_status.py --config prod.toml --schedule mc-linux-weekly --fail-on-issues
      use another config, require a specific schedule, exit 1 if anything is flagged

exit codes:
  0  report written (VMs may still need attention unless --fail-on-issues is set)
  1  no Linux VMs found, or --fail-on-issues and at least one VM needs attention
  2  bad arguments or config file, or a subscription or resource group
     that doesn't exist or can't be read

Authentication uses DefaultAzureCredential (az login, a pipeline service connection, or
AZURE_CLIENT_ID/AZURE_CLIENT_SECRET/AZURE_TENANT_ID). Reader access is enough.
See README.md for the config file format."""


def build_parser() -> argparse.ArgumentParser:
    """Define the command line."""
    parser = argparse.ArgumentParser(
        prog="patch_status.py",
        description=HELP_DESCRIPTION,
        epilog=HELP_EPILOG,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    scope = parser.add_argument_group("what to scan")
    scope.add_argument(
        "-s",
        "--subscription",
        action="append",
        metavar="ID",
        help="subscription ID to scan. Repeat for several. Overrides the config file and "
        "$AZURE_SUBSCRIPTION_ID.",
    )
    scope.add_argument(
        "-g",
        "--resource-group",
        action="append",
        metavar="NAME",
        help="only scan this resource group. Repeat for several. Applies to every "
        "subscription being scanned. (default: the whole subscription)",
    )

    sources = parser.add_argument_group("configuration")
    sources.add_argument(
        "-c",
        "--config",
        type=Path,
        metavar="FILE",
        help="TOML config file listing subscriptions and thresholds. "
        "(default: $PATCH_STATUS_CONFIG, or ./patch-status.toml if it exists)",
    )
    sources.add_argument(
        "--env-file",
        type=Path,
        metavar="FILE",
        help="env file to load before anything else. (default: ./.env if it exists)",
    )

    checks = parser.add_argument_group("checks")
    checks.add_argument(
        "--schedule",
        metavar="NAME",
        help="name of the maintenance configuration every VM should be assigned to. Without "
        "it, any weekly schedule passes.",
    )
    checks.add_argument(
        "--stale-days",
        type=int,
        metavar="DAYS",
        help="flag running VMs whose last patch assessment is older than this. (default: 2)",
    )
    checks.add_argument(
        "--max-run-age",
        type=int,
        metavar="DAYS",
        help="flag VMs whose last patch run is older than this. A weekly schedule should "
        "never exceed 7. (default: 8)",
    )
    checks.add_argument(
        "--eol-warn-days",
        type=int,
        metavar="DAYS",
        help="flag OS releases that reach end of life within this many days. (default: 180)",
    )
    checks.add_argument(
        "--no-eol-check",
        action="store_true",
        help="skip the endoflife.date lookup, e.g. where the network blocks it. The osSupport "
        "column then reads 'not checked'.",
    )

    output = parser.add_argument_group("output")
    output.add_argument(
        "--format",
        choices=["xlsx", "csv"],
        help="report format: an Excel table with an About sheet, or plain CSV. (default: xlsx)",
    )
    output.add_argument(
        "-o",
        "--output",
        type=Path,
        metavar="FILE",
        help="where to write the report. (default: patch-status-YYYYMMDD.xlsx or .csv)",
    )
    output.add_argument(
        "--fail-on-issues",
        action="store_true",
        help="exit with status 1 if any VM needs attention, so a pipeline run fails.",
    )
    return parser


def load_sources(
    parser: argparse.ArgumentParser, args: argparse.Namespace
) -> tuple[dict[str, Any], Path]:
    """Load the .env file into the environment, then read the config file if there is one."""
    env_file = args.env_file or DEFAULT_ENV_FILE
    if args.env_file and not env_file.is_file():
        parser.error(f"env file not found: {env_file}")
    if env_file.is_file():
        load_env_file(env_file)

    config_path = args.config or Path(os.environ.get("PATCH_STATUS_CONFIG", DEFAULT_CONFIG))
    explicit = args.config is not None or "PATCH_STATUS_CONFIG" in os.environ
    config: dict[str, Any] = {}
    if config_path.is_file():
        try:
            config = load_config(config_path)
        except (tomllib.TOMLDecodeError, ValueError, TypeError) as exc:
            parser.error(str(exc))
    elif explicit:
        parser.error(f"config file not found: {config_path}")
    return config, config_path


def validate_targets(parser: argparse.ArgumentParser, targets: list[Target]) -> list[Target]:
    """Check subscription IDs and resource group names before they go into KQL."""
    for t in targets:
        if not GUID_RE.fullmatch(t.id):
            parser.error(f"invalid subscription ID: {t.id!r}")
        for rg in t.resource_groups:
            if not RG_NAME_RE.fullmatch(rg):
                parser.error(f"invalid resource group name: {rg!r}")
    return [Target(t.id.lower(), t.name, t.resource_groups, t.schedule) for t in targets]


def parse_args(argv: list[str] | None = None) -> Settings:
    """Parse the command line, .env and config file into resolved settings."""
    parser = build_parser()
    args = parser.parse_args(argv)
    config, config_path = load_sources(parser, args)

    targets, source = build_targets(args, config, config_path)
    if not targets:
        parser.error(
            "nothing to scan: pass -s/--subscription, add [[subscriptions]] to "
            f"{config_path}, or set AZURE_SUBSCRIPTION_ID"
        )
    targets = validate_targets(parser, targets)

    def pick(name: str) -> Any:  # noqa: ANN401 - mixed setting types
        value = getattr(args, name)
        return value if value is not None else config.get(name, DEFAULTS[name])

    output_format = pick("format")
    if output_format not in {"xlsx", "csv"}:
        parser.error(f"format must be xlsx or csv, not {output_format!r}")
    output = args.output
    if output is None:
        stamp = date.today().strftime("%Y%m%d")  # noqa: DTZ011 - local date for a file name
        output = Path(f"patch-status-{stamp}.{output_format}")

    return Settings(
        targets=tuple(targets),
        output_format=output_format,
        output=output,
        stale_days=int(pick("stale_days")),
        max_run_age=int(pick("max_run_age")),
        eol_warn_days=int(pick("eol_warn_days")),
        eol_check=not args.no_eol_check,
        fail_on_issues=args.fail_on_issues,
        config_source=source,
    )


def main(argv: list[str] | None = None) -> int:
    """Entry point."""
    settings = parse_args(argv)
    try:
        rows = collect(settings)
    except LookupError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return EXIT_CONFIG
    if not rows:
        print("No Linux VMs found", file=sys.stderr)
        return EXIT_ISSUES
    if settings.output_format == "csv":
        write_csv(rows, settings.output)
    else:
        write_xlsx(rows, settings.output, settings)
    print_summary(rows)
    print(f"Report written to {settings.output}")
    attention = any(r["status"] == "Attention" for r in rows)
    return EXIT_ISSUES if attention and settings.fail_on_issues else 0


if __name__ == "__main__":
    sys.exit(main())
