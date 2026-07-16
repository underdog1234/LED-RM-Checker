#!/usr/bin/env python3
"""Warn in Slack when upcoming Rentman projects use watched LED equipment."""

from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any


DEFAULT_EQUIPMENT_FILE = Path("data/Export_Equipment_20260715.xlsx")
DEFAULT_MATCH_COLUMNS = ("Code", "Name (in database)", "General Name")
DEFAULT_PROJECT_FIELDS = (
    "id,name,number,reference,displayname,planperiod_start,planperiod_end,"
    "usageperiod_start,usageperiod_end,tags,status,project_type,account_manager"
)
DEFAULT_PROJECT_EQUIPMENT_FIELDS = (
    "id,name,displayname,quantity,quantity_total,equipment,planperiod_start,planperiod_end"
)
DEFAULT_SUBPROJECT_FIELDS = (
    "id,name,status,planperiod_start,planperiod_end,usageperiod_start,usageperiod_end"
)
DEFAULT_REQUIRED_TAG = "tec\u2714\ufe0f"
DEFAULT_PROJECT_EXPAND = "project_type,account_manager"
DEFAULT_SUBPROJECT_EXPAND = "status"
DEFAULT_PROJECT_URL_TEMPLATE = "https://multimedia.rentmanapp.com/#/projects/{id}/details"
STATUS_DETAILS = {
    "/statuses/1": ("\u23f3", "Pending"),
    "/statuses/2": ("\u274c", "Cancelled"),
    "/statuses/3": ("\U0001f44d", "Confirmed"),
    "/statuses/4": ("\U0001f4e6", "Prepped"),
    "/statuses/5": ("\U0001f4cd", "On location"),
    "/statuses/6": ("\U0001f69a", "Returned"),
    "/statuses/7": ("\U0001f5d3\ufe0f", "Inquiry"),
    "/statuses/8": ("\U0001f4dd", "Concept"),
}
PROJECT_TYPE_COLORS = {
    "011 Production": "cf9dff",
    "Dryhire": "8F53F2",
    "Internal rental project": "000000",
    "Transferproject": "c8c8c8",
    "Conference": "ec00a3",
    "Award Show": "08d400",
    "Meeting": "1e27f2",
    "Product launch": "e1ed10",
    "Dinner": "4a36f9",
    "Expo": "47dcdc",
    "Warehouse": "8db0b8",
    "Conference and Dinner": "221be3",
    "Equipment Hire/ Delivery": "000000",
    "AGM": "2fe8ec",
    "Cocktail": "fade1f",
    "Roadshow": "faf833",
    "Training": "000000",
    "Wedding": "012340",
    "003 Christchurch": "ff1424",
    "001 Auckland": "2c4bf9",
    "002 Wellington": "43d226",
    "SOA sent": "bc85c1",
    "Completed": "40d228",
    "001c GMA": "da9694",
    "005 Te Pae": "fc5d58",
    "006 Parliament": "fbf373",
    "001b DDEC": "ffff00",
    "Choose a Location": "000000",
    "001a Auckland Museum": "fc9b25",
    "Public Holiday": "ff9ee4",
    "012 Long Term Hire": "000000",
    "001aa Akl Dry Hire": "b8b8cc",
    "002a WLG Dry Hire": "dedef6",
    "003a CHC Dry Hire": "dedef6",
    "000a Multiple Locations": "6713B0",
    "001d Staff Dry Hire": "c8c8c8",
    "Production": "cf9dff",
    "CHC": "ff1424",
    "Newmarket": "2c4bf9",
    "WEL": "43d226",
    "GMA": "da9694",
    "Te Pae": "fc5d58",
    "Parliament": "fbf373",
    "DDEC": "ffff00",
    "AWMMM": "fc9b25",
    "Long Term Hire": "000000",
    "DH": "b8b8cc",
    "\U0001f4cd\U0001f4cd": "6713B0",
    "??": "000000",
}
NS_MAIN = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
NS_REL = {
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}


class ConfigError(RuntimeError):
    """Raised when required configuration is missing or invalid."""


@dataclass(frozen=True)
class EquipmentTarget:
    row_number: int
    values: dict[str, str]
    identifiers: frozenset[str]


@dataclass(frozen=True)
class EquipmentMatch:
    project_equipment_id: Any
    display_name: str
    quantity: Any
    matched_identifier: str
    target_row: int


@dataclass(frozen=True)
class Finding:
    project: dict[str, Any]
    matches: tuple[EquipmentMatch, ...]


def normalize(value: Any) -> str:
    text = "" if value is None else str(value)
    return re.sub(r"\s+", " ", text).strip().casefold()


def split_csv_env(name: str, default: tuple[str, ...]) -> tuple[str, ...]:
    raw = os.environ.get(name)
    if not raw:
        return default
    return tuple(part.strip() for part in raw.split(",") if part.strip())


def parse_int_env(name: str, default: int | None) -> int | None:
    raw = os.environ.get(name)
    if raw is None or raw == "":
        return default
    value = int(raw)
    return value if value > 0 else None


def cell_column_index(cell_ref: str) -> int:
    letters = re.match(r"[A-Z]+", cell_ref.upper())
    if not letters:
        return 0
    index = 0
    for char in letters.group(0):
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index - 1


def read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    strings: list[str] = []
    for item in root.findall("m:si", NS_MAIN):
        text_parts = [node.text or "" for node in item.findall(".//m:t", NS_MAIN)]
        strings.append("".join(text_parts))
    return strings


def first_worksheet_path(archive: zipfile.ZipFile) -> str:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    first_sheet = workbook.find("m:sheets/m:sheet", NS_MAIN)
    if first_sheet is None:
        raise ConfigError("Workbook does not contain any sheets.")

    rel_id = first_sheet.attrib.get(f"{{{NS_REL['r']}}}id")
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    for rel in rels.findall("pr:Relationship", NS_REL):
        if rel.attrib.get("Id") == rel_id:
            target = rel.attrib["Target"].lstrip("/")
            return target if target.startswith("xl/") else f"xl/{target}"
    raise ConfigError("Could not resolve first worksheet path.")


def read_xlsx_rows(path: Path) -> list[list[str]]:
    if not path.exists():
        raise ConfigError(f"Equipment file not found: {path}")

    with zipfile.ZipFile(path) as archive:
        shared_strings = read_shared_strings(archive)
        worksheet_path = first_worksheet_path(archive)
        root = ET.fromstring(archive.read(worksheet_path))

    rows: list[list[str]] = []
    for row in root.findall(".//m:sheetData/m:row", NS_MAIN):
        values: list[str] = []
        for cell in row.findall("m:c", NS_MAIN):
            col_index = cell_column_index(cell.attrib.get("r", "A1"))
            while len(values) <= col_index:
                values.append("")

            cell_type = cell.attrib.get("t")
            raw_value = cell.find("m:v", NS_MAIN)
            inline_string = cell.find("m:is", NS_MAIN)
            value = ""
            if cell_type == "s" and raw_value is not None and raw_value.text is not None:
                value = shared_strings[int(raw_value.text)]
            elif cell_type == "inlineStr" and inline_string is not None:
                value = "".join(node.text or "" for node in inline_string.findall(".//m:t", NS_MAIN))
            elif raw_value is not None and raw_value.text is not None:
                value = raw_value.text
            values[col_index] = value
        if any(normalize(value) for value in values):
            rows.append(values)
    return rows


def load_equipment_targets(path: Path, match_columns: tuple[str, ...]) -> list[EquipmentTarget]:
    rows = read_xlsx_rows(path)
    if not rows:
        raise ConfigError(f"Equipment file is empty: {path}")

    headers = [str(value).strip() for value in rows[0]]
    header_lookup = {normalize(header): index for index, header in enumerate(headers)}
    requested_indexes: list[tuple[str, int]] = []
    missing_headers: list[str] = []
    for column in match_columns:
        index = header_lookup.get(normalize(column))
        if index is None:
            missing_headers.append(column)
        else:
            requested_indexes.append((column, index))

    if not requested_indexes:
        raise ConfigError(
            "None of the configured match columns exist in the workbook. "
            f"Configured: {', '.join(match_columns)}. Headers: {', '.join(headers)}"
        )
    if missing_headers:
        print(
            "Warning: missing equipment match column(s): " + ", ".join(missing_headers),
            file=sys.stderr,
        )

    targets: list[EquipmentTarget] = []
    for row_index, row in enumerate(rows[1:], start=2):
        values: dict[str, str] = {}
        identifiers: set[str] = set()
        for column, col_index in requested_indexes:
            raw = row[col_index] if col_index < len(row) else ""
            text = str(raw).strip()
            if text:
                values[column] = text
                identifiers.add(normalize(text))
        if identifiers:
            targets.append(
                EquipmentTarget(
                    row_number=row_index,
                    values=values,
                    identifiers=frozenset(identifiers),
                )
            )
    return targets


def parse_datetime(value: Any, timezone: dt.tzinfo) -> dt.datetime | None:
    if not value:
        return None
    text = str(value).strip()
    if not text:
        return None
    try:
        parsed = dt.datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError:
        return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone)
    try:
        return parsed.astimezone(timezone)
    except OverflowError:
        return None


def project_is_upcoming(
    project: dict[str, Any],
    today: dt.date,
    timezone: dt.tzinfo,
    lookahead_days: int | None,
) -> bool:
    start_of_today = dt.datetime.combine(today, dt.time.min, tzinfo=timezone)
    horizon = (
        start_of_today + dt.timedelta(days=lookahead_days)
        if lookahead_days is not None
        else None
    )
    ranges = (
        ("planperiod_start", "planperiod_end"),
        ("usageperiod_start", "usageperiod_end"),
    )
    for start_key, end_key in ranges:
        start = parse_datetime(project.get(start_key), timezone)
        end = parse_datetime(project.get(end_key), timezone) or start
        if end is None:
            continue
        if end < start_of_today:
            continue
        if horizon is not None and start is not None and start > horizon:
            continue
        return True
    return False


def tag_is_present(raw_tags: Any, required_tag: str) -> bool:
    required = normalize(required_tag)
    if raw_tags is None:
        return False
    if isinstance(raw_tags, list):
        tag_values = raw_tags
    else:
        tag_values = re.split(r"[,;\n]", str(raw_tags))
    return any(normalize(tag) == required for tag in tag_values)


def resource_path(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, dict):
        for key in ("path", "url", "href"):
            if value.get(key):
                return str(value[key]).strip()
        if value.get("id") not in (None, ""):
            return f"/statuses/{value['id']}"
    return str(value).strip()


def effective_project_status_value(project: dict[str, Any]) -> Any:
    if project.get("status") not in (None, ""):
        return project.get("status")
    subprojects = project.get("_subprojects")
    if isinstance(subprojects, list) and subprojects:
        return subprojects[0].get("status")
    return None


def project_status_key(project: dict[str, Any]) -> str:
    status = effective_project_status_value(project)
    path = resource_path(status)
    if path:
        return path
    if isinstance(status, dict):
        return str(status.get("name") or status.get("displayname") or "").strip()
    return str(status or "").strip()


def project_status_details(project: dict[str, Any]) -> tuple[str, str]:
    status = effective_project_status_value(project)
    key = project_status_key(project)
    normalized_key = normalize(key)
    for status_key, details in STATUS_DETAILS.items():
        if normalized_key == normalize(status_key):
            return details
    if isinstance(status, dict):
        label = str(status.get("name") or status.get("displayname") or status.get("id") or "").strip()
        if label:
            return ("\u2139\ufe0f", label)
    if key:
        return ("\u2139\ufe0f", key)
    return ("\u2139\ufe0f", "Status unknown")


def project_status(project: dict[str, Any]) -> str:
    _emoji, label = project_status_details(project)
    return normalize(label)


def project_is_cancelled(project: dict[str, Any]) -> bool:
    emoji, label = project_status_details(project)
    return emoji == "\u274c" or normalize(label) == "cancelled"


def project_type_name(project: dict[str, Any]) -> str:
    project_type = project.get("project_type")
    if isinstance(project_type, dict):
        return str(
            project_type.get("name")
            or project_type.get("displayname")
            or project_type.get("label")
            or project_type.get("id")
            or ""
        ).strip()
    return str(project_type or "").strip()


def nearest_color_emoji(hex_color: str) -> str:
    raw = re.sub(r"[^0-9a-fA-F]", "", str(hex_color or ""))
    if not raw or int(raw or "0", 16) == 0:
        return "\u2b1b"
    raw = raw[-6:].zfill(6)
    rgb = tuple(int(raw[index : index + 2], 16) for index in (0, 2, 4))
    palette = (
        ("\U0001f7e5", (255, 0, 0)),
        ("\U0001f7e7", (255, 128, 0)),
        ("\U0001f7e8", (255, 230, 0)),
        ("\U0001f7e9", (0, 180, 0)),
        ("\U0001f7e6", (0, 120, 255)),
        ("\U0001f7ea", (150, 65, 200)),
        ("\u2b1b", (0, 0, 0)),
        ("\u2b1c", (220, 220, 220)),
        ("\U0001f7eb", (120, 70, 25)),
    )
    return min(
        palette,
        key=lambda item: sum((rgb[index] - item[1][index]) ** 2 for index in range(3)),
    )[0]


def project_type_summary(project: dict[str, Any]) -> str:
    name = project_type_name(project)
    if not name:
        return "\u2b1b Type unknown"
    project_type = project.get("project_type")
    color = None
    if isinstance(project_type, dict):
        color = (
            project_type.get("color")
            or project_type.get("colour")
            or project_type.get("hex_color")
            or project_type.get("hexColour")
        )
    if color is None:
        color = PROJECT_TYPE_COLORS.get(name)
    if color is None:
        color = PROJECT_TYPE_COLORS.get(str(project_type or ""))
    return f"{nearest_color_emoji(color or '')} {name}"


def project_account_manager_name(project: dict[str, Any]) -> str:
    account_manager = project.get("account_manager")
    if isinstance(account_manager, dict):
        for key in ("displayname", "name", "vt_fullname", "full_name", "fullname", "email", "id"):
            value = account_manager.get(key)
            if value not in (None, ""):
                return str(value).strip()
    if account_manager not in (None, ""):
        return str(account_manager).strip()
    return "Not assigned"


def candidate_equipment_identifiers(item: dict[str, Any]) -> list[tuple[str, str]]:
    equipment = item.get("equipment") if isinstance(item.get("equipment"), dict) else {}
    candidates = {
        "project equipment name": item.get("name"),
        "project equipment display name": item.get("displayname"),
        "equipment code": equipment.get("code"),
        "equipment name": equipment.get("name"),
        "equipment display name": equipment.get("displayname"),
    }
    output: list[tuple[str, str]] = []
    for label, value in candidates.items():
        normalized = normalize(value)
        if normalized:
            output.append((label, normalized))
    return output


def match_equipment_item(
    item: dict[str, Any],
    targets: list[EquipmentTarget],
    match_mode: str,
) -> EquipmentMatch | None:
    candidates = candidate_equipment_identifiers(item)
    for target in targets:
        for _label, candidate in candidates:
            if match_mode == "contains":
                matched = any(
                    len(identifier) >= 3
                    and (identifier in candidate or candidate in identifier)
                    for identifier in target.identifiers
                )
            else:
                matched = candidate in target.identifiers
            if matched:
                display_name = (
                    str(item.get("displayname") or item.get("name") or candidate).strip()
                )
                quantity = item.get("quantity_total", item.get("quantity", ""))
                return EquipmentMatch(
                    project_equipment_id=item.get("id"),
                    display_name=display_name,
                    quantity=quantity,
                    matched_identifier=candidate,
                    target_row=target.row_number,
                )
    return None


class RentmanClient:
    def __init__(self, base_url: str, token: str, auth_header: str | None = None) -> None:
        self.base_url = base_url.rstrip("/")
        self.token = token
        self.auth_header = auth_header or f"Bearer {token}"

    def get_json(self, url_or_path: str, params: dict[str, Any] | None = None) -> dict[str, Any]:
        if url_or_path.startswith("http://") or url_or_path.startswith("https://"):
            url = url_or_path
        else:
            url = f"{self.base_url}/{url_or_path.lstrip('/')}"
        if params:
            url = f"{url}?{urllib.parse.urlencode(params)}"

        request = urllib.request.Request(
            url,
            headers={
                "Authorization": self.auth_header,
                "Accept": "application/json",
                "User-Agent": "rentman-led-check/1.0",
            },
        )
        for attempt in range(4):
            try:
                with urllib.request.urlopen(request, timeout=45) as response:
                    return json.loads(response.read().decode("utf-8"))
            except urllib.error.HTTPError as error:
                body = error.read().decode("utf-8", errors="replace")
                if error.code in {429, 500, 502, 503, 504} and attempt < 3:
                    retry_after = error.headers.get("Retry-After")
                    delay = int(retry_after) if retry_after and retry_after.isdigit() else 2**attempt
                    time.sleep(delay)
                    continue
                raise RuntimeError(f"Rentman API error {error.code}: {body}") from error
            except urllib.error.URLError as error:
                if attempt < 3:
                    time.sleep(2**attempt)
                    continue
                raise RuntimeError(f"Rentman API request failed: {error}") from error
        raise RuntimeError("Rentman API request failed after retries.")

    def list_endpoint(
        self,
        path: str,
        params: dict[str, Any] | None = None,
        path_vars: dict[str, Any] | None = None,
    ) -> list[dict[str, Any]]:
        items: list[dict[str, Any]] = []
        next_url: str | None = path
        next_params = dict(params or {})
        next_params.setdefault("limit", 300)

        while next_url:
            payload = self.get_json(next_url, next_params)
            data = payload if isinstance(payload, list) else payload.get("data", [])
            if not isinstance(data, list):
                raise RuntimeError(f"Unexpected Rentman list response for {path}: {payload}")
            items.extend(data)

            next_page_url = None if isinstance(payload, list) else payload.get("next_page_url")
            if next_page_url:
                if path_vars:
                    for key, value in path_vars.items():
                        next_page_url = next_page_url.replace("{" + key + "}", str(value))
                next_url = next_page_url
                next_params = {}
            else:
                next_url = None
        return items


def build_rentman_client() -> RentmanClient:
    token = os.environ.get("RENTMAN_API_TOKEN") or os.environ.get("RENTMAN_TOKEN")
    if not token:
        raise ConfigError("Set RENTMAN_API_TOKEN as a GitHub Actions secret.")
    return RentmanClient(
        base_url=os.environ.get("RENTMAN_API_BASE", "https://api.rentman.net"),
        token=token,
        auth_header=os.environ.get("RENTMAN_AUTH_HEADER"),
    )


def find_projects_to_warn(
    client: RentmanClient,
    targets: list[EquipmentTarget],
    today: dt.date,
    timezone: dt.tzinfo,
    lookahead_days: int | None,
    required_tag: str,
    match_mode: str,
) -> list[Finding]:
    excluded_statuses = {normalize(value) for value in split_csv_env("EXCLUDED_STATUSES", tuple())}
    projects = client.list_endpoint(
        "projects",
        {
            "fields": os.environ.get("PROJECT_FIELDS", DEFAULT_PROJECT_FIELDS),
            "expand": os.environ.get("PROJECT_EXPAND", DEFAULT_PROJECT_EXPAND),
        },
    )
    findings: list[Finding] = []
    for project in projects:
        if excluded_statuses and project_status(project) in excluded_statuses:
            continue
        if not project_is_upcoming(project, today, timezone, lookahead_days):
            continue
        project["_subprojects"] = client.list_endpoint(
            f"projects/{project['id']}/subprojects",
            {
                "fields": os.environ.get("SUBPROJECT_FIELDS", DEFAULT_SUBPROJECT_FIELDS),
                "expand": os.environ.get("SUBPROJECT_EXPAND", DEFAULT_SUBPROJECT_EXPAND),
            },
            path_vars={"id": project["id"]},
        )
        if project_is_cancelled(project):
            continue
        if excluded_statuses and project_status(project) in excluded_statuses:
            continue
        if tag_is_present(project.get("tags"), required_tag):
            continue

        equipment_items = client.list_endpoint(
            f"projects/{project['id']}/projectequipment",
            {
                "fields": os.environ.get("PROJECT_EQUIPMENT_FIELDS", DEFAULT_PROJECT_EQUIPMENT_FIELDS),
                "expand": "equipment",
            },
            path_vars={"id": project["id"]},
        )
        matches = tuple(
            match
            for item in equipment_items
            if (match := match_equipment_item(item, targets, match_mode)) is not None
        )
        if matches:
            findings.append(Finding(project=project, matches=matches))
    return findings


def project_label(project: dict[str, Any]) -> str:
    number = project.get("number")
    name = project.get("displayname") or project.get("name") or "Unnamed project"
    url = project_url(project)
    if number not in (None, ""):
        number_label = f"#{number}"
        if url:
            number_label = f"<{url}|{number_label}>"
        return f"{number_label} {name}"
    if url:
        return f"<{url}|{name}>"
    return str(name)


def format_human_datetime(value: Any, timezone: dt.tzinfo) -> str:
    parsed = parse_datetime(value, timezone)
    if parsed is None:
        return ""
    return parsed.strftime("%d/%m/%Y %H:%M")


def project_date_summary(project: dict[str, Any], timezone: dt.tzinfo) -> str:
    start = format_human_datetime(
        project.get("planperiod_start") or project.get("usageperiod_start"),
        timezone,
    )
    end = format_human_datetime(
        project.get("planperiod_end") or project.get("usageperiod_end"),
        timezone,
    )
    if start and end:
        return f"{start} to {end}"
    return start or end or "date not set"


def project_url(project: dict[str, Any]) -> str | None:
    template = os.environ.get("RENTMAN_PROJECT_URL_TEMPLATE", DEFAULT_PROJECT_URL_TEMPLATE)
    if not template:
        return None
    return template.format(**{key: value or "" for key, value in project.items()})


def format_slack_message(findings: list[Finding], required_tag: str, timezone: dt.tzinfo) -> str:
    if not findings:
        return "Rentman LED check: no upcoming projects need the LED tech tag."

    lines = [
        f":warning: Rentman LED check: {len(findings)} upcoming project(s) include watched LED equipment "
        f"but are missing the `{required_tag}` tag."
    ]
    for finding in findings:
        project = finding.project
        label = project_label(project)
        lines.append(f"\n*{label}*")
        status_emoji, status_label = project_status_details(project)
        lines.append(f"Status: {status_emoji} {status_label}")
        lines.append(f"Type: {project_type_summary(project)}")
        lines.append(f"Account manager: {project_account_manager_name(project)}")
        lines.append(f"Planning period: {project_date_summary(project, timezone)}")
        lines.append("Matched equipment:")
        for match in finding.matches[:10]:
            qty = f" x{match.quantity}" if match.quantity not in (None, "") else ""
            lines.append(f"- {match.display_name}{qty} (equipment sheet row {match.target_row})")
        if len(finding.matches) > 10:
            lines.append(f"- ...and {len(finding.matches) - 10} more")
    return "\n".join(lines)


def send_slack_webhook(webhook_url: str, text: str) -> None:
    payload = json.dumps({"text": text}).encode("utf-8")
    request = urllib.request.Request(
        webhook_url,
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=30) as response:
        if response.status >= 300:
            raise RuntimeError(f"Slack webhook returned HTTP {response.status}")


def send_slack_bot_message(bot_token: str, user_id: str, text: str) -> None:
    payload = json.dumps({"channel": user_id, "text": text}).encode("utf-8")
    request = urllib.request.Request(
        "https://slack.com/api/chat.postMessage",
        data=payload,
        headers={
            "Authorization": f"Bearer {bot_token}",
            "Content-Type": "application/json; charset=utf-8",
        },
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=30) as response:
        body = json.loads(response.read().decode("utf-8"))
    if not body.get("ok"):
        raise RuntimeError(f"Slack chat.postMessage failed: {body}")


def send_slack_message(text: str) -> None:
    bot_token = os.environ.get("SLACK_BOT_TOKEN")
    user_id = os.environ.get("SLACK_USER_ID")
    webhook_url = os.environ.get("SLACK_WEBHOOK_URL")

    if bot_token and user_id:
        send_slack_bot_message(bot_token, user_id, text)
        return
    if webhook_url:
        send_slack_webhook(webhook_url, text)
        return
    raise ConfigError(
        "Set either SLACK_BOT_TOKEN + SLACK_USER_ID, or SLACK_WEBHOOK_URL, as GitHub Actions secrets/variables."
    )


def local_timezone() -> dt.tzinfo:
    timezone_name = os.environ.get("TIMEZONE", "Pacific/Auckland")
    try:
        from zoneinfo import ZoneInfo

        return ZoneInfo(timezone_name)
    except Exception:
        print(f"Warning: could not load timezone {timezone_name}; using UTC.", file=sys.stderr)
        return dt.timezone.utc


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--equipment-file",
        default=os.environ.get("EQUIPMENT_FILE", str(DEFAULT_EQUIPMENT_FILE)),
        help="Path to the XLSX containing watched equipment.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        default=os.environ.get("DRY_RUN", "").casefold() == "true",
        help="Print the Slack message instead of posting it.",
    )
    parser.add_argument(
        "--print-targets",
        action="store_true",
        help="Print parsed equipment targets and exit.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    equipment_file = Path(args.equipment_file)
    match_columns = split_csv_env("EQUIPMENT_MATCH_COLUMNS", DEFAULT_MATCH_COLUMNS)
    match_mode = os.environ.get("MATCH_MODE", "exact").strip().casefold()
    if match_mode not in {"exact", "contains"}:
        raise ConfigError("MATCH_MODE must be either 'exact' or 'contains'.")

    targets = load_equipment_targets(equipment_file, match_columns)
    if args.print_targets:
        print(
            json.dumps(
                [
                    {
                        "row_number": target.row_number,
                        "values": target.values,
                        "identifiers": sorted(target.identifiers),
                    }
                    for target in targets
                ],
                indent=2,
            )
        )
        return 0

    timezone = local_timezone()
    today = dt.datetime.now(timezone).date()
    lookahead_days = parse_int_env("LOOKAHEAD_DAYS", None)
    required_tag = os.environ.get("REQUIRED_TAG", DEFAULT_REQUIRED_TAG)

    client = build_rentman_client()
    findings = find_projects_to_warn(
        client=client,
        targets=targets,
        today=today,
        timezone=timezone,
        lookahead_days=lookahead_days,
        required_tag=required_tag,
        match_mode=match_mode,
    )
    message = format_slack_message(findings, required_tag, timezone)

    print(
        json.dumps(
            {
                "date": today.isoformat(),
                "targets": len(targets),
                "findings": len(findings),
                "dry_run": args.dry_run,
            },
            indent=2,
        )
    )

    should_notify_on_empty = os.environ.get("SLACK_NOTIFY_ON_EMPTY", "").casefold() == "true"
    if args.dry_run:
        print("\n--- Slack message preview ---")
        print(message)
        return 0
    if findings or should_notify_on_empty:
        send_slack_message(message)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except ConfigError as error:
        print(f"Configuration error: {error}", file=sys.stderr)
        raise SystemExit(2)
