#!/usr/bin/env python3

import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import date
from pathlib import Path


NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def infer_funnel(campaign_name):
    if "_UF_" in campaign_name:
        return "UF"
    if "_MF_" in campaign_name:
        return "MF"
    if "_LF_" in campaign_name:
        return "LF"
    return "Other"


def infer_status(spend, purchases):
    if spend <= 0:
        return "Off"
    if spend > 0 or purchases > 0:
        return "Live"
    return "Paused"


def get_inactivity_thresholds(funnel):
    if funnel == "MF":
        return {"spend_days": 3, "impression_days": 5}
    return {"spend_days": 1, "impression_days": 3}


def parse_iso_date(value):
    try:
        year, month, day = (int(part) for part in value.split("-"))
        return date(year, month, day)
    except Exception:
        return None


def simplify_campaign_name(campaign_name):
    if campaign_name.startswith("RTG_"):
        return "RTG"
    if campaign_name.startswith("DABA_"):
        return "DABA"

    parts = campaign_name.split("_")
    funnel = next((part for part in parts if part in {"UF", "MF", "LF"}), "")

    if "Active" in parts:
        audience = "Active"
    elif "Prospecting" in parts:
        audience = "Prospecting"
    elif "Inactive" in parts:
        audience = "Inactive"
    else:
        audience = ""

    tactic = ""
    if "DABA" in parts:
        tactic = "DABA"
    elif "Remarketing" in parts:
        tactic = "Remarketing"
    elif "Reactivation" in parts:
        tactic = "Reactivation"
    elif "Traffic" in parts:
        tactic = "Traffic"
    elif "Awareness" in parts:
        tactic = "Awareness"

    variant = ""
    if "H&W+Core" in parts:
        variant = "H&W+Core"
    elif "Core BAU" in campaign_name:
        variant = "Core BAU"

    label = " ".join(part for part in [funnel, audience, tactic, variant] if part)
    return label or campaign_name


def simplify_ad_name(ad_name):
    parts = [part.strip() for part in ad_name.split("_") if part.strip()]
    creative_title = parts[-1] if parts else ad_name.strip()
    has_inf = re.search(r"(^|[^A-Za-z])(inf|infl|influencer)([^A-Za-z]|$)", ad_name, re.IGNORECASE)

    length_match = re.search(r"_Dynamic_(\d+)_", ad_name)
    if not length_match:
        length_match = re.search(r"_(\d+)(?:_|$)", ad_name)

    if not length_match:
        return f"INF {creative_title}" if has_inf else creative_title

    length_value = length_match.group(1)
    length_label = "Static" if length_value == "0" else length_value
    label = f"{creative_title} | {length_label}"
    return f"INF {label}" if has_inf else label


def derive_version_campaign(ad_name):
    if "_Food_" in ad_name and "_No Food_" not in ad_name:
        return "Food"

    parts = [part.strip() for part in ad_name.split("_") if part.strip()]
    if not parts:
        return ""

    if len(parts) >= 3 and parts[0] == "IMC" and parts[1] == "Studio":
        words = parts[2].split()
        return " ".join(words[:2]) if words else parts[2]

    return parts[0]


def derive_fiscal_launch(ad_name):
    match = re.search(r"_(FY\d{2}P\d{2}W\d)_", ad_name)
    if match:
        return match.group(1)
    return ""


def read_shared_strings(archive):
    shared = []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    for si in root.findall("a:si", NS):
        text = "".join(node.text or "" for node in si.iterfind(".//a:t", NS))
        shared.append(text)
    return shared


def cell_value(cell, shared):
    cell_type = cell.attrib.get("t")
    value = cell.find("a:v", NS)
    inline = cell.find("a:is", NS)

    if cell_type == "s" and value is not None and value.text is not None:
        return shared[int(value.text)]
    if cell_type == "inlineStr" and inline is not None:
        return "".join(node.text or "" for node in inline.iterfind(".//a:t", NS))
    if value is not None and value.text is not None:
        return value.text
    return ""


def parse_sheet(path):
    with zipfile.ZipFile(path) as archive:
        shared = read_shared_strings(archive)
        header = None
        rows = []

        for _, elem in ET.iterparse(archive.open("xl/worksheets/sheet1.xml"), events=("end",)):
            if not elem.tag.endswith("row"):
                continue

            values_by_col = {}
            for cell in elem.findall("a:c", NS):
                ref = cell.attrib.get("r", "")
                match = re.match(r"([A-Z]+)", ref)
                if not match:
                    continue
                values_by_col[match.group(1)] = cell_value(cell, shared)

            ordered_cols = sorted(values_by_col.keys(), key=lambda key: (len(key), key))
            ordered_values = [values_by_col[col] for col in ordered_cols]

            if header is None:
                header = ordered_values
            else:
                rows.append(dict(zip(header, ordered_values)))

            elem.clear()

    return rows


def build_preview_lookup(preview_rows):
    preview_lookup = {}

    for row in preview_rows:
        ad_name = row.get("Ad name", "").strip()
        preview_link = row.get("Preview link", "").strip()
        if not ad_name or not preview_link:
            continue

        display_ad_name = simplify_ad_name(ad_name)
        key = display_ad_name
        preview_lookup.setdefault(key, preview_link)

    return preview_lookup


def to_float(value):
    try:
        return float(value or 0)
    except ValueError:
        return 0.0


def build_creative_rows(raw_rows, preview_lookup=None):
    latest_export_day = None
    launch_dates_by_visible_creative = {}
    preview_lookup = preview_lookup or {}
    grouped = defaultdict(
        lambda: {
            "launchDate": "",
            "launchCode": "",
            "funnel": "",
            "versionCampaign": "",
            "campaign": "",
            "adSet": "",
            "adName": "",
            "previewUrl": "",
            "spend": 0.0,
            "purchases": 0.0,
            "revenue": 0.0,
            "status": "Paused",
            "notes": "",
            "lastSpendDate": "",
            "lastImpressionDate": "",
            "firstSeenDate": "",
            "firstSpendDate": ""
        }
    )

    for row in raw_rows:
        campaign = row.get("Campaign name", "").strip()
        ad_set = row.get("Ad set name", "").strip()
        ad_name = row.get("Ad name", "").strip()
        day = row.get("Day", "").strip()
        spend = to_float(row.get("Amount spent (USD)", 0))
        impressions = to_float(row.get("Impressions", 0))
        purchases = to_float(row.get("Purchases", 0))
        revenue = to_float(row.get("Purchases conversion value", 0))
        day_value = parse_iso_date(day)

        if day_value and (latest_export_day is None or day_value > latest_export_day):
            latest_export_day = day_value

        if not campaign or not ad_set or not ad_name:
            continue

        funnel = infer_funnel(campaign)
        version_campaign = derive_version_campaign(ad_name)
        display_campaign = simplify_campaign_name(campaign)
        display_ad_name = simplify_ad_name(ad_name)
        key = (funnel, version_campaign, display_campaign, ad_set, display_ad_name)
        launch_key = (funnel, version_campaign, ad_set, display_ad_name)
        record = grouped[key]
        record["funnel"] = funnel
        record["launchCode"] = derive_fiscal_launch(ad_name)
        record["versionCampaign"] = version_campaign
        record["campaign"] = display_campaign
        record["adSet"] = ad_set
        record["adName"] = display_ad_name
        preview_key = display_ad_name
        if not record["previewUrl"]:
            record["previewUrl"] = preview_lookup.get(preview_key, "")
        record["spend"] += spend
        record["purchases"] += purchases
        record["revenue"] += revenue

        if day:
            if not record["firstSeenDate"] or day < record["firstSeenDate"]:
                record["firstSeenDate"] = day

        if spend > 0 and day:
            if not record["firstSpendDate"] or day < record["firstSpendDate"]:
                record["firstSpendDate"] = day
            record["lastSpendDate"] = max(record["lastSpendDate"], day)

            earliest_launch = launch_dates_by_visible_creative.get(launch_key)
            if not earliest_launch or day < earliest_launch:
                launch_dates_by_visible_creative[launch_key] = day

        if impressions > 0 and day:
            record["lastImpressionDate"] = max(record["lastImpressionDate"], day)

    creative_rows = []
    for record in grouped.values():
        launch_key = (
            record["funnel"],
            record["versionCampaign"],
            record["adSet"],
            record["adName"]
        )
        record["launchDate"] = (
            launch_dates_by_visible_creative.get(launch_key)
            or record["firstSpendDate"]
            or record["firstSeenDate"]
        )
        record["spend"] = round(record["spend"], 2)
        record["purchases"] = int(record["purchases"])
        record["revenue"] = round(record["revenue"], 2)
        last_spend_day = parse_iso_date(record["lastSpendDate"])
        last_impression_day = parse_iso_date(record["lastImpressionDate"])
        thresholds = get_inactivity_thresholds(record["funnel"])

        if record["spend"] <= 0:
            record["status"] = "Off"
        elif (
            latest_export_day
            and last_spend_day
            and last_impression_day
            and (latest_export_day - last_spend_day).days >= thresholds["spend_days"]
            and (latest_export_day - last_impression_day).days >= thresholds["impression_days"]
        ):
            record["status"] = "Off"
        else:
            record["status"] = infer_status(record["spend"], record["purchases"])

        record["notes"] = "Imported from Meta export. Funnel inferred from campaign naming."
        if record["status"] == "Off" and record["spend"] <= 0:
            record["notes"] += " Zero spend in the export window."
        elif record["status"] == "Off" and record["lastSpendDate"]:
            impression_note = (
                f" and no impressions after {record['lastImpressionDate']}"
                if record["lastImpressionDate"]
                else ""
            )
            record["notes"] += (
                f" No spend for {thresholds['spend_days']}+ day"
                f"{'s' if thresholds['spend_days'] != 1 else ''} after {record['lastSpendDate']}"
                f"{impression_note}. Threshold uses {thresholds['impression_days']}+ days without impressions."
            )

        record.pop("lastSpendDate", None)
        record.pop("lastImpressionDate", None)
        record.pop("firstSeenDate", None)
        record.pop("firstSpendDate", None)
        creative_rows.append(record)

    creative_rows.sort(key=lambda row: (row["funnel"], row["campaign"], row["adSet"], row["adName"]))
    return creative_rows, latest_export_day.isoformat() if latest_export_day else ""


def write_data_js(output_path, rows, last_data_date):
    payload = {
        "meta": {
            "lastDataDate": last_data_date
        },
        "rows": rows
    }
    output = "window.metaCreativeData = " + json.dumps(payload, indent=2) + ";\n"
    output += "window.metaCreativeSeedRows = window.metaCreativeData.rows;\n"
    output_path.write_text(output, encoding="utf-8")


def main():
    if len(sys.argv) < 2:
        print("Usage: import_meta_xlsx.py <xlsx-path> [output-js-path]")
        raise SystemExit(1)

    source = Path(sys.argv[1]).expanduser()
    output = Path(sys.argv[2]).expanduser() if len(sys.argv) > 2 else Path("data.js")
    preview_source = Path(sys.argv[3]).expanduser() if len(sys.argv) > 3 else None

    rows = parse_sheet(source)
    preview_lookup = {}
    if preview_source and preview_source.exists():
        preview_lookup = build_preview_lookup(parse_sheet(preview_source))

    creative_rows, last_data_date = build_creative_rows(rows, preview_lookup)
    write_data_js(output, creative_rows, last_data_date)

    print(f"Imported {len(creative_rows)} creative rows into {output}")


if __name__ == "__main__":
    main()
