from pathlib import Path
from datetime import datetime
import pandas as pd
from PIL import Image

# =============================================================================
# 0. PATHS
# =============================================================================

BASE_DIR = Path(r"C:\Users\agonz\OneDrive\Documentos\White Door Closet")

EXCEL_PATH = BASE_DIR / "inventory.xlsx"
JPG_FOLDER = BASE_DIR / "jpg_photos"

OUTPUT_HTML = BASE_DIR / "photo_audit_gallery.html"
OUTPUT_ISSUES = BASE_DIR / "photo_audit_issues.xlsx"

PHOTO_ARCHIVE_SHEET = "photo_archive"
INVENTORY_SHEET = "inventory_vf"

RUN_TIME = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

DROP_PHOTO_TYPE = "drop"


# =============================================================================
# 1. READ EXCEL
# =============================================================================

photo_archive = pd.read_excel(EXCEL_PATH, sheet_name=PHOTO_ARCHIVE_SHEET)
inventory = pd.read_excel(EXCEL_PATH, sheet_name=INVENTORY_SHEET)

photo_archive.columns = photo_archive.columns.str.strip()
inventory.columns = inventory.columns.str.strip()

photo_archive["ITEM_ID"] = photo_archive["ITEM_ID"].astype(str).str.strip()
photo_archive["photo_type"] = photo_archive["photo_type"].astype(str).str.strip().str.lower()
inventory["ITEM_ID"] = inventory["ITEM_ID"].astype(str).str.strip()

# Make sure NOTAG_REQUIRED exists
if "NOTAG_REQUIRED" not in inventory.columns:
    inventory["NOTAG_REQUIRED"] = ""

inventory_cols = [
    "ITEM_ID", "DESCRIPTION", "CATEGORY", "BRAND", "SIZE", "SIZE (SML)",
    "PRICE", "PHOTO", "NOTAG_REQUIRED", "POST_DATE", "SOLD_DATE",
    "BUYER", "PICKUP", "URL",
]

inventory_cols = [c for c in inventory_cols if c in inventory.columns]

df_raw = photo_archive.merge(
    inventory[inventory_cols],
    on="ITEM_ID",
    how="left",
    indicator=True
)


# =============================================================================
# 2. DROP PHOTOS THAT SHOULD NOT ENTER THE AUDIT
# =============================================================================

df = df_raw[df_raw["photo_type"] != DROP_PHOTO_TYPE].copy()

sort_cols = ["ITEM_ID", "photo_type"]
if "datetime" in df.columns:
    sort_cols.append("datetime")

df = df.sort_values(sort_cols, na_position="last")
df = df.drop_duplicates(subset=["ITEM_ID", "photo_type"], keep="first").copy()


# =============================================================================
# 3. IMAGE DIMENSIONS
# =============================================================================

def get_image_info(filename):
    if pd.isna(filename):
        return pd.Series([None, None, "missing_filename", "missing_filename"])

    path = JPG_FOLDER / str(filename)

    if not path.exists():
        return pd.Series([None, None, "missing_file", "missing_file"])

    try:
        with Image.open(path) as img:
            width, height = img.size

        if height > width:
            orientation = "vertical"
        elif width > height:
            orientation = "horizontal"
        else:
            orientation = "square"

        return pd.Series([width, height, orientation, "found"])

    except Exception as e:
        return pd.Series([None, None, "error", f"error: {e}"])


df[["width", "height", "orientation", "file_status"]] = (
    df["converted_name"].apply(get_image_info)
)


# =============================================================================
# 4. CREATE ITEM-LEVEL ISSUES
# =============================================================================

def clean(x):
    if pd.isna(x):
        return ""
    return str(x)


def flag_is_1(x):
    if pd.isna(x):
        return False
    return str(x).strip() in ["1", "1.0", "yes", "YES", "true", "TRUE", "x", "X"]


def make_issue(group, issue_type, issue_detail, suggested_action):
    first = group.iloc[0]

    photo_types_present = ", ".join(
        sorted(group["photo_type"].dropna().astype(str).unique())
    ) if "photo_type" in group.columns else ""

    files_present = ", ".join(
        sorted(group["converted_name"].dropna().astype(str).unique())
    ) if "converted_name" in group.columns else ""

    return {
        "ITEM_ID": clean(first.get("ITEM_ID", "")),
        "issue_type": issue_type,
        "issue_status": "open",
        "DESCRIPTION": clean(first.get("DESCRIPTION", "")),
        "CATEGORY": clean(first.get("CATEGORY", "")),
        "BRAND": clean(first.get("BRAND", "")),
        "SIZE": clean(first.get("SIZE", "")),
        "PRICE": clean(first.get("PRICE", "")),
        "suggested_action": suggested_action,
        "issue_detail": issue_detail,
        "photo_types_present": photo_types_present,
        "files_present": files_present,
        "manual_issue": "",
        "notes": "",
        "last_seen": RUN_TIME,
    }


issues = []

# A) Items in photo_archive but not inventory_vf
for item_id, group in df.groupby("ITEM_ID", sort=True):
    if (group["_merge"] != "both").any():
        issues.append(make_issue(
            group,
            "item_not_found_in_inventory",
            "ITEM_ID appears in photo_archive but not in inventory_vf.",
            "Correct ITEM_ID in photo_archive or add the item to inventory_vf."
        ))

# B) Missing JPG files
for item_id, group in df.groupby("ITEM_ID", sort=True):
    if (group["file_status"] == "missing_file").any():
        issues.append(make_issue(
            group,
            "missing_jpg_file",
            "At least one converted_name does not exist in jpg_photos.",
            "Check whether the JPG was moved, deleted, or incorrectly named."
        ))

# C) Missing required photo types
# Rules:
# - front is always required
# - back is never required
# - tag is required unless:
#     1) CATEGORY is zapatos/shoes, OR
#     2) NOTAG_REQUIRED == 1 in inventory_vf

for item_id, group in df.groupby("ITEM_ID", sort=True):
    first = group.iloc[0]
    category = clean(first.get("CATEGORY", "")).strip().lower()
    existing_types = set(group["photo_type"].dropna().astype(str).str.lower())

    if "front" not in existing_types:
        issues.append(make_issue(
            group,
            "missing_front",
            "Missing required photo_type: front.",
            "Take or assign a front photo for this item."
        ))

    is_shoes = category in ["zapatos", "shoes"]
    notag_required = flag_is_1(first.get("NOTAG_REQUIRED", ""))

    if (not is_shoes) and (not notag_required) and ("tag" not in existing_types):
        issues.append(make_issue(
            group,
            "missing_tag",
            "Missing required photo_type: tag.",
            "Take or assign a tag photo for this item, or set NOTAG_REQUIRED = 1 if no tag photo is needed."
        ))

# D) Mixed orientation only
for item_id, group in df.groupby("ITEM_ID", sort=True):
    orientations = set(
        group.loc[
            group["orientation"].isin(["vertical", "horizontal", "square"]),
            "orientation"
        ]
    )

    if len(orientations) > 1:
        issues.append(make_issue(
            group,
            "mixed_orientation",
            "Item has mixed vertical/horizontal/square images.",
            "Edit/crop/retake the inconsistent photo so all photos have the same orientation."
        ))

# E) Inventory items with no photos
# Rule:
# - If item has SOLD_DATE, it is not an issue.

photo_items = set(df["ITEM_ID"].dropna().astype(str))

for _, row in inventory.iterrows():
    item_id = clean(row.get("ITEM_ID", ""))
    sold_date = row.get("SOLD_DATE", "")

    item_is_sold = (
        pd.notna(sold_date)
        and clean(sold_date).strip() not in ["", "NaT", "nan", "None"]
    )

    if item_id and item_id not in photo_items and not item_is_sold:
        fake_group = pd.DataFrame([row])
        issues.append(make_issue(
            fake_group,
            "inventory_item_without_photos",
            "ITEM_ID appears in inventory_vf but has no photos in photo_archive.",
            "Take photos or check whether ITEM_ID was mistyped."
        ))

issues_df = pd.DataFrame(issues)

issue_cols = [
    "ITEM_ID", "issue_type", "issue_status", "DESCRIPTION", "CATEGORY",
    "BRAND", "SIZE", "PRICE", "suggested_action", "issue_detail",
    "photo_types_present", "files_present", "manual_issue", "notes",
    "last_seen",
]

if issues_df.empty:
    issues_df = pd.DataFrame(columns=issue_cols)


# =============================================================================
# 5. PRESERVE OLD MANUAL NOTES IF FILE EXISTS AND THERE ARE ISSUES
# =============================================================================

if not issues_df.empty and OUTPUT_ISSUES.exists():
    try:
        old = pd.read_excel(OUTPUT_ISSUES, sheet_name="item_issues")
        old.columns = old.columns.str.strip()

        key_cols = ["ITEM_ID", "issue_type"]
        keep_cols = key_cols + [
            c for c in ["issue_status", "manual_issue", "notes"]
            if c in old.columns
        ]

        old_keep = old[keep_cols].drop_duplicates(key_cols)

        issues_df = issues_df.merge(
            old_keep,
            on=key_cols,
            how="left",
            suffixes=("", "_old")
        )

        for col in ["issue_status", "manual_issue", "notes"]:
            old_col = f"{col}_old"
            if old_col in issues_df.columns:
                issues_df[col] = issues_df[old_col].combine_first(issues_df[col])
                issues_df = issues_df.drop(columns=[old_col])

    except Exception as e:
        print(f"Warning: could not preserve old manual notes: {e}")


if not issues_df.empty:
    issues_df = issues_df.drop_duplicates(["ITEM_ID", "issue_type"], keep="first")
    issues_df = issues_df.sort_values(["ITEM_ID", "issue_type"]).reset_index(drop=True)


# =============================================================================
# 6. SAVE OR DELETE ISSUES EXCEL
# =============================================================================

if issues_df.empty:
    if OUTPUT_ISSUES.exists():
        OUTPUT_ISSUES.unlink()
        issues_file_message = "No issues found. Existing issues Excel was deleted."
    else:
        issues_file_message = "No issues found. No issues Excel was created."
else:
    summary_df = (
        issues_df
        .groupby(["issue_type", "issue_status"], dropna=False)
        .size()
        .reset_index(name="count")
        .sort_values(["issue_type", "issue_status"])
    )

    with pd.ExcelWriter(OUTPUT_ISSUES, engine="openpyxl") as writer:
        issues_df.to_excel(writer, sheet_name="item_issues", index=False)
        summary_df.to_excel(writer, sheet_name="summary", index=False)

    issues_file_message = f"Issues Excel created/updated: {OUTPUT_ISSUES}"


# =============================================================================
# 7. BUILD HTML
# =============================================================================

open_issues = issues_df[
    ~issues_df["issue_status"].astype(str).str.lower().isin(["fixed", "ignore"])
].copy()

html = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Photo Audit Gallery</title>

<style>
body {
    font-family: Arial, sans-serif;
    margin: 24px;
    background: #f7f7f7;
    color: #222;
}

h1 { margin-bottom: 4px; }

.summary {
    margin-bottom: 24px;
    color: #555;
}

.item {
    background: white;
    border: 1px solid #ddd;
    border-radius: 12px;
    padding: 14px;
    margin-bottom: 18px;
}

.item-warning { border-left: 8px solid #b00020; }

.header {
    font-weight: bold;
    font-size: 17px;
    margin-bottom: 6px;
}

.meta {
    color: #555;
    font-size: 13px;
    line-height: 1.5;
    margin-bottom: 8px;
}

.ok {
    color: #137333;
    font-weight: bold;
}

.warning {
    color: #b00020;
    font-weight: bold;
}

.photos {
    display: flex;
    gap: 12px;
    flex-wrap: wrap;
    margin-top: 12px;
}

.photo-card {
    width: 145px;
    font-size: 12px;
    text-align: center;
    word-wrap: break-word;
}

.photo-card img {
    max-width: 145px;
    max-height: 190px;
    border: 1px solid #ccc;
    border-radius: 8px;
    object-fit: contain;
    background: #eee;
}

.missing-box {
    width: 145px;
    height: 190px;
    background: #ddd;
    border: 1px solid #bbb;
    border-radius: 8px;
    display: flex;
    align-items: center;
    justify-content: center;
    color: #777;
}

.small {
    color: #666;
    font-size: 11px;
}
</style>
</head>

<body>
<h1>Photo Audit Gallery</h1>
<div class="summary">
Visual check for <b>photo_archive</b> merged with <b>inventory_vf</b>.<br>
Photos marked as <b>drop</b> and duplicated <b>ITEM_ID + photo_type</b> rows are excluded from this gallery.<br>
If <b>NOTAG_REQUIRED = 1</b>, a missing tag photo is not treated as an issue.
</div>
"""

html += f"""
<div class="item">
    <div class="header">Summary</div>
    <div class="meta">
        Total visible items: <b>{df["ITEM_ID"].nunique()}</b><br>
        Total visible photo rows: <b>{len(df)}</b><br>
        Open / non-ignored issues: <b>{len(open_issues)}</b><br>
        Issues file status: <b>{issues_file_message}</b>
    </div>
</div>
"""

for item_id, group in df.groupby("ITEM_ID", sort=True):
    first = group.iloc[0]

    item_issue_types = open_issues.loc[
        open_issues["ITEM_ID"].astype(str) == str(item_id),
        "issue_type"
    ].dropna().astype(str).tolist()

    has_issues = len(item_issue_types) > 0
    card_class = "item item-warning" if has_issues else "item"

    status = (
        "<span class='ok'>✓ OK</span>"
        if not has_issues
        else "<span class='warning'>⚠ " + " | ".join(item_issue_types) + "</span>"
    )

    html += f"""
    <div class="{card_class}">
        <div class="header">{item_id} — {first.get("DESCRIPTION", "")}</div>

        <div class="meta">
            Category: <b>{first.get("CATEGORY", "")}</b> |
            Brand: <b>{first.get("BRAND", "")}</b> |
            Size: <b>{first.get("SIZE", "")}</b> |
            SML: <b>{first.get("SIZE (SML)", "")}</b> |
            Price: <b>${first.get("PRICE", "")}</b><br>

            Expected photos: <b>{first.get("PHOTO", "")}</b> |
            NOTAG_REQUIRED: <b>{first.get("NOTAG_REQUIRED", "")}</b> |
            Visible photo rows: <b>{len(group)}</b> |
            Post date: <b>{first.get("POST_DATE", "")}</b> |
            Sold date: <b>{first.get("SOLD_DATE", "")}</b> |
            Buyer: <b>{first.get("BUYER", "")}</b> |
            Pickup: <b>{first.get("PICKUP", "")}</b><br>

            URL: <a href="{first.get("URL", "")}" target="_blank">{first.get("URL", "")}</a>
        </div>

        <div>{status}</div>

        <div class="photos">
    """

    for _, row in group.sort_values("photo_type").iterrows():
        filename = row.get("converted_name", "")
        img_path = JPG_FOLDER / str(filename)

        if img_path.exists():
            img_uri = img_path.as_uri()
            img_tag = f"<a href='{img_uri}' target='_blank'><img src='{img_uri}'></a>"
        else:
            img_tag = "<div class='missing-box'>Missing</div>"

        html += f"""
            <div class="photo-card">
                {img_tag}
                <div><b>{row.get("photo_type", "")}</b></div>
                <div class="small">{filename}</div>
                <div class="small">{row.get("width", "")} x {row.get("height", "")}</div>
                <div class="small">{row.get("orientation", "")}</div>
                <div class="small">{row.get("file_status", "")}</div>
            </div>
        """

    html += """
        </div>
    </div>
    """

html += """
</body>
</html>
"""


# =============================================================================
# 8. SAVE HTML
# =============================================================================

OUTPUT_HTML.write_text(html, encoding="utf-8")

print("Done!")
print("Open HTML:")
print(OUTPUT_HTML)
print()
print(issues_file_message)