from pathlib import Path
import re
import html
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# =============================================================================
# 0. PATHS
# =============================================================================

BASE_DIR = Path(r"C:\Users\agonz\OneDrive\Documentos\White Door Closet")

EXCEL_PATH = BASE_DIR / "inventory.xlsx"
JPG_FOLDER = BASE_DIR / "jpg_photos"
COLLAGE_FOLDER = BASE_DIR / "collages"
OUTPUT_HTML = BASE_DIR / "collage_gallery.html"

PHOTO_ARCHIVE_SHEET = "photo_archive"
INVENTORY_SHEET = "inventory_vf"

COLLAGE_FOLDER.mkdir(exist_ok=True)

# =============================================================================
# 1. SETTINGS
# =============================================================================

TEST_RUN = 0
TEST_N_ITEMS = 5

if TEST_RUN == 1 and (TEST_N_ITEMS is None or TEST_N_ITEMS <= 0):
    raise ValueError("If TEST_RUN = 1, you must set TEST_N_ITEMS to a positive number, e.g. 5.")

CANVAS_W = 2400
CANVAS_H = 1000

MARGIN_X = 80
GAP = 30

HEADER_H = 140
PHOTO_H = 520
LABEL_H = 60
FOOTER_H = 130

PHOTO_Y = HEADER_H
LABEL_Y = PHOTO_Y + PHOTO_H + 10
FOOTER_Y = LABEL_Y + LABEL_H
LEGEND_Y = FOOTER_Y + 75

SLOT_W = int((CANVAS_W - 2 * MARGIN_X - 2 * GAP) / 3)
SLOT_H = PHOTO_H

WHITE = "white"
BLACK = "black"
GRAY = (245, 245, 245)


# =============================================================================
# 2. FONTS
# =============================================================================

def get_font(size, bold=False):
    font_dir = Path(r"C:\Windows\Fonts")

    candidates = [
        font_dir / ("arialbd.ttf" if bold else "arial.ttf"),
        font_dir / ("Arial Bold.ttf" if bold else "Arial.ttf"),
    ]

    for font_path in candidates:
        if font_path.exists():
            return ImageFont.truetype(str(font_path), size=size)

    return ImageFont.load_default()


FONT_TITLE = get_font(58, bold=True)
FONT_LABEL = get_font(32, bold=True)
FONT_FOOTER = get_font(38, bold=False)
FONT_FOOTER_BOLD = get_font(38, bold=True)
FONT_LEGEND = get_font(30, bold=True)


# =============================================================================
# 3. HELPERS
# =============================================================================

def clean(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def is_blank(x):
    return pd.isna(x) or clean(x) in ["", "NaT", "nan", "None"]


def flag_is_1(x):
    if pd.isna(x):
        return False
    return clean(x).lower() in ["1", "1.0", "yes", "true", "x"]


def safe_folder_name(x):
    name = clean(x)
    if name == "":
        name = "SIN_CATEGORIA"
    name = re.sub(r'[<>:"/\\|?*]', "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name


def fit_image_to_box(img, box_w, box_h):
    img = img.convert("RGB")
    img.thumbnail((box_w, box_h), Image.LANCZOS)

    background = Image.new("RGB", (box_w, box_h), WHITE)
    x = (box_w - img.width) // 2
    y = (box_h - img.height) // 2
    background.paste(img, (x, y))

    return background


def draw_centered_text(draw, text, box_x, box_y, box_w, font, fill=BLACK):
    bbox = draw.textbbox((0, 0), text, font=font)
    text_w = bbox[2] - bbox[0]
    x = box_x + (box_w - text_w) // 2
    draw.text((x, box_y), text, font=font, fill=fill)


def draw_centered_rich_text(draw, parts, y):
    widths = []
    total_w = 0

    for text, font in parts:
        bbox = draw.textbbox((0, 0), text, font=font)
        w = bbox[2] - bbox[0]
        widths.append(w)
        total_w += w

    x = (CANVAS_W - total_w) // 2

    for (text, font), w in zip(parts, widths):
        draw.text((x, y), text, font=font, fill=BLACK)
        x += w


# =============================================================================
# 4. READ DATA
# =============================================================================

inventory = pd.read_excel(EXCEL_PATH, sheet_name=INVENTORY_SHEET)
photo_archive = pd.read_excel(EXCEL_PATH, sheet_name=PHOTO_ARCHIVE_SHEET)

inventory.columns = inventory.columns.str.strip()
photo_archive.columns = photo_archive.columns.str.strip()

inventory["ITEM_ID"] = inventory["ITEM_ID"].astype(str).str.strip()
photo_archive["ITEM_ID"] = photo_archive["ITEM_ID"].astype(str).str.strip()
photo_archive["photo_type"] = photo_archive["photo_type"].astype(str).str.strip().str.lower()

if "NOTAG_REQUIRED" not in inventory.columns:
    inventory["NOTAG_REQUIRED"] = ""

photo_archive = photo_archive[
    photo_archive["photo_type"].isin(["front", "back", "tag"])
].copy()

photo_wide = (
    photo_archive
    .sort_values(["ITEM_ID", "photo_type"])
    .drop_duplicates(["ITEM_ID", "photo_type"], keep="first")
    .pivot(index="ITEM_ID", columns="photo_type", values="converted_name")
    .reset_index()
)

df = inventory.merge(photo_wide, on="ITEM_ID", how="left")

df = df[df["SOLD_DATE"].apply(is_blank)].copy()

if "CATEGORY" in df.columns:
    df = df[
        ~df["CATEGORY"].astype(str).str.strip().str.lower().isin(["zapatos", "shoes"])
    ].copy()

if TEST_RUN == 1:
    df = df.head(TEST_N_ITEMS).copy()


# =============================================================================
# 5. MAKE COLLAGES
# =============================================================================

created = 0
skipped = []
gallery_rows = []

for _, row in df.iterrows():
    item_id = clean(row.get("ITEM_ID", ""))
    description = clean(row.get("DESCRIPTION", ""))
    category = clean(row.get("CATEGORY", ""))
    category_safe = safe_folder_name(category)
    price = clean(row.get("PRICE", ""))
    brand = clean(row.get("BRAND", ""))
    size = clean(row.get("SIZE", ""))
    size_sml = clean(row.get("SIZE (SML)", ""))
    notag_required = flag_is_1(row.get("NOTAG_REQUIRED", ""))

    front_file = clean(row.get("front", ""))
    back_file = clean(row.get("back", ""))
    tag_file = clean(row.get("tag", ""))

    has_front = front_file != "" and (JPG_FOLDER / front_file).exists()
    has_back = back_file != "" and (JPG_FOLDER / back_file).exists()
    has_tag = tag_file != "" and (JPG_FOLDER / tag_file).exists()

    if not has_front:
        skipped.append((item_id, "Missing: front"))
        continue

    if not notag_required and not has_tag:
        skipped.append((item_id, "Missing: tag"))
        continue

    category_folder = COLLAGE_FOLDER / category_safe
    category_folder.mkdir(parents=True, exist_ok=True)

    canvas = Image.new("RGB", (CANVAS_W, CANVAS_H), WHITE)
    draw = ImageDraw.Draw(canvas)

    title = f"{description} | ${price}"
    draw_centered_text(draw, title, 0, 35, CANVAS_W, FONT_TITLE)

    photo_types = ["front", "back", "tag"]
    labels = ["Parte delantera", "Parte trasera", "Etiqueta"]
    files = [front_file, back_file, tag_file]

    for i, (photo_type, label, filename) in enumerate(zip(photo_types, labels, files)):
        x = MARGIN_X + i * (SLOT_W + GAP)

        # If back is missing, leave the middle photo slot blank,
        # but still write "Parte trasera".
        # If NOTAG_REQUIRED == 1, leave third slot fully blank.
        if photo_type == "tag" and notag_required:
            continue

        if photo_type == "back" and not has_back:
            draw_centered_text(draw, label, x, LABEL_Y, SLOT_W, FONT_LABEL)
            continue

        img_path = JPG_FOLDER / filename

        if img_path.exists():
            with Image.open(img_path) as img:
                fitted = fit_image_to_box(img, SLOT_W, SLOT_H)
            canvas.paste(fitted, (x, PHOTO_Y))
        else:
            draw.rectangle(
                [x, PHOTO_Y, x + SLOT_W, PHOTO_Y + SLOT_H],
                fill=GRAY
            )

        draw_centered_text(draw, label, x, LABEL_Y, SLOT_W, FONT_LABEL)

    footer_parts = [
        ("MARCA: ", FONT_FOOTER_BOLD),
        (f"{brand} / ", FONT_FOOTER),
        ("TALLA: ", FONT_FOOTER_BOLD),
        (f"{size} / ", FONT_FOOTER),
        ("LE QUEDA A: ", FONT_FOOTER_BOLD),
        (f"{size_sml}", FONT_FOOTER),
    ]

    draw_centered_rich_text(draw, footer_parts, FOOTER_Y + 20)

    legend_1 = "ENTREGAS SOLO EN LÍNEA 2 TAXQUEÑA - HIDALGO"
    legend_2 = "PRIORIDAD A NATIVITAS / PORTALES · SE ENTREGA POR ORDEN DE CONFIRMACIÓN"

    draw_centered_text(draw, legend_1, 0, LEGEND_Y, CANVAS_W, FONT_LEGEND)
    draw_centered_text(draw, legend_2, 0, LEGEND_Y + 42, CANVAS_W, FONT_LEGEND)

    out_path = category_folder / f"{item_id}_collage.jpg"
    canvas.save(out_path, quality=95)

    gallery_rows.append({
        "ITEM_ID": item_id,
        "DESCRIPTION": description,
        "CATEGORY": category_safe,
        "PRICE": price,
        "COLLAGE_PATH": out_path,
    })

    created += 1


# =============================================================================
# 6. BUILD HTML GALLERY WITH CATEGORY FILTER
# =============================================================================

categories = sorted(set(row["CATEGORY"] for row in gallery_rows))

html_parts = []

html_parts.append("""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Collage Gallery</title>

<style>
body {
    font-family: Arial, sans-serif;
    margin: 24px;
    background: #f7f7f7;
    color: #222;
}

h1 {
    margin-bottom: 6px;
}

.controls {
    background: white;
    border: 1px solid #ddd;
    border-radius: 12px;
    padding: 14px;
    margin-bottom: 20px;
    position: sticky;
    top: 0;
    z-index: 10;
}

select {
    font-size: 16px;
    padding: 8px;
}

.grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(520px, 1fr));
    gap: 18px;
}

.card {
    background: white;
    border: 1px solid #ddd;
    border-radius: 12px;
    padding: 12px;
}

.card img {
    width: 100%;
    border-radius: 8px;
    border: 1px solid #ddd;
}

.meta {
    font-size: 14px;
    color: #555;
    margin-bottom: 8px;
    line-height: 1.4;
}

.hidden {
    display: none;
}
</style>

<script>
function filterCategory() {
    const selected = document.getElementById("categoryFilter").value;
    const cards = document.querySelectorAll(".card");

    cards.forEach(card => {
        const cat = card.getAttribute("data-category");

        if (selected === "ALL" || cat === selected) {
            card.classList.remove("hidden");
        } else {
            card.classList.add("hidden");
        }
    });

    const visible = document.querySelectorAll(".card:not(.hidden)").length;
    document.getElementById("visibleCount").textContent = visible;
}
</script>
</head>

<body>
<h1>Collage Gallery</h1>

<div class="controls">
    <b>Filter by category:</b>
    <select id="categoryFilter" onchange="filterCategory()">
        <option value="ALL">All categories</option>
""")

for cat in categories:
    html_parts.append(
        f'        <option value="{html.escape(cat)}">{html.escape(cat)}</option>\n'
    )

html_parts.append(f"""
    </select>
    <br><br>
    Showing <b><span id="visibleCount">{len(gallery_rows)}</span></b> of <b>{len(gallery_rows)}</b> collages.
</div>

<div class="grid">
""")

for row in gallery_rows:
    img_uri = row["COLLAGE_PATH"].as_uri()

    html_parts.append(f"""
    <div class="card" data-category="{html.escape(row["CATEGORY"])}">
        <div class="meta">
            <b>{html.escape(row["ITEM_ID"])}</b> — {html.escape(row["DESCRIPTION"])}<br>
            Category: <b>{html.escape(row["CATEGORY"])}</b> |
            Price: <b>${html.escape(row["PRICE"])}</b>
        </div>
        <a href="{img_uri}" target="_blank">
            <img src="{img_uri}">
        </a>
    </div>
    """)

html_parts.append("""
</div>
</body>
</html>
""")

OUTPUT_HTML.write_text("".join(html_parts), encoding="utf-8")


# =============================================================================
# 7. SUMMARY
# =============================================================================

print("Done!")
print(f"Test run: {TEST_RUN}")
print(f"Items selected: {len(df)}")
print(f"Collages created: {created}")
print(f"Collage folder: {COLLAGE_FOLDER}")
print(f"HTML gallery: {OUTPUT_HTML}")

if skipped:
    print()
    print("Skipped items:")
    for item_id, reason in skipped:
        print(f"- {item_id}: {reason}")