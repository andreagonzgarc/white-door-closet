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
POSTER_FOLDER = BASE_DIR / "collages_one_photo"
OUTPUT_HTML = BASE_DIR / "collage_gallery_one_photo.html"

PHOTO_ARCHIVE_SHEET = "photo_archive"
INVENTORY_SHEET = "inventory_vf"

POSTER_FOLDER.mkdir(exist_ok=True)

# =============================================================================
# 1. SETTINGS
# =============================================================================

TEST_RUN = 0
TEST_N_ITEMS = 5

if TEST_RUN == 1 and (TEST_N_ITEMS is None or TEST_N_ITEMS <= 0):
    raise ValueError("If TEST_RUN = 1, set TEST_N_ITEMS to a positive number.")

JPEG_QUALITY = 95

TITLE_Y_PCT = 0.04
FOOTER_Y_PCT = 0.78
LEGEND_1_Y_PCT = 0.88
LEGEND_2_Y_PCT = 0.93

WHITE = "white"
BOX_OPACITY = 128  # 50% black

# =============================================================================
# 2. FONTS
# =============================================================================

def get_font(size, bold=True):
    font_dir = Path(r"C:\Windows\Fonts")
    candidates = [
        font_dir / ("arialbd.ttf" if bold else "arial.ttf"),
        font_dir / ("Arial Bold.ttf" if bold else "Arial.ttf"),
    ]

    for font_path in candidates:
        if font_path.exists():
            return ImageFont.truetype(str(font_path), size=size)

    return ImageFont.load_default()


def scaled_fonts(img_w):
    return {
        "title": get_font(max(34, int(img_w * 0.045)), bold=True),
        "footer": get_font(max(26, int(img_w * 0.032)), bold=True),
        "legend": get_font(max(22, int(img_w * 0.026)), bold=True),
    }

# =============================================================================
# 3. HELPERS
# =============================================================================

def clean(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def is_blank(x):
    return pd.isna(x) or clean(x) in ["", "NaT", "nan", "None"]


def safe_folder_name(x):
    name = clean(x) or "SIN_CATEGORIA"
    name = re.sub(r'[<>:"/\\|?*]', "_", name)
    return re.sub(r"\s+", " ", name).strip()


def draw_text_box(base_img, draw, text, x, y, font, padding_x=18, padding_y=10):
    bbox = draw.textbbox((0, 0), text, font=font)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]

    box = [
        x - padding_x,
        y - padding_y,
        x + text_w + padding_x,
        y + text_h + padding_y,
    ]

    overlay = Image.new("RGBA", base_img.size, (0, 0, 0, 0))
    overlay_draw = ImageDraw.Draw(overlay)

    overlay_draw.rounded_rectangle(
        box,
        radius=10,
        fill=(0, 0, 0, BOX_OPACITY)
    )

    base_img.alpha_composite(overlay)

    draw.text(
        (x, y),
        text,
        font=font,
        fill=WHITE,
    )


def draw_wrapped_centered_text(base_img, draw, text, y, img_w, font, max_width_pct=0.86, line_gap=10):
    max_width = int(img_w * max_width_pct)
    words = text.split()
    lines = []
    current = ""

    for word in words:
        test = f"{current} {word}".strip()
        bbox = draw.textbbox((0, 0), test, font=font)
        test_w = bbox[2] - bbox[0]

        if test_w <= max_width:
            current = test
        else:
            if current:
                lines.append(current)
            current = word

    if current:
        lines.append(current)

    line_h = draw.textbbox((0, 0), "Ag", font=font)[3]

    for i, line in enumerate(lines):
        bbox = draw.textbbox((0, 0), line, font=font)
        text_w = bbox[2] - bbox[0]

        x = (img_w - text_w) // 2
        line_y = y + i * (line_h + line_gap)

        draw_text_box(base_img, draw, line, x, line_y, font)

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

photo_archive = photo_archive[photo_archive["photo_type"].eq("front")].copy()

front_photos = (
    photo_archive
    .sort_values(["ITEM_ID"])
    .drop_duplicates(["ITEM_ID"], keep="first")
    [["ITEM_ID", "converted_name"]]
    .rename(columns={"converted_name": "front"})
)

df = inventory.merge(front_photos, on="ITEM_ID", how="left")

df = df[df["SOLD_DATE"].apply(is_blank)].copy()

if "CATEGORY" in df.columns:
    df = df[
        ~df["CATEGORY"].astype(str).str.strip().str.lower().isin(["zapatos", "shoes"])
    ].copy()

if TEST_RUN == 1:
    df = df.head(TEST_N_ITEMS).copy()

# =============================================================================
# 5. MAKE ONE-PHOTO POSTERS
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

    front_file = clean(row.get("front", ""))
    front_path = JPG_FOLDER / front_file

    if front_file == "" or not front_path.exists():
        skipped.append((item_id, "Missing front photo"))
        continue

    category_folder = POSTER_FOLDER / category_safe
    category_folder.mkdir(parents=True, exist_ok=True)

    with Image.open(front_path) as img:
        poster = img.convert("RGBA").copy()

    draw = ImageDraw.Draw(poster)
    fonts = scaled_fonts(poster.width)

    title = f"{description} | ${price}"
    footer = f"MARCA: {brand} / TALLA: {size} / LE QUEDA A: {size_sml}"
    legend_1 = "ENTREGAS SOLO EN LÍNEA 2 TAXQUEÑA - HIDALGO"
    legend_2 = "PRIORIDAD A NATIVITAS / PORTALES · SE ENTREGA POR ORDEN DE CONFIRMACIÓN"

    draw_wrapped_centered_text(
        poster, draw, title,
        int(poster.height * TITLE_Y_PCT),
        poster.width,
        fonts["title"],
    )

    draw_wrapped_centered_text(
        poster, draw, footer,
        int(poster.height * FOOTER_Y_PCT),
        poster.width,
        fonts["footer"],
    )

    draw_wrapped_centered_text(
        poster, draw, legend_1,
        int(poster.height * LEGEND_1_Y_PCT),
        poster.width,
        fonts["legend"],
    )

    draw_wrapped_centered_text(
        poster, draw, legend_2,
        int(poster.height * LEGEND_2_Y_PCT),
        poster.width,
        fonts["legend"],
    )

    out_path = category_folder / f"{item_id}_one_photo.jpg"
    poster.convert("RGB").save(out_path, quality=JPEG_QUALITY)

    gallery_rows.append({
        "ITEM_ID": item_id,
        "DESCRIPTION": description,
        "CATEGORY": category_safe,
        "PRICE": price,
        "POSTER_PATH": out_path,
    })

    created += 1

# =============================================================================
# 6. BUILD HTML GALLERY
# =============================================================================

categories = sorted(set(row["CATEGORY"] for row in gallery_rows))

html_parts = ["""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>One Photo Poster Gallery</title>

<style>
body {
    font-family: Arial, sans-serif;
    margin: 24px;
    background: #f7f7f7;
    color: #222;
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
    grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
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

.hidden { display: none; }
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

    document.getElementById("visibleCount").textContent =
        document.querySelectorAll(".card:not(.hidden)").length;
}
</script>
</head>

<body>
<h1>One Photo Poster Gallery</h1>

<div class="controls">
    <b>Filter by category:</b>
    <select id="categoryFilter" onchange="filterCategory()">
        <option value="ALL">All categories</option>
"""]

for cat in categories:
    html_parts.append(f'<option value="{html.escape(cat)}">{html.escape(cat)}</option>\n')

html_parts.append(f"""
    </select>
    <br><br>
    Showing <b><span id="visibleCount">{len(gallery_rows)}</span></b> of <b>{len(gallery_rows)}</b> posters.
</div>

<div class="grid">
""")

for row in gallery_rows:
    img_uri = row["POSTER_PATH"].as_uri()

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
print(f"One-photo posters created: {created}")
print(f"Poster folder: {POSTER_FOLDER}")
print(f"HTML gallery: {OUTPUT_HTML}")

if skipped:
    print()
    print("Skipped items:")
    for item_id, reason in skipped:
        print(f"- {item_id}: {reason}")