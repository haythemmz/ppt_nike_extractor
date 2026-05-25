from __future__ import annotations

import re
from pathlib import Path
from typing import Iterable

import pandas as pd
from pptx import Presentation


# -----------------------------
# Regex / cleaning rules
# -----------------------------

STYLE_RE = re.compile(r"\b[A-Z]{2}\d{4}\b")
STYLE_COLOR_RE = re.compile(r"\b([A-Z]{2}\d{4})[-–—](\d{3})\b")

# Special Nike accessories/bags style-color format, for example: N.100.3478.091
ACCESSORY_STYLE_COLOR_RE = re.compile(r"\bN\.\d{3}\.\d{4}\.\d{3}\b", re.I)

STATUS_RE = re.compile(r"\b(NEW|CARRYOVER)\b", re.I)

# IMPORTANT: use word boundaries so WSH is not interpreted as "WS".
PRICE_CONTEXT_RE = re.compile(r"\b(RETAIL|WS|WHSE|WHOLESALE)\b", re.I)

STOP_RE = re.compile(
    r"^(Material|Material Content|Features and Benefits|Key Features|Overview|Product Details|Keep it|FOR INTERNAL USE|©|Product Images|"
    r"Performance Benefits|FIT:|RIDE:|TRACTION:|Product Description)",
    re.I,
)

# Lines that often appear near the style code but are NOT product names.
BAD_PRODUCT_NAME_RE = re.compile(
    r"^(?:"
    r"Dri-Fit|Therma-Fit|Aero-Fit(?: Technology)?|Storm-Fit(?: ADV)?|"
    r"Repel Technology|UV Protective|UV Protection|"
    r"Premium knit fabric|Premium knit jacquard fabric|Premium hand feel|"
    r"Standard Fit|Slim Fit|Loose Fit|Oversized Fit|Tight fit|Secure Fit|"
    r"Material:?|Key Features:?|Product Information|TBD|LAUNCH TBD|"
    r"SP\d+|SPSU\d+|"
    r"Stretch to provide full range of motion|Thumbholes|Snapback closure|"
    r"Stretchy, elastic closure|100% Cotton|100% Polyester|"
    r"Mid-depth|Max-depth|High depth|Low depth"
    r")$",
    re.I,
)

# Strong title patterns for Nike product names.
PRODUCT_TITLE_RE = re.compile(
    r"^(?:M|W|B|G|K|U)\s+NK\b|"
    r"^NK\s+|"
    r"^K\s+NSW\b|"
    r"^W\s+NSW\b|"
    r"^NIKE\s+(?:AIR|INFINITY|EVERYDAY|UNICORN|BRASILIA)\b|"
    r"^Nike Brasilia\b|"
    r"^VICTORY\b|"
    r"^NEXT%\b|"
    r"^PEGASUS\b|"
    r"^TEMPO\b|"
    r"^AIR MAX\b|"
    r"^WOMEN['’]S\s+",
    re.I,
)

PRODUCT_KEYWORDS = {
    "POLO", "PANT", "SHORT", "JACKET", "JKT", "HOODIE", "VEST",
    "CAP", "SKIRT", "SKRT", "SHOE", "SOCK", "SOCKS", "BACKPACK",
    "BAG", "TOP", "TEE", "CREW", "DRESS", "VISOR", "BEANIE",
    "DUFFEL", "DRAWSTRING", "FOOTIE", "NO-SHOW", "ANKLE", "TRAINING",
}

CATEGORY_PATTERNS: list[tuple[str, re.Pattern]] = [
    ("Men's Apparel", re.compile(r"\bMEN'?S\s+APPAREL\b", re.I)),
    ("Women's Apparel", re.compile(r"\bWOMEN'?S\s+APPAREL\b", re.I)),
    ("Kids' Apparel", re.compile(r"\b(?:KIDS?|YOUTH|BOYS?|GIRLS?)\s+APPAREL\b", re.I)),
    ("Apparel", re.compile(r"^APPAREL$", re.I)),
    ("Bags", re.compile(r"\b(?:BAGS?|BACKPACKS?|DUFFELS?|DRAWSTRING)\b", re.I)),
    ("Gloves", re.compile(r"\bGLOVES?\b", re.I)),
    ("Headcovers", re.compile(r"\bHEAD\s*COVERS?\b|\bHEADCOVERS?\b", re.I)),
    ("Footwear", re.compile(r"\bFOOTWEAR\b|\bSHOES?\b", re.I)),
    ("Socks", re.compile(r"\bSOCKS?\b", re.I)),
    ("Caps & Headwear", re.compile(r"\b(?:CAPS?|HATS?|HEADWEAR|VISORS?|BEANIES?)\b", re.I)),
    ("Accessories", re.compile(r"\bACCESSORIES\b", re.I)),
]


# -----------------------------
# Helper functions
# -----------------------------

def clean_text(value: str) -> str:
    value = str(value)
    value = value.replace("\xa0", " ")
    value = value.replace("–", "-").replace("—", "-")
    value = value.replace("’", "'")
    return re.sub(r"\s+", " ", value).strip()


def unique_preserve_order(values: list[str]) -> list[str]:
    result: list[str] = []
    for value in values:
        if value and value not in result:
            result.append(value)
    return result


def extract_style_color_pairs(text: str) -> list[tuple[str, str]]:
    text = clean_text(text)
    return [
        (match.group(1).upper(), match.group(2).zfill(3))
        for match in STYLE_COLOR_RE.finditer(text)
    ]


def extract_colors_from_style_color_tokens(text: str) -> list[str]:
    return unique_preserve_order([color for _, color in extract_style_color_pairs(text)])


def detect_slide_category(lines: list[str]) -> str | None:
    """
    Detect section divider/category slides such as WOMEN'S APPAREL or HEADCOVERS.
    Product-heavy lines are ignored so product names do not accidentally reset category.
    """
    candidates = lines[:12]

    for line in candidates:
        line_clean = clean_text(line)
        if not line_clean:
            continue

        if STYLE_RE.search(line_clean) or ACCESSORY_STYLE_COLOR_RE.search(line_clean):
            continue

        if looks_like_price_context(line_clean):
            continue

        for category, pattern in CATEGORY_PATTERNS:
            if pattern.search(line_clean):
                return category

    return None


def iter_all_shapes(shapes):
    """Yield shapes recursively, including shapes inside groups."""
    for shape in shapes:
        yield shape
        if hasattr(shape, "shapes"):
            yield from iter_all_shapes(shape.shapes)


def get_slide_text_items(slide) -> list[dict]:
    """
    Extract text from slide shapes with their visual position.
    Sorting by top/left makes the extraction closer to what we see on the slide.
    """
    items: list[dict] = []

    for shape_index, shape in enumerate(iter_all_shapes(slide.shapes)):
        if not hasattr(shape, "text") or not shape.text:
            continue

        text = shape.text.strip()
        if not text:
            continue

        for line_index, raw_line in enumerate(text.splitlines()):
            line = clean_text(raw_line)
            if not line:
                continue

            items.append(
                {
                    "shape_index": shape_index,
                    "line_index": line_index,
                    "text": line,
                    "left": int(getattr(shape, "left", 0)),
                    "top": int(getattr(shape, "top", 0)),
                    "width": int(getattr(shape, "width", 0)),
                    "height": int(getattr(shape, "height", 0)),
                }
            )

    # Visual order: top to bottom, then left to right
    return sorted(items, key=lambda x: (x["top"], x["left"], x["shape_index"]))


def get_slide_lines(slide) -> list[str]:
    return [item["text"] for item in get_slide_text_items(slide)]


def looks_like_price_context(text: str) -> bool:
    return bool(PRICE_CONTEXT_RE.search(clean_text(text)))


def is_new_product_start(line: str, lookahead: str = "") -> bool:
    """
    A real product detail line must contain a style code and WHOLESALE/RETAIL context.

    This intentionally ignores overview/product-grid lines like:
        WS: IR4565
        $140/70
    because those are not detailed product blocks and usually do not contain colors.
    """
    line_clean = clean_text(line)
    combined = clean_text(line + " " + lookahead)

    if not STYLE_RE.search(line_clean):
        return False

    # Ignore overview/grid lines.
    if re.match(r"^(?:WS|AS)\s*:", line_clean, flags=re.I):
        return False

    return bool(re.search(r"\b(WHOLESALE|RETAIL)\b", combined, flags=re.I))


def extract_prices(text: str, aggressive: bool = False) -> tuple[float | None, float | None]:
    """
    Extract wholesale and retail prices.
    Most product lines are like:
        IR4565 | $70.00 WHOLESALE | $140.00 RETAIL | NEW
    Fallback keeps the first two dollar values as wholesale/retail.
    
    Args:
        text: Text to extract prices from
        aggressive: If True, will assign first smaller price to wholesale and larger to retail
                   even if labels are missing or unclear
    """
    text_clean = clean_text(text)

    wholesale = None
    retail = None
    debug_mode = False  # Set to True to see extraction details

    # Explicit label near value: $70.00 WHOLESALE / $140.00 RETAIL or $70.00Wholesale
    for match in re.finditer(
        r"\$\s*(\d+(?:\.\d{2})?)\s*(WHOLESALE|WS|WHSE|RETAIL)\b(?!\s*:)",
        text_clean,
        flags=re.I,
    ):
        value = float(match.group(1))
        label = match.group(2).upper()
        if label in {"WHOLESALE", "WS", "WHSE"} and wholesale is None:
            wholesale = value
            if debug_mode: print(f"  Found WHOLESALE (after $): {value}")
        elif label == "RETAIL" and retail is None:
            retail = value
            if debug_mode: print(f"  Found RETAIL (after $): {value}")

    # Explicit label before value: RETAIL $26.00 or Retail: $26.00
    for match in re.finditer(
        r"\b(WHOLESALE|WS|WHSE|RETAIL)\b\s*:?\s*\$\s*(\d+(?:\.\d{2})?)",
        text_clean,
        flags=re.I,
    ):
        label = match.group(1).upper()
        value = float(match.group(2))
        if label in {"WHOLESALE", "WS", "WHSE"} and wholesale is None:
            wholesale = value
            if debug_mode: print(f"  Found WHOLESALE (before $): {value}")
        elif label == "RETAIL" and retail is None:
            retail = value
            if debug_mode: print(f"  Found RETAIL (before $): {value}")

    # Same labeled formats without a dollar sign: 70.00 WHOLESALE / Retail: 140.00
    for match in re.finditer(
        r"(?<![A-Z0-9$])(\d+\.\d{2})\s*(WHOLESALE|WS|WHSE|RETAIL)\b(?!\s*:)",
        text_clean,
        flags=re.I,
    ):
        value = float(match.group(1))
        label = match.group(2).upper()
        if label in {"WHOLESALE", "WS", "WHSE"} and wholesale is None:
            wholesale = value
            if debug_mode: print(f"  Found WHOLESALE (no $ after value): {value}")
        elif label == "RETAIL" and retail is None:
            retail = value
            if debug_mode: print(f"  Found RETAIL (no $ after value): {value}")

    for match in re.finditer(
        r"\b(WHOLESALE|WS|WHSE|RETAIL)\b\s*:?\s*(?!\$)(\d+\.\d{2})",
        text_clean,
        flags=re.I,
    ):
        label = match.group(1).upper()
        value = float(match.group(2))
        if label in {"WHOLESALE", "WS", "WHSE"} and wholesale is None:
            wholesale = value
            if debug_mode: print(f"  Found WHOLESALE (no $ before value): {value}")
        elif label == "RETAIL" and retail is None:
            retail = value
            if debug_mode: print(f"  Found RETAIL (no $ before value): {value}")

    # Fallback: all explicit dollar amounts in order.
    values: list[float] = []
    for match in re.finditer(r"\$\s*(\d+(?:\.\d{2})?)", text_clean):
        value = float(match.group(1))
        if value not in values:
            values.append(value)
    
    if debug_mode and values: print(f"  All dollar values found: {values}")

    # Backup: values like 20.00 WS without $
    if len(values) < 2:
        for match in re.finditer(
            r"(?<![A-Z0-9])(\d+\.\d{2})(?=\s*(?:WS|WHSE|WHOLESALE|RETAIL|\|))",
            text_clean,
            flags=re.I,
        ):
            value = float(match.group(1))
            if value not in values:
                values.append(value)

    # First pass fallback for completely unlabeled prices only.
    # Do not copy a single explicitly labeled retail price into wholesale.
    if wholesale is None and retail is None:
        if len(values) >= 1:
            wholesale = values[0]
            if debug_mode: print(f"  Fallback WHOLESALE: {wholesale}")
        if len(values) >= 2:
            retail = values[1]
            if debug_mode: print(f"  Fallback RETAIL: {retail}")
    elif retail is None and len(values) >= 2:
        for value in values:
            if value != wholesale:
                retail = value
                if debug_mode: print(f"  Fallback RETAIL: {retail}")
                break
    
    if debug_mode:
        print(f"  After initial extraction - WS: {wholesale}, RT: {retail}")
    
    # Aggressive mode: intelligently assign prices by value when labels are unclear
    if aggressive and len(values) >= 2:
        unique_values = sorted(set(values))
        if debug_mode: print(f"  Aggressive mode: unique sorted values: {unique_values}")
        
        # If we found two different prices, assign smaller to wholesale, larger to retail
        if len(unique_values) >= 2:
            # Make sure wholesale is the smaller price
            if wholesale is None or retail is None:
                # At least one is missing, fill both intelligently
                wholesale = unique_values[0]  # Smaller
                retail = unique_values[1]     # Larger
                if debug_mode: print(f"  Aggressive: filled missing - WS: {wholesale}, RT: {retail}")
            elif wholesale > retail:
                # Prices are backwards, swap them
                wholesale, retail = retail, wholesale
                if debug_mode: print(f"  Aggressive: swapped backwards - WS: {wholesale}, RT: {retail}")

    if debug_mode:
        print(f"  Final result - WS: {wholesale}, RT: {retail}\n")

    return wholesale, retail


def is_likely_product_name(line: str) -> bool:
    line = clean_text(line)

    if not line:
        return False

    if STYLE_RE.search(line):
        return False

    if STOP_RE.search(line):
        return False

    if BAD_PRODUCT_NAME_RE.search(line):
        return False

    # Color-code-only line: 010 or 010 100 274
    if re.fullmatch(r"[0-9]{3}(\s+[0-9]{3})*", line):
        return False

    # Size ranges like 3.5-13,14,15,16
    if re.fullmatch(r"\d+(?:\.\d+)?-\d+.*", line):
        return False

    if looks_like_price_context(line):
        return False

    # Reject obvious paragraph text.
    if len(line.split()) > 12 and not PRODUCT_TITLE_RE.search(line):
        return False

    return len(line) > 2


def product_title_score(line: str) -> int:
    line_clean = clean_text(line)
    line_upper = line_clean.upper()
    tokens = set(re.findall(r"[A-Z0-9%'/.-]+", line_upper))

    score = 0

    # Strong Nike product title pattern
    if PRODUCT_TITLE_RE.search(line_clean):
        score += 100

    # Product category words
    if tokens & PRODUCT_KEYWORDS:
        score += 25

    # Many product names are uppercase
    if line_clean == line_upper and len(tokens) >= 3:
        score += 10

    # Product names are usually compact, not long descriptions.
    if 2 <= len(tokens) <= 10:
        score += 5

    # Penalize weak words that are often features/materials.
    if BAD_PRODUCT_NAME_RE.search(line_clean):
        score -= 100

    return score


def previous_product_name(lines: list[str], idx: int) -> str | None:
    candidates: list[dict] = []

    # Search farther back because some slides have:
    # title -> material/key features -> style code
    for j in range(idx - 1, max(-1, idx - 25), -1):
        line = lines[j]

        # Stop if we hit a previous product block.
        if STYLE_RE.search(line) and looks_like_price_context(line):
            break

        if is_likely_product_name(line):
            candidates.append(
                {
                    "line": line,
                    "score": product_title_score(line),
                    "distance": abs(idx - j),
                }
            )

    if not candidates:
        return None

    candidates = sorted(
        candidates,
        key=lambda x: (x["score"], -x["distance"]),
        reverse=True,
    )

    return candidates[0]["line"]


def next_product_name(lines: list[str], idx: int) -> str | None:
    candidates: list[dict] = []

    for j in range(idx + 1, min(len(lines), idx + 8)):
        line = lines[j]

        if STYLE_RE.search(line) and looks_like_price_context(line):
            break

        if is_likely_product_name(line):
            candidates.append(
                {
                    "line": line,
                    "score": product_title_score(line),
                    "distance": abs(idx - j),
                }
            )

    if not candidates:
        return None

    candidates = sorted(
        candidates,
        key=lambda x: (x["score"], -x["distance"]),
        reverse=True,
    )

    return candidates[0]["line"]


def build_product_name(line: str, lines: list[str], idx: int, first_style: str) -> str | None:
    pos = line.find(first_style)

    # Case where product name and style code are on the same line:
    # "M NK DF PAR POLO SS SOLID IB0233 | $45..."
    prefix = clean_text(line[:pos].replace("Style", "")) if pos > 0 else ""

    if prefix and is_likely_product_name(prefix):
        return prefix

    return previous_product_name(lines, idx) or next_product_name(lines, idx) or None


# -----------------------------
# Nike accessories / bags format
# -----------------------------

def split_accessory_style_color(full_code: str) -> tuple[str, str]:
    """
    Converts Nike accessory/bag codes like N.100.3478.091 into:
        style_code = N.100.3478
        color_code = 091
    """
    code = clean_text(full_code).upper()
    parts = code.split(".")
    if len(parts) == 4:
        return ".".join(parts[:3]), parts[3].zfill(3)
    return code, ""


def has_accessory_style_color(text: str) -> bool:
    return bool(ACCESSORY_STYLE_COLOR_RE.search(clean_text(text)))


def get_text_box_lines(items: list[dict], shape_index: int) -> list[str]:
    shape_items = [item for item in items if item["shape_index"] == shape_index]
    shape_items = sorted(shape_items, key=lambda x: x.get("line_index", 0))
    return [item["text"] for item in shape_items]


def get_accessory_price_context(items: list[dict], code_item: dict, lines: list[str], line_idx: int) -> str:
    """
    Accessory slides often place each product caption in one text box:
        product name
        N.101.4761.016
        Wholesale: $17.50
        Retail: $35.00

    Use that text box first so prices from neighboring captions are not mixed in.
    """
    same_box_lines = get_text_box_lines(items, code_item["shape_index"])
    same_box_text = clean_text(" ".join(same_box_lines))
    wholesale, retail = extract_prices(same_box_text, aggressive=True)

    if wholesale is not None and retail is not None:
        return same_box_text

    left = int(code_item.get("left", 0))
    top = int(code_item.get("top", 0))
    width = max(int(code_item.get("width", 0)), 1)
    height = max(int(code_item.get("height", 0)), 1)
    center_x = left + width / 2

    nearby: list[dict] = []
    for item in items:
        item_left = int(item.get("left", 0))
        item_top = int(item.get("top", 0))
        item_width = max(int(item.get("width", 0)), 1)
        item_center_x = item_left + item_width / 2

        same_column = abs(item_center_x - center_x) <= max(width, item_width) * 0.65
        close_below = top - height <= item_top <= top + height * 3.5

        if same_column and close_below:
            nearby.append(item)

    nearby_text = clean_text(" ".join(item["text"] for item in sorted(nearby, key=lambda x: (x["top"], x["left"], x.get("line_index", 0)))))
    wholesale, retail = extract_prices(nearby_text, aggressive=True)
    if wholesale is not None and retail is not None:
        return nearby_text

    return same_box_text or nearby_text or clean_text(" ".join(lines[line_idx : min(line_idx + 10, len(lines))]))


def build_style_color_key(style: str, color: str | None) -> str:
    style = clean_text(style)
    color = clean_text(color) if color is not None else ""

    if not color:
        return style

    # Accessories already use dot notation in the original Nike code.
    if style.upper().startswith("N."):
        return f"{style}.{color}"

    # Apparel uses the regular style-color convention.
    return f"{style}-{color}"


# -----------------------------
# Color extraction
# -----------------------------

def is_color_code_line(line: str, slide_number: int | None = None) -> bool:
    line = clean_text(line)

    if not line:
        return False

    if slide_number is not None and line == str(slide_number):
        return False

    if STYLE_RE.search(line):
        return False

    if looks_like_price_context(line):
        return False

    if STOP_RE.search(line):
        return False

    # Ignore size ranges
    if re.fullmatch(r"\d+(?:\.\d+)?-\d+.*", line):
        return False

    # Ignore years
    if re.search(r"\b20\d{2}\b", line):
        return False

    # Exact color code: 010
    if re.fullmatch(r"\d{3}", line):
        return True

    # Multiple exact color codes: 010 100 274
    if re.fullmatch(r"(\d{3}\s*)+", line):
        return True

    # Color code with description: 010 - BLACK, 010 BLACK
    if re.match(r"^\d{3}\s*(?:[-/]\s*)?[A-Za-z]", line):
        return True

    return False


def extract_color_codes_from_text(line: str, slide_number: int | None = None) -> list[str]:
    line = clean_text(line)

    if not is_color_code_line(line, slide_number=slide_number):
        return []

    # Exact code / multiple exact codes
    if re.fullmatch(r"(\d{3}\s*)+", line):
        return [code.zfill(3) for code in re.findall(r"\d{3}", line)]

    # Code + description
    match = re.match(r"^(\d{3})\s*(?:[-/]\s*)?[A-Za-z]", line)
    if match:
        return [match.group(1).zfill(3)]

    return []


def extract_color_codes_from_lines(lines: list[str], start_idx: int) -> list[str]:
    """
    Extract color codes near the product block.
    This is useful, but not always enough because PowerPoint shape order can be strange.
    """
    color_codes: list[str] = []

    for line in lines[start_idx : start_idx + 30]:
        if STOP_RE.search(line):
            break

        if STYLE_RE.search(line) and looks_like_price_context(line):
            break

        color_codes.extend(extract_color_codes_from_text(line))

    return unique_preserve_order(color_codes)


def extract_color_codes_from_slide_shapes(slide, slide_number: int) -> list[str]:
    """
    Extract color codes from all text boxes on the slide.
    This catches separate labels under product images, like 010, 274, 393, 419.
    """
    color_codes: list[str] = []
    items = get_slide_text_items(slide)

    for item in items:
        color_codes.extend(extract_color_codes_from_text(item["text"], slide_number=slide_number))

    return unique_preserve_order(color_codes)


# -----------------------------
# Section extraction
# -----------------------------

def extract_section(lines: list[str], start_idx: int, label: str, stop_labels: Iterable[str]) -> str:
    results: list[str] = []
    active = False
    label_norm = label.lower().rstrip(":")
    stop_norm = tuple(s.lower().rstrip(":") for s in stop_labels)

    for line in lines[start_idx : start_idx + 60]:
        line_norm = line.lower().rstrip(":")

        if active:
            if any(line_norm.startswith(s) for s in stop_norm):
                break
            if STYLE_RE.search(line) and looks_like_price_context(line):
                break
            results.append(line)

        elif line_norm.startswith(label_norm):
            after_label = line[len(label) :].strip()
            if after_label:
                results.append(after_label)
            active = True

    return " | ".join(results)


# -----------------------------
# Main extraction
# -----------------------------

def extract_products_from_pptx(
    pptx_path: str | Path,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    prs = Presentation(str(pptx_path))

    raw_rows: list[dict] = []
    product_rows: list[dict] = []
    used_line_indices: set[tuple[int, int]] = set()
    current_category: str | None = None

    for slide_number, slide in enumerate(prs.slides, start=1):
        text_items = get_slide_text_items(slide)
        lines = get_slide_lines(slide)
        detected_category = detect_slide_category(lines)
        if detected_category:
            current_category = detected_category

        raw_rows.append(
            {
                "slide_number": slide_number,
                "category": current_category,
                "slide_text": "\n".join(lines),
            }
        )

        # Find all product-start lines on the slide.
        product_start_indices: list[int] = []
        for idx, slide_line in enumerate(lines):
            lookahead = " ".join(lines[idx + 1 : idx + 3])
            if is_new_product_start(slide_line, lookahead):
                product_start_indices.append(idx)

        slide_shape_colors = extract_color_codes_from_slide_shapes(slide, slide_number)

        # Special handling for Nike accessories/bags slides.
        # These use codes like N.100.3478.091 and prices on the following lines.
        accessory_line_items = {
            (item["shape_index"], item.get("line_index", 0)): idx
            for idx, item in enumerate(text_items)
        }

        for item in text_items:
            line = item["text"]
            i = accessory_line_items.get((item["shape_index"], item.get("line_index", 0)), 0)
            accessory_codes = ACCESSORY_STYLE_COLOR_RE.findall(line)
            if not accessory_codes:
                continue

            combined = get_accessory_price_context(text_items, item, lines, i)
            # Use aggressive mode for bags to find both prices even if labels are unclear
            wholesale, retail = extract_prices(combined, aggressive=True)
            status_match = STATUS_RE.search(combined)
            status = status_match.group(1).upper() if status_match else None
            product_name = previous_product_name(lines, i) or next_product_name(lines, i)

            material = extract_section(
                lines,
                i + 1,
                "Material Content",
                ["Features and Benefits", "Key Features:", "Overview:", "Product Details:"],
            )

            key_features = extract_section(
                lines,
                i + 1,
                "Features and Benefits",
                ["Material Content", "Material:", "Overview:", "Product Details:"],
            )

            for full_code in accessory_codes:
                style_code, color_code = split_accessory_style_color(full_code)

                product_rows.append(
                    {
                        "slide_number": slide_number,
                        "category": current_category,
                        "product_name": product_name,
                        "style_codes": style_code,
                        "wholesale_price": wholesale,
                        "retail_price": retail,
                        "status": status,
                        "color_codes": color_code,
                        "color_count": 1 if color_code else 0,
                        "material": material,
                        "key_features": key_features,
                        "raw_product_line": combined,
                    }
                )

            used_line_indices.add((slide_number, i))

        for i, line in enumerate(lines):
            lookahead = " ".join(lines[i + 1 : i + 3])

            if (slide_number, i) in used_line_indices:
                continue

            if not is_new_product_start(line, lookahead):
                continue

            style_codes = STYLE_RE.findall(line)
            if not style_codes:
                continue

            first_style = style_codes[0]
            end_idx = i
            combined = line

            # Combine nearby lines if price/status is split from style code.
            for j in range(i + 1, min(i + 4, len(lines))):
                if looks_like_price_context(combined):
                    break
                combined += " " + lines[j]
                end_idx = j

            if not looks_like_price_context(line):
                for j in range(i + 1, min(i + 4, len(lines))):
                    if looks_like_price_context(lines[j]):
                        combined = line + " " + lines[j]
                        end_idx = j
                        break

            style_codes = unique_preserve_order(STYLE_RE.findall(combined))
            style_token_colors = extract_colors_from_style_color_tokens(combined)
            wholesale, retail = extract_prices(combined)

            status_match = STATUS_RE.search(combined)
            status = status_match.group(1).upper() if status_match else None

            product_name = build_product_name(line, lines, i, first_style)

            # Original line-based extraction.
            line_colors = unique_preserve_order(
                style_token_colors + extract_color_codes_from_lines(lines, end_idx + 1)
            )

            # If the slide has only one product, use all color labels found on the slide.
            # This catches colors stored as separate text boxes.
            if len(product_start_indices) <= 1:
                color_codes = unique_preserve_order(line_colors + slide_shape_colors)
            else:
                # Multi-product slides are harder.
                # Do NOT assign all slide colors to every product, because that creates false extra rows.
                # Keep only the safer line-based colors.
                color_codes = line_colors

            material = extract_section(
                lines,
                end_idx + 1,
                "Material:",
                ["Key Features:", "Overview:", "Product Details:"],
            )

            key_features = extract_section(
                lines,
                end_idx + 1,
                "Key Features:",
                ["Overview:", "Material:", "Product Details:"],
            )

            product_rows.append(
                {
                    "slide_number": slide_number,
                    "category": current_category,
                    "product_name": product_name,
                    "style_codes": ",".join(style_codes),
                    "wholesale_price": wholesale,
                    "retail_price": retail,
                    "status": status,
                    "color_codes": ",".join(color_codes),
                    "color_count": len(color_codes),
                    "material": material,
                    "key_features": key_features,
                    "raw_product_line": combined,
                }
            )

            for j in range(i, end_idx + 1):
                used_line_indices.add((slide_number, j))

    products_df = pd.DataFrame(product_rows)
    raw_text_df = pd.DataFrame(raw_rows)

    exploded_rows: list[dict] = []

    for _, row in products_df.iterrows():
        styles = [x for x in str(row["style_codes"]).split(",") if x]
        colors = [x for x in str(row["color_codes"]).split(",") if x]

        if not styles:
            continue

        if colors:
            for style in styles:
                for color in colors:
                    exploded_rows.append(
                        {
                            "slide_number": row["slide_number"],
                            "category": row.get("category"),
                            "product_name": row["product_name"],
                            "style_code": style,
                            "color_code": color,
                            "style_color_key": build_style_color_key(style, color),
                            "wholesale_price": row["wholesale_price"],
                            "retail_price": row["retail_price"],
                            "status": row["status"],
                        }
                    )
        else:
            for style in styles:
                exploded_rows.append(
                    {
                        "slide_number": row["slide_number"],
                        "category": row.get("category"),
                        "product_name": row["product_name"],
                        "style_code": style,
                        "color_code": None,
                        "style_color_key": style,
                        "wholesale_price": row["wholesale_price"],
                        "retail_price": row["retail_price"],
                        "status": row["status"],
                    }
                )

    style_color_df = pd.DataFrame(exploded_rows)

    return products_df, style_color_df, raw_text_df


# -----------------------------
# Optional validation report
# -----------------------------

def write_validation_report(style_colors: pd.DataFrame, expected_csv: Path, output_folder: Path) -> None:
    """
    Compare the extracted style-color output to a validation CSV.

    The validation CSV should have at least:
        slide_number, product_name, style_code, color_code, style_color_key
    """
    if not expected_csv.exists():
        print("No validation CSV found. Skipping validation report.")
        return

    expected = pd.read_csv(expected_csv)
    actual = style_colors.copy()

    for df in (expected, actual):
        df["slide_number"] = pd.to_numeric(df["slide_number"], errors="coerce").astype("Int64")
        df["style_color_key"] = df["style_color_key"].astype(str).str.strip()

    compare = expected.merge(
        actual,
        on=["slide_number", "style_color_key"],
        how="outer",
        suffixes=("_expected", "_actual"),
        indicator=True,
    )

    missing = compare[compare["_merge"] == "left_only"].copy()
    extra = compare[compare["_merge"] == "right_only"].copy()

    product_name_mismatch = compare[
        (compare["_merge"] == "both")
        & (
            compare["product_name_expected"].fillna("").astype(str).str.strip()
            != compare["product_name_actual"].fillna("").astype(str).str.strip()
        )
    ].copy()

    price_mismatch = compare[
        (compare["_merge"] == "both")
        & (
            (compare["wholesale_price_expected"].fillna(-1) != compare["wholesale_price_actual"].fillna(-1))
            | (compare["retail_price_expected"].fillna(-1) != compare["retail_price_actual"].fillna(-1))
        )
    ].copy()

    status_mismatch = compare[
        (compare["_merge"] == "both")
        & (
            compare["status_expected"].fillna("").astype(str).str.strip().str.upper()
            != compare["status_actual"].fillna("").astype(str).str.strip().str.upper()
        )
    ].copy()

    missing.to_csv(output_folder / "validation_missing_rows.csv", index=False)
    extra.to_csv(output_folder / "validation_extra_rows.csv", index=False)
    product_name_mismatch.to_csv(output_folder / "validation_product_name_mismatch.csv", index=False)
    price_mismatch.to_csv(output_folder / "validation_price_mismatch.csv", index=False)
    status_mismatch.to_csv(output_folder / "validation_status_mismatch.csv", index=False)

    print("\nValidation report:")
    print(f"Missing rows: {len(missing)}")
    print(f"Extra rows: {len(extra)}")
    print(f"Product-name mismatches: {len(product_name_mismatch)}")
    print(f"Price mismatches: {len(price_mismatch)}")
    print(f"Status mismatches: {len(status_mismatch)}")
