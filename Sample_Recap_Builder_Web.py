"""
Sample Recap Builder — Web App
Generates the Apparel Store Bought Samples (SBS) meeting deck.

Workflow:
  Setup → Presenters → Analyze → Catalog → Build PPTX

Slides are grouped: Presenter → Category → Brand → Gender.
"""

import streamlit as st
import os, io, json, base64, re, time, tempfile, hashlib, logging
from datetime import date, datetime
from concurrent.futures import ThreadPoolExecutor, as_completed, wait, FIRST_COMPLETED
import threading

try:
    from PIL import Image as PILImage, ImageOps as PILImageOps
    PIL_OK = True
except ImportError:
    PIL_OK = False

try:
    import pillow_heif
    pillow_heif.register_heif_opener()
    HEIC_OK = True
except ImportError:
    HEIC_OK = False

try:
    import anthropic
    ANTHROPIC_OK = True
except ImportError:
    ANTHROPIC_OK = False

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE
    PPTX_OK = True
except ImportError:
    PPTX_OK = False


APP_VERSION = "1.0.0"
APP_TITLE   = "Sample Recap Builder"

# ── Slide / style constants ──────────────────────────────────────────────────
if PPTX_OK:
    SW = Inches(13.33)
    SH = Inches(7.5)
    P_BLACK       = RGBColor(0x00, 0x00, 0x00)   # pure black backgrounds (source style)
    P_BG          = RGBColor(0xFA, 0xF8, 0xF5)   # warm off-white (sample-slide BG)
    P_DARK        = RGBColor(0x1A, 0x1A, 0x1A)
    P_INK         = RGBColor(0x0F, 0x10, 0x13)
    P_MUTED       = RGBColor(0x6B, 0x6F, 0x76)
    P_LINE        = RGBColor(0xD8, 0xD3, 0xCB)
    P_CARD        = RGBColor(0xFF, 0xFF, 0xFF)
    P_ACCENT      = RGBColor(0xC8, 0x10, 0x2E)
    P_TINT        = RGBColor(0xEE, 0xEA, 0xE0)
    P_FOOTER_BG   = RGBColor(0xFF, 0xFF, 0xFF)   # white footer bar
    P_FOOTER_TXT  = RGBColor(0x1A, 0x1A, 0x1A)

# Fonts — match the source SBS deck. Impact/Arial Black ship on every Windows
# install so PowerPoint will render them exactly. Don't change to a "fancier"
# face without the user's say-so — the source uses condensed bold display.
TITLE_FONT   = "Impact"
HEADER_FONT  = "Arial Black"
LABEL_FONT   = "Arial"
BODY_FONT    = "Arial"

ASSETS_DIR  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
# Dark-mode logo: transparent background, white wordmark, red heart preserved.
# Generated from haddad_logo.png by replacing the connected white background
# (corner-flood-fill) with transparency and inverting the dark wordmark text
# to white. Looks correct on the black header bar.
LOGO_PATH   = os.path.join(ASSETS_DIR, "haddad_logo_dark.png")


# ── Canonical orderings ──────────────────────────────────────────────────────
CATEGORY_ORDER = [
    "Tops",
    "Bottoms",
    "Sets",
    "Dresses",
    "Outerwear",
    "Activewear",
    "Sleepwear",
    "Swimwear",
    "Underwear",
    "Socks",
    "Baby",
    "Accessories",
    "Footwear",
    "Other",
]
CATEGORY_INDEX = {c.lower(): i for i, c in enumerate(CATEGORY_ORDER)}

GENDER_ORDER = [
    "Boys",
    "Girls",
    "Unisex Kids",
    "Baby Boys",
    "Baby Girls",
    "Baby Unisex",
    "Mens",
    "Womens",
    "Adult Unisex",
    "Unspecified",
]
GENDER_INDEX = {g.lower(): i for i, g in enumerate(GENDER_ORDER)}


# ── Pricing for cost estimate ─────────────────────────────────────────────────
PRICING = {
    "haiku":  (1.00,  5.00),   # input, output per 1M tokens
    "sonnet": (3.00, 15.00),
}
MODEL_IDS = {
    "haiku":  "claude-haiku-4-5-20251001",
    "sonnet": "claude-sonnet-4-6",
}


def _price_key_for_model(model_id):
    """Pick the PRICING bucket that matches a model_id string. Sonnet's
    rates are ~3× Haiku's, so charging Haiku rates against Sonnet usage
    understated the displayed cost by roughly two-thirds."""
    mid = (model_id or "").lower()
    if "sonnet" in mid:
        return "sonnet"
    return "haiku"

# ── Scaling caps (enforced at upload to prevent runtime OOMs / timeouts) ─────
# At ~150KB per re-encoded JPEG these limits keep peak temp-dir usage under
# ~75MB and keep the 4-stage analysis pipeline finishing in roughly 4 minutes
# of wall time on the parallelized path. If the team needs more, the right
# move is a queued/persisted background job, not just bigger numbers.
MAX_PHOTOS_PER_PRESENTER = 150
MAX_PHOTOS_TOTAL         = 500


# ── Data models ──────────────────────────────────────────────────────────────
class SampleRecord:
    def __init__(self, original_path, preview_path, filename):
        self.original_path = original_path
        self.preview_path  = preview_path
        self.filename      = filename
        # User / meeting fields
        self.presenter      = ""
        self.meeting_date   = ""
        self.bought_from    = ""
        self.price          = ""
        # AI-extracted classification fields
        self.brand          = ""
        self.category       = ""
        self.gender         = ""
        self.product        = ""
        self.colorways      = ""
        self.fabric         = ""
        self.details        = ""
        self.confidence     = "MEDIUM"
        # How sure the AI was about the rotation it suggested at analysis time.
        # "HIGH" = confident, "LOW" = guess. LOW triggers Stage-4 verification.
        self.rotation_confidence = "HIGH"
        # True when the auto-rotator pipeline finished without confirming the
        # orientation is correct. Surfaces as a yellow warning in the catalog
        # UI so the user knows to double-check this photo.
        self.orientation_flag = False
        self.analysis_error = ""
        self.notes          = ""
        # Color palette — list of #RRGGBB hex strings the user has pinned
        # from the AI-suggested dominant colors of this sample's image.
        self.color_palette  = []
        # Suggested colors (auto-extracted) cached so we don't re-quantize on every render
        self.suggested_colors = []
        # If True, this sample's image will be added to the PREVIOUS sample's
        # slide instead of getting its own slide. Useful for back-view, detail,
        # or alternate-color shots of the same physical garment. The previous
        # sample's metadata is what shows on the combined slide.
        self.merge_with_previous = False
        # Stable identity of the PRIMARY sample this sample is merged into
        # (the primary's `filename`). Empty string = this sample is its own
        # primary and gets its own slide. We resolve groups by filename
        # rather than by adjacency so that re-sorting after a category /
        # brand / gender edit can't reassign a merged photo to an unrelated
        # primary. merge_with_previous remains as the catalog checkbox
        # state; merge_into is the source of truth for build_deck.
        self.merge_into = ""


class PresenterConfig:
    def __init__(self, name=""):
        self.name        = name
        self.samples     = []   # list of SampleRecord
        self.bought_from_hint = ""  # optional default if presenter has a single source


class MeetingDeck:
    def __init__(self):
        self.meeting_date = ""
        self.title        = "APPAREL STORE BOUGHT SAMPLE MEETING"
        self.presenters   = []   # list of PresenterConfig


# ── Utilities ────────────────────────────────────────────────────────────────
def file_hash(uploaded_file):
    data = uploaded_file.read()
    uploaded_file.seek(0)
    return hashlib.md5(data).hexdigest()


def _exif_orient_then_rgb(im):
    """Apply EXIF rotation aggressively. Tries three paths because HEIC/iPhone
    photos sometimes store orientation in places PIL's default exif_transpose
    misses, leading to sideways saves and sideways AI inputs."""
    # Path 1: pre-load orientation BEFORE any transform (some HEIC images lose
    # EXIF after .convert/.copy/.thumbnail).
    pre_ori = None
    try:
        raw_exif = im.getexif() if hasattr(im, "getexif") else None
        if raw_exif:
            pre_ori = raw_exif.get(0x0112)
    except Exception:
        pre_ori = None
    # Path 2: standard exif_transpose
    try:
        im = PILImageOps.exif_transpose(im)
    except Exception:
        pass
    # Path 3: if exif_transpose didn't change anything (image still tagged as
    # rotated), apply manually.
    if pre_ori in (3, 6, 8):
        try:
            still_tagged = False
            now_exif = im.getexif() if hasattr(im, "getexif") else None
            if now_exif and now_exif.get(0x0112) in (3, 6, 8):
                still_tagged = True
            if still_tagged:
                if pre_ori == 3:
                    im = im.rotate(180, expand=True)
                elif pre_ori == 6:
                    im = im.rotate(270, expand=True)
                elif pre_ori == 8:
                    im = im.rotate(90, expand=True)
        except Exception:
            pass
    return im.convert("RGB")


# Per-category expected aspect ratio bands (w/h) for laid-flat or hung apparel
# photos. Used by the Stage-3 sanity check to decide whether to spend a
# Stage-4 verification API call. Categories not in this map are skipped — for
# those (Footwear, Accessories, Socks, Baby, Other) the AR signal is too noisy
# to be useful and the false-positive rate would burn API calls without
# improving accuracy. 99%+ of sample photos are in the apparel categories
# below, where the AR signal is strong.
ORIENTATION_AR_BANDS = {
    "Tops":       (0.50, 1.05),
    "Bottoms":    (0.55, 1.20),
    "Sets":       (0.55, 1.20),
    "Dresses":    (0.40, 1.00),
    "Outerwear":  (0.50, 1.10),
    "Activewear": (0.50, 1.15),
    "Sleepwear":  (0.55, 1.20),
    "Swimwear":   (0.45, 1.20),
    "Underwear":  (0.50, 1.30),
}


def _orientation_ar_suspicious(rec):
    """True if the post-rotation aspect ratio is outside the expected band
    for this category. Returns False for categories without an AR prior, for
    missing files, and for borderline ratios."""
    if not PIL_OK or not os.path.exists(rec.preview_path):
        return False
    band = ORIENTATION_AR_BANDS.get(rec.category)
    if not band:
        return False
    try:
        with PILImage.open(rec.preview_path) as im:
            w, h = im.size
    except Exception:
        return False
    if not w or not h:
        return False
    ar = w / h
    lo, hi = band
    return ar < lo or ar > hi


def _encode_one_upload(raw_bytes, dest_path):
    """Decode (HEIC/JPEG/PNG), exif-rotate, downscale to 1568px, save JPEG."""
    if not raw_bytes:
        return False
    if PIL_OK:
        try:
            with PILImage.open(io.BytesIO(raw_bytes)) as im:
                im = _exif_orient_then_rgb(im)
                w, h = im.size
                if max(w, h) > 1568:
                    s = 1568 / max(w, h)
                    im = im.resize((int(w * s), int(h * s)), PILImage.LANCZOS)
                buf = io.BytesIO()
                im.save(buf, "JPEG", quality=85)
            with open(dest_path, "wb") as fh:
                fh.write(buf.getvalue())
            return True
        except Exception as e:
            logging.warning("encode failed for %s: %s", dest_path, e)
    try:
        with open(dest_path, "wb") as fh:
            fh.write(raw_bytes)
        return True
    except Exception:
        return False


def rotate_image_file(path, degrees):
    """Rotate the image at `path` by the given degrees (90, 180, 270 — clockwise
    when negative). Overwrites in place. Used by the Review step's rotate buttons.

    For the 90/180/270 case (the only values that ever come in from the UI or
    the AI auto-rotator) we use PIL's `transpose` which is a pixel-exact axis
    swap — no resampling, no interpolation. We still re-encode JPEG, which is
    lossy, but bumping quality to 95 keeps repeated rotation cycles visually
    indistinguishable from the source for the kind of editing the user is
    likely to do (one or two rotations to fix a sideways photo).
    """
    if not PIL_OK or not os.path.exists(path):
        return False
    try:
        deg = int(degrees) % 360
        with PILImage.open(path) as im:
            if deg == 0:
                return True
            if deg == 90:
                rotated = im.transpose(PILImage.Transpose.ROTATE_270)  # CW 90
            elif deg == 180:
                rotated = im.transpose(PILImage.Transpose.ROTATE_180)
            elif deg == 270:
                rotated = im.transpose(PILImage.Transpose.ROTATE_90)   # CW 270 = CCW 90
            else:
                # Arbitrary angle (rare path): resampled rotate, expand.
                rotated = im.rotate(-deg, expand=True)
            rotated.convert("RGB").save(path, "JPEG", quality=95)
        return True
    except Exception as e:
        logging.warning("rotate failed for %s: %s", path, e)
        return False


def _hex(rgb):
    return "#{:02X}{:02X}{:02X}".format(*rgb)


def _hex_to_rgb_tuple(h):
    h = (h or "").lstrip("#")
    if len(h) != 6:
        return (0, 0, 0)
    try:
        return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except Exception:
        return (0, 0, 0)


def extract_dominant_colors(path, n=8, ignore_near_white=True):
    """Quick dominant-color extraction using PIL median-cut quantization.
    Returns up to `n` hex strings ordered by frequency. Ignores near-pure-white
    pixels (the photo backdrop) when `ignore_near_white` is True."""
    if not PIL_OK or not path or not os.path.exists(path):
        return []
    try:
        with PILImage.open(path) as im:
            im = im.convert("RGB")
            # Downsample for speed
            im.thumbnail((220, 220), PILImage.LANCZOS)
            # Quantize via median cut to ~3*n colors then sort by area
            q = im.quantize(colors=max(8, n * 3), method=0)
            pal = q.getpalette()
            counts = sorted(q.getcolors(), key=lambda c: -c[0])
            out = []
            for cnt, idx in counts:
                r, g, b = pal[idx*3], pal[idx*3+1], pal[idx*3+2]
                if ignore_near_white and r > 240 and g > 240 and b > 240:
                    continue
                # Dedupe near-identical entries
                if any(abs(r - rr) + abs(g - gg) + abs(b - bb) < 24
                       for rr, gg, bb in [_hex_to_rgb_tuple(h) for h in out]):
                    continue
                out.append(_hex((r, g, b)))
                if len(out) >= n:
                    break
            return out
    except Exception as e:
        logging.warning("color extract failed for %s: %s", path, e)
        return []


def _parallel_encode_uploads(jobs, progress_label="Processing photos"):
    if not jobs:
        return []
    saved = [None] * len(jobs)
    workers = min(8, max(1, len(jobs)))
    progress = None
    try:
        progress = st.progress(0.0, text=f"{progress_label} (0/{len(jobs)})")
    except Exception:
        progress = None
    completed = 0
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {pool.submit(_encode_one_upload, raw, dest): (i, name, dest)
                   for i, (name, raw, dest) in enumerate(jobs)}
        for fut in as_completed(futures):
            i, name, dest = futures[fut]
            try:
                ok = fut.result()
            except Exception:
                ok = False
            if ok and os.path.exists(dest):
                saved[i] = {"name": name, "path": dest}
            completed += 1
            if progress is not None:
                try:
                    progress.progress(completed / len(jobs),
                                      text=f"{progress_label} ({completed}/{len(jobs)})")
                except Exception:
                    pass
    if progress is not None:
        try:
            progress.empty()
        except Exception:
            pass
    return [s for s in saved if s is not None]


def encode_image_b64(path, max_px=1568):
    if not PIL_OK:
        with open(path, "rb") as fh:
            return base64.b64encode(fh.read()).decode("utf-8")
    try:
        with PILImage.open(path) as im:
            im = PILImageOps.exif_transpose(im)
            w, h = im.size
            if max(w, h) > max_px:
                s = max_px / max(w, h)
                im = im.resize((int(w * s), int(h * s)), PILImage.LANCZOS)
            buf = io.BytesIO()
            im.convert("RGB").save(buf, "JPEG", quality=85)
        return base64.b64encode(buf.getvalue()).decode("utf-8")
    except Exception:
        with open(path, "rb") as fh:
            return base64.b64encode(fh.read()).decode("utf-8")


def parse_json_response(text):
    if not text:
        return None
    s = text.strip()
    fence = re.search(r"```(?:json)?\s*(.+?)\s*```", s, re.DOTALL)
    if fence:
        s = fence.group(1).strip()
    try:
        return json.loads(s)
    except Exception:
        m = re.search(r"\{.*\}", s, re.DOTALL)
        if m:
            try:
                return json.loads(m.group(0))
            except Exception:
                pass
    return None


def get_temp_dir():
    if "_tmp_dir" not in st.session_state:
        st.session_state._tmp_dir = tempfile.mkdtemp(prefix="srb_samples_")
    return st.session_state._tmp_dir


def _get_secret(name, default=""):
    try:
        v = st.secrets.get(name, default)
        if v and v != "PASTE-YOUR-KEY-HERE":
            return v
    except Exception:
        pass
    return os.environ.get(name, default) or default


def _get_anthropic_api_key():
    return _get_secret("ANTHROPIC_API_KEY", "") or st.session_state.get("api_key", "")


def normalize_category(value):
    if not value:
        return "Other"
    v = value.strip().lower()
    aliases = {
        "tops": "Tops", "top": "Tops", "tee": "Tops", "t-shirt": "Tops",
        "tshirt": "Tops", "shirt": "Tops", "polo": "Tops", "hoody": "Tops",
        "hoodie": "Tops", "sweater": "Tops", "tank": "Tops",
        "bottoms": "Bottoms", "bottom": "Bottoms", "pants": "Bottoms",
        "shorts": "Bottoms", "jogger": "Bottoms", "joggers": "Bottoms",
        "jeans": "Bottoms", "skirt": "Bottoms", "leggings": "Bottoms",
        "sets": "Sets", "set": "Sets", "2-piece": "Sets", "2 piece": "Sets",
        "matching set": "Sets",
        "dress": "Dresses", "dresses": "Dresses", "romper": "Dresses",
        "jumpsuit": "Dresses", "coverall": "Dresses",
        "outerwear": "Outerwear", "jacket": "Outerwear", "coat": "Outerwear",
        "puffer": "Outerwear", "vest": "Outerwear", "windbreaker": "Outerwear",
        "activewear": "Activewear", "athletic": "Activewear", "performance": "Activewear",
        "sleepwear": "Sleepwear", "pajamas": "Sleepwear", "pjs": "Sleepwear",
        "swim": "Swimwear", "swimwear": "Swimwear", "swimsuit": "Swimwear",
        "trunks": "Swimwear", "bikini": "Swimwear",
        "underwear": "Underwear", "briefs": "Underwear", "boxers": "Underwear",
        "bra": "Underwear",
        "socks": "Socks", "sock": "Socks",
        "baby": "Baby", "infant": "Baby", "newborn": "Baby", "bodysuit": "Baby",
        "onesie": "Baby",
        "accessories": "Accessories", "hat": "Accessories", "cap": "Accessories",
        "bag": "Accessories", "backpack": "Accessories", "belt": "Accessories",
        "footwear": "Footwear", "shoes": "Footwear", "sneakers": "Footwear",
        "boots": "Footwear", "sandals": "Footwear",
    }
    if v in aliases:
        return aliases[v]
    for k, mapped in aliases.items():
        if k in v:
            return mapped
    cap = v.title()
    if cap in CATEGORY_ORDER:
        return cap
    return "Other"


def normalize_gender(value):
    if not value:
        return "Unspecified"
    v = value.strip().lower()
    if "baby" in v or "infant" in v or "newborn" in v:
        if "boy" in v: return "Baby Boys"
        if "girl" in v: return "Baby Girls"
        return "Baby Unisex"
    if "men" in v and "wo" not in v: return "Mens"
    if "women" in v: return "Womens"
    if "adult" in v and ("uni" in v or "neutral" in v): return "Adult Unisex"
    if "boy" in v and "wo" not in v: return "Boys"
    if "girl" in v: return "Girls"
    if "uni" in v or "neutral" in v: return "Unisex Kids"
    return "Unspecified"


def category_sort_key(cat):
    return CATEGORY_INDEX.get((cat or "").lower(), 999)


def gender_sort_key(g):
    return GENDER_INDEX.get((g or "").lower(), 999)


# ── AI analysis ──────────────────────────────────────────────────────────────
ANALYSIS_SYSTEM = (
    "You are a meticulous apparel sample analyst for Haddad Brands, a children's and adult "
    "apparel design house. Samples include kids, adult men's, adult women's, and unisex "
    "items across every category — DO NOT default to baby/infant/diaper/onesie/thermal "
    "labels just because a garment is laid flat or has a soft silhouette. You read hangtags, "
    "neck labels, packaging, and visible logos to extract structured product metadata. You "
    "are precise, conservative, and never invent information. When uncertain about a "
    "specific field, leave that field empty rather than guessing. Critical: identify the "
    "garment by its STRUCTURAL features — neckline shape, collar type, presence of "
    "drawstring/zipper/buttons, sleeve length, hemline, waistband — NOT by its proportions "
    "in the photo (which can mislead because of camera angle and how the garment is laid)."
)


def _build_analysis_prompt(presenter_name, batch_size, default_bought_from=""):
    cat_str = ", ".join(CATEGORY_ORDER)
    gender_str = ", ".join(GENDER_ORDER)
    bought_hint = (
        f"\nIf no store name is visible, you MAY default to '{default_bought_from}' (the "
        f"presenter's noted source). Otherwise leave bought_from empty."
    ) if default_bought_from else ""
    return f"""Presenter: {presenter_name}.{bought_hint}

Analyze each of the {batch_size} sample photos above IN ORDER. Each photo shows a single
store-bought apparel sample being presented at a sample meeting. Read EVERY tag, label,
sticker, and visible logo. Also evaluate the photo's ORIENTATION and the GARMENT'S COLORS.
Return a JSON object with exactly {batch_size} entries:

{{"samples": [
  {{
    "rotation_degrees": "INTEGER 0, 90, 180, or 270. How many degrees the image needs to be rotated CLOCKWISE so the garment appears in its natural upright orientation. 0 = already correct. 90 = the photo is currently rotated 90° counter-clockwise (sideways with top-of-garment on the right) and needs +90 CW to fix. 180 = upside-down. 270 = rotated 90° clockwise (top-of-garment on the left) and needs 270 CW (i.e. -90) to fix. JUDGE FROM: position of neckline/collar (should be at top), readability of any text on the garment or tag, position of waistband/hem (should be at bottom). For sweatshirts/tees the neckline/shoulder area should be at the TOP of the photo. For bottoms the waistband should be at TOP. For dresses the shoulders/neckline at TOP. If unclear, return 0.",
    "rotation_confidence": "HIGH or LOW. HIGH only when you can clearly identify the garment's natural top from a confident reference point (visible neckline, waistband, readable text/logo, hanger). LOW when the garment is ambiguous (square fold, unusual silhouette, no clear reference point) OR when the rotation might be wrong. A LOW rating triggers a more careful verification pass downstream — be honest, not optimistic.",
    "garment_colors": ["UP TO 6 hex color codes (#RRGGBB) representing the ACTUAL GARMENT'S colors — the fabric, prints, embroidery, and trim of the apparel item itself. EXCLUDE the photo backdrop (table/floor/paper), shadows, hangtags, hangers, and packaging. Order by visual prominence on the garment (most-dominant first). Each entry is a hex string like '#1A2B3C'. Even a 'solid' garment usually has 2-3 distinct tones (highlight, mid, shadow) — include them so the palette captures the real fabric appearance. For prints/graphics, include each notable color. Use the actual visible color values in the photo (not the platonic ideal of the color)."],
    "brand": "PRODUCT BRAND in CAPS — the MANUFACTURER/label brand, NOT the store. ONLY return a brand if you can READ it directly from a hangtag, neck label, sewn-in label, screen-printed logo, embroidered logo, or printed packaging — clearly and unambiguously. If the brand is partially obscured, blurry, cropped off, or you are inferring it from style/silhouette/store association — RETURN EMPTY STRING. Do NOT guess based on the look of the item, the retailer, or similar items in the batch. Examples of acceptable identifications (only if literally visible): NIKE, JORDAN, CONVERSE, HURLEY, LEVI'S, LACOSTE, POLO, ADIDAS, PUMA, GAP, OLD NAVY, PRIMARK, H&M, CAT & JACK, WONDER NATION, GARANIMALS, CARTER'S, OSHKOSH, 3BRAND, ROXY, QUIKSILVER. NOTE: store-house brands ARE brands (Cat & Jack is Target's, Wonder Nation is Walmart's) — but still only label them when the actual logo/text is visible. WHEN IN DOUBT, RETURN EMPTY STRING.",
    "category": "EXACTLY ONE OF: {cat_str}. See category definitions below.",
    "gender": "EXACTLY ONE OF: {gender_str}. Determine from sizing tag, styling, packaging.",
    "product": "Short PO-style label naming the actual garment type. Use straightforward, common apparel terminology. e.g. 'S/S Tee', 'L/S Tee', 'Polo S/S', 'Tank Top', 'Fleece Hoody', 'Zip Hoody', 'Crewneck Sweatshirt', 'Mesh Short', 'Sweat Short', 'Denim Jean', 'Fleece Jogger', 'Puffer Jacket', 'Bomber Jacket', 'Knit Sweater', 'Ribbed Tank', 'Slip Dress', 'Bodysuit', 'Onesie', 'Two-Piece Set'. Under 6 words. CRITICAL — DO NOT default to niche labels: NEVER use 'Leg Warmers', 'Thermal Underwear', 'Diaper Cover', 'Long Underwear', 'Compression Sleeve', or any other obscure label unless the hangtag literally says so. If you see two cylinders with elastic at one end → SHORTS (athletic / mesh / sweat short), NOT leg warmers. If you see a zip-front grey garment with a brand logo → ZIP HOODY or TRACK JACKET, NOT thermal underwear. If you see an elastic-waist garment laid flat → SHORT or PANT, NOT diaper cover. Default to mainstream apparel categories.",
    "colorways": "Plain-language color description. e.g. 'Black, Cream' or 'Tie-dye Pink'.",
    "fabric": "Material if labeled. e.g. '100% Cotton', 'Polyester Fleece'. Empty if not visible.",
    "details": "Notable features that are CLEARLY VISIBLE in the photo. One short phrase. e.g. 'Puff print front graphic', 'Acid wash finish', 'Drop shoulder oversized fit', 'Embroidered chenille patch'. ONLY describe features you can actually see — do NOT speculate about construction, lining, or hidden detailing. If nothing distinctive is clearly visible, RETURN EMPTY STRING. Do not pad with generic phrases like 'standard fit' or 'classic styling'.",
    "bought_from": "Retailer/store where sample was purchased — usually printed on the hangtag or price sticker. e.g. 'Primark', 'H&M', 'Old Navy', 'TJ Maxx', 'Target', 'Walmart', 'Macy's', 'Ross'. Empty if not visible.",
    "price": "Selling price ON THE TAG. Format '$X.99'. Check hangtags, stickers, packaging. Empty if not visible.",
    "confidence": "HIGH only when brand AND product AND category are all derived from clearly visible labels/tags/logos in the photo (no guessing). MEDIUM when most fields are confident but one or two are inferred. LOW when significant fields are guesses or the photo is partially obscured. Be strict — when uncertain, choose MEDIUM not HIGH."
  }}
]}}

ROTATION DETECTION — STEP BY STEP:
1. Find the garment's anatomical reference points. Use whichever apply:
   - Tops/Outerwear/Sets/Sleepwear: neckline, collar, shoulders → TOP. Hem → BOTTOM.
   - Bottoms: waistband (wider, with belt loops or elastic) → TOP. Leg openings → BOTTOM.
   - Dresses/Rompers/Jumpsuits: straps/neckline → TOP. Hem/leg openings → BOTTOM.
   - Swimwear (one-piece, bikini top, trunks): straps/waistband/back of suit → TOP.
     Leg openings, halter ties, or rear flap → BOTTOM. A one-piece swimsuit lying with
     the leg openings on the LEFT or RIGHT (instead of bottom) IS sideways.
   - Footwear: toe → LEFT or RIGHT (typically LEFT for one shoe), heel opposite.
   - Accessories (hats/bags): wearable opening → BOTTOM (the side that goes against the
     body or head). Logo text on the item should read left-to-right.
2. Where are those reference points ACTUALLY in the photo?
   - If the natural-top points to the RIGHT side of the photo → rotated CCW, answer 90.
   - If the natural-top points to the LEFT side  → rotated CW,  answer 270.
   - If the natural-top is at the BOTTOM        → upside-down,  answer 180.
   - If the natural-top is at the TOP           → already correct, answer 0.
3. Cross-check: any text/logo on the garment or tag should read horizontally
   left-to-right when oriented correctly. If the text is sideways, the photo is rotated.
4. If the garment has no clear top/bottom (a folded scarf, a flat square accessory),
   default to 0 — DO NOT guess a rotation.

GARMENT COLOR EXTRACTION:
- Look at the GARMENT ONLY. Ignore the photo backdrop completely.
- For SOLID color garments: pick 2-3 hex codes for the highlight, mid-tone, and shadow.
- For MULTI-color: pick the dominant base color first, then accent/print colors.
- Use the actual hue/saturation/lightness you see in the photo (not what the color "would be" under perfect lighting).
- Example: a navy hoody might be ['#1B2540', '#2E3B5C', '#0E1424'] for highlight/mid/shadow.
- Example: a tie-dye tee might be ['#F2C2D8', '#E899B8', '#9F4A70', '#FFFFFF'] for the fabric tones.

CATEGORY DEFINITIONS — USE EXACTLY ONE:
- Tops = upper-body garments (tee, polo, hoody, sweatshirt, crew, tank, henley, sweater).
- Bottoms = lower-body garments (pants, jeans, joggers, shorts, leggings, skirts) NOT in a set.
- Sets = TWO+ separate garments sold together (tee+short, hoody+jogger, polo+short).
- Dresses = ONE-piece garments: dress, romper, jumpsuit, coverall.
- Outerwear = jackets, coats, puffers, vests, windbreakers worn OVER other clothes.
- Activewear = explicit athletic/performance garments not better classified above.
- Sleepwear = pajamas, PJ sets, nightgowns, robes, loungewear.
- Swimwear = swimsuits, trunks, board shorts, rash guards, bikinis.
- Underwear = briefs, boxers, undershirts, bras, thermals.
- Socks = ALL socks (crew, ankle, no-show, sport).
- Baby = infant/newborn-specific apparel (sized 0-24M): bodysuits, onesies, layette.
- Accessories = hats, caps, bags, backpacks, belts, jewelry, headbands, gloves.
- Footwear = shoes, sneakers, boots, sandals.
- Other = anything that genuinely does not fit (very rare).

BRAND vs BOUGHT_FROM — these are DIFFERENT fields:
- brand = the manufacturer label (e.g. Cat & Jack)
- bought_from = the retail store where it was purchased (e.g. Target)
A sample bought at Target with a Cat & Jack tag → brand="CAT & JACK", bought_from="Target".

PRICE: read the tag exactly. Look for '$X.99' on hangtags, hanger stickers, or packaging.
NEVER invent a price. Empty string is acceptable when no price is visible.

CRITICAL: Exactly {batch_size} objects in order, matching the photos one-for-one."""


def _run_analysis_batch(client, model_id, samples, presenter_name, default_bought_from=""):
    if not samples:
        return 0.0
    content = []
    for i, rec in enumerate(samples):
        content.append({"type": "text", "text": f"[Sample {i+1} of {len(samples)}]"})
        b64 = encode_image_b64(rec.preview_path)
        content.append({"type": "image", "source": {"type": "base64",
                                                     "media_type": "image/jpeg",
                                                     "data": b64}})
    content.append({"type": "text", "text": _build_analysis_prompt(
        presenter_name, len(samples), default_bought_from)})

    cost = 0.0
    for attempt in range(4):
        try:
            resp = client.messages.create(
                model=model_id, max_tokens=8192,
                system=[{"type": "text", "text": ANALYSIS_SYSTEM,
                         "cache_control": {"type": "ephemeral"}}],
                messages=[{"role": "user", "content": content}],
                timeout=90.0,
            )
            if hasattr(resp, "usage"):
                u = resp.usage
                # Use the actual model's pricing — Sonnet is ~3× Haiku so a
                # hardcoded "haiku" bucket would understate Sonnet cost.
                p_in, p_out = PRICING[_price_key_for_model(model_id)]
                cost += (u.input_tokens * p_in + u.output_tokens * p_out) / 1_000_000
            if not resp.content:
                raise ValueError("empty response")
            parsed = parse_json_response(resp.content[0].text)
            if not parsed or "samples" not in parsed:
                raise ValueError("invalid JSON shape")
            # Truncation guard: if the model returned fewer entries than we
            # asked for (max_tokens cutoff, malformed JSON salvaged by the
            # parser), retry rather than silently leaving the tail of the
            # batch as defaults — that's what produced "Other" on the last
            # 6 slides when the final batch hit max_tokens.
            if not isinstance(parsed["samples"], list) or \
                    len(parsed["samples"]) < len(samples):
                raise ValueError(
                    f"short response: got {len(parsed.get('samples') or [])} "
                    f"of {len(samples)} entries")
            for j, info in enumerate(parsed["samples"]):
                if j >= len(samples):
                    break
                rec = samples[j]
                rec.brand       = (info.get("brand") or "").strip().upper()
                rec.category    = normalize_category(info.get("category"))
                rec.gender      = normalize_gender(info.get("gender"))
                rec.product     = (info.get("product") or "").strip()
                rec.colorways   = (info.get("colorways") or "").strip()
                rec.fabric      = (info.get("fabric") or "").strip()
                rec.details     = (info.get("details") or "").strip()
                rec.confidence  = (info.get("confidence") or "MEDIUM").strip().upper()
                bf = (info.get("bought_from") or "").strip()
                if bf:
                    rec.bought_from = bf
                elif default_bought_from and not rec.bought_from:
                    rec.bought_from = default_bought_from
                pr = (info.get("price") or "").strip()
                if pr and not rec.price:
                    rec.price = pr
                # ── Auto-rotate the photo if the AI says it's misoriented.
                # The model returns the CW degrees needed to make the garment
                # right-side-up. We apply it directly to the on-disk preview.
                try:
                    rot_raw = info.get("rotation_degrees", 0)
                    rot = int(rot_raw) if rot_raw not in (None, "") else 0
                    rot = rot % 360
                    if rot in (90, 180, 270):
                        rotate_image_file(rec.preview_path, rot)
                    rec.rotation_confidence = (
                        info.get("rotation_confidence") or "HIGH"
                    ).strip().upper()
                    if rec.rotation_confidence not in ("HIGH", "LOW"):
                        rec.rotation_confidence = "HIGH"
                except Exception as ex:
                    logging.warning("rotation parse failed: %s", ex)
                    rec.rotation_confidence = "LOW"
                # ── Save garment colors directly into color_palette so they
                # appear on the slide automatically. Keep up to 6, normalize
                # to #RRGGBB. Only overwrite if the user hasn't already pinned
                # a custom palette.
                try:
                    raw_colors = info.get("garment_colors") or []
                    cleaned = []
                    if isinstance(raw_colors, list):
                        for c in raw_colors:
                            if not isinstance(c, str):
                                continue
                            s = c.strip()
                            if not s.startswith("#"):
                                s = "#" + s
                            if re.fullmatch(r"#[0-9A-Fa-f]{6}", s):
                                cleaned.append(s.upper())
                    if cleaned:
                        rec.suggested_colors = cleaned[:6]
                        if not rec.color_palette:
                            rec.color_palette = list(cleaned[:6])
                except Exception as ex:
                    logging.warning("color parse failed: %s", ex)
                # Stage 3 (AR sanity check) and Stage 4 (verification) run as
                # a separate post-batch pass — see auto_verify_orientation.
            return cost
        except Exception as e:
            err = str(e)
            if "401" in err or "authentication" in err.lower() or "invalid x-api-key" in err.lower():
                raise RuntimeError(f"API_KEY_INVALID: {err}") from e
            is_rate = "429" in err or "rate" in err.lower() or "overloaded" in err.lower()
            # If the model hit max_tokens and returned fewer than `len(samples)`
            # entries, retrying the same batch with the same content won't help
            # — the input is identical and we'll truncate again. Split the
            # batch in half and recurse so each half fits inside max_tokens.
            # We do this on the LAST attempt only (after rate-limit / transient
            # retries have failed) and only when len(samples) > 1.
            is_short = "short response" in err.lower()
            if is_short and attempt >= 1 and len(samples) > 1:
                mid = len(samples) // 2
                logging.info("splitting truncated batch %d→%d+%d for %s",
                             len(samples), mid, len(samples) - mid, presenter_name)
                cost_a = _run_analysis_batch(client, model_id, samples[:mid],
                                              presenter_name, default_bought_from)
                cost_b = _run_analysis_batch(client, model_id, samples[mid:],
                                              presenter_name, default_bought_from)
                return cost + (cost_a or 0.0) + (cost_b or 0.0)
            if attempt == 3:
                for r in samples:
                    r.analysis_error = f"ERROR: {err[:80]}"
                logging.warning("analysis batch failed: %s", err)
                return cost
            time.sleep((4 * (2 ** attempt)) if is_rate else 2)
    return cost


# ── Stage 4: targeted orientation verification ───────────────────────────────
# A single-image, single-question follow-up call for samples that Stage 2 left
# uncertain (LOW rotation_confidence) or whose post-rotation aspect ratio
# looks wrong for their predicted category. One call per flagged sample, with
# a single retry after applying the suggested fix. Runs in parallel.

ORIENTATION_VERIFY_SYSTEM = (
    "You are a single-image orientation reviewer for apparel sample photos. "
    "Given one photo and the garment's category, you decide whether the "
    "garment is shown with its anatomical TOP at the top of the image, and "
    "if not, you specify exactly how many degrees clockwise the photo must "
    "be rotated to make it correct. Be decisive. When in doubt about the "
    "rotation amount, prefer 0 (don't rotate)."
)

# Per-category orientation cue used in the verification prompt. Tells the
# model what to look for so a focused single-image judgment is cleaner than
# the multi-field batch judgment.
_ORIENT_CAT_HINT = {
    "Tops":       "Neckline/collar at TOP, hem at BOTTOM. Sleeves extend left and right.",
    "Bottoms":    "Waistband (wider, with belt loops or elastic) at TOP, leg openings at BOTTOM.",
    "Sets":       "The two pieces' natural tops (necklines/waistbands) at TOP, hems at BOTTOM.",
    "Dresses":    "Straps/neckline at TOP, hem at BOTTOM.",
    "Outerwear":  "Collar/neckline at TOP, hem at BOTTOM. Sleeves extend left and right.",
    "Activewear": "Neckline at TOP for tops, waistband at TOP for bottoms.",
    "Sleepwear":  "Neckline or waistband at TOP, hem at BOTTOM.",
    "Swimwear": (
        "ONE-PIECE swimsuits/leotards: straps/neckline at TOP, leg openings at BOTTOM. "
        "TRUNKS / BOARD SHORTS / BIKINI BOTTOMS: waistband at TOP, leg openings at BOTTOM. "
        "If the leg openings are on the LEFT or RIGHT side (not at the bottom), the photo IS sideways."
    ),
    "Underwear":  "Waistband or shoulder straps at TOP, leg openings or hem at BOTTOM.",
}


def _orient_verify_call(client, model_id, image_path, category):
    """One verification API call. Returns (parsed_dict_or_None, cost)."""
    if not os.path.exists(image_path):
        return None, 0.0
    cat_hint = _ORIENT_CAT_HINT.get(category,
        "The garment's natural top should be at the top of the image.")
    b64 = encode_image_b64(image_path)
    prompt = (
        f"This is a single photo of a {category or 'garment'}.\n"
        f"Orientation rule: {cat_hint}\n\n"
        "Is the garment oriented correctly (its natural TOP at the TOP of the image)?\n"
        "Reply with JSON ONLY (no preamble, no commentary):\n"
        '{"oriented":"YES"|"NO","fix_cw":0|90|180|270}\n\n'
        "Use oriented=YES with fix_cw=0 when the photo is already correct. "
        "Use oriented=NO with fix_cw set to the number of degrees CLOCKWISE "
        "the photo must be rotated to make the garment upright."
    )
    try:
        resp = client.messages.create(
            model=model_id, max_tokens=120,
            system=[{"type": "text", "text": ORIENTATION_VERIFY_SYSTEM,
                     "cache_control": {"type": "ephemeral"}}],
            messages=[{"role": "user", "content": [
                {"type": "image", "source": {"type": "base64",
                                              "media_type": "image/jpeg",
                                              "data": b64}},
                {"type": "text", "text": prompt},
            ]}],
            timeout=60.0,
        )
        cost = 0.0
        if hasattr(resp, "usage"):
            u = resp.usage
            p_in, p_out = PRICING[_price_key_for_model(model_id)]
            cost = (u.input_tokens * p_in + u.output_tokens * p_out) / 1_000_000
        if not resp.content:
            return None, cost
        parsed = parse_json_response(resp.content[0].text)
        return parsed, cost
    except Exception as ex:
        logging.warning("orientation verify call failed: %s", ex)
        return None, 0.0


def _verify_one_sample_orientation(client, model_id, rec):
    """Run Stage 4 on a single record. Mutates rec.orientation_flag to True
    iff the sample is still suspect after one round of correction. Returns
    the dollar cost spent on this sample's verification calls."""
    cost = 0.0
    res, c1 = _orient_verify_call(client, model_id, rec.preview_path, rec.category)
    cost += c1
    if not res:
        # API failure — leave the photo alone, don't flag (avoid noise)
        return cost

    oriented = (res.get("oriented") or "").strip().upper() == "YES"
    try:
        fix = int(res.get("fix_cw", 0)) % 360
    except Exception:
        fix = 0

    if oriented:
        rec.orientation_flag = False
        return cost

    # Try the suggested fix once.
    if fix in (90, 180, 270):
        if rotate_image_file(rec.preview_path, fix):
            # Color palette was sampled from the pre-rotation image; let it
            # regenerate so swatches still match what's on screen.
            rec.suggested_colors = []
        # Single retry — verify the corrected image.
        res2, c2 = _orient_verify_call(
            client, model_id, rec.preview_path, rec.category)
        cost += c2
        if not res2:
            # Couldn't re-verify. We applied a rotation we couldn't confirm —
            # surface that to the user so they can sanity-check.
            rec.orientation_flag = True
            return cost
        oriented2 = (res2.get("oriented") or "").strip().upper() == "YES"
        rec.orientation_flag = not oriented2
    else:
        # Said NO but didn't suggest a usable fix. Leave the image as-is and flag.
        rec.orientation_flag = True
    return cost


def auto_verify_orientation(client, model_id, samples, on_progress=None):
    """Stage 4 orchestrator. Runs the targeted single-image verification call
    on every apparel-category sample (anything in ORIENTATION_AR_BANDS),
    regardless of Stage-2 confidence. We learned the hard way that Stage 2
    can return rotation=0 with HIGH confidence on an upside-down photo: the
    aspect ratio is identical, the per-field-confidence signal didn't trigger,
    and the user ended up rotating ~12 photos by hand.

    The trade-off: ~1 API call per apparel photo (so a 30-photo deck does
    ~30 calls in waves of 4 workers ≈ 15-25s extra wall time, ~$0.10-0.15
    extra). Footwear / Accessories / Socks / Baby / Other are skipped — they
    are rare in this app (<1% of samples) and their AR/orientation cues are
    too varied for a reliable verification prompt.

    Mutates rec.orientation_flag (True = still suspect after one retry,
    False = confirmed correct). Returns total dollar cost."""
    if not samples:
        return 0.0
    flagged = [r for r in samples
               if os.path.exists(r.preview_path)
               and r.category in ORIENTATION_AR_BANDS]
    if not flagged:
        return 0.0

    # Stage-4 worker pool. With 3 presenters running concurrently in
    # show_analyze and 6 verify workers each, global Stage-4 ceiling is
    # ~18 in-flight calls. Keeps wall time manageable on 100-photo presenters
    # (~30s instead of ~60s) without saturating rate limits.
    workers = min(6, len(flagged))
    total_cost = 0.0
    done = 0
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {pool.submit(_verify_one_sample_orientation,
                               client, model_id, r): r
                   for r in flagged}
        for fut in as_completed(futures):
            try:
                total_cost += fut.result() or 0.0
            except Exception as ex:
                logging.warning("orientation verify worker failed: %s", ex)
                # On worker failure, raise the flag — user should review.
                futures[fut].orientation_flag = True
            done += 1
            if on_progress:
                on_progress(done, len(flagged))
    return total_cost


def analyze_presenter_samples(client, model_id, presenter, on_progress=None):
    BATCH_SIZE = 6
    # Per-presenter Stage-2 batch concurrency. With show_analyze running up to
    # 3 presenters in parallel, the global Stage-2 ceiling is 3*3 = 9 in-flight
    # batches — comfortably inside Anthropic tier-4 limits.
    MAX_CONCURRENT = 3
    total = len(presenter.samples)
    if total == 0:
        return 0.0
    ranges = [(s, min(s + BATCH_SIZE, total)) for s in range(0, total, BATCH_SIZE)]
    completed = 0
    total_cost = 0.0

    with ThreadPoolExecutor(max_workers=MAX_CONCURRENT) as pool:
        futures = []
        for s, e in ranges:
            futures.append(pool.submit(
                _run_analysis_batch, client, model_id,
                presenter.samples[s:e], presenter.name, presenter.bought_from_hint))
        for fut in as_completed(futures):
            try:
                total_cost += fut.result() or 0.0
            except RuntimeError:
                raise
            except Exception as ex:
                logging.warning("batch error: %s", ex)
            completed += 1
            if on_progress:
                on_progress(completed, len(ranges), presenter.name)

    # Targeted retry for any record whose batch failed (transient API errors,
    # truncation that survived all in-batch retries, etc). Re-batches just
    # the failures in groups of BATCH_SIZE — much cheaper than asking the user
    # to re-run the whole presenter, and far better than letting failures
    # silently land as "Other" / blank fields on their slides.
    failed = [r for r in presenter.samples if r.analysis_error]
    if failed:
        logging.info("retrying %d failed records for %s",
                     len(failed), presenter.name)
        retry_ranges = [(s, min(s + BATCH_SIZE, len(failed)))
                        for s in range(0, len(failed), BATCH_SIZE)]
        with ThreadPoolExecutor(max_workers=MAX_CONCURRENT) as pool:
            retry_futs = []
            for s, e in retry_ranges:
                # Clear analysis_error on the records we're retrying so a
                # successful retry won't leave the red banner up.
                for r in failed[s:e]:
                    r.analysis_error = ""
                retry_futs.append(pool.submit(
                    _run_analysis_batch, client, model_id,
                    failed[s:e], presenter.name, presenter.bought_from_hint))
            for fut in as_completed(retry_futs):
                try:
                    total_cost += fut.result() or 0.0
                except RuntimeError:
                    raise
                except Exception as ex:
                    logging.warning("retry batch error: %s", ex)

    # Stage 4: targeted orientation verification on suspect photos. Only runs
    # on samples Stage 2 wasn't confident about or whose post-rotation aspect
    # ratio doesn't match the predicted category. For a typical 30-photo
    # presenter this is usually 0–3 photos and adds ~1–2s + a few cents.
    try:
        verify_cost = auto_verify_orientation(
            client, model_id, presenter.samples,
            on_progress=(
                lambda d, t: on_progress(len(ranges), len(ranges),
                                          f"{presenter.name} (verifying {d}/{t})")
                if on_progress else None))
        total_cost += verify_cost or 0.0
    except Exception as ex:
        logging.warning("orientation verification stage failed for %s: %s",
                        presenter.name, ex)
    return total_cost


# ── Duplicate detection (auto-merge similar photos onto one slide) ───────────
DUPLICATE_DETECTION_SYSTEM = (
    "You are a sample-photography QA reviewer. You look at a sequence of product "
    "photos and identify which ones show the SAME PHYSICAL GARMENT — same fabric, "
    "same construction, same colorway — just photographed from different angles "
    "(front, back, detail close-up, on hanger vs flat). You are EXTREMELY "
    "conservative: false-merging two distinct products onto one slide is a much "
    "worse mistake than leaving genuine duplicates separate. The default answer "
    "is 'separate'. Only merge when EVERY one of these match: garment silhouette, "
    "neckline/collar shape, sleeve/hem treatment, fabric texture, dominant color, "
    "and any visible graphic or print. Same brand or same retail backdrop is NOT "
    "evidence — most photos in a presenter's batch share both."
)


def _ai_detect_duplicate_groups(client, model_id, samples_subset):
    """Send a subset of consecutive samples to Claude. Returns (groups, cost):
    `groups` is a list of group lists — each inner list is a set of 1-indexed
    positions within the subset that all show the same garment. Singletons
    appear as 1-element lists. `cost` is dollars spent on this call.

    Cost: one API call per ~12 photos, so ~$0.03/presenter on Sonnet.
    """
    if len(samples_subset) < 2:
        return [[i + 1] for i in range(len(samples_subset))], 0.0
    content = []
    for i, rec in enumerate(samples_subset):
        content.append({"type": "text", "text": f"[Image {i+1}]"})
        b64 = encode_image_b64(rec.preview_path)
        content.append({"type": "image", "source": {"type": "base64",
                                                     "media_type": "image/jpeg",
                                                     "data": b64}})
    n = len(samples_subset)
    content.append({"type": "text", "text": f"""
Look at all {n} photos above IN ORDER. Each photo shows a single product sample.
SOME of these photos may show the SAME PHYSICAL GARMENT photographed multiple
times (front view, back view, detail shot, on-hanger vs flat). Group them.

Return JSON:
{{"groups": [[1], [2,3], [4], [5,6,7]]}}

Where each inner array lists the 1-indexed image numbers that show the same
garment. Singletons stay in their own one-element group.

GROUPING RULES — DEFAULT TO SEPARATE. Only merge when ALL of the following are
true. If even one is uncertain, the photos are SEPARATE garments:
- Same garment silhouette (e.g. both are quarter-zips, both are graphic tees,
  both are joggers — NOT one quarter-zip + one tee, even if same brand).
- Same neckline/collar (crew, V, polo, hood, mock, zip — must match exactly).
- Same sleeve length and same hem treatment.
- Same fabric texture and weight (jersey vs fleece vs french terry — distinct).
- Same dominant base color (a navy tee and a light-blue zip-up are SEPARATE,
  even if the brand label and backdrop look identical).
- Same print/graphic (different chest graphics = SEPARATE garments, period).

ALLOWED MERGES — only these scenarios:
- Front view + back view of the obviously identical garment.
- Garment on a hanger + the same garment laid flat next to it.
- A wide shot + a tight detail close-up of the same garment (label, stitch,
  or print) where the close-up is clearly cropped from the same item.

NOT MERGES — common false positives to AVOID:
- Two items from the same brand on the same backdrop = SEPARATE.
- Two items with similar tags or hangers = SEPARATE.
- Two items in similar (but not identical) colors = SEPARATE.
- Two tees with different chest graphics = SEPARATE.
- A quarter-zip and a tee from the same brand = SEPARATE.
- One garment in two sizes or two colorways = SEPARATE (each is its own sample).

When in doubt — DEFAULT TO SEPARATE.

Output JSON ONLY, no preamble. Every image 1..{n} must appear in exactly one group.
"""})
    try:
        resp = client.messages.create(
            model=model_id, max_tokens=2048,
            system=[{"type": "text", "text": DUPLICATE_DETECTION_SYSTEM,
                     "cache_control": {"type": "ephemeral"}}],
            messages=[{"role": "user", "content": content}],
            timeout=90.0,
        )
        cost = 0.0
        if hasattr(resp, "usage"):
            u = resp.usage
            p_in, p_out = PRICING[_price_key_for_model(model_id)]
            cost = (u.input_tokens * p_in + u.output_tokens * p_out) / 1_000_000
        if not resp.content:
            return [[i + 1] for i in range(n)], cost
        parsed = parse_json_response(resp.content[0].text)
        if not parsed or "groups" not in parsed:
            return [[i + 1] for i in range(n)], cost
        groups = parsed["groups"]
        # Validate every index appears exactly once
        seen = set()
        clean = []
        for g in groups:
            if not isinstance(g, list):
                continue
            valid_ids = []
            for x in g:
                try:
                    xi = int(x)
                    if 1 <= xi <= n and xi not in seen:
                        seen.add(xi)
                        valid_ids.append(xi)
                except Exception:
                    pass
            if valid_ids:
                clean.append(valid_ids)
        # Anything missing → its own group
        for i in range(1, n + 1):
            if i not in seen:
                clean.append([i])
        return clean, cost
    except Exception as ex:
        logging.warning("duplicate detection failed: %s", ex)
        return [[i + 1] for i in range(n)], 0.0


def ai_suggest_merges_for_presenter(client, model_id, presenter,
                                     on_progress=None):
    """Run AI duplicate detection on a presenter's samples (in their sorted
    order) and SET merge_with_previous on samples that the model groups with
    the previous one. Returns total cost in dollars.
    """
    sorted_samples = sort_samples(presenter.samples)
    if len(sorted_samples) < 2:
        return 0.0
    # Reset existing merge flags so we don't compound previous decisions
    for r in sorted_samples:
        r.merge_with_previous = False
        r.merge_into = ""
    CHUNK = 12
    total_cost = 0.0
    chunks = [sorted_samples[i:i + CHUNK]
              for i in range(0, len(sorted_samples), CHUNK)]

    def _process(chunk):
        # Each chunk is independent — the merge flags are only ever set on
        # records inside the chunk, never across chunks. Safe to parallelize.
        groups, cost = _ai_detect_duplicate_groups(client, model_id, chunk)
        for g in groups:
            if len(g) <= 1:
                continue
            ordered = sorted(g)
            # Only honor groups whose members are CONSECUTIVE in the chunk's
            # sorted order. Non-adjacent groups are dropped — protects against
            # the AI accidentally bridging unrelated items that happen to sit
            # between two visually similar ones.
            if ordered != list(range(ordered[0], ordered[-1] + 1)):
                continue
            primary = chunk[ordered[0] - 1]
            for pos in ordered[1:]:
                chunk[pos - 1].merge_with_previous = True
                # Stable identity: store the primary's filename so the merge
                # survives re-sort triggered by user edits to sort keys.
                chunk[pos - 1].merge_into = primary.filename
        return cost

    workers = min(4, max(1, len(chunks)))
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {pool.submit(_process, c): i for i, c in enumerate(chunks)}
        completed = 0
        for fut in as_completed(futures):
            try:
                total_cost += fut.result() or 0.0
            except Exception as ex:
                logging.warning("dup-detect chunk failed for %s: %s",
                                presenter.name, ex)
            completed += 1
            if on_progress:
                on_progress(completed, len(chunks), presenter.name)
    return total_cost


# ── PPTX generators ──────────────────────────────────────────────────────────
def _blank_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def _bg(slide, color):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    bg.fill.solid(); bg.fill.fore_color.rgb = color
    bg.line.fill.background()


def _box(slide, l, t, w, h, fill, line=None, line_w=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        if line_w is not None:
            sh.line.width = line_w
    return sh


def _txt(slide, text, l, t, w, h, *, size=14, bold=False,
         color=None, font="Corbel", align="left", anchor="top"):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    if anchor == "middle":
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    elif anchor == "bottom":
        tf.vertical_anchor = MSO_ANCHOR.BOTTOM
    else:
        tf.vertical_anchor = MSO_ANCHOR.TOP
    p = tf.paragraphs[0]
    p.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
                   "right": PP_ALIGN.RIGHT}.get(align, PP_ALIGN.LEFT)
    run = p.add_run()
    run.text = text or ""
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
    return box


def _img_fit(slide, path, l, t, max_w, max_h, anchor="center", fit="smart"):
    """Place an image at (l, t) within a (max_w, max_h) box.

    fit modes:
      - "contain": preserve AR, letterbox (empty bands fill background) — original behavior.
      - "cover":   preserve AR, crop overflow so the box is fully filled.
      - "smart":   cover-crop when image AR is within 30% of box AR (cleans up
                   the awkward thin-image-with-wide-bands look), otherwise
                   letterbox. Prevents chopping huge sections off a sideways
                   or strongly-mismatched photo.
    """
    if not path or not os.path.exists(path):
        return None
    if not PIL_OK:
        return slide.shapes.add_picture(path, l, t, max_w, max_h)
    try:
        with PILImage.open(path) as im:
            iw, ih = im.size
    except Exception:
        return slide.shapes.add_picture(path, l, t, max_w, max_h)
    if iw <= 0 or ih <= 0:
        return slide.shapes.add_picture(path, l, t, max_w, max_h)

    box_ratio = max_w / max_h
    img_ratio = iw / ih

    use_cover = False
    if fit == "cover":
        use_cover = True
    elif fit == "smart":
        # Mismatch as a ratio of larger/smaller AR. 1.0 = identical, 1.3 = 30% off.
        ratio_mismatch = max(box_ratio, img_ratio) / max(min(box_ratio, img_ratio), 1e-6)
        if ratio_mismatch <= 1.30:
            use_cover = True

    if use_cover:
        # Crop the source image to the box's aspect ratio (centered), then
        # re-encode the cropped bytes and embed at exact box size. This keeps
        # the on-disk preview untouched (so review UI / re-runs still see the
        # full image) but the embedded PPTX picture is cropped, not letterboxed.
        try:
            with PILImage.open(path) as im:
                im = PILImageOps.exif_transpose(im)
                pw, ph = im.size
                if img_ratio > box_ratio:
                    # image wider than box → crop sides
                    new_pw = int(ph * box_ratio)
                    x0 = max(0, (pw - new_pw) // 2)
                    im = im.crop((x0, 0, x0 + new_pw, ph))
                elif img_ratio < box_ratio:
                    # image taller than box → crop top/bottom (centered)
                    new_ph = int(pw / box_ratio)
                    y0 = max(0, (ph - new_ph) // 2)
                    im = im.crop((0, y0, pw, y0 + new_ph))
                buf = io.BytesIO()
                im.convert("RGB").save(buf, "JPEG", quality=88)
                buf.seek(0)
            return slide.shapes.add_picture(buf, l, t, max_w, max_h)
        except Exception as ex:
            logging.warning("cover-fit failed for %s, falling back to letterbox: %s",
                            path, ex)
            # fall through to contain

    # contain (letterbox)
    if img_ratio > box_ratio:
        new_w = max_w
        new_h = int(max_w / img_ratio)
    else:
        new_h = max_h
        new_w = int(max_h * img_ratio)
    if anchor == "center":
        x = l + (max_w - new_w) // 2
        y = t + (max_h - new_h) // 2
    else:
        x, y = l, t
    return slide.shapes.add_picture(path, x, y, new_w, new_h)


HEADER_H = Inches(0.55)


def _header_bar(slide, deck, section_label):
    """Black header bar at the top of the slide with white section label, white
    date, and the HaddadBrands logo on the left. The logo file has a white
    JPG background so it appears as a clean white inset against the black bar
    — that's intentional, gives the brand mark its own pill of breathing room.
    Consistent on every slide (dark or light body)."""
    y = 0
    _box(slide, 0, y, SW, HEADER_H, P_BLACK)
    if os.path.exists(LOGO_PATH):
        try:
            _img_fit(slide, LOGO_PATH,
                     Inches(0.18), y + Inches(0.06),
                     Inches(1.4), Inches(0.43),
                     anchor="left", fit="contain")
        except Exception:
            pass
    _txt(slide, (section_label or "").upper(),
         Inches(1.75), y, Inches(8.0), HEADER_H,
         size=14, bold=True, color=P_CARD,
         font=HEADER_FONT, align="left", anchor="middle")
    _txt(slide, deck.meeting_date or "",
         Inches(9.0), y, Inches(4.15), HEADER_H,
         size=14, bold=True, color=P_CARD,
         font=HEADER_FONT, align="right", anchor="middle")


def build_title_slide(prs, deck):
    s = _blank_slide(prs)
    _bg(s, P_BLACK)
    _header_bar(s, deck, "")  # consistent header on every slide
    # Big condensed bold title — wraps to two lines if long. Centered.
    _txt(s, deck.title.upper() or "APPAREL STORE BOUGHT SAMPLE MEETING",
         Inches(0.5), Inches(2.4), Inches(12.33), Inches(2.5),
         size=72, bold=False, color=P_CARD,
         font=TITLE_FONT, align="center", anchor="middle")
    # Date below — same condensed font, smaller
    _txt(s, deck.meeting_date or "",
         Inches(0.5), Inches(4.7), Inches(12.33), Inches(0.6),
         size=24, bold=False, color=P_CARD,
         font=TITLE_FONT, align="center", anchor="middle")


def build_presentation_order_slide(prs, deck):
    s = _blank_slide(prs)
    _bg(s, P_BLACK)
    _header_bar(s, deck,"PRESENTATION ORDER")

    # Heading
    _txt(s, "PRESENTATION ORDER", Inches(0.5), Inches(0.85),
         Inches(12.33), Inches(0.9),
         size=36, bold=False, color=P_CARD,
         font=TITLE_FONT, align="center", anchor="middle")

    presenters = [p for p in deck.presenters if p.name.strip()]
    n = len(presenters)
    col_h_each = Inches(0.5)
    top = Inches(1.95)
    if n <= 10:
        for i, p in enumerate(presenters):
            y = top + col_h_each * i
            _txt(s, f"{i+1}.", Inches(3.0), y, Inches(0.7), col_h_each,
                 size=22, bold=False, color=P_CARD,
                 font=TITLE_FONT, align="left", anchor="middle")
            _txt(s, p.name.upper(), Inches(3.7), y, Inches(7.2), col_h_each,
                 size=22, bold=False, color=P_CARD,
                 font=TITLE_FONT, align="left", anchor="middle")
            sample_n = len(p.samples)
            _txt(s, f"{sample_n}", Inches(10.9), y, Inches(0.6), col_h_each,
                 size=14, bold=True, color=RGBColor(0xC8, 0xC4, 0xBC),
                 font=HEADER_FONT, align="right", anchor="middle")
    else:
        per = (n + 1) // 2
        for i, p in enumerate(presenters):
            col = 0 if i < per else 1
            row = i if col == 0 else i - per
            x_num   = Inches(0.9) if col == 0 else Inches(7.0)
            x_name  = Inches(1.7) if col == 0 else Inches(7.8)
            x_count = Inches(6.5) if col == 0 else Inches(12.6)
            y = top + col_h_each * row
            _txt(s, f"{i+1}.", x_num, y, Inches(0.7), col_h_each,
                 size=18, bold=False, color=P_CARD,
                 font=TITLE_FONT, align="left", anchor="middle")
            _txt(s, p.name.upper(), x_name, y, Inches(4.6), col_h_each,
                 size=18, bold=False, color=P_CARD,
                 font=TITLE_FONT, align="left", anchor="middle")
            sample_n = len(p.samples)
            _txt(s, f"{sample_n}", x_count - Inches(0.3), y, Inches(0.5), col_h_each,
                 size=12, bold=True, color=RGBColor(0xC8, 0xC4, 0xBC),
                 font=HEADER_FONT, align="right", anchor="middle")


def build_presenter_cover(prs, presenter, deck, sample_count):
    s = _blank_slide(prs)
    _bg(s, P_BLACK)
    _header_bar(s, deck,f"{presenter.name.upper()} — GRID")
    # Big presenter name centered
    _txt(s, presenter.name.upper(),
         Inches(0.5), Inches(2.6), Inches(12.33), Inches(1.6),
         size=72, bold=False, color=P_CARD,
         font=TITLE_FONT, align="center", anchor="middle")
    _txt(s, "GRID",
         Inches(0.5), Inches(4.2), Inches(12.33), Inches(0.6),
         size=30, bold=False, color=P_CARD,
         font=TITLE_FONT, align="center", anchor="middle")
    _txt(s, f"{sample_count} SAMPLE{'S' if sample_count != 1 else ''} · {deck.meeting_date or ''}",
         Inches(0.5), Inches(4.9), Inches(12.33), Inches(0.4),
         size=14, bold=True, color=RGBColor(0xC8, 0xC4, 0xBC),
         font=HEADER_FONT, align="center", anchor="middle")


def _draw_color_palette(slide, palette, x, y, w, h):
    """Render a horizontal row of color swatches inside the (x,y,w,h) box."""
    if not palette:
        return
    palette = palette[:8]
    n = len(palette)
    gap = Inches(0.06)
    swatch_w = (w - gap * (n - 1)) / n
    swatch_h = h
    for i, hex_str in enumerate(palette):
        cx = x + (swatch_w + gap) * i
        rgb = _hex_to_rgb_tuple(hex_str)
        _box(slide, cx, y, swatch_w, swatch_h,
             RGBColor(*rgb), line=P_LINE, line_w=Pt(0.5))


def _draw_image_card(slide, image_paths, l, t, w, h):
    """Render 1-4 product images inside a (l,t,w,h) card. The outer card has a
    light border; each image is centered/letterboxed inside its own cell.
    Layouts: 1 = single big, 2 = stacked top/bottom, 3 = stacked vertical
    thirds, 4 = 2×2 grid."""
    # outer card
    _box(slide, l, t, w, h, P_CARD, line=P_LINE, line_w=Pt(0.5))
    paths = [p for p in (image_paths or []) if p and os.path.exists(p)][:4]
    if not paths:
        return
    pad = Inches(0.12)
    inner_l = l + pad
    inner_t = t + pad
    inner_w = w - 2 * pad
    inner_h = h - 2 * pad
    n = len(paths)
    if n == 1:
        _img_fit(slide, paths[0], inner_l, inner_t, inner_w, inner_h, anchor="center")
        return
    gap = Inches(0.1)
    if n == 2:
        # stacked vertical halves
        cell_h = (inner_h - gap) / 2
        for i, p in enumerate(paths):
            y = inner_t + (cell_h + gap) * i
            _img_fit(slide, p, inner_l, y, inner_w, cell_h, anchor="center")
        return
    if n == 3:
        # one large on top, two below side-by-side
        top_h = inner_h * 0.55 - gap / 2
        bot_h = inner_h * 0.45 - gap / 2
        _img_fit(slide, paths[0], inner_l, inner_t, inner_w, top_h, anchor="center")
        cell_w = (inner_w - gap) / 2
        y_bot = inner_t + top_h + gap
        for i in range(2):
            x = inner_l + (cell_w + gap) * i
            _img_fit(slide, paths[i + 1], x, y_bot, cell_w, bot_h, anchor="center")
        return
    # n == 4: 2×2 grid
    cell_w = (inner_w - gap) / 2
    cell_h = (inner_h - gap) / 2
    for i, p in enumerate(paths):
        col = i % 2
        row = i // 2
        x = inner_l + (cell_w + gap) * col
        y = inner_t + (cell_h + gap) * row
        _img_fit(slide, p, x, y, cell_w, cell_h, anchor="center")


def build_sample_slide(prs, deck, sample, extra_images=None):
    """Render one product slide. `extra_images` is an optional list of paths
    for additional views (back, detail, alternate color) of the same garment;
    when present they're laid out alongside the primary image on the same
    slide instead of consuming separate slides."""
    s = _blank_slide(prs)
    _bg(s, P_BG)
    _header_bar(s, deck,"APPAREL STORE BOUGHT SAMPLES")

    # ── Image area: starts BELOW the header (which occupies 0–0.55"), with
    # a small breathing gap. Card stops well above the slide bottom too.
    img_left = Inches(0.4)
    img_top  = Inches(0.75)
    img_w    = Inches(7.0)
    img_h    = Inches(6.45)
    image_paths = [sample.preview_path] + list(extra_images or [])
    _draw_image_card(s, image_paths, img_left, img_top, img_w, img_h)

    # ── Right metadata column — also pushed down to clear the header
    meta_x = Inches(7.7)
    meta_w = Inches(5.3)

    # Brand title — only show if AI is fully confident the label/logo was visible.
    # Anything less and we'd risk printing the wrong brand on the deck.
    is_high_conf = (sample.confidence or "").upper() == "HIGH"
    brand = sample.brand if (is_high_conf and sample.brand) else ""
    _txt(s, brand or "—", meta_x, Inches(0.85), meta_w, Inches(0.7),
         size=30, bold=False, color=P_DARK,
         font=TITLE_FONT, align="left", anchor="middle")

    # Product label
    product = sample.product or sample.category or ""
    if product:
        _txt(s, product, meta_x, Inches(1.55), meta_w, Inches(0.45),
             size=18, bold=False, color=P_MUTED,
             font=BODY_FONT, align="left", anchor="middle")

    # Divider
    _box(s, meta_x, Inches(2.15), meta_w, Inches(0.025), P_LINE)

    # Field rows — bigger labels, bigger values
    rows = [
        ("DATE OF PRESENTING",  sample.meeting_date or deck.meeting_date or ""),
        ("SAMPLE BOUGHT BY",    sample.presenter or ""),
        ("SAMPLE BOUGHT FROM",  sample.bought_from or ""),
        ("SAMPLE PRICE",        sample.price or ""),
    ]
    row_y = Inches(2.35)
    row_h = Inches(0.7)
    for label, value in rows:
        _txt(s, label, meta_x, row_y, meta_w, Inches(0.26),
             size=10, bold=True, color=P_MUTED,
             font=HEADER_FONT, align="left")
        _txt(s, value or "—", meta_x, row_y + Inches(0.26),
             meta_w, Inches(0.42),
             size=18, bold=True, color=P_DARK,
             font=BODY_FONT, align="left")
        row_y += row_h

    # ── Color palette (only if pinned)
    palette_y = row_y + Inches(0.1)
    if sample.color_palette:
        _txt(s, "COLOR PALETTE", meta_x, palette_y, meta_w, Inches(0.25),
             size=10, bold=True, color=P_MUTED,
             font=HEADER_FONT, align="left")
        _draw_color_palette(s, sample.color_palette,
                            meta_x, palette_y + Inches(0.28),
                            meta_w, Inches(0.45))
        palette_y += Inches(0.85)

    # ── Notes/details line at bottom of right column (within slide body, above footer)
    desc_top = palette_y + Inches(0.05)
    desc_bottom = SH - Inches(0.2)
    desc_h = desc_bottom - desc_top
    if desc_h > Inches(0.3):
        details_lines = []
        if sample.colorways:
            details_lines.append(f"Colorways: {sample.colorways}")
        if sample.fabric:
            details_lines.append(f"Fabric: {sample.fabric}")
        # AI-generated details only render at HIGH confidence — when not 100%
        # visible we'd rather show nothing than risk a wrong description.
        if sample.details and is_high_conf:
            details_lines.append(sample.details)
        if sample.notes:
            details_lines.append(sample.notes)
        if details_lines:
            _txt(s, "  ·  ".join(details_lines),
                 meta_x, desc_top, meta_w, desc_h,
                 size=11, bold=False, color=P_MUTED,
                 font=BODY_FONT, align="left")


def sort_samples(samples):
    """Sort by Category → Brand (alpha) → Gender."""
    def key(r):
        return (
            category_sort_key(r.category),
            (r.brand or "").upper(),
            gender_sort_key(r.gender),
            (r.product or "").upper(),
        )
    return sorted(samples, key=key)


def resolve_merge_groups(sorted_samples):
    """Build slide groups from a sorted sample list.

    Returns a list of (primary_record, [extra_preview_paths]) tuples, one
    entry per slide. Honors `merge_into` (stable filename identity) so that
    a sample merged into garment X stays with X even if user edits to
    category/brand/gender change the sort order. Falls back to adjacency
    via merge_with_previous for legacy records that pre-date merge_into.

    Group order = order in which each primary first appears in sorted_samples.
    """
    by_name = {s.filename: s for s in sorted_samples}

    def root(s, seen=None):
        seen = seen or set()
        cur = s
        # Walk merge_into chain to the ultimate primary, with a cycle guard.
        while (cur.merge_into
               and cur.merge_into != cur.filename
               and cur.merge_into in by_name
               and cur.filename not in seen):
            seen.add(cur.filename)
            cur = by_name[cur.merge_into]
        return cur

    groups = []
    idx_by_primary = {}
    for s in sorted_samples:
        primary = root(s)
        if primary is s:
            # Primary (either standalone, or the root of a merge chain)
            if s.merge_with_previous and not s.merge_into and groups:
                # Legacy adjacency fallback: merge_with_previous=True but no
                # stable identity. Append to the most recently seen group.
                groups[-1][1].append(s.preview_path)
            else:
                if s.filename not in idx_by_primary:
                    idx_by_primary[s.filename] = len(groups)
                    groups.append([s, []])
        else:
            # Merged via merge_into (stable identity)
            if primary.filename not in idx_by_primary:
                idx_by_primary[primary.filename] = len(groups)
                groups.append([primary, []])
            groups[idx_by_primary[primary.filename]][1].append(s.preview_path)
    return groups


def build_deck(deck, output_path, on_progress=None):
    prs = Presentation()
    prs.slide_width = SW
    prs.slide_height = SH

    presenters = [p for p in deck.presenters if p.name.strip() and p.samples]

    build_title_slide(prs, deck)
    build_presentation_order_slide(prs, deck)

    total_samples = sum(len(p.samples) for p in presenters)
    done = 0

    for p in presenters:
        sorted_samples = sort_samples(p.samples)
        build_presenter_cover(prs, p, deck, len(sorted_samples))

        # Resolve groups by stable merge_into identity (with legacy adjacency
        # fallback). This means a sample merged into garment X stays with X
        # even if the user changed X's category and the sort order shifted.
        groups = resolve_merge_groups(sorted_samples)

        for primary, extras in groups:
            # Auto-populate color palette from the primary photo if needed.
            if not primary.color_palette:
                if not primary.suggested_colors:
                    primary.suggested_colors = extract_dominant_colors(
                        primary.preview_path, n=6)
                primary.color_palette = list(primary.suggested_colors[:6])
            build_sample_slide(prs, deck, primary, extra_images=extras)
            done += 1 + len(extras)
            if on_progress and total_samples:
                on_progress(done, total_samples)

    prs.save(output_path)
    return output_path


# ── Streamlit UI ─────────────────────────────────────────────────────────────
APP_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&display=swap');
html, body, [class*="css"]  { font-family: 'DM Sans', system-ui, sans-serif; }
.stApp { background: #FAFAF8; }
section[data-testid="stSidebar"] { background: #15161A; }
section[data-testid="stSidebar"] * { color: #F2F0EB !important; }
section[data-testid="stSidebar"] h1, section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 { color: #F2F0EB !important; }
.srb-hero { background: #15161A; color: #F2F0EB; padding: 28px 32px; border-radius: 18px; }
.srb-hero h1 { color: #F2F0EB; margin: 0 0 6px 0; font-weight: 700; letter-spacing: -.02em; }
.srb-hero p { color: #BDB9B0; margin: 0; }
.srb-card { background: #fff; border: 1px solid #E8E4DA; border-radius: 14px; padding: 18px 20px; }
.srb-stat { font-size: 28px; font-weight: 700; color: #1A1A1A; }
.srb-stat-label { font-size: 11px; text-transform: uppercase; letter-spacing: .1em; color: #6B6F76; }
.srb-pill { display:inline-block; padding:3px 10px; border-radius: 999px; background:#1A1A1A; color:#FAFAF8; font-size:11px; font-weight:600; letter-spacing:.05em; }
.stButton > button { border-radius: 10px; font-weight: 600; }
.stButton > button[kind="primary"] { background: #C8102E; color: white; border-color: #C8102E; }
.stButton > button[kind="primary"]:hover { background: #A30E26; border-color: #A30E26; }
</style>
"""


def init_state():
    if "step" not in st.session_state:
        st.session_state.step = "setup"
    if "deck" not in st.session_state:
        d = MeetingDeck()
        try:
            today = date.today()
            d.meeting_date = f"{today.month}/{today.day}/{str(today.year)[2:]}"
        except Exception:
            d.meeting_date = ""
        d.presenters = [PresenterConfig()]
        st.session_state.deck = d
    if "model_name" not in st.session_state:
        # Default to Sonnet: classification of flat-lay apparel needs the
        # stronger vision model — Haiku misreads mesh shorts as "leg warmers"
        # and zip hoodies as "thermal underwear" too often.
        st.session_state.model_name = "sonnet"
    else:
        # Force-upgrade existing sessions sitting on Haiku — descriptions are
        # bad enough on Haiku that we shouldn't keep someone on it accidentally.
        if st.session_state.model_name == "haiku":
            st.session_state.model_name = "sonnet"
    if "api_key" not in st.session_state:
        st.session_state.api_key = ""
    if "_seen_hashes" not in st.session_state:
        st.session_state._seen_hashes = {}
    if "auto_merge_duplicates" not in st.session_state:
        st.session_state.auto_merge_duplicates = False


def go(step):
    st.session_state.step = step
    st.rerun()


def show_nav():
    steps = [("setup", "Setup"), ("presenters", "Presenters"),
             ("analyze", "Analyze"), ("catalog", "Review"), ("build", "Build")]
    cur_i = next((i for i, (k, _) in enumerate(steps) if k == st.session_state.step), 0)
    cols = st.columns(len(steps))
    for i, (k, label) in enumerate(steps):
        if i < cur_i:
            color = "#1F8A3B"; dot = "●"
        elif i == cur_i:
            color = "#C8102E"; dot = "●"
        else:
            color = "#C7C2B8"; dot = "○"
        cols[i].markdown(
            f"<div style='text-align:center'>"
            f"<div style='color:{color}; font-size:18px;'>{dot}</div>"
            f"<div style='font-size:11px; color:#6B6F76; text-transform:uppercase; letter-spacing:.05em;'>{label}</div>"
            f"</div>", unsafe_allow_html=True)


def show_sidebar():
    with st.sidebar:
        st.markdown(f"### {APP_TITLE}")
        st.caption(f"v{APP_VERSION}")
        st.markdown("---")
        st.markdown("**Model**")
        model = st.radio(
            "Model",
            options=["haiku", "sonnet"],
            index=0 if st.session_state.model_name == "haiku" else 1,
            format_func=lambda v: "Haiku — fast" if v == "haiku" else "Sonnet — accurate",
            label_visibility="collapsed",
        )
        st.session_state.model_name = model
        st.markdown("---")
        if st.button("↻ Start over", use_container_width=True):
            for k in list(st.session_state.keys()):
                if k != "step":
                    del st.session_state[k]
            st.session_state.step = "setup"
            st.rerun()


def show_setup():
    st.markdown(
        '<div class="srb-hero">'
        '<h1>SAMPLE RECAP BUILDER</h1>'
        '<p>Build the monthly Apparel Store Bought Samples deck — '
        'upload, AI-classify, sort by Category › Brand › Gender, export PPTX.</p>'
        '</div>', unsafe_allow_html=True)
    st.write("")
    api_key_in_secrets = bool(_get_secret("ANTHROPIC_API_KEY", ""))

    if api_key_in_secrets:
        st.success("Anthropic API key loaded from secrets — you're ready to go.")
        if st.button("Continue", type="primary"):
            go("presenters")
        return

    with st.container():
        st.markdown('<div class="srb-card">', unsafe_allow_html=True)
        st.markdown("**Anthropic API key**")
        st.caption("Pasted key is held in this session only. Add to `.streamlit/secrets.toml` to skip this step.")
        key = st.text_input("API key", value=st.session_state.api_key,
                            type="password", label_visibility="collapsed",
                            placeholder="sk-ant-…")
        if st.button("Continue", type="primary", disabled=not key.strip()):
            st.session_state.api_key = key.strip()
            go("presenters")
        st.markdown("</div>", unsafe_allow_html=True)


def show_presenters():
    deck = st.session_state.deck
    st.markdown("### Meeting details")
    c1, c2 = st.columns([1, 2])
    with c1:
        date_in = st.text_input("Meeting date", value=deck.meeting_date,
                                placeholder="4/23/26",
                                help="Shown on every slide. Use the format you prefer (e.g. 4/23/26).")
        deck.meeting_date = date_in.strip()
    with c2:
        title_in = st.text_input("Deck title", value=deck.title,
                                 placeholder="APPAREL STORE BOUGHT SAMPLES MEETING")
        deck.title = title_in.strip() or "APPAREL STORE BOUGHT SAMPLES MEETING"

    st.markdown("---")
    st.markdown("### Presenters & samples")
    st.caption("Add a row per presenter (in the order they'll present). "
               "Upload all photos for that presenter — the order doesn't matter; "
               "samples are auto-sorted by Category › Brand › Gender within each presenter.")

    if not deck.presenters:
        deck.presenters = [PresenterConfig()]

    to_remove = None
    for i, p in enumerate(deck.presenters):
        st.markdown(
            f'<div class="srb-card" style="margin-top:12px;">'
            f'<div class="srb-pill">PRESENTER {i+1}</div>'
            f'</div>',
            unsafe_allow_html=True)
        c1, c2 = st.columns([2, 1])
        with c1:
            p.name = st.text_input("Presenter name", value=p.name,
                                   placeholder="e.g. Rina Fera",
                                   key=f"pname_{i}")
        with c2:
            p.bought_from_hint = st.text_input(
                "Default bought-from (optional)",
                value=p.bought_from_hint,
                placeholder="leave blank to use tag",
                help="If most of this presenter's samples are from one store, set it here. "
                     "AI still extracts the store name from each tag when visible.",
                key=f"pbf_{i}")

        # File uploader
        upl = st.file_uploader(
            f"Upload sample photos for {p.name or f'Presenter {i+1}'}",
            type=["jpg", "jpeg", "png", "heic", "heif", "webp"],
            accept_multiple_files=True, key=f"pu_{i}",
        )
        if upl:
            # Enforce hard caps before doing the expensive encode pass. The
            # caps protect against runtime OOMs / Streamlit timeouts on very
            # large uploads — we'd rather reject up front than fail 8 minutes
            # into analysis.
            current_total = sum(len(pp.samples) for pp in deck.presenters)
            presenter_room = MAX_PHOTOS_PER_PRESENTER - len(p.samples)
            total_room     = MAX_PHOTOS_TOTAL - current_total
            room = max(0, min(presenter_room, total_room))
            if presenter_room <= 0:
                st.error(
                    f"This presenter already has {len(p.samples)} photos "
                    f"(max {MAX_PHOTOS_PER_PRESENTER}). Remove some before adding more.")
            elif total_room <= 0:
                st.error(
                    f"Deck already has {current_total} photos "
                    f"(max {MAX_PHOTOS_TOTAL} total). Remove some before adding more.")
            elif len(upl) > room:
                st.warning(
                    f"Only the first {room} of {len(upl)} uploaded photos will be kept "
                    f"(per-presenter cap {MAX_PHOTOS_PER_PRESENTER}, total cap {MAX_PHOTOS_TOTAL}).")
            new_jobs = []
            tmp_dir = os.path.join(get_temp_dir(), f"presenter_{i}")
            os.makedirs(tmp_dir, exist_ok=True)
            existing_names = {s.filename for s in p.samples}
            accepted = 0
            for f in upl:
                if accepted >= room:
                    break
                h = file_hash(f)
                key = f"{i}:{h}"
                if key in st.session_state._seen_hashes:
                    continue
                st.session_state._seen_hashes[key] = True
                base, _ext = os.path.splitext(f.name)
                safe_name = re.sub(r"[^A-Za-z0-9_.-]", "_", base) + ".jpg"
                if safe_name in existing_names:
                    safe_name = f"{base}_{h[:6]}.jpg"
                dest = os.path.join(tmp_dir, safe_name)
                new_jobs.append((f.name, f.read(), dest))
                f.seek(0)
                accepted += 1
            saved = _parallel_encode_uploads(new_jobs, progress_label=f"Processing {p.name or 'presenter'} samples")
            for s in saved:
                rec = SampleRecord(original_path=s["path"],
                                   preview_path=s["path"],
                                   filename=os.path.basename(s["path"]))
                rec.presenter = p.name
                rec.meeting_date = deck.meeting_date
                if p.bought_from_hint:
                    rec.bought_from = p.bought_from_hint
                p.samples.append(rec)

        # status row
        c3, c4, c5 = st.columns([1, 1, 1])
        c3.markdown(
            f'<div class="srb-stat">{len(p.samples)}</div>'
            f'<div class="srb-stat-label">Samples uploaded</div>',
            unsafe_allow_html=True)
        if p.samples and st.session_state.get(f"_show_thumbs_{i}", False):
            grid = st.columns(8)
            for j, rec in enumerate(p.samples[:24]):
                with grid[j % 8]:
                    if os.path.exists(rec.preview_path):
                        st.image(rec.preview_path, use_container_width=True)
            if len(p.samples) > 24:
                st.caption(f"+ {len(p.samples) - 24} more")
        with c4:
            if st.button("Show/hide previews", key=f"th_{i}", use_container_width=True):
                st.session_state[f"_show_thumbs_{i}"] = not st.session_state.get(f"_show_thumbs_{i}", False)
                st.rerun()
        with c5:
            if st.button("Clear samples", key=f"cl_{i}", use_container_width=True,
                         disabled=not p.samples):
                p.samples = []
                # also clear seen hashes for this presenter so re-upload works
                st.session_state._seen_hashes = {
                    k: v for k, v in st.session_state._seen_hashes.items()
                    if not k.startswith(f"{i}:")
                }
                st.rerun()
        if len(deck.presenters) > 1:
            if st.button("Remove presenter", key=f"rm_{i}"):
                to_remove = i

    if to_remove is not None:
        deck.presenters.pop(to_remove)
        st.rerun()

    st.write("")
    cc1, cc2, cc3 = st.columns([1, 1, 2])
    with cc1:
        if st.button("➕ Add presenter", use_container_width=True):
            deck.presenters.append(PresenterConfig())
            st.rerun()
    with cc2:
        st.write("")  # spacer
    with cc3:
        ready = (deck.meeting_date.strip()
                 and any(p.name.strip() and p.samples for p in deck.presenters))
        if st.button("Continue → Analyze", type="primary",
                     use_container_width=True, disabled=not ready):
            # Push presenter name & meeting_date onto each sample (for any that were
            # added before the user finalized the name).
            for p in deck.presenters:
                for s in p.samples:
                    s.presenter = p.name
                    if not s.meeting_date:
                        s.meeting_date = deck.meeting_date
            go("analyze")


def show_analyze():
    deck = st.session_state.deck
    api_key = _get_anthropic_api_key()
    model = st.session_state.model_name
    model_id = MODEL_IDS[model]

    presenters_with_samples = [p for p in deck.presenters if p.name.strip() and p.samples]
    total_samples = sum(len(p.samples) for p in presenters_with_samples)

    st.markdown("### AI analysis")
    st.caption("Claude reads each photo, extracts brand / category / gender / store / price / "
               "fabric / details and prefills the slide fields. You can edit anything in the next step.")
    cA, cB, cC = st.columns(3)
    cA.markdown(f'<div class="srb-stat">{len(presenters_with_samples)}</div>'
                f'<div class="srb-stat-label">Presenters</div>', unsafe_allow_html=True)
    cB.markdown(f'<div class="srb-stat">{total_samples}</div>'
                f'<div class="srb-stat-label">Samples</div>', unsafe_allow_html=True)
    cC.markdown(f'<div class="srb-stat">{model.upper()}</div>'
                f'<div class="srb-stat-label">Model</div>', unsafe_allow_html=True)

    if not api_key:
        st.error("No Anthropic API key — go back to Setup and add one.")
        if st.button("← Setup"): go("setup")
        return

    if not ANTHROPIC_OK:
        st.error("anthropic package missing. `pip install -r requirements.txt` and reload.")
        return

    # Auto-merge toggle — opt-in. When on, after the main analysis we run a
    # second pass that asks Claude to detect photos showing the same garment
    # (front/back/detail) and combine them onto one slide. Adds ~$0.03 per
    # presenter on Sonnet but saves a lot of manual checkbox clicking.
    st.session_state.auto_merge_duplicates = st.checkbox(
        "🤖 Auto-detect duplicate photos (combine front/back/detail shots onto one slide)",
        value=st.session_state.auto_merge_duplicates,
        help="After classification, asks Claude to identify which photos show "
             "the same physical garment from different angles, and merges them. "
             "You can still manually adjust in the Review step.")

    st.write("")
    bL, bR = st.columns([1, 1])
    with bL:
        if st.button("← Back", use_container_width=True):
            go("presenters")
    with bR:
        run = st.button("Run analysis", type="primary", use_container_width=True)

    if run:
        client = anthropic.Anthropic(api_key=api_key, timeout=90.0)
        progress = st.progress(0.0, text="Starting…")
        log = st.empty()
        total_cost = 0.0

        # Shared state for thread-safe progress updates. Worker threads update
        # this dict; the main thread (this function body) is the only thing
        # that touches Streamlit UI elements — that keeps us out of trouble
        # with Streamlit's not-quite-thread-safe rendering.
        #
        # Key by id(presenter) instead of presenter name so two presenters
        # who share a name (or are both blank) still get separate progress
        # slots. We keep the display name alongside for the UI.
        state_lock = threading.Lock()
        presenter_progress = {id(p): [p.name, 0, 1] for p in presenters_with_samples}

        def make_progress_cb(p_obj):
            def cb(done, total_b, _name=None):
                with state_lock:
                    presenter_progress[id(p_obj)][1] = done
                    presenter_progress[id(p_obj)][2] = total_b
            return cb

        # Up to 3 presenters analyzed concurrently. Each presenter's internal
        # pools (3 Stage-2 + 6 Stage-4 workers) sit beneath this. Worst-case
        # in-flight: 3 * (3 + 6) = 27 calls, well inside Anthropic tier-4 limits.
        PRESENTER_CONCURRENCY = min(3, len(presenters_with_samples))

        def _format_eta(elapsed, frac):
            if frac < 0.05 or elapsed < 5:
                return ""
            est_total = elapsed / frac
            remaining = max(0, int(est_total - elapsed))
            if remaining < 60:
                return f" · ETA ~{remaining}s"
            return f" · ETA ~{remaining // 60}m {remaining % 60}s"

        try:
            start_time = time.time()
            with ThreadPoolExecutor(max_workers=PRESENTER_CONCURRENCY) as pool:
                futures = {
                    pool.submit(analyze_presenter_samples, client, model_id, p,
                                make_progress_cb(p)): p
                    for p in presenters_with_samples
                }
                pending = set(futures.keys())
                completed_samples = 0
                while pending:
                    # Poll: collect any completed futures, render UI, repeat.
                    done_now, pending = wait(pending, timeout=0.6,
                                             return_when=FIRST_COMPLETED)
                    for fut in done_now:
                        p = futures[fut]
                        try:
                            total_cost += fut.result() or 0.0
                            completed_samples += len(p.samples)
                        except RuntimeError:
                            raise
                        except Exception as ex:
                            logging.warning("presenter analysis error for %s: %s",
                                            p.name, ex)
                            completed_samples += len(p.samples)  # avoid stuck progress
                    # Render snapshot from main thread.
                    with state_lock:
                        snap = [(name, d, t)
                                for (name, d, t) in presenter_progress.values()]
                    elapsed = time.time() - start_time
                    frac = completed_samples / max(1, total_samples)
                    eta = _format_eta(elapsed, frac)
                    progress.progress(min(0.95, frac),
                                      text=f"{completed_samples}/{total_samples} samples analyzed{eta}")
                    in_flight_lines = [f"  {name or '(unnamed)'}: batch {d}/{t}"
                                       for name, d, t in snap if 0 < d < t]
                    if in_flight_lines:
                        log.write("Analyzing:\n" + "\n".join(in_flight_lines))

            # Optional second pass: AI duplicate detection. Also parallelized
            # across presenters now (was: serial), and across chunks within
            # each presenter (handled inside ai_suggest_merges_for_presenter).
            if st.session_state.auto_merge_duplicates:
                progress.progress(0.95, text="Detecting duplicate photos…")
                with ThreadPoolExecutor(max_workers=PRESENTER_CONCURRENCY) as pool:
                    dup_futs = {
                        pool.submit(ai_suggest_merges_for_presenter,
                                    client, model_id, p): p
                        for p in presenters_with_samples
                    }
                    dup_done = 0
                    for fut in as_completed(dup_futs):
                        p = dup_futs[fut]
                        try:
                            total_cost += fut.result() or 0.0
                        except Exception as ex:
                            logging.warning("dup-detect error for %s: %s", p.name, ex)
                        dup_done += 1
                        progress.progress(
                            min(0.99, 0.95 + 0.04 * dup_done / len(dup_futs)),
                            text=f"Duplicate detection {dup_done}/{len(dup_futs)} presenters")

            progress.progress(1.0, text=f"Done. Estimated cost ≈ ${total_cost:.3f}")
            elapsed_total = time.time() - start_time
            log.success(
                f"Analysis complete in {int(elapsed_total)}s. "
                f"${total_cost:.3f} estimated.")
            time.sleep(0.4)
            go("catalog")
        except RuntimeError as e:
            if "API_KEY_INVALID" in str(e):
                st.error("Anthropic API key is invalid. Update it in Setup.")
            else:
                st.error(str(e))
        except Exception as e:
            st.error(f"Analysis failed: {e}")


def _render_color_palette_picker(rec, key_prefix):
    """Show suggested swatches (auto-extracted) and user-pinned palette.
    Click suggested → adds to palette. Click pinned → removes."""
    if not rec.suggested_colors:
        rec.suggested_colors = extract_dominant_colors(rec.preview_path, n=8)

    st.caption("Color palette — click any color sampled from the photo to add to the slide.")

    # Pinned palette row
    if rec.color_palette:
        pinned_cols = st.columns(min(8, max(1, len(rec.color_palette))))
        for i, hex_str in enumerate(rec.color_palette):
            with pinned_cols[i % len(pinned_cols)]:
                st.markdown(
                    f'<div style="background:{hex_str};height:34px;'
                    f'border-radius:6px;border:1px solid #D8D3CB;'
                    f'display:flex;align-items:center;justify-content:center;'
                    f'color:#fff;font-size:10px;font-weight:600;'
                    f'text-shadow:0 1px 2px rgba(0,0,0,.5);">{hex_str}</div>',
                    unsafe_allow_html=True)
                if st.button("Remove", key=f"{key_prefix}_rm_{i}",
                             use_container_width=True):
                    rec.color_palette.pop(i)
                    st.rerun()
    else:
        st.caption("_no colors pinned yet_")

    # Suggested swatches row (skip ones already pinned)
    suggestions = [c for c in rec.suggested_colors if c not in rec.color_palette]
    if suggestions:
        st.caption("Suggested from photo:")
        sug_cols = st.columns(min(8, max(1, len(suggestions))))
        for i, hex_str in enumerate(suggestions):
            with sug_cols[i % len(sug_cols)]:
                st.markdown(
                    f'<div style="background:{hex_str};height:30px;'
                    f'border-radius:6px;border:1px dashed #B5B0A6;'
                    f'display:flex;align-items:center;justify-content:center;'
                    f'color:rgba(255,255,255,.85);font-size:9px;font-weight:600;'
                    f'text-shadow:0 1px 2px rgba(0,0,0,.5);">{hex_str}</div>',
                    unsafe_allow_html=True)
                if st.button("+ Add", key=f"{key_prefix}_add_{i}",
                             use_container_width=True):
                    rec.color_palette.append(hex_str)
                    st.rerun()

    # Manual color picker
    new_hex = st.color_picker("Or pick a custom color",
                              value="#1A1A1A",
                              key=f"{key_prefix}_pick",
                              label_visibility="collapsed")
    if st.button("+ Add custom color", key=f"{key_prefix}_addcustom"):
        if new_hex and new_hex.upper() not in [c.upper() for c in rec.color_palette]:
            rec.color_palette.append(new_hex.upper())
            st.rerun()


def show_catalog():
    deck = st.session_state.deck
    st.markdown("### Review & edit")
    st.caption("Edit any field. Use the rotate buttons if a photo is sideways. "
               "Click colors to pin them for the slide's palette.")

    presenters_with_samples = [p for p in deck.presenters if p.name.strip() and p.samples]
    if not presenters_with_samples:
        st.info("No samples yet.")
        if st.button("← Back to presenters"): go("presenters")
        return

    tabs = st.tabs([f"{p.name} ({len(p.samples)})" for p in presenters_with_samples])
    for p_i, (tab, p) in enumerate(zip(tabs, presenters_with_samples)):
        with tab:
            st.caption("**Tip:** check _Combine with previous slide_ on a sample if it's "
                       "a back / detail / alt-color shot of the **same garment** as the "
                       "sample above it. Up to 4 images can share one slide.")
            # Per-presenter AI auto-merge button
            mc1, mc2, mc3 = st.columns([2, 2, 1])
            with mc1:
                if st.button("🤖 AI auto-merge similar samples",
                             key=f"automerge_{p_i}",
                             use_container_width=True,
                             help="Ask Claude to detect which photos show the "
                                  "same physical garment and combine them onto "
                                  "one slide. ~$0.03 per presenter on Sonnet."):
                    api_key = _get_anthropic_api_key()
                    if not api_key:
                        st.error("No Anthropic API key — go back to Setup.")
                    else:
                        client = anthropic.Anthropic(api_key=api_key, timeout=90.0)
                        model_id = MODEL_IDS[st.session_state.model_name]
                        with st.spinner(f"Detecting duplicates in {p.name}'s samples…"):
                            try:
                                cost = ai_suggest_merges_for_presenter(
                                    client, model_id, p)
                                merged_n = sum(1 for s in p.samples
                                               if s.merge_with_previous or s.merge_into)
                                st.success(f"Detected {merged_n} duplicate "
                                            f"photo{'s' if merged_n != 1 else ''}. "
                                            f"Cost: ${cost:.3f}")
                                time.sleep(0.6)
                                st.rerun()
                            except Exception as ex:
                                st.error(f"Auto-merge failed: {ex}")
            with mc2:
                if any(s.merge_with_previous or s.merge_into for s in p.samples):
                    if st.button("↺ Reset all merges",
                                 key=f"resetmerge_{p_i}",
                                 use_container_width=True):
                        for s in p.samples:
                            s.merge_with_previous = False
                            s.merge_into = ""
                        st.rerun()
            sorted_samples = sort_samples(p.samples)

            # Pagination: at scale (the app's caps allow 150 samples per
            # presenter / 500 total) rendering every row at once stalls
            # Streamlit — every script run re-builds every image preview,
            # form, and expander. Page by PAGE_SIZE so edits stay snappy.
            # idx values still refer to position in the full sorted list so
            # the merge-into-previous logic remains correct.
            CATALOG_PAGE_SIZE = 25
            n_samples = len(sorted_samples)
            page_key = f"_catpage_{p_i}"
            if n_samples > CATALOG_PAGE_SIZE:
                pages = (n_samples + CATALOG_PAGE_SIZE - 1) // CATALOG_PAGE_SIZE
                page = max(0, min(st.session_state.get(page_key, 0), pages - 1))
                st.session_state[page_key] = page
                nav_a, nav_b, nav_c, nav_d = st.columns([1, 4, 1, 1])
                if nav_a.button("← Prev", key=f"prev_{p_i}",
                                disabled=(page == 0),
                                use_container_width=True):
                    st.session_state[page_key] = page - 1
                    st.rerun()
                lo = page * CATALOG_PAGE_SIZE + 1
                hi = min(n_samples, (page + 1) * CATALOG_PAGE_SIZE)
                nav_b.markdown(
                    f"<div style='text-align:center;padding-top:8px;color:#6B6F76;"
                    f"font-size:13px;'>Page <b>{page + 1}</b> of <b>{pages}</b> "
                    f"— samples {lo}–{hi} of {n_samples}</div>",
                    unsafe_allow_html=True)
                if nav_c.button("Next →", key=f"next_{p_i}",
                                disabled=(page >= pages - 1),
                                use_container_width=True):
                    st.session_state[page_key] = page + 1
                    st.rerun()
                # Quick jump for big decks
                with nav_d:
                    new_page = st.selectbox(
                        "Jump", options=list(range(1, pages + 1)),
                        index=page, key=f"jump_{p_i}",
                        label_visibility="collapsed") - 1
                    if new_page != page:
                        st.session_state[page_key] = new_page
                        st.rerun()
                visible_start = page * CATALOG_PAGE_SIZE
                visible_end = min(n_samples, visible_start + CATALOG_PAGE_SIZE)
            else:
                visible_start = 0
                visible_end = n_samples

            for idx in range(visible_start, visible_end):
                rec = sorted_samples[idx]
                rec_key = f"{p.name}_{idx}"
                # If this sample is merged into another sample's slide, mark
                # visually. Derived from merge_into so the indicator is
                # correct even after the sort changed (i.e. the primary may
                # no longer be the immediately-preceding row).
                is_merged = bool(rec.merge_into) or (idx > 0 and rec.merge_with_previous)
                if idx > 0 and is_merged:
                    st.markdown(
                        '<div style="border-left:3px solid #C8102E; padding:6px 12px; '
                        'margin:6px 0; background:#FAF1F2; color:#7A0E1F; font-size:11px; '
                        'font-weight:700; text-transform:uppercase; letter-spacing:.05em;">'
                        '↳ Merged into another slide</div>',
                        unsafe_allow_html=True)
                # Red banner if AI analysis failed entirely on this record —
                # without this, the user just sees "Other" / blank fields and
                # doesn't know why. Re-run analysis to retry these.
                if rec.analysis_error:
                    st.markdown(
                        f'<div style="border-left:3px solid #C8102E; padding:8px 12px; '
                        f'margin:6px 0; background:#FAF1F2; color:#7A0E1F; font-size:12px; '
                        f'font-weight:700;">'
                        f'❌ Analysis failed for this photo: {rec.analysis_error}<br>'
                        f'<span style="font-weight:400;">Click ← Re-run analysis to retry, '
                        f'or fill the fields manually below.</span></div>',
                        unsafe_allow_html=True)
                # Yellow warning bar above the row when the auto-rotator
                # couldn't confirm orientation. Lets the user spot the photos
                # that need a manual sanity-check at a glance.
                if rec.orientation_flag:
                    st.markdown(
                        '<div style="border-left:3px solid #E0A100; padding:6px 12px; '
                        'margin:6px 0; background:#FFF7E0; color:#7A5800; font-size:11px; '
                        'font-weight:700; text-transform:uppercase; letter-spacing:.05em;">'
                        '⚠ Orientation may be wrong — verify and rotate if needed</div>',
                        unsafe_allow_html=True)
                with st.container():
                    cA, cB = st.columns([1, 3])
                    with cA:
                        if os.path.exists(rec.preview_path):
                            st.image(rec.preview_path, use_container_width=True)
                        # Rotate buttons. Manual rotation = the user has
                        # personally confirmed the orientation, so clear the
                        # auto-flag.
                        rL, rR = st.columns(2)
                        if rL.button("↺", key=f"rotL_{rec_key}",
                                     help="Rotate left 90°",
                                     use_container_width=True):
                            if rotate_image_file(rec.preview_path, -90):
                                rec.suggested_colors = []
                                rec.orientation_flag = False
                                st.rerun()
                        if rR.button("↻", key=f"rotR_{rec_key}",
                                     help="Rotate right 90°",
                                     use_container_width=True):
                            if rotate_image_file(rec.preview_path, 90):
                                rec.suggested_colors = []
                                rec.orientation_flag = False
                                st.rerun()
                        # Merge toggle (not available for the very first sample)
                        if idx > 0:
                            # Display the checkbox as "on" whenever the
                            # sample is merged into anything — even if a
                            # later sort change made its primary no longer
                            # adjacent. The stable identity is merge_into.
                            checkbox_val = bool(rec.merge_with_previous or rec.merge_into)
                            new_val = st.checkbox(
                                "Combine with previous slide",
                                value=checkbox_val,
                                key=f"merge_{rec_key}",
                                help="Add this image to the previous sample's slide "
                                     "(use for back / detail / alternate-color shots).")
                            if new_val != checkbox_val:
                                rec.merge_with_previous = new_val
                                if new_val:
                                    # Snapshot the primary's identity NOW so the
                                    # merge survives later re-sorts.
                                    prev = sorted_samples[idx - 1]
                                    rec.merge_into = (
                                        prev.merge_into or prev.filename)
                                else:
                                    rec.merge_into = ""
                                st.rerun()
                    with cB:
                        # Confidence indicator + "verify" override. Brand and
                        # details only print on the slide when confidence is HIGH.
                        # If the AI rated this MEDIUM/LOW, we show a hint and a
                        # checkbox to force-print after the user sanity-checks.
                        conf = (rec.confidence or "MEDIUM").upper()
                        if conf == "HIGH":
                            st.caption("✓ HIGH confidence — brand & details will print on slide.")
                        else:
                            cf_col1, cf_col2 = st.columns([3, 2])
                            cf_col1.caption(
                                f"⚠ {conf} confidence — brand & details are HIDDEN on slide "
                                "(AI couldn't read them clearly). Edit fields below if needed.")
                            verified = cf_col2.checkbox(
                                "Verified — print anyway",
                                value=False, key=f"vf_{rec_key}",
                                help="Check after you've confirmed brand & details "
                                     "are correct. Forces them to print on the slide.")
                            if verified:
                                rec.confidence = "HIGH"
                        rA, rB, rC = st.columns(3)
                        rec.brand = rA.text_input("Brand", value=rec.brand,
                                                  key=f"br_{rec_key}").upper()
                        rec.category = rB.selectbox(
                            "Category",
                            options=CATEGORY_ORDER,
                            index=CATEGORY_ORDER.index(rec.category)
                            if rec.category in CATEGORY_ORDER
                            else CATEGORY_ORDER.index("Other"),
                            key=f"ca_{rec_key}")
                        rec.gender = rC.selectbox(
                            "Gender",
                            options=GENDER_ORDER,
                            index=GENDER_ORDER.index(rec.gender)
                            if rec.gender in GENDER_ORDER
                            else GENDER_ORDER.index("Unspecified"),
                            key=f"ge_{rec_key}")
                        rD, rE, rF = st.columns(3)
                        rec.bought_from = rD.text_input("Bought from",
                                                        value=rec.bought_from,
                                                        key=f"bf_{rec_key}")
                        rec.price = rE.text_input("Price", value=rec.price,
                                                  key=f"pr_{rec_key}",
                                                  placeholder="$X.99")
                        rec.product = rF.text_input("Product", value=rec.product,
                                                    key=f"pd_{rec_key}")
                        rG, rH = st.columns(2)
                        rec.colorways = rG.text_input("Colorways", value=rec.colorways,
                                                      key=f"cw_{rec_key}")
                        rec.fabric = rH.text_input("Fabric", value=rec.fabric,
                                                   key=f"fb_{rec_key}")
                        rec.details = st.text_input("Details / notable features",
                                                    value=rec.details,
                                                    key=f"dt_{rec_key}")
                        rec.notes = st.text_input("Free-form notes (optional)",
                                                  value=rec.notes,
                                                  key=f"nt_{rec_key}")

                        with st.expander("🎨 Color palette (click to add to slide)",
                                          expanded=bool(rec.color_palette)):
                            _render_color_palette_picker(rec, key_prefix=f"cp_{rec_key}")
                st.markdown("---")

    st.write("")
    bL, bR = st.columns([1, 1])
    with bL:
        if st.button("← Re-run analysis", use_container_width=True):
            go("analyze")
    with bR:
        if st.button("Build deck →", type="primary", use_container_width=True):
            go("build")


def show_build():
    deck = st.session_state.deck
    st.markdown("### Build deck")
    presenters_with_samples = [p for p in deck.presenters if p.name.strip() and p.samples]
    total_samples = sum(len(p.samples) for p in presenters_with_samples)

    cA, cB, cC = st.columns(3)
    cA.markdown(f'<div class="srb-stat">{len(presenters_with_samples)}</div>'
                f'<div class="srb-stat-label">Presenters</div>', unsafe_allow_html=True)
    cB.markdown(f'<div class="srb-stat">{total_samples}</div>'
                f'<div class="srb-stat-label">Samples</div>', unsafe_allow_html=True)
    # Match build_deck: title + presentation-order + per-presenter
    # (cover + one slide per resolved primary group).
    expected_slides = 2 + sum(
        1 + len(resolve_merge_groups(sort_samples(p.samples)))
        for p in presenters_with_samples)
    cC.markdown(f'<div class="srb-stat">{expected_slides}</div>'
                f'<div class="srb-stat-label">Slides</div>', unsafe_allow_html=True)

    if "_pptx_bytes" not in st.session_state:
        st.session_state._pptx_bytes = None
    if "_pptx_name" not in st.session_state:
        st.session_state._pptx_name = ""

    if st.button("Generate PPTX", type="primary", use_container_width=True):
        progress = st.progress(0.0, text="Building deck…")
        out_path = os.path.join(get_temp_dir(), f"SBS_Recap_{int(time.time())}.pptx")

        def on_p(done, total):
            progress.progress(min(1.0, done / total), text=f"Slide {done}/{total}")

        try:
            build_deck(deck, out_path, on_progress=on_p)
            with open(out_path, "rb") as fh:
                st.session_state._pptx_bytes = fh.read()
            date_clean = re.sub(r"[^0-9._-]", "", deck.meeting_date.replace("/", "."))
            st.session_state._pptx_name = f"APPAREL SBS DECK {date_clean or 'recap'}.pptx"
            progress.progress(1.0, text="Done")
            st.success(f"Built {expected_slides} slides.")
        except Exception as e:
            st.error(f"Build failed: {e}")
            logging.exception("build failed")
            st.session_state._pptx_bytes = None

    if st.session_state._pptx_bytes:
        st.download_button(
            "⬇ Download deck",
            data=st.session_state._pptx_bytes,
            file_name=st.session_state._pptx_name,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
            type="primary",
        )

    st.write("")
    bL, _ = st.columns([1, 3])
    with bL:
        if st.button("← Back to review", use_container_width=True):
            go("catalog")


def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide", page_icon="🧷")
    st.markdown(APP_CSS, unsafe_allow_html=True)
    init_state()
    show_sidebar()
    show_nav()
    st.write("")
    step = st.session_state.step
    if step == "setup":
        show_setup()
    elif step == "presenters":
        show_presenters()
    elif step == "analyze":
        show_analyze()
    elif step == "catalog":
        show_catalog()
    elif step == "build":
        show_build()
    else:
        show_setup()


if __name__ == "__main__":
    main()
