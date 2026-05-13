"""Received Samples Deck Builder — second tab.

Weekly deck format (per the reference PDFs in `Wk *.pdf`):

    Title slide: black BG, "SALES SAMPLES" / "SALESMAN SAMPLES" + week range.
    Section slides: white BG, black header bar with the brand logo (or text
    label) on the left and "Season: SPRING 27" on the right. Below the bar,
    a 6x2 grid of sample photos (12 max per slide; pages auto-paginate).

Architecture notes — read me before editing
-------------------------------------------
- This module is rendered inside an `st.tabs(...)` container, so its
  render_tab() is called on every script run. Keep work behind explicit
  button clicks — do NOT do AI work or PPTX building at render time.
- Session-state keys are ALL prefixed `rs_` so SBS state can't leak in
  and our state can't leak out. The SBS tab uses unprefixed keys.
- Utilities (image encode, parallel encode, rotate, color extract) are
  imported lazily from Sample_Recap_Builder_Web inside functions, so this
  module loads cleanly during the entry script's import cycle even though
  the entry script in turn imports this one (deferred at main()).
"""

from __future__ import annotations

import os
import io
import re
import time
import logging
import threading
import base64
from concurrent.futures import ThreadPoolExecutor, as_completed, wait, FIRST_COMPLETED

import streamlit as st

# Deferred imports inside functions for utilities owned by the SBS module.

try:
    from PIL import Image as PILImage, ImageOps as PILImageOps
    PIL_OK = True
except ImportError:
    PIL_OK = False

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


# ── Constants ────────────────────────────────────────────────────────────────
APP_VERSION = "1.0.0"

ASSETS_DIR     = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
BRAND_LOGOS_DIR = os.path.join(ASSETS_DIR, "brand_logos")

# Common section/brand quick-picks for the Sections step. Order matters —
# this is how they'll appear in the dropdown. "Other / custom" lets the
# user type any value.
COMMON_SECTIONS = [
    "Jordan",
    "Nike",
    "ACG",
    "Nike SB",
    "Hurley",
    "Converse",
    "Levi's",
    "abercrombie kids",
    "Lacoste",
    "3BRAND",
    "ALL BRANDS HOSIERY",
    "ALL BRANDS ACCESSORIES",
]

# Hard caps so a runaway upload doesn't OOM Streamlit. The reference decks
# run 13-35 pages at 12 photos/page, so 500/section and 1500/deck cover
# any realistic weekly volume.
MAX_PHOTOS_PER_SECTION = 500
MAX_PHOTOS_TOTAL       = 1500

# Pricing for cost estimate. Mirrors SBS so the model selector lines up.
PRICING = {
    "haiku":  (1.00,  5.00),
    "sonnet": (3.00, 15.00),
}
MODEL_IDS = {
    "haiku":  "claude-haiku-4-5-20251001",
    "sonnet": "claude-sonnet-4-6",
}


# Slide colors / fonts mirror the SBS deck so the two builders feel like
# the same product. Title slide is pure black; section slides are white
# with a black header bar.
if PPTX_OK:
    SW            = Inches(13.33)
    SH            = Inches(7.5)
    P_BLACK       = RGBColor(0x00, 0x00, 0x00)
    P_BG          = RGBColor(0xFF, 0xFF, 0xFF)   # white body (matches refs)
    P_DARK        = RGBColor(0x1A, 0x1A, 0x1A)
    P_MUTED       = RGBColor(0x6B, 0x6F, 0x76)
    P_LINE        = RGBColor(0xD8, 0xD3, 0xCB)
    P_CARD        = RGBColor(0xFF, 0xFF, 0xFF)

TITLE_FONT  = "Impact"
HEADER_FONT = "Arial Black"
BODY_FONT   = "Arial"


# ── Helpers ──────────────────────────────────────────────────────────────────
# Canonical brand-name aliases. AI brand detection (and user typing) can
# produce a half-dozen spellings for the same brand — "Nike ACG", "ACG",
# "ACG by Nike" all refer to one bucket. We normalize to a single
# canonical spelling so they don't fragment into separate sections.
# Add new entries here when we see another spelling collision in the wild.
BRAND_ALIASES = {
    # ACG (Nike's outdoor line — gets its own section, not "Nike")
    "nike acg":                "ACG",
    "acg":                     "ACG",
    "acg by nike":             "ACG",
    "acg nike":                "ACG",
    "nike acg all conditions gear": "ACG",
    "all conditions gear":     "ACG",
    # Jordan
    "jordan":                  "Jordan",
    "air jordan":              "Jordan",
    "jumpman":                 "Jordan",
    "nike jordan":             "Jordan",
    "jordan brand":            "Jordan",
    # Nike SB
    "nike sb":                 "Nike SB",
    "sb":                      "Nike SB",
    "nike skateboarding":      "Nike SB",
    # Ralph Lauren
    "polo ralph lauren":       "Ralph Lauren",
    "polo by ralph lauren":    "Ralph Lauren",
    "ralph lauren polo":       "Ralph Lauren",
    "polo rl":                 "Ralph Lauren",
    "rl":                      "Ralph Lauren",
    "ralph lauren":            "Ralph Lauren",
    # Common kid/family brand spellings
    "levis":                   "Levi's",
    "levi's":                  "Levi's",
    "carters":                 "Carter's",
    "carter's":                "Carter's",
    "osh kosh":                "OshKosh",
    "oshkosh":                 "OshKosh",
    "oshkosh b'gosh":          "OshKosh",
    "cat and jack":            "Cat & Jack",
    "cat & jack":              "Cat & Jack",
    "abercrombie & fitch kids":"abercrombie kids",
    "abercrombie kids":        "abercrombie kids",
    "hurley":                  "Hurley",
    "converse":                "Converse",
    "lacoste":                 "Lacoste",
    "3brand":                  "3BRAND",
    "nike":                    "Nike",
}


def _normalize_brand(brand: str) -> str:
    """Map a raw brand string (AI output or user typing) to its canonical
    name. Falls back to the original (trimmed) when no alias matches —
    new brands flow through unchanged so we don't lose anything."""
    if not brand:
        return ""
    key = brand.strip().lower()
    if key in BRAND_ALIASES:
        return BRAND_ALIASES[key]
    # Also try without trailing 's, or with collapsed whitespace, as a
    # second-chance match for minor variations.
    alt = re.sub(r"\s+", " ", key).strip()
    if alt in BRAND_ALIASES:
        return BRAND_ALIASES[alt]
    return brand.strip()


def _slugify_section_name(name: str) -> str:
    """Lowercase + collapse non-alphanumeric runs to underscores. Used to
    resolve a section's brand-logo file: `assets/brand_logos/<slug>.png`."""
    s = (name or "").strip().lower()
    s = re.sub(r"[^a-z0-9]+", "_", s)
    s = s.strip("_")
    return s


def _resolve_brand_logo(section_name: str, override_path: str = "") -> str:
    """Return a path to the logo file the section should display, or "" if
    nothing is available (in which case the header falls back to text).

    Resolution order: explicit override > slug-based library lookup > none.
    Override wins because a user uploaded it deliberately."""
    if override_path and os.path.exists(override_path):
        return override_path
    slug = _slugify_section_name(section_name)
    if not slug:
        return ""
    candidate = os.path.join(BRAND_LOGOS_DIR, f"{slug}.png")
    if os.path.exists(candidate):
        return candidate
    return ""


def _price_key_for_model(model_id: str) -> str:
    mid = (model_id or "").lower()
    return "sonnet" if "sonnet" in mid else "haiku"


# ── Data models ──────────────────────────────────────────────────────────────
class ReceivedSample:
    """One photo of a received sample. The brand is either user-assigned
    (manual-mode upload) or AI-detected (bulk-mode upload)."""

    def __init__(self, original_path, preview_path, filename):
        self.original_path = original_path
        self.preview_path  = preview_path
        self.filename      = filename
        # AI- or user-assigned brand. Drives which section this sample lands
        # in. Empty string before classification (bulk mode), or set
        # immediately on upload (manual mode).
        self.brand         = ""
        # AI- or user-assigned category. Distinct from `brand` — used by
        # _bucket_samples_into_sections to route hosiery/accessory/cap
        # photos into cross-brand "ALL BRANDS X" sections so a Nike sock
        # doesn't end up with Nike apparel. Defaults to apparel when set
        # manually (user already picked the section); set by AI in bulk.
        self.category      = "apparel"
        # Brand-detection confidence so we can surface low-confidence rows
        # in the Review step. HIGH = unambiguous logo/tag; LOW = guess.
        self.brand_confidence = ""
        self.analysis_error   = ""
        # Same opt-in orientation handling as the SBS app: rotation_confidence
        # is set by the AI; we only rotate the on-disk image when the user
        # opts in via the Analyze step's auto-rotate checkbox.
        self.rotation_confidence = "HIGH"
        self.orientation_flag    = False


class ReceivedSection:
    """A group of samples that will render together. Usually a brand
    (Jordan, Nike) but can also be a cross-brand category ("ALL BRANDS
    HOSIERY") — the slug-based logo resolver works either way."""

    def __init__(self, name: str = ""):
        self.name              = name
        # Optional per-section logo override. Defaults to "" and the slide
        # builder resolves via the library; this only kicks in when the
        # user uploads one on the Sections step.
        self.logo_override     = ""
        # Per-section season string. Empty = use deck.default_season.
        self.season            = ""
        # Optional sub-label that prints under the brand on the header
        # ("outerwear", "MENS LOUNGE", "HO26 LATE ADDS"). Rarely used.
        self.sub_label         = ""
        self.samples           = []


class ReceivedDeck:
    """Top-level container. Built once per week. No persistence — the user
    rebuilds from scratch each Monday."""

    def __init__(self):
        # Deck title — "SALES SAMPLES" or "SALESMAN SAMPLES". Drives the
        # title slide. Title style otherwise identical between the two.
        self.deck_type     = "SALES SAMPLES"
        # Week range, free text. Examples from the reference decks:
        # "5.4 to 5.8", "4/27 – 5/1", "4/13 – 4/17". Whatever style the
        # team is using that month.
        self.week_label    = ""
        # Default season; sections can override.
        self.default_season = "SPRING 27"
        self.sections      = []


# ── State / utility shims (deferred so we don't fight the import cycle) ─────
def _sbs():
    """Lazy handle to the SBS module for utility reuse."""
    import Sample_Recap_Builder_Web as sbs
    return sbs


def _init_state():
    """All RS state lives under `rs_*` keys. The first time the user opens
    the tab we hydrate defaults; thereafter we leave existing state alone."""
    if "rs_step" not in st.session_state:
        st.session_state.rs_step = "setup"
    if "rs_deck" not in st.session_state:
        st.session_state.rs_deck = ReceivedDeck()
    if "rs_model_name" not in st.session_state:
        # Default to Haiku for the RS tab — brand detection from a single
        # photo's hangtag/logo is well within Haiku's capability and runs
        # ~3x cheaper than Sonnet. User can flip to Sonnet if needed.
        st.session_state.rs_model_name = "haiku"
    if "rs_upload_mode" not in st.session_state:
        st.session_state.rs_upload_mode = "manual"   # or "bulk"
    if "rs_auto_rotate" not in st.session_state:
        st.session_state.rs_auto_rotate = False
    if "rs_seen_hashes" not in st.session_state:
        st.session_state.rs_seen_hashes = {}
    if "rs_pptx_bytes" not in st.session_state:
        st.session_state.rs_pptx_bytes = None
    if "rs_pptx_name" not in st.session_state:
        st.session_state.rs_pptx_name = ""


def _go(step: str):
    st.session_state.rs_step = step
    st.rerun()


def _get_temp_dir():
    """Per-session temp dir for RS uploads. Distinct from the SBS temp dir."""
    if "rs_tmp_dir" not in st.session_state:
        import tempfile
        st.session_state.rs_tmp_dir = tempfile.mkdtemp(prefix="rs_samples_")
    return st.session_state.rs_tmp_dir


# ── AI: brand detection ──────────────────────────────────────────────────────
BRAND_DETECT_SYSTEM = (
    "You are a children's and adult apparel sample reviewer. You look at "
    "one sample photo and identify (a) the BRAND on the item (from "
    "hangtag, neck label, sewn-in label, printed logo, embroidery, "
    "packaging) and (b) the CATEGORY of the item — apparel, hosiery, "
    "accessory, cap, or footwear. You are conservative: if a field is "
    "not clearly readable, you say 'unknown' rather than guess."
)


# Canonical category set. Drives the bucketing in
# _bucket_samples_into_sections — anything not 'apparel' or 'footwear'
# routes into a cross-brand "ALL BRANDS X" section to match the
# reference deck layout where hosiery, accessory, and caps are grouped
# across brands rather than under the manufacturer.
RS_CATEGORIES = ("apparel", "hosiery", "accessory", "cap", "footwear", "unknown")


def _brand_detect_call(client, model_id, image_path, candidate_brands):
    """Single-image API call: return (brand, category, confidence, cost).

    `candidate_brands` is a list of brand names the user has already
    introduced (existing sections). We pass these as hints so the model
    prefers matching an existing section over inventing a new one — same
    photo of "Jordan" stays in the same bucket as last week's. Free-text
    output still allowed when no candidate fits.

    Category is one of RS_CATEGORIES. The caller routes hosiery /
    accessory / cap into the corresponding "ALL BRANDS X" section so a
    Nike sock and a Jordan sock both land on the hosiery slides, not
    in the brand's apparel slides — matching the reference decks.
    """
    import Sample_Recap_Builder_Web as sbs
    if not os.path.exists(image_path):
        return "", "unknown", "", 0.0
    b64 = sbs.encode_image_b64(image_path)
    candidates_clause = ""
    if candidate_brands:
        clean = ", ".join(b for b in candidate_brands if b.strip())
        candidates_clause = (
            f"\nSPELLING CONSISTENCY: the deck already has these brand "
            f"sections: [{clean}]. If you READ a brand on the photo that "
            f"matches one of those (case-insensitive), use that exact "
            f"spelling so the photo joins the existing section. This list "
            f"is for SPELLING ONLY — never assign a brand from this list "
            f"unless you can actually read it on the photo. If you can't "
            f"read a brand, the answer is 'unknown', NOT one of these."
        )
    prompt = (
        "Identify both the BRAND and the CATEGORY of the sample in this "
        "photo.\n\n"
        "BRAND: read hangtags, neck labels, sewn-in labels, screen-printed "
        "or embroidered logos, and packaging. Return the brand as you "
        "read it from the photo (title case). Use 'unknown' if no brand "
        "text or logo is clearly readable. Do not infer the brand from "
        "garment style.\n\n"
        "SUB-BRANDS: Some labels live under a parent brand but ship as "
        "their own distinct collection. Return the SUB-BRAND name only "
        "when its own mark is LITERALLY VISIBLE on the photo — never "
        "infer a sub-brand from style, color, or vibe. When in doubt, "
        "return the parent brand (or 'unknown').\n"
        "  - ACG → return 'ACG' ONLY when you can clearly see one of: "
        "the letters 'ACG', the ACG triangle logo, or the phrase 'All "
        "Conditions Gear'. Tactical/outdoor styling is NOT enough. If "
        "the only ACG-related signal is your assumption, return 'Nike' "
        "(or the actual visible brand). Do NOT return 'ACG' for "
        "non-Nike items — Ralph Lauren, Polo, Carter's etc. are NEVER "
        "ACG no matter what they look like.\n"
        "  - Jordan → return 'Jordan' only when the Jumpman silhouette "
        "or 'AIR JORDAN' wordmark is visible.\n"
        "  - Nike SB → return 'Nike SB' only when 'SB' is visible on "
        "the tag/logo.\n"
        "Return 'Nike' (no sub-brand) for plain Nike items: swoosh "
        "alone, plain NIKE wordmark, no sub-brand marker visible.\n\n"
        "BRAND IDENTIFICATION RULE — read it or skip it. If you cannot "
        "literally READ a brand name or RECOGNIZE a logo on the photo, "
        "return brand='unknown'. Never guess based on garment style, "
        "color, fabric, or because it 'looks like' a brand.\n\n"
        "CATEGORY: classify the item itself, NOT the brand. Use exactly "
        "one of these values:\n"
        "  - apparel    = tops, bottoms, dresses, sets, outerwear, "
        "activewear, sleepwear, swimwear, underwear, baby clothing\n"
        "  - hosiery    = socks (any length), tights, leg warmers\n"
        "  - accessory  = bags, backpacks, belts, headbands, gloves, "
        "scarves, jewelry, fanny packs, crossbody bags, totes — anything "
        "worn or carried that isn't apparel/footwear/hosiery/caps\n"
        "  - cap        = baseball caps, snapbacks, bucket hats, "
        "beanies, dad hats\n"
        "  - footwear   = shoes, sneakers, boots, sandals, slides\n"
        "  - unknown    = ambiguous or unreadable\n\n"
        "Reply with JSON only:\n"
        '  {"brand": "<brand>", "category": "<one of: apparel, hosiery, '
        'accessory, cap, footwear, unknown>", '
        '"confidence": "HIGH" | "MEDIUM" | "LOW"}\n\n'
        "Confidence is HIGH only when both brand and category are read "
        "from clearly visible labels/logos/shape (no guessing)."
        + candidates_clause
    )
    try:
        resp = client.messages.create(
            model=model_id, max_tokens=220,
            system=[{"type": "text", "text": BRAND_DETECT_SYSTEM,
                     "cache_control": {"type": "ephemeral"}}],
            messages=[{"role": "user", "content": [
                {"type": "image", "source": {"type": "base64",
                                              "media_type": "image/jpeg",
                                              "data": b64}},
                {"type": "text", "text": prompt},
            ]}],
            timeout=60.0,
        )
    except Exception as ex:
        logging.warning("brand detect call failed: %s", ex)
        return "", "unknown", "", 0.0
    cost = 0.0
    if hasattr(resp, "usage"):
        u = resp.usage
        p_in, p_out = PRICING[_price_key_for_model(model_id)]
        cost = (u.input_tokens * p_in + u.output_tokens * p_out) / 1_000_000
    if not resp.content:
        return "", "unknown", "", cost
    parsed = sbs.parse_json_response(resp.content[0].text)
    if not parsed:
        return "", "unknown", "", cost
    brand = (parsed.get("brand") or "").strip()
    category = (parsed.get("category") or "unknown").strip().lower()
    if category not in RS_CATEGORIES:
        category = "unknown"
    conf = (parsed.get("confidence") or "MEDIUM").strip().upper()
    if conf not in ("HIGH", "MEDIUM", "LOW"):
        conf = "MEDIUM"
    if brand.lower() == "unknown":
        brand = ""
    # Canonicalize: "Nike ACG", "ACG by Nike" → "ACG"; "Polo Ralph Lauren"
    # → "Ralph Lauren"; "Levis" → "Levi's"; etc. Done as the last step so
    # the rest of the pipeline only ever sees the canonical spelling.
    brand = _normalize_brand(brand)
    return brand, category, conf, cost


def _detect_brands_for_samples(client, model_id, samples,
                                candidate_brands, on_progress=None):
    """Run brand+category detection in parallel for a flat list of
    samples. Mutates sample.brand / sample.category / sample.brand_confidence.
    Returns total dollar cost."""
    if not samples:
        return 0.0
    workers = min(6, len(samples))
    done = 0
    total_cost = 0.0
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futs = {
            pool.submit(_brand_detect_call, client, model_id,
                        s.preview_path, candidate_brands): s
            for s in samples
        }
        for fut in as_completed(futs):
            s = futs[fut]
            try:
                brand, category, conf, cost = fut.result()
            except Exception as ex:
                logging.warning("brand detect worker failed: %s", ex)
                brand, category, conf, cost = "", "unknown", "", 0.0
            s.brand = brand
            s.category = category or "unknown"
            s.brand_confidence = conf
            total_cost += cost
            done += 1
            if on_progress:
                on_progress(done, len(samples))
    return total_cost


# Cross-brand category sections — accessories of all kinds (bags, caps,
# belts, hats, etc.) collapse into ONE "ALL BRANDS ACCESSORIES" bucket,
# matching how the team wants the deck organized. Hosiery is still its
# own bucket because the reference decks separate socks from accessories
# (sock displays vs. bag/belt/cap displays).
ALL_BRANDS_SECTIONS = {
    "hosiery":   "ALL BRANDS HOSIERY",
    "accessory": "ALL BRANDS ACCESSORIES",
    "cap":       "ALL BRANDS ACCESSORIES",
}


def _target_section_name(samp: ReceivedSample) -> str:
    """Where should this sample land? Hosiery/accessory/cap photos route
    into the cross-brand ALL-BRANDS sections. Apparel and footwear route
    to a brand-named section. Empty-brand apparel falls back to Unsorted."""
    cat = (samp.category or "").lower()
    if cat in ALL_BRANDS_SECTIONS:
        return ALL_BRANDS_SECTIONS[cat]
    brand = (samp.brand or "").strip()
    if brand:
        return brand
    return "Unsorted"


def _bucket_samples_into_sections(deck: ReceivedDeck, samples):
    """Distribute samples into deck.sections by routing rule:
       - hosiery   -> ALL BRANDS HOSIERY
       - accessory -> ALL BRANDS ACCESSORY
       - cap       -> ALL BRANDS caps
       - apparel / footwear -> brand-named section (or "Unsorted" if AI
         couldn't read the brand).
    Creates sections on demand."""
    by_lower = {s.name.lower(): s for s in deck.sections}
    for samp in samples:
        target_name = _target_section_name(samp)
        sec = by_lower.get(target_name.lower())
        if sec is None:
            sec = ReceivedSection(name=target_name)
            deck.sections.append(sec)
            by_lower[target_name.lower()] = sec
        sec.samples.append(samp)


# ── PPTX layout ──────────────────────────────────────────────────────────────
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
         color=None, font="Arial", align="left", anchor="top"):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    tf.vertical_anchor = {
        "middle": MSO_ANCHOR.MIDDLE,
        "bottom": MSO_ANCHOR.BOTTOM,
        "top":    MSO_ANCHOR.TOP,
    }.get(anchor, MSO_ANCHOR.TOP)
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


def _img_contain(slide, path, l, t, max_w, max_h):
    """Letterbox-fit an image inside a (max_w, max_h) box, centered.
    Matches the reference deck — photos sit on white with their natural
    aspect ratio preserved (no cropping, no border)."""
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
    if img_ratio > box_ratio:
        new_w = max_w
        new_h = int(max_w / img_ratio)
    else:
        new_h = max_h
        new_w = int(max_h * img_ratio)
    x = l + (max_w - new_w) // 2
    y = t + (max_h - new_h) // 2
    return slide.shapes.add_picture(path, x, y, new_w, new_h)


# Header bar measurements — match the reference Wk *.pdf decks. The bar
# is a consistent height on every slide (so they read as one deck), and
# the logo/text sits vertically centered with comfortable left padding.
# 0.65" at the 7.5" slide height is ~8.7% — same proportion as the source.
HEADER_H        = Inches(0.65)
HEADER_PAD_X    = Inches(0.25)   # left/right padding inside the bar
LOGO_BOX_W      = Inches(2.20)   # max width the logo may occupy
LOGO_BOX_H      = Inches(0.50)   # max height (leaves 0.075" top+bottom)
LOGO_TOP        = Inches(0.075)  # vertical centering inside the 0.65" bar
SEASON_BOX_W    = Inches(4.20)


def _section_header(slide, section: ReceivedSection, deck: ReceivedDeck):
    """Black bar across the top of every section slide. Left: brand logo
    if available, else section name as bold white text. Right: 'Season: X'
    bold white right-aligned. Mirrors the reference decks exactly.

    The bar height (HEADER_H) is constant across every section slide in
    the deck — that consistency is the look the team is going for. If a
    section has no logo, the text fallback uses the same vertical center
    so the header still feels uniform."""
    _box(slide, 0, 0, SW, HEADER_H, P_BLACK)

    # Logo or text on the LEFT
    logo_path = _resolve_brand_logo(section.name, section.logo_override)
    if logo_path:
        try:
            _img_contain(slide, logo_path,
                         HEADER_PAD_X, LOGO_TOP,
                         LOGO_BOX_W,   LOGO_BOX_H)
        except Exception:
            # Defensive: a corrupt asset must not break the whole deck.
            logo_path = ""
    if not logo_path:
        _txt(slide, section.name or "",
             HEADER_PAD_X, 0, Inches(8.0), HEADER_H,
             size=22, bold=True, color=P_CARD,
             font=HEADER_FONT, align="left", anchor="middle")

    # Optional sub-label between the brand block and the season — used
    # for things like "outerwear" or "HO26 LATE ADDS" in the reference.
    if section.sub_label:
        _txt(slide, section.sub_label,
             HEADER_PAD_X + LOGO_BOX_W + Inches(0.30), 0,
             Inches(6.0), HEADER_H,
             size=16, bold=True, color=P_CARD,
             font=HEADER_FONT, align="left", anchor="middle")

    # "Season: SPRING 27" on the RIGHT
    season = section.season.strip() or deck.default_season.strip()
    if season:
        season_left = SW - SEASON_BOX_W - HEADER_PAD_X
        _txt(slide, f"Season: {season}",
             season_left, 0, SEASON_BOX_W, HEADER_H,
             size=16, bold=True, color=P_CARD,
             font=HEADER_FONT, align="right", anchor="middle")


def _build_title_slide(prs, deck: ReceivedDeck):
    s = _blank_slide(prs)
    _bg(s, P_BLACK)
    _txt(s, (deck.deck_type or "SALES SAMPLES").upper(),
         Inches(0.5), Inches(2.8), Inches(12.33), Inches(1.5),
         size=72, bold=False, color=P_CARD,
         font=TITLE_FONT, align="center", anchor="middle")
    if deck.week_label:
        _txt(s, f"WEEK {deck.week_label}",
             Inches(0.5), Inches(4.3), Inches(12.33), Inches(0.6),
             size=24, bold=False, color=P_CARD,
             font=BODY_FONT, align="center", anchor="middle")


def _build_grid_slide(prs, deck: ReceivedDeck, section: ReceivedSection,
                      samples_subset):
    """One 12-up grid slide. `samples_subset` is the up-to-12 ReceivedSample
    instances destined for this page. Pagination is the caller's job."""
    s = _blank_slide(prs)
    _bg(s, P_BG)
    _section_header(s, section, deck)

    # Grid math: 6 columns x 2 rows. Use the area below the header bar with
    # symmetric padding so the grid feels framed on a white page.
    grid_top    = HEADER_H + Inches(0.35)
    grid_bottom = SH - Inches(0.4)
    grid_left   = Inches(0.4)
    grid_right  = SW - Inches(0.4)
    grid_w      = grid_right - grid_left
    grid_h      = grid_bottom - grid_top
    cols, rows  = 6, 2
    gap_x       = Inches(0.15)
    gap_y       = Inches(0.18)
    cell_w      = (grid_w - gap_x * (cols - 1)) / cols
    cell_h      = (grid_h - gap_y * (rows - 1)) / rows

    for i, samp in enumerate(samples_subset[:cols * rows]):
        col = i % cols
        row = i // cols
        cx = grid_left + (cell_w + gap_x) * col
        cy = grid_top  + (cell_h + gap_y) * row
        # Thin border, matches the reference. Image letterboxed centered.
        _box(s, cx, cy, cell_w, cell_h, P_CARD, line=P_LINE, line_w=Pt(0.5))
        if samp.preview_path and os.path.exists(samp.preview_path):
            pad = Inches(0.06)
            _img_contain(s, samp.preview_path,
                         cx + pad, cy + pad,
                         cell_w - pad * 2, cell_h - pad * 2)


def _build_deck(deck: ReceivedDeck, output_path: str, on_progress=None):
    prs = Presentation()
    prs.slide_width  = SW
    prs.slide_height = SH

    sections_with_samples = [s for s in deck.sections
                             if s.name.strip() and s.samples]

    _build_title_slide(prs, deck)

    # Page-count for progress: title + ceil(N/12) per section.
    PER_SLIDE = 12
    total_slides = 1 + sum(
        max(1, (len(sec.samples) + PER_SLIDE - 1) // PER_SLIDE)
        for sec in sections_with_samples)
    done = 1
    if on_progress:
        on_progress(done, total_slides)

    for sec in sections_with_samples:
        samples = list(sec.samples)
        # 12-up pagination. An empty section won't reach here because we
        # filtered above; a section with <12 still gets one slide.
        for start in range(0, max(1, len(samples)), PER_SLIDE):
            page = samples[start:start + PER_SLIDE]
            _build_grid_slide(prs, deck, sec, page)
            done += 1
            if on_progress:
                on_progress(done, total_slides)

    prs.save(output_path)
    return output_path


# ── UI: step indicator ───────────────────────────────────────────────────────
def _show_nav():
    steps = [("setup", "Setup"),
             ("sections", "Sections"),
             ("review", "Review"),
             ("build", "Build")]
    cur_i = next((i for i, (k, _) in enumerate(steps)
                  if k == st.session_state.rs_step), 0)
    cols = st.columns(len(steps))
    for i, (k, label) in enumerate(steps):
        if i < cur_i:
            color, dot = "#1F8A3B", "●"
        elif i == cur_i:
            color, dot = "#C8102E", "●"
        else:
            color, dot = "#C7C2B8", "○"
        cols[i].markdown(
            f"<div style='text-align:center'>"
            f"<div style='color:{color}; font-size:18px;'>{dot}</div>"
            f"<div style='font-size:11px; color:#6B6F76; "
            f"text-transform:uppercase; letter-spacing:.05em;'>{label}</div>"
            f"</div>", unsafe_allow_html=True)


# ── UI: Step 1 — Setup ───────────────────────────────────────────────────────
def _show_setup():
    deck = st.session_state.rs_deck
    st.markdown(
        '<div class="srb-hero">'
        '<h1>RECEIVED SAMPLES DECK BUILDER</h1>'
        '<p>Builds the weekly Sales / Salesman Samples deck — upload photos '
        'by brand (or bulk + AI auto-detect), generate the 12-up grid deck.</p>'
        '</div>', unsafe_allow_html=True)
    st.write("")
    st.markdown("### Deck details")
    c1, c2, c3 = st.columns([1, 1, 1])
    with c1:
        deck.deck_type = st.radio(
            "Deck title",
            options=["SALES SAMPLES", "SALESMAN SAMPLES"],
            index=0 if deck.deck_type != "SALESMAN SAMPLES" else 1,
            help="Matches the title-slide wording. Both decks use the same "
                 "layout otherwise.")
    with c2:
        deck.week_label = st.text_input(
            "Week range", value=deck.week_label,
            placeholder="5.4 to 5.8",
            help="Prints on the title slide as 'WEEK <your text>'.")
    with c3:
        deck.default_season = st.text_input(
            "Default season", value=deck.default_season,
            placeholder="SPRING 27",
            help="Prints in the right side of every section's header bar. "
                 "Sections can override individually.")
    st.write("")
    ready = bool(deck.week_label.strip())
    cL, cR = st.columns([1, 1])
    if cR.button("Continue → Sections", type="primary",
                 use_container_width=True, disabled=not ready):
        _go("sections")


# ── UI: Step 2 — Sections (manual + bulk modes) ─────────────────────────────
def _show_brand_logo_library():
    """Collapsible UI for managing the persistent brand-logo library at
    assets/brand_logos/. Logos uploaded here are written under their
    slugified name and survive across sessions, so the user only ever
    has to upload a given brand's logo once. Replaces the previous
    workflow of dropping PNGs into the folder by hand."""
    import os
    import glob
    with st.expander("📚 Brand Logo Library — upload once, reuse forever",
                     expanded=False):
        st.caption(
            "Logos saved here are matched to your sections by name. Drop a "
            "PNG and pick which brand it represents — we save it as "
            "`assets/brand_logos/<slug>.png` and every future deck that uses "
            "that brand will pick it up automatically.")

        existing = sorted(glob.glob(os.path.join(BRAND_LOGOS_DIR, "*.png")))
        if existing:
            st.caption("Currently in library:")
            cols = st.columns(min(6, max(1, len(existing))))
            for j, f in enumerate(existing):
                with cols[j % len(cols)]:
                    try:
                        st.image(f, caption=os.path.basename(f),
                                 use_container_width=True)
                    except Exception:
                        st.caption(f"_{os.path.basename(f)}_")
                    if st.button("Delete", key=f"rs_lib_del_{j}",
                                 use_container_width=True):
                        try:
                            os.remove(f)
                            st.rerun()
                        except Exception as ex:
                            st.error(f"Couldn't delete: {ex}")
        else:
            st.caption("_Library is empty._")

        st.markdown("**Upload a new logo**")
        c1, c2 = st.columns([2, 3])
        with c1:
            brand_name = st.text_input(
                "Brand name", value="",
                placeholder="e.g. Hurley",
                help="The exact brand spelling. We'll slugify it for the "
                     "filename (e.g. 'Hurley' → hurley.png).",
                key="rs_lib_brand_in")
        with c2:
            up = st.file_uploader(
                "PNG file (white logo on transparent background works best)",
                type=["png"], accept_multiple_files=False,
                key="rs_lib_up")
        if up is not None and brand_name.strip():
            slug = _slugify_section_name(brand_name)
            if not slug:
                st.error("Couldn't derive a slug from that brand name.")
            else:
                dest = os.path.join(BRAND_LOGOS_DIR, f"{slug}.png")
                try:
                    with open(dest, "wb") as fh:
                        fh.write(up.read())
                    st.success(f"Saved → {os.path.basename(dest)}")
                    st.rerun()
                except Exception as ex:
                    st.error(f"Couldn't save: {ex}")
        elif up is not None and not brand_name.strip():
            st.warning("Enter a brand name above before saving.")


def _show_sections():
    deck = st.session_state.rs_deck
    st.markdown("### Sections & photos")
    _show_brand_logo_library()
    st.caption("Each section becomes a header on its own slides. Pick a mode:")

    mode = st.radio(
        "Upload mode",
        options=["manual", "bulk"],
        format_func=lambda v: (
            "📁 Manual — I'll set the brand per section"
            if v == "manual"
            else "🤖 Bulk + AI — drop all photos, AI detects brands"
        ),
        index=0 if st.session_state.rs_upload_mode == "manual" else 1,
        horizontal=True,
        key="rs_upload_mode_radio",
    )
    st.session_state.rs_upload_mode = mode

    if mode == "manual":
        _show_sections_manual()
    else:
        _show_sections_bulk()

    st.write("")
    cL, cR = st.columns([1, 1])
    if cL.button("← Back", use_container_width=True, key="rs_secback"):
        _go("setup")
    has_anything = any(s.name.strip() and s.samples for s in deck.sections)
    if cR.button("Continue → Review", type="primary",
                 use_container_width=True, disabled=not has_anything,
                 key="rs_secnext"):
        _go("review")


def _show_sections_manual():
    """Manual mode: user creates sections (brand or label) and uploads
    photos directly into each. No AI involved."""
    deck = st.session_state.rs_deck
    if not deck.sections:
        # Seed with one empty section so the user has something to populate
        # without clicking "Add" first.
        deck.sections.append(ReceivedSection())

    to_remove = None
    for i, sec in enumerate(deck.sections):
        st.markdown(
            f'<div class="srb-card" style="margin-top:12px;">'
            f'<div class="srb-pill">SECTION {i+1}</div></div>',
            unsafe_allow_html=True)
        cA, cB, cC = st.columns([2, 1, 1])
        with cA:
            # The quick-pick dropdown + free-text combo: user can pick a
            # common section in one click, or type a custom name.
            opts = ["(custom)"] + COMMON_SECTIONS
            cur_idx = (opts.index(sec.name) if sec.name in COMMON_SECTIONS
                       else 0)
            picked = st.selectbox(
                "Quick-pick brand / category", options=opts, index=cur_idx,
                key=f"rs_pick_{i}",
                help="Pick a common brand or section, or choose '(custom)' "
                     "and type below.")
            if picked != "(custom)":
                sec.name = picked
            else:
                sec.name = st.text_input(
                    "Section name (custom)", value=sec.name,
                    placeholder="e.g. Hurley",
                    key=f"rs_name_{i}")
        with cB:
            sec.season = st.text_input(
                "Season override", value=sec.season,
                placeholder=f"defaults to '{deck.default_season}'",
                key=f"rs_season_{i}",
                help="Leave blank to use the deck-level default. Set when "
                     "this section is a different season (HO26 LATE ADDS, "
                     "moved to FALL, etc.)")
        with cC:
            sec.sub_label = st.text_input(
                "Sub-label (optional)", value=sec.sub_label,
                placeholder="e.g. outerwear",
                key=f"rs_sub_{i}",
                help="Small label next to the brand in the header. Examples "
                     "from prior decks: 'outerwear', 'MENS LOUNGE', "
                     "'Golf capsule (dropped)', 'HO26 LATE ADDS'.")

        # Logo handling: show what the library auto-resolved (if anything)
        # and offer an uploader override. We keep override paths in the
        # session tmp dir so they survive reruns within the session.
        resolved = _resolve_brand_logo(sec.name, sec.logo_override)
        lcol1, lcol2 = st.columns([3, 2])
        with lcol1:
            if resolved:
                src = (
                    "library" if resolved.endswith(
                        f"{_slugify_section_name(sec.name)}.png")
                        and resolved.startswith(BRAND_LOGOS_DIR)
                    else "uploaded"
                )
                st.caption(
                    f"Logo: **{os.path.basename(resolved)}** "
                    f"({src}) — will print on this section's slides.")
            else:
                st.caption(
                    "Logo: _not found_ — header will print the section name "
                    "in white text. Drop a PNG named "
                    f"**{_slugify_section_name(sec.name) or 'name'}.png** "
                    f"into `assets/brand_logos/` to auto-resolve next time.")
        with lcol2:
            up = st.file_uploader(
                "Override logo (PNG)", type=["png"],
                accept_multiple_files=False, key=f"rs_logo_{i}",
                label_visibility="collapsed")
            if up is not None:
                dest = os.path.join(_get_temp_dir(),
                                    f"logo_section_{i}_{up.name}")
                with open(dest, "wb") as fh:
                    fh.write(up.read())
                sec.logo_override = dest
                st.success(f"Saved override → {os.path.basename(dest)}")

        # Photos
        sec_label = sec.name or f"Section {i+1}"
        upl = st.file_uploader(
            f"Upload sample photos for {sec_label}",
            type=["jpg", "jpeg", "png", "heic", "heif", "webp"],
            accept_multiple_files=True, key=f"rs_pu_{i}")
        if upl:
            _ingest_uploads(sec_label, sec, i, upl)

        # Status row
        sA, sB, sC = st.columns([1, 1, 1])
        sA.markdown(
            f'<div class="srb-stat">{len(sec.samples)}</div>'
            f'<div class="srb-stat-label">Photos uploaded</div>',
            unsafe_allow_html=True)
        with sB:
            if st.button("Show/hide previews", key=f"rs_th_{i}",
                         use_container_width=True):
                k = f"rs_show_thumbs_{i}"
                st.session_state[k] = not st.session_state.get(k, False)
                st.rerun()
        with sC:
            if st.button("Clear photos", key=f"rs_cl_{i}",
                         use_container_width=True,
                         disabled=not sec.samples):
                sec.samples = []
                st.session_state.rs_seen_hashes = {
                    k: v for k, v in st.session_state.rs_seen_hashes.items()
                    if not k.startswith(f"sec{i}:")
                }
                st.rerun()
        if sec.samples and st.session_state.get(f"rs_show_thumbs_{i}", False):
            grid = st.columns(8)
            for j, samp in enumerate(sec.samples[:24]):
                with grid[j % 8]:
                    if os.path.exists(samp.preview_path):
                        st.image(samp.preview_path, use_container_width=True)
            if len(sec.samples) > 24:
                st.caption(f"+ {len(sec.samples) - 24} more")

        if len(deck.sections) > 1:
            if st.button("Remove section", key=f"rs_rm_{i}"):
                to_remove = i

    if to_remove is not None:
        deck.sections.pop(to_remove)
        st.rerun()

    if st.button("➕ Add section", key="rs_addsec"):
        deck.sections.append(ReceivedSection())
        st.rerun()


def _show_sections_bulk():
    """Bulk mode: user uploads ALL photos into one bucket. We hold them in
    a pending list. When the user clicks Detect, we run brand detection
    in parallel, then move them into auto-created sections."""
    deck = st.session_state.rs_deck

    if "rs_bulk_pending" not in st.session_state:
        # Pending bucket lives outside deck.sections so a re-run of detection
        # doesn't move photos around behind the user's back.
        st.session_state.rs_bulk_pending = []  # list[ReceivedSample]

    pending = st.session_state.rs_bulk_pending

    st.caption(
        "Drop every photo for this week into one bucket. Click **Detect "
        "brands** and Claude will read each photo's hangtag/logo and bucket "
        "the samples into sections for you. Unreadable photos go into an "
        "**Unsorted** section you can fix in Review.")

    upl = st.file_uploader(
        "Upload all sample photos",
        type=["jpg", "jpeg", "png", "heic", "heif", "webp"],
        accept_multiple_files=True, key="rs_bulk_pu")
    if upl:
        _ingest_uploads("bulk", None, "bulk", upl, pending=pending)

    if pending or any(s.samples for s in deck.sections):
        sA, sB, sC = st.columns([1, 1, 1])
        sA.markdown(
            f'<div class="srb-stat">{len(pending)}</div>'
            f'<div class="srb-stat-label">Pending detection</div>',
            unsafe_allow_html=True)
        total_sectioned = sum(len(s.samples) for s in deck.sections)
        sB.markdown(
            f'<div class="srb-stat">{total_sectioned}</div>'
            f'<div class="srb-stat-label">Already sectioned</div>',
            unsafe_allow_html=True)
        sC.markdown(
            f'<div class="srb-stat">{len(deck.sections)}</div>'
            f'<div class="srb-stat-label">Sections so far</div>',
            unsafe_allow_html=True)

    if not pending:
        st.info("Upload photos above, then click Detect.")
        return

    # Optional: let the user constrain detection to a known brand list,
    # otherwise we use the deck's existing section names + COMMON_SECTIONS
    # as candidates so detection prefers consistent spellings.
    existing_names = [s.name for s in deck.sections if s.name.strip()]
    candidates = list({*existing_names, *COMMON_SECTIONS})

    api_key = _sbs()._get_anthropic_api_key()
    model = st.session_state.rs_model_name
    model_id = MODEL_IDS[model]
    mcA, mcB = st.columns([1, 1])
    with mcA:
        new_model = st.selectbox(
            "Brand-detection model",
            options=["haiku", "sonnet"],
            index=0 if model == "haiku" else 1,
            format_func=lambda v: ("Haiku — fast/cheap (~$0.005/photo)"
                                    if v == "haiku"
                                    else "Sonnet — accurate (~$0.015/photo)"),
            key="rs_model_pick",
        )
        st.session_state.rs_model_name = new_model
        model = new_model
        model_id = MODEL_IDS[model]
    with mcB:
        st.session_state.rs_auto_rotate = st.checkbox(
            "🔄 Also auto-rotate (experimental — can be wrong)",
            value=st.session_state.rs_auto_rotate,
            help="Same opt-in auto-rotation as the SBS tab. Off by default.")

    if not api_key or not ANTHROPIC_OK:
        st.error("No Anthropic API key. Add it in the Sample Recap Builder "
                 "tab's Setup step (it's shared across both tabs).")
        return

    if st.button("🤖 Detect brands", type="primary",
                 use_container_width=True, key="rs_detect_btn"):
        _run_brand_detection(pending, candidates, model_id)


def _run_brand_detection(pending, candidates, model_id):
    """Click handler: runs the parallel detection pass, then moves the
    pending samples into deck.sections by detected brand."""
    deck = st.session_state.rs_deck
    client = anthropic.Anthropic(
        api_key=_sbs()._get_anthropic_api_key(), timeout=90.0)
    progress = st.progress(0.0, text="Detecting brands…")
    state_lock = threading.Lock()
    done_shared = [0]

    def on_p(done, total):
        with state_lock:
            done_shared[0] = done
        try:
            progress.progress(min(0.99, done / max(1, total)),
                              text=f"Detecting brands {done}/{total}")
        except Exception:
            pass

    try:
        cost = _detect_brands_for_samples(
            client, model_id, pending, candidates, on_progress=on_p)
    except Exception as ex:
        st.error(f"Detection failed: {ex}")
        return

    if st.session_state.rs_auto_rotate:
        # Lightweight orientation pass using the SBS pipeline. Re-uses the
        # same opt-in mechanism so behavior matches across both tabs.
        try:
            sbs = _sbs()
            sbs.auto_verify_orientation(client, model_id, pending)
        except Exception as ex:
            logging.warning("RS auto-rotate failed: %s", ex)

    _bucket_samples_into_sections(deck, pending)
    progress.progress(1.0, text=f"Done. Estimated cost ≈ ${cost:.3f}")
    st.session_state.rs_bulk_pending = []
    time.sleep(0.3)
    st.rerun()


# ── Upload ingestion (shared between manual and bulk modes) ─────────────────
def _ingest_uploads(label, sec, sec_idx, upl, pending=None):
    """Re-encode uploaded files (HEIC→JPEG, resize) and either append to
    `sec.samples` or to the bulk-pending list, depending on which mode is
    active. Enforces per-section / total caps."""
    sbs = _sbs()
    deck = st.session_state.rs_deck
    current_total = (
        sum(len(s.samples) for s in deck.sections)
        + len(st.session_state.get("rs_bulk_pending", []))
    )
    if pending is not None:
        # bulk mode
        section_room = MAX_PHOTOS_PER_SECTION - len(pending)
    else:
        section_room = MAX_PHOTOS_PER_SECTION - len(sec.samples)
    total_room = MAX_PHOTOS_TOTAL - current_total
    room = max(0, min(section_room, total_room))
    if section_room <= 0:
        st.error(f"This section already has the max "
                 f"{MAX_PHOTOS_PER_SECTION} photos.")
        return
    if total_room <= 0:
        st.error(f"Deck already has {current_total} photos "
                 f"(max {MAX_PHOTOS_TOTAL}).")
        return
    if len(upl) > room:
        st.warning(f"Only the first {room} of {len(upl)} uploaded photos "
                   f"will be kept (per-section cap "
                   f"{MAX_PHOTOS_PER_SECTION}, total cap "
                   f"{MAX_PHOTOS_TOTAL}).")

    tmp_dir = os.path.join(_get_temp_dir(), f"section_{sec_idx}")
    os.makedirs(tmp_dir, exist_ok=True)
    target_list = pending if pending is not None else sec.samples
    existing_names = {s.filename for s in target_list}
    jobs = []
    accepted = 0
    for f in upl:
        if accepted >= room:
            break
        h = sbs.file_hash(f)
        key = f"sec{sec_idx}:{h}"
        if key in st.session_state.rs_seen_hashes:
            continue
        st.session_state.rs_seen_hashes[key] = True
        base, _ext = os.path.splitext(f.name)
        safe_name = re.sub(r"[^A-Za-z0-9_.-]", "_", base) + ".jpg"
        if safe_name in existing_names:
            safe_name = f"{base}_{h[:6]}.jpg"
        dest = os.path.join(tmp_dir, safe_name)
        jobs.append((f.name, f.read(), dest))
        f.seek(0)
        accepted += 1
    saved = sbs._parallel_encode_uploads(
        jobs, progress_label=f"Processing {label} photos")
    for s in saved:
        rec = ReceivedSample(original_path=s["path"],
                             preview_path=s["path"],
                             filename=os.path.basename(s["path"]))
        if pending is None and sec is not None:
            # Manual mode: brand is the section name (user already told us).
            # Set category from the section name so the data stays
            # internally consistent — if the user later edits via Review
            # and changes the section, the routing-by-category still works.
            rec.brand = sec.name
            rec.brand_confidence = "HIGH"
            sname = (sec.name or "").lower()
            if "hosiery" in sname:
                rec.category = "hosiery"
            elif "accessor" in sname or "cap" in sname:
                # Caps now ride with accessories per the new spec — one
                # "ALL BRANDS ACCESSORIES" bucket holds bags + caps + belts.
                rec.category = "accessory"
            else:
                rec.category = "apparel"
        target_list.append(rec)


# ── UI: Step 3 — Review ─────────────────────────────────────────────────────
def _show_review():
    deck = st.session_state.rs_deck
    st.markdown("### Review")
    st.caption("Per-section thumbnails. Move a photo to a different section "
               "if a detection landed wrong. Rotate or remove as needed.")

    sections_with_samples = [s for s in deck.sections
                             if s.name.strip() and s.samples]
    if not sections_with_samples:
        st.info("No sections with photos. Go back and upload.")
        if st.button("← Back to sections"):
            _go("sections")
        return

    section_names = [s.name for s in sections_with_samples]
    tabs = st.tabs([f"{s.name} ({len(s.samples)})"
                    for s in sections_with_samples])
    sbs = _sbs()
    for t_i, (tab, sec) in enumerate(zip(tabs, sections_with_samples)):
        with tab:
            PAGE = 24
            n = len(sec.samples)
            page_key = f"rs_revpage_{t_i}"
            if n > PAGE:
                pages = (n + PAGE - 1) // PAGE
                page = max(0, min(st.session_state.get(page_key, 0),
                                    pages - 1))
                st.session_state[page_key] = page
                nav_a, nav_b, nav_c = st.columns([1, 6, 1])
                if nav_a.button("← Prev", key=f"rs_prev_{t_i}",
                                disabled=(page == 0),
                                use_container_width=True):
                    st.session_state[page_key] = page - 1
                    st.rerun()
                lo = page * PAGE + 1
                hi = min(n, (page + 1) * PAGE)
                nav_b.markdown(
                    f"<div style='text-align:center;padding-top:8px;"
                    f"color:#6B6F76;font-size:13px;'>"
                    f"Page <b>{page+1}</b> of <b>{pages}</b> "
                    f"— samples {lo}–{hi} of {n}</div>",
                    unsafe_allow_html=True)
                if nav_c.button("Next →", key=f"rs_next_{t_i}",
                                disabled=(page >= pages - 1),
                                use_container_width=True):
                    st.session_state[page_key] = page + 1
                    st.rerun()
                visible = range(page * PAGE, hi)
            else:
                visible = range(n)

            cols = st.columns(6)
            for j in visible:
                samp = sec.samples[j]
                col = cols[j % 6]
                with col:
                    if os.path.exists(samp.preview_path):
                        st.image(samp.preview_path, use_container_width=True)
                    bc1, bc2 = st.columns(2)
                    if bc1.button("↺", key=f"rs_rotL_{t_i}_{j}",
                                  help="Rotate left 90°",
                                  use_container_width=True):
                        sbs.rotate_image_file(samp.preview_path, -90)
                        samp.orientation_flag = False
                        st.rerun()
                    if bc2.button("↻", key=f"rs_rotR_{t_i}_{j}",
                                  help="Rotate right 90°",
                                  use_container_width=True):
                        sbs.rotate_image_file(samp.preview_path, 90)
                        samp.orientation_flag = False
                        st.rerun()
                    # Move-to-section dropdown (excludes current section)
                    other_names = [n for n in section_names if n != sec.name]
                    move_opts = ["(stay)"] + other_names
                    pick = st.selectbox(
                        "Move", options=move_opts, index=0,
                        key=f"rs_move_{t_i}_{j}",
                        label_visibility="collapsed")
                    if pick != "(stay)":
                        target = next(s for s in sections_with_samples
                                      if s.name == pick)
                        target.samples.append(samp)
                        sec.samples.remove(samp)
                        st.rerun()
                    if st.button("✕ Remove", key=f"rs_del_{t_i}_{j}",
                                 use_container_width=True):
                        sec.samples.remove(samp)
                        st.rerun()
                    # Category + low-confidence indicator. Showing the
                    # AI-detected category lets the user spot a wrongly
                    # routed sample at a glance (e.g. a sock that landed
                    # in "Nike" instead of "ALL BRANDS HOSIERY").
                    if samp.category and samp.category != "unknown":
                        cat_label = samp.category.upper()
                    else:
                        cat_label = "—"
                    if samp.brand_confidence == "LOW":
                        st.caption(f"{cat_label} · ⚠ low confidence")
                    else:
                        st.caption(cat_label)

    st.write("")
    bL, bR = st.columns([1, 1])
    if bL.button("← Back", use_container_width=True, key="rs_revback"):
        _go("sections")
    if bR.button("Build deck →", type="primary",
                 use_container_width=True, key="rs_revnext"):
        _go("build")


# ── UI: Step 4 — Build ──────────────────────────────────────────────────────
def _show_build():
    deck = st.session_state.rs_deck
    st.markdown("### Build")
    sections_with_samples = [s for s in deck.sections
                             if s.name.strip() and s.samples]
    total_samples = sum(len(s.samples) for s in sections_with_samples)
    PER_SLIDE = 12
    expected_slides = 1 + sum(
        max(1, (len(s.samples) + PER_SLIDE - 1) // PER_SLIDE)
        for s in sections_with_samples)

    cA, cB, cC = st.columns(3)
    cA.markdown(f'<div class="srb-stat">{len(sections_with_samples)}</div>'
                f'<div class="srb-stat-label">Sections</div>',
                unsafe_allow_html=True)
    cB.markdown(f'<div class="srb-stat">{total_samples}</div>'
                f'<div class="srb-stat-label">Photos</div>',
                unsafe_allow_html=True)
    cC.markdown(f'<div class="srb-stat">{expected_slides}</div>'
                f'<div class="srb-stat-label">Slides</div>',
                unsafe_allow_html=True)

    if st.button("Generate PPTX", type="primary", use_container_width=True,
                 key="rs_buildbtn"):
        progress = st.progress(0.0, text="Building deck…")
        out_path = os.path.join(_get_temp_dir(),
                                 f"SalesSamples_{int(time.time())}.pptx")

        def on_p(done, total):
            try:
                progress.progress(min(1.0, done / total),
                                  text=f"Slide {done}/{total}")
            except Exception:
                pass

        try:
            _build_deck(deck, out_path, on_progress=on_p)
            with open(out_path, "rb") as fh:
                st.session_state.rs_pptx_bytes = fh.read()
            week_clean = re.sub(r"[^0-9._-]", "",
                                deck.week_label.replace("/", "."))
            st.session_state.rs_pptx_name = (
                f"{deck.deck_type.title()} WK {week_clean or 'recap'}.pptx")
            progress.progress(1.0, text="Done")
            st.success(f"Built {expected_slides} slides.")
        except Exception as e:
            st.error(f"Build failed: {e}")
            logging.exception("RS build failed")
            st.session_state.rs_pptx_bytes = None

    if st.session_state.rs_pptx_bytes:
        st.download_button(
            "⬇ Download deck",
            data=st.session_state.rs_pptx_bytes,
            file_name=st.session_state.rs_pptx_name,
            mime="application/vnd.openxmlformats-officedocument."
                 "presentationml.presentation",
            use_container_width=True,
            type="primary",
            key="rs_dl")

    st.write("")
    if st.button("← Back to review", use_container_width=True,
                 key="rs_buildback"):
        _go("review")


# ── Entry point (called from Sample_Recap_Builder_Web.main) ─────────────────
def render_tab():
    """Render the Received Samples Deck Builder tab body. Safe to call on
    every script run — work only happens inside button-click handlers."""
    _init_state()
    _show_nav()
    st.write("")
    step = st.session_state.rs_step
    if step == "setup":
        _show_setup()
    elif step == "sections":
        _show_sections()
    elif step == "review":
        _show_review()
    elif step == "build":
        _show_build()
    else:
        _show_setup()
