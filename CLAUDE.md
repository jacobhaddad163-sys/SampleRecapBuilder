# CLAUDE.md — Sample Recap Builder

This file provides guidance to Claude Code when working with this app.

## What it is

Sister app to **StoreRecapBuilder**. Builds the monthly *Apparel Store Bought Samples* (SBS)
meeting deck for the sample team.

Workflow inside the company: salespeople and design go store-shopping monthly, photograph
products and *also buy* samples (for inspiration — silhouette, fabrication, perceived value,
detailing). At the end of each month they present samples; afterwards the sample team
photographs every sample and enters metadata. This app replaces the manual catalog step.

## Running

```bash
pip install -r requirements.txt
streamlit run Sample_Recap_Builder_Web.py
```

Secrets in `.streamlit/secrets.toml` (not present by default — share with parent app or copy):
- `ANTHROPIC_API_KEY` — skips the setup screen.

## Architecture

Single file: **`Sample_Recap_Builder_Web.py`** (~1350 lines). No DB, no auth, no history (kept
intentionally simple — sample meetings happen monthly, not constantly).

### 5-step state machine (`st.session_state.step`)

| Step | Function | Purpose |
|------|----------|---------|
| `setup` | `show_setup()` | API key entry (skipped if in secrets) |
| `presenters` | `show_presenters()` | Add presenters, upload sample photos per presenter |
| `analyze` | `show_analyze()` | Claude reads tags → extracts brand / category / gender / store / price / fabric / details |
| `catalog` | `show_catalog()` | Review & edit AI fields per sample |
| `build` | `show_build()` | Generate + download PPTX |

`go(step)` = set step + `st.rerun()`. Same pattern as parent app.

### Data models

- **`SampleRecord`** — one sample. Fields: `presenter`, `meeting_date`, `bought_from`,
  `price`, `brand`, `category`, `gender`, `product`, `colorways`, `fabric`, `details`,
  `notes`, `confidence`.
- **`PresenterConfig`** — `name`, `samples[]`, `bought_from_hint` (optional default
  retailer for the presenter — used as fallback when AI can't read the tag).
- **`MeetingDeck`** — `meeting_date`, `title`, `presenters[]`.

### Sorting / grouping invariant

Within each presenter, samples are sorted by:
1. `CATEGORY_ORDER` index (canonical order: Tops, Bottoms, Sets, Dresses, Outerwear,
   Activewear, Sleepwear, Swimwear, Underwear, Socks, Baby, Accessories, Footwear, Other)
2. Brand (alphabetical, uppercased)
3. `GENDER_ORDER` index (Boys, Girls, Unisex Kids, Baby Boys, Baby Girls, Baby Unisex,
   Mens, Womens, Adult Unisex, Unspecified)
4. Product label (alpha, tiebreak)

`sort_samples()` is the single source of truth. `build_deck()` groups by category to insert
section dividers but keeps the brand/gender order from `sort_samples()`.

`normalize_category()` and `normalize_gender()` map free-form AI output (or user typing) to
canonical values. Add new aliases there, not in the prompt.

### Claude integration

- `_run_analysis_batch()` — batch of up to 6 photos, asks Claude for JSON with one entry
  per photo. Includes brand, category, gender, product, colorways, fabric, details,
  bought_from, price, confidence.
- `analyze_presenter_samples()` — wraps `_run_analysis_batch` with a ThreadPoolExecutor,
  2 concurrent batches per presenter.
- Prompt distinguishes **brand** (manufacturer) from **bought_from** (retailer). Critical
  because Cat & Jack ≠ Target. Don't merge them.
- `presenter.bought_from_hint` is passed into the prompt as a *fallback* — Claude still
  prefers what's on the tag. This handles the common case where an entire presenter's
  haul is from one store.
- Models: `haiku` = `claude-haiku-4-5-20251001`, `sonnet` = `claude-sonnet-4-6`.
- `parse_json_response()` accepts fenced or raw JSON.

### PPTX layout

Slide width 13.33" × 7.5" (16:9). 5 slide types:

1. **Title** — dark ink BG, deck title in cream, meeting date in muted gray. Thin red
   accent strip near footer.
2. **Presentation Order** — light BG, dark left band, numbered list of presenters with
   sample counts. Two-column layout if >9 presenters.
3. **Presenter cover** — dark ink BG, big presenter name + "GRID" subtitle + meeting date
   + sample count.
4. **Category divider** — light BG with dark left band; presenter name (small, top-left),
   category name in big type, sample count, meeting date.
5. **Sample slide** — light BG, top header bar (matches the source PDF format), big image
   on left, metadata column on right with: brand title, product label, category & gender
   chips, then four field rows (DATE OF PRESENTING / SAMPLE BOUGHT BY / SAMPLE BOUGHT
   FROM / SAMPLE PRICE), then a wrapped details line (colorways · fabric · details · notes).

Color palette is fixed and small — change by editing the `P_*` constants near the top of
the file. No theme system (parent app has 5 themes; this one doesn't need them).

Fonts: title font Century Gothic, body font Corbel. Match parent app for visual continuity.

### Photo encoding

`_parallel_encode_uploads()` re-encodes via PIL: EXIF rotation → resize to ≤1568px → JPEG
quality 85. Same logic as parent app. HEIC handled via `pillow_heif`.

`encode_image_b64()` is used at API call time and re-encodes if the on-disk image is still
larger than max_px (it shouldn't be, but the safety net is cheap).

### Defaults that get propagated

When a sample is created during upload, we set:
- `rec.presenter = p.name`
- `rec.meeting_date = deck.meeting_date`
- `rec.bought_from = p.bought_from_hint` (only if hint is set)

We re-sync these on the way into `analyze` (in case the user typed the presenter's name
*after* uploading photos). Do not rely on the upload-time snapshot alone.

### Things to NOT do

- Don't add per-store accent colors. The deck is presenter-organized, not store-organized.
- Don't add trends / takeaway / comparison slides. SBS is a flat sample catalog, not a
  competitive recap.
- Don't merge brand and bought_from. They're distinct fields and the AI prompt depends on
  it. Cat & Jack = brand, Target = bought_from.
- Don't introduce Supabase, history, or auth — keep the app stateless and simple. If the
  team wants persistence later, mirror the parent app's pattern.

## Deployment

Deploy independently from StoreRecapBuilder. Streamlit Cloud:
1. Push folder to a separate repo (or use a subdirectory deploy).
2. Set entry point to `Sample_Recap_Builder_Web.py`.
3. Add `ANTHROPIC_API_KEY` to the app's secrets.

## Reference

Source format: `c:\Users\jacobh\OneDrive - Haddad Brands\APPAREL SBS DECK 4.23.26 .pdf`.
Each sample slide has the same four fields the source PDF used; we just add image
processing, AI extraction, and automated sorting.
