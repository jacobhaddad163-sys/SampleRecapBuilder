Brand logos for the Received Samples Deck Builder.

Drop PNG files in this folder named after the section/brand slug.
The Received Samples builder will auto-resolve a section's logo by
slugifying its name and looking for <slug>.png here.

Slug rules: lowercase, replace any non-alphanumeric run with a single
underscore, strip leading/trailing underscores.

Examples:
  "Jordan"                  -> jordan.png
  "Nike"                    -> nike.png
  "abercrombie kids"        -> abercrombie_kids.png
  "ALL BRANDS HOSIERY"      -> all_brands_hosiery.png
  "ALL BRANDS ACCESSORY"    -> all_brands_accessory.png
  "ALL BRANDS caps"         -> all_brands_caps.png

Logos render in the black header bar at the top of each slide. White
or light-colored logos on a transparent background look best (the
header bar is solid black).

If no matching PNG is found, the header falls back to rendering the
section name as bold white text. You can also override the resolved
logo per section by uploading one in the Sections step.
