---
name: new-pics
description: Process new Silver Spirits product photos from a folder. Use when Codex needs to rename top-level product pictures from label names, match each product to a Business Central item number, create background-removed copies with rembg, resize images to 1920px, and export final JPGs while preserving originals.
---

# New Pics

## Workflow

Use this skill for new product-image batches like the Marco folder workflow.

1. Work only on the folder requested by the user. Unless they explicitly ask otherwise, process only image files directly in that folder and ignore subfolders.
2. Inventory top-level image files with `find . -maxdepth 1 -type f ...`. Preserve original image content; renaming source files is allowed only as part of the requested naming workflow, and processed pixels must always go to copies.
3. Create contact sheets with filenames visible using `scripts/process_new_pics.py contact-sheets --input <folder> --output <tmp/contact-folder>`.
4. Read product labels from the contact sheets and group multiple shots of the same product in photo order.
5. Match product names to Business Central item numbers:
   - Prefer a live read-only BC lookup through `/Users/ph/Documents/Silver Spirits Project Codex/src/business_central` when network permission is available.
   - Use `/Users/ph/Library/CloudStorage/OneDrive-IZComputer/SilverSpirits/Claude/ABsUpdated/bc_items_20260416_145736.csv` as the local fallback item export.
   - When several sizes exist, use visible bottle size or GTIN from the photo if available. If still ambiguous, choose the best BC match and mention the assumption.
6. Rename originals to `Product Name-ProductID.ext` for one image, or `Product Name1-ProductID.ext`, `Product Name2-ProductID.ext`, etc. for multiple images.
   - Use the product name as matched in BC when practical.
   - Avoid filename characters illegal or awkward on macOS (`/`, `:`, control characters).
   - Before renaming, generate or inspect a mapping and check for missing sources, duplicate targets, and existing target collisions.
7. Remove backgrounds only from copies, never from the originals. Use `rembg` via the helper script.
8. Resize output so the longest side is exactly `1920px`, preserving aspect ratio.
9. Export final JPGs. Since JPG cannot store transparency, flatten the background-removed PNGs onto white.
10. Verify counts and dimensions at the end.

## Helper Script

The bundled script is at `scripts/process_new_pics.py`.

Useful commands:

```bash
python3 /Users/ph/.codex/skills/new-pics/scripts/process_new_pics.py contact-sheets \
  --input "/path/to/folder" \
  --output /private/tmp/new-pics-contacts
```

```bash
python3 /Users/ph/.codex/skills/new-pics/scripts/process_new_pics.py apply-rename-map \
  --input "/path/to/folder" \
  --map "/path/to/rename-map.csv"
```

```bash
NUMBA_DISABLE_JIT=1 TMPDIR=/private/tmp \
python3 /Users/ph/.codex/skills/new-pics/scripts/process_new_pics.py rembg-resize \
  --input "/path/to/folder" \
  --output "/path/to/folder/rembg-1920" \
  --size 1920
```

```bash
python3 /Users/ph/.codex/skills/new-pics/scripts/process_new_pics.py jpg \
  --input "/path/to/folder/rembg-1920" \
  --output "/path/to/folder/rembg-1920-jpg" \
  --quality 95
```

```bash
python3 /Users/ph/.codex/skills/new-pics/scripts/process_new_pics.py verify \
  --input "/path/to/folder/rembg-1920-jpg" \
  --size 1920 \
  --format jpg
```

## Rename Map Format

Use CSV with headers:

```csv
source,target
20260416_111204.jpg,Bulleit Bourbon 10 yo1-206890.jpg
20260416_111218.jpg,Bulleit Bourbon 10 yo2-206890.jpg
```

The script validates sources, duplicate targets, and existing target files before renaming.

## Rembg Notes

Use Daniel Gatis `rembg`. In this environment, the stable invocation is usually:

```bash
NUMBA_DISABLE_JIT=1 TMPDIR=/private/tmp python3 ...
```

If ONNX Runtime tries to compile using an unavailable provider, force CPU provider. The helper script already calls `new_session("u2net", providers=["CPUExecutionProvider"])`.

The expected intermediate output is transparent PNG because background removal needs alpha. The final JPG output is a separate white-background folder.

## Final Response

Report:

- source folder processed
- number of originals
- output folder for transparent PNGs, if created
- output folder for JPGs, if created
- verification results, including count and dimensions
- any BC match assumptions or ambiguous items
