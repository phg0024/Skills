---
name: pics-folder
description: Build verified T1 shipment photo folders from Business Central purchase orders by matching non-HEIF two-bottle product images by exact item ID. Use for BOZ customs shipments; do not use for creating or posting BC orders.
---

# Pics folder

Create a customs-photo folder for a user-identified T1 shipment arriving at the BOZ warehouse.

## Scope and safety

- Work only on the shipment and purchase order named by the user. Do not create, release, receive, ship, invoice, or post anything in Business Central; the BC step is read-only.
- Use `/Users/ph/Documents/Silver Spirits Project Codex` as the canonical Silver Spirits workspace, its `.env`, and `src/business_central` for live BC access. Never print credentials or raw secret/configuration values.
- Preserve all source images. Copy selected files; do not rename, convert, crop, remove backgrounds, or overwrite the originals.
- The requested destination root is:
  `/Users/ph/Library/CloudStorage/OneDrive-IZComputer/SilverSpirits/Lagermitarbeiter/T1 Shipment Pics`

## Workflow

1. Confirm that the user has identified the shipment as a T1 shipment for BOZ, and note the BC purchase-order number and supplier. If the supplier is not supplied, take the vendor name from BC and use its customary short name.
2. Read the PO from live Business Central. Filter `purchaseOrders` by the exact PO `number`, then fetch `purchaseOrderLines` for that PO. Record each distinct product line's `lineObjectNumber`, `description`, and quantity. Include tangible product lines with `lineType: Item`, but skip transport, freight, service, discount, and other charge lines even when BC labels them as `Item` (for example, a line described as `Transport cost`). Retain the BC item ID exactly as shown for each photographed product.
3. Immediately run a computer-wide local filename scan for every retained BC item ID, before narrowing by folder, selecting candidates, or declaring any item missing. Include the canonical Silver Spirits archive and other local photo folders on the computer in this first pass. On macOS, use a broad indexed filename search such as `mdfind`, with a filesystem-wide fallback when needed. Build a candidate inventory per exact item ID; do not limit the initial scan to the main Silver Spirits archive.
4. Build the folder name as `<SupplierShortName><DDMMYYYY>` using the current local Europe/Zurich date, with spaces and punctuation removed. Example: `Maxbrands24082026`. Do not add the PO number to the folder name.
5. Before writing, preflight the destination and all selected sources. If the target folder already exists and it is not clearly the same incomplete job, do not overwrite it; report the collision. If it is clearly the same job, add only missing files and never replace existing files without user direction.
6. Review candidates from the computer-wide inventory. Prefer the canonical Silver Spirits photo folders when candidates are otherwise equivalent, but inspect viable candidates from other local photo folders in the same workflow. Search and match by **filename containing the exact BC item ID**; do not match by product name alone when the item ID is available.
7. Inspect candidate images visually. Accept only a clear image showing two bottles of the same product, normally one front-facing and one back-facing for customs. Reject one-bottle images, bottle-plus-box images, mixed products, or unclear shots.
8. Always exclude HEIC/HEIF. Exclude `.heic` and `.heif` extensions and also reject files whose actual content is HEIF/HEVC even when the filename incorrectly ends in `.jpg`, `.jpeg`, or another extension. Use normal non-HEIF image formats such as JPG/JPEG, PNG, WEBP, or TIFF.
9. When several qualifying non-HEIF images exist for an item, choose the smallest file size in bytes. If file sizes tie, choose the smaller pixel area, provided the two bottles remain clearly legible.
10. Create the target folder only after preflight, then copy the chosen files with their original filenames and extensions. Do not include a manifest or other extra files unless the user requests one.
11. Verify the result:
    - every distinct BC item ID with an available qualifying image has one copied file whose filename contains that exact ID;
    - selected source and destination files have matching byte hashes;
    - the folder contains no HEIC/HEIF files by extension or actual content;
    - the folder has no unrequested extra files; and
    - each copied image still visibly shows the required two bottles.

## Missing-photo handling

If no qualifying non-HEIF two-bottle image is available, do not substitute a one-bottle image, a bottle-plus-box image, or an HEIC/HEIF image. Base the conclusion on the completed computer-wide scan, not only the canonical archive. Report the BC item ID, product description, and the reason, distinguishing between “no matching image found” and “images found but none met the two-bottle/non-HEIF rule.”

## Final response

Provide a link to the created folder and report the PO, supplier, number of copied images, and verification result. List every missing item with its BC item ID and description. State explicitly that HEIC/HEIF files were excluded and that the source images were preserved.
