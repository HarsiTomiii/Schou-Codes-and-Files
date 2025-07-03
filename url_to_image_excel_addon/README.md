# Excel Image Insertion Macros

This README describes two VBA macros for inserting images into Excel worksheets from URLs. Both methods support images hosted on a NetSuite instance and are optimized for consistent formatting and alignment within Excel cells.

---

## 1. `InsertImageFromURL` (Download Method)

### Overview
Downloads each image file to a temporary folder, inserts it into the Excel sheet, resizes and centers it within a cell, then deletes the image file.

### Features
- Downloads images from URLs containing `/core/media/` path segment.
- Supports full or partial NetSuite-hosted paths.
- Resizes and centers images within cell bounds (max 160px height, 180px width).
- Clears the image link after embedding the image.
- Temporarily disables screen updates and auto-calculation for speed.

### Usage
1. Select the column containing the image URLs.
2. Run the macro `InsertImageFromURL` via Ribbon or developer console.
3. Images will be downloaded and placed in their corresponding cells.

---

## 2. `URLPictureInsert` (Formula Method)

### Overview
Transforms valid URLs into Excel's `=IMAGE()` formula for embedded web rendering.

### Features
- Converts links into formulas like `=IMAGE("<URL>")`.
- Supports both full and relative paths from NetSuite.
- Automatically saves workbook in `.xlsx` format to support Excel's IMAGE formula.
- Deletes the original file afterward to avoid duplication.
- Formats selected column width and row height.
- Aligns images using cell properties.

### Usage
1. Select the column with URLs.
2. Run the macro `URLPictureInsert`.
3. Workbook will be saved as `.xlsx`, and image formulas will be placed in cells.

---

## Requirements
- Excel 365 or Excel 2019+ (for `=IMAGE()` formula support).
- Internet access to fetch the images.

## Error Handling
- Displays a message box if image retrieval or saving fails.
- Automatically restores Excel calculation and screen updating settings.

---

## Customization
- You can change `BASE_URL` and path logic if your media URLs differ.
- Adjust `MAX_HEIGHT`, `MAX_WIDTH`, or `DEFAULT_ROW_HEIGHT` to change image layout.

## Notes
- The formula-based method is faster but requires Excel support for `=IMAGE()`.
- The download-based method provides more flexibility and consistent rendering.

---

Both macros enhance Excel sheets by embedding visual content directly next to textual data, improving readability and presentation for reporting and cataloguing tasks.

