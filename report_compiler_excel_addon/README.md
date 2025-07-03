# Excel Report Formatting Macro

This VBA macro automates the process of transforming and formatting data from a worksheet named `SUPPLYRecyclingreportingResult` into a new Pivot Table report on a newly created worksheet named `PivotSheet`.

## Features

- Adds a new worksheet named `PivotSheet`.
- Converts the data in `SUPPLYRecyclingreportingResult` into a structured Excel Table.
- Creates and configures a Pivot Table from the data.
- Sets Pivot Table fields, layouts, and properties.
- Adds slicers for easy filtering by:
  - Sold-to country
  - Customer exclusive
  - Detail type
  - Subtype
  - Definition
  - Battery chemistry
- Applies formatting to number columns.
- Ensures subtotal fields are disabled for cleaner reports.
- Prompts the user before execution and displays a custom form (`SaveForm`) at the end.

## How to Use

1. Open the Excel workbook containing the `SUPPLYRecyclingreportingResult` sheet.
2. Run the macro using the custom Ribbon button linked to the `ReportFormatting` subroutine.
3. Confirm the action in the pop-up dialog.
4. After processing, a new `PivotSheet` will appear with a fully formatted Pivot Table and slicers.

## Prerequisites

- A worksheet named `SUPPLYRecyclingreportingResult` must exist.
- A reference to the Ribbon control (if invoked through a custom Ribbon).
- A UserForm named `SaveForm` must be available in the VBA project.

## File Outputs

- `PivotSheet`: A new worksheet with the Pivot Table and slicers.

## Notes

- The macro uses `On Error Resume Next`, so silent failures might occur if certain assumptions are not met (e.g., missing fields).
- It assumes specific field names exist in the source sheet. Adjust field names in the macro if they differ.
- The macro applies comma-style formatting to certain columns (`F`, `G`, `I`).

## Error Handling

If an error occurs during macro execution, a message box displays the error description.

---

For best results, ensure your source sheet matches the expected structure and field names.

