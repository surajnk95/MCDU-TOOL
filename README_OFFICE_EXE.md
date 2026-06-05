# 777-9 MCDU Grid Tool - Office User Guide

This guide is for using the shared `MCDU Tool.exe` file.

The tool extracts the text from a 777-9 MCDU display photo and places it into a 13-row by 40-column grid. It can also export the final grid to a Word document.

## How To Start

1. Open the folder where `MCDU Tool.exe` is saved.
2. Double-click `MCDU Tool.exe`.
3. Wait for the tool window or browser page to open.
4. If Windows asks for permission to run the file, allow it only if you received the file from the approved office location.

## First Run

On first run, the tool automatically creates its own working folder/files.

You do not need to create anything manually.

The tool may create:

```text
data/
corrections.json
templates.json
exports/
```

These files help the tool remember corrections and save Word exports.

Do not delete these files unless you want the tool to forget learned corrections.

## Basic Workflow

1. Click `Choose Image`.
2. Select the MCDU photo.
3. The tool will try to detect the black display area automatically.
4. Check whether the grid is aligned with the display text.
5. If needed, use `Photo View` and drag the four corner handles to the exact black display corners.
6. Use `Flattened View` to check the straightened display.
7. Use the inset sliders if the text is slightly outside the grid:
   - `Left inset`
   - `Right inset`
   - `Top inset`
   - `Bottom inset`
8. Click `Analyze Grid`.
9. Check the extracted grid.
10. Manually correct any wrong or missing characters.
11. Click `Remember Corrections`.
12. Click `Export Word File` to create the Word document.

## Grid Rule

The MCDU extraction grid has 13 rows and 40 physical columns.

Only columns 1 through 38 are used for text.

The first and last physical columns are intentionally blank.

If text appears in the first or last blank column, adjust the grid and analyze again.

## Correcting The Grid

After `Analyze Grid`, review the extracted table carefully.

If a character is wrong:

1. Click inside the wrong cell.
2. Replace it with the correct character.
3. Repeat for all wrong cells.
4. Click `Remember Corrections`.

The tool uses remembered corrections to improve future results.

## Exporting To Word

After the grid looks correct:

1. Click `Export Word File`.
2. Choose where to save the `.docx` file if prompted.
3. Open the Word file and verify the table.

Exports may also be saved in the tool's `exports` folder.

## Tips For Better Accuracy

- Use a clear photo.
- Avoid glare on the screen.
- Keep the full black MCDU display visible.
- Make sure the image is not overly blurred.
- Align the four screen corners carefully.
- In `Flattened View`, make sure each character sits inside its own grid cell.
- Correct mistakes before clicking `Remember Corrections`.

## If Something Does Not Work

Try these steps:

1. Close the tool.
2. Open it again.
3. Choose the image again.
4. Re-align the corners.
5. Analyze again.

If OCR does not run or the tool shows an error, contact the person who shared the tool.

## Important Notes

- The tool processes images locally on the computer.
- The tool is intended for MCDU grid extraction support only.
- Do not delete the correction files unless you want to reset the tool's learning.
- Keep the `.exe` and its created folders together if your team wants corrections to remain available.
