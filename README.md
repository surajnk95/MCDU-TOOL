# 777-9 MCDU Grid Tool

The 777-9 MCDU Grid Tool extracts the text shown on a Boeing 777-9 MCDU display from a phone photo and places it into a 13-row by 40-column reference grid.

The tool is designed for tilted or angled phone photos. It lets the user detect the black display, flatten the screen, resize the grid, correct OCR results, remember corrections, and export the final grid to a Word document.

## What The Tool Does

- Loads a phone photo of the MCDU.
- Detects the black MCDU display area.
- Lets the user drag the four screen corners if detection needs fine tuning.
- Flattens angled images into a rectangular screen view.
- Overlays a 13 x 40 grid.
- Keeps physical column 0 and column 39 blank.
- Labels only the usable columns as 1 through 38.
- Labels rows 1 through 13.
- Runs OCR locally using Tesseract.
- Places detected characters into the editable grid.
- Lets the user correct any wrong cells.
- Remembers corrected row patterns for future images.
- Exports the final grid to a Word `.docx` file.

## Using The Shared EXE Version

When this tool is packaged as an `.exe`, teammates should receive a file such as:

```text
MCDU Tool.exe
```

To use it:

1. Double-click `MCDU Tool.exe`.
2. Wait for the browser window to open automatically.
3. Click `Choose Image`.
4. Select the MCDU photo.
5. Confirm the detected display area.
6. Use `Photo View` if you need to drag the four screen corners.
7. Use `Flattened View` to check the straightened screen and grid.
8. Adjust `Left inset`, `Right inset`, `Top inset`, or `Bottom inset` if the text does not sit inside the grid cells.
9. Click `Analyze Grid`.
10. Correct any wrong cells in the extracted table.
11. Click `Remember Corrections` so future images improve.
12. Click `Export Word File` to create the `.docx` output.

On first run, the tool automatically creates its working files and folders. The user does not need to create them manually.

```text
data/
├── corrections.json
├── templates.json
└── exports/
```

If the user starts from a fresh copy, `corrections.json` and `templates.json` will be blank and the tool will learn from that computer's corrections.

## Important Grid Rule

The tool always uses 40 physical grid columns:

- Physical column 0 is blank.
- Physical column 39 is blank.
- The displayed/labeled columns are 1 through 38.

This means text should appear only in the labeled 1-38 columns.

## Tips For Good Accuracy

- Use the clearest photo available.
- Avoid glare on the display.
- Make sure the black MCDU screen is fully visible.
- In `Photo View`, drag the four corner handles exactly to the black display corners.
- In `Flattened View`, use the inset sliders until each character sits inside its grid cell.
- Correct OCR mistakes before clicking `Remember Corrections`.
- Use several corrected images to make the tool more reliable over time.

## Running From Source

If running from the GitHub source instead of an `.exe`, copy or clone the project and run:

```bash
python app.py
```

Then open:

```text
http://127.0.0.1:8766
```

Required Python packages:

```bash
python -m pip install pillow numpy python-docx
```

Tesseract OCR must also be installed on the computer.

On macOS with Homebrew, Tesseract is usually:

```text
/opt/homebrew/bin/tesseract
```

On Windows, it is commonly:

```text
C:\Program Files\Tesseract-OCR\tesseract.exe
```

If OCR does not run, check that Tesseract is installed and that the tool points to the correct Tesseract path.

## Minimum Source Files

For a fresh source-code setup with no previous learning, these are the required files:

```text
app.py
static/index.html
static/app.js
static/styles.css
```

The `data` folder is created automatically when the tool starts.

## Privacy

The tool runs locally on the user's computer. Images are processed locally and are not uploaded to a cloud service by this tool.

## Output

Word exports are saved in:

```text
data/exports/
```

Each export is a `.docx` file containing the 13-row by 40-column MCDU grid.
