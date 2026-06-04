# 777-9 MCMDU Grid Tool

Local office utility for extracting a 13-row by 40-column character grid from a phone photo of a Boeing 777-9 MCMDU screen.

## What it does

- Loads a phone photo in the browser.
- Auto-detects the black MCMDU display region in the photo.
- Lets the user drag the four MCMDU screen corners to correct tilted or non-perpendicular photos.
- Flattens angled screen photos into a rectangular display preview.
- Draws an equal-size 13 x 40 perspective grid over the screen.
- Labels rows 1 to 13.
- Labels physical grid columns 2 through 39 as columns 1 through 38, leaving the first and last columns unlabeled.
- Sends the corrected screen area to local Tesseract OCR.
- Maps detected text into an editable table.
- Exports the table as a Word `.docx` file.
- Stores corrected row text in `data/corrections.json` so repeated OCR mistakes can be corrected on future analyses.
- Stores corrected per-cell character image templates in `data/templates.json` so the fixed MCMDU font can be recognized better after user correction.

## Run

```bash
chmod +x run.sh
./run.sh
```

Open:

```text
http://localhost:8766
```

## Workflow

1. Click `Choose Image`.
2. Let `Auto Detect Display` place the corners on the black display, or click it manually if needed.
3. Use `Photo View` to drag the numbered corner handles to the exact black display corners.
4. Use `Flattened View` to see the corrected rectangular display with the straight grid over it.
5. Use the inset sliders if the text grid starts slightly inside the detected screen area.
6. Click `Analyze Grid`.
7. Edit any cells that OCR missed or misread.
8. Click `Remember Corrections` after editing cells so the tool learns MCMDU character samples.
9. Click `Export Word File` to download the Word reference table.

## Notes

- OCR runs locally through `/opt/homebrew/bin/tesseract`; no image is uploaded to the cloud.
- The first OCR pass may need correction because phone glare, blur, and MCMDU font rendering vary. After corrections are remembered, the tool also saves per-cell character templates for future analyses.
- Generated files are written to `data/exports/`.
