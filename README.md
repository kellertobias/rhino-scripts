# Rhino CAD Workflow Tools

A collection of Python scripts for automating CAD workflows in Rhino 8 for macOS, focusing on DWG file processing, layout assembly, and export operations.

## Overview

This project provides tools for:

- **Exporting deck sections** from Rhino models as DWG files
- **Assembling layouts** from exported DWG files into a multi-page PDF

## Requirements

- **Rhino 8 for macOS** (for `export.py` and `assemble_layouts.py`)

## Scripts

### `export.py`

**Purpose**: Generates ClippingDrawings for each DECK\_\* section in a Rhino model and exports them as DWG files.

**Usage**: Run this script from within Rhino 8 (macOS) using the Python script editor or command line.

**Behavior**:

- Discovers all clipping planes starting with `DECK_`
- For each section:
  1. Generates a ClippingDrawing using parallel projection
  2. Moves the drawing 3× model width to the right of the model center
  3. Exports the drawing layer (including sublayers) to `~/Desktop/<sectionName>-Export.dwg`
  4. Cleans up temporary layers and drawing objects

**ClippingDrawings Settings**:

- Angle: 0
- Projection: Parallel
- AddSilhouette: Yes
- ShowHatch: Yes
- ShowSolid: Yes
- AddBackground: Yes
- ShowLabel: No

**Output**: DWG files saved to `~/Desktop/<sectionName>-Export.dwg`

**Logging**: Logs are written to `~/Desktop/RhinoClippingExport.log`

### `assemble_layouts.py`

**Purpose**: Assembles multiple exported DWG files into a single multi-page PDF with proper layouts and scaling.

**Usage**: Run this script from within Rhino 8 (macOS) in a mostly empty Rhino file.

**Behavior**:

- Discovers DWG files (by default: `~/Desktop/*-Export.dwg`) or accepts explicit paths
- Optionally duplicates a master layout (if present) as the base for each page
- Creates one page layout per DWG and inserts a single detail viewport
- Sets the detail scale so that 1 mm on paper equals 0.2 m in the imported DWG by default
- Attempts to export a vector PDF of all layouts

**Scaling**: Default scale is 1 mm on paper = 0.2 m in DWG (adjustable via script arguments)

**Output**: Vector PDF with all layouts

**Logging**: Logs are written to `~/Desktop/LayoutAssembly.log`

## Workflow

Typical workflow:

1. **Export deck sections** (`export.py`):

   - Open your Rhino model with clipping planes named `DECK_*`
   - Run `export.py` to generate DWG files for each deck section
   - DWG files are saved to `~/Desktop/`

2. **Assemble layouts** (`assemble_layouts.py`):
   - Open a new or mostly empty Rhino file
   - Optionally create a master layout template
   - Run `assemble_layouts.py` to import DWG files and create layouts
   - Export a multi-page PDF

## File Structure

```
rhino/
├── assemble_layouts.py    # Rhino script: Assemble DWG files into PDF layouts
├── export.py              # Rhino script: Export deck sections as DWG files
├── dwg_to_png.py          # Standalone: Convert DWG/DXF to PNG
├── requirements.txt       # Python dependencies
└── README.md              # This file
```

## Notes

- **Rhino Scripts**: `export.py` and `assemble_layouts.py` are designed to run within Rhino 8 for macOS and use Rhino's Python scripting API (`rhinoscriptsyntax` and `Rhino` modules).
- **Master Layouts**: If a master layout exists in your Rhino file when running `assemble_layouts.py`, its contents will be copied to every generated page as a template.
- **Scaling**: Default scaling in `assemble_layouts.py` is 1 mm on paper = 0.2 m in DWG. Adjust via script arguments if needed.

## License

MIT License
