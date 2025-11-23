# turnitin-rubric-converter

A Python script and a Shiny app to convert rubric files between Excel and two common formats:
- **Turnitin's Custom rubric** (.rbc files)
- **IMS Global specification** (.json files)

## Features

- Convert Turnitin .rbc files to Excel for editing
- Convert IMS specification .json files to Excel for editing
- Convert Excel files back to Turnitin .rbc or IMS .json format
- Support for variable numbers of levels per criterion (IMS format)
- Web-based Shiny app for easy conversions
- Command-line tool for batch processing

## Format Differences

### Turnitin Format
- All criteria must have the same set of scales (grid structure)
- Uses named scales (e.g., "Excellent", "Good", "Fair", "Poor")
- Fixed structure across all criteria

### IMS Format
- Each criterion can have a different number of levels
- More flexible structure
- Follows IMS Global Learning Consortium standards

## Quick Start

### Command Line Usage

**Convert from Turnitin/IMS to Excel:**
```bash
python rubric_converter.py yourrubric.rbc
# or
python rubric_converter.py yourrubric.json
```

**Convert from Excel to Turnitin format (default):**
```bash
python rubric_converter.py yourrubric.xlsx
```

**Convert from Excel to IMS format:**
```bash
python rubric_converter.py yourrubric.xlsx -f ims
```

### Web App Usage

Run the Shiny app:
```bash
shiny run app.py
```

Then open your browser and upload files for conversion.

## Requirements

See `requirements.txt` for dependencies.
