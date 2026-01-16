# XML Comparator

High-performance XML file comparison application between two configurable environments.

## Project Structure

```
XCT/
├── input/
│   ├── DEV68/          # Place folder 1 XML files here (configurable name)
│   └── DEV78/          # Place folder 2 XML files here (configurable name)
├── output/
│   └── run_YYYYMMDD_HHMMSS/      # One folder per execution with timestamp
│       ├── DEV68/                 # Renamed files (matchX.xml or unmatchedX.xml)
│       ├── DEV78/                 # Renamed files (matchX.xml or unmatchedX.xml)
│       ├── diff/                  # XML difference files (optional)
│       └── summary_YYYYMMDD_HHMMSS.xlsx  # Excel summary
├── xml_comparator.py   # Main script
├── run.bat             # Launcher file
├── conf.json           # Folder names and options configuration
├── requirements.txt    # Python dependencies
└── README.md          # This file
```

## Configuration (conf.json)

The `conf.json` file allows you to customize folder names and options:

```json
{
  "folder1_name": "DEV68",
  "folder2_name": "DEV78",
  "generate_diff": true
}
```

**Configuration Options:**
- **folder1_name**: Name of the first folder (default: "DEV68")
- **folder2_name**: Name of the second folder (default: "DEV78")
- **generate_diff**: Generate detailed XML diff files (true/false)
  - `true`: Creates detailed difference files for each match
  - `false`: Only finds matches without generating diff files (faster)

These names are used everywhere: input, output, Excel, etc.

## Installation

1. Assurez-vous d'avoir Python installé (version 3.7 ou supérieure)
2. Aucune installation manuelle nécessaire - le fichier .bat installe automatiquement les dépendances

## Installation

1. Ensure Python is installed (version 3.7 or higher)
2. No manual installation needed - the .bat file automatically installs dependencies

## Usage

1. **Configure folder names (optional):**
   - Edit `conf.json` to customize folder names
   - Default: DEV68 and DEV78
   - Set `generate_diff` to `false` if you only need matching without diff files (faster)

2. **Place your XML files:**
   - Copy your XML files to `input/[folder1_name]/`
   - Copy your XML files to `input/[folder2_name]/`

3. **Run the program:**
   - Double-click on `run.bat`
   - OR execute from terminal: `python xml_comparator.py`

4. **View results:**
   - A new folder `output/run_YYYYMMDD_HHMMSS/` is created for each execution
   - Matched files are renamed `match1.xml`, `match2.xml`, etc.
   - Unmatched files are named `unmatched1.xml`, `unmatched2.xml`, etc.
   - Detailed differences are in the `diff/` subfolder (if enabled)
   - An Excel summary is created in the run folder with timestamp

## Features

### Performance Optimization
- **Parallel Processing**: Uses all available CPU cores for faster execution
- **Multi-threaded Normalization**: XML files are processed in parallel
- **Concurrent Similarity Calculation**: All comparisons run simultaneously
- **Optimized File Operations**: Parallel file copying for better performance

### Flexible Configuration
- Customizable folder names via `conf.json`
- Optional diff file generation (can be disabled for faster execution)
- Each execution creates a unique folder with timestamp
- All results (files, diff, Excel) in one folder per run

### Intelligent Comparison
- Compares all XML files between two configured folders
- Calculates similarity rate with 10 decimal precision
- Finds optimal matches (each file can only match once)
- Handles cases where one folder has more files than the other
- Content-based matching (filenames don't matter)

### Excel Summary
The Excel file contains 3 sheets:

1. **Summary**: Global statistics
   - Processing time
   - Number of matches found
   - Number of unmatched files per folder
   - Diff generation status

2. **Matches**: Details for each match
   - Match number
   - Original filenames (with configurable names)
   - Precise similarity rate
   - Number of differences
   - Difference details

3. **Unmatched**: List of unmatched files
   - Original and new names for each folder

### Difference Files (Optional)
For each match, an XML file is created in `diff/` containing:
- Compared filenames
- Comparison timestamp
- Detailed list of all differences found

Set `"generate_diff": false` in conf.json to skip this step and improve performance.

## Matching Algorithm

1. Normalizes all XML contents (parallel processing)
2. Calculates similarity matrix (all vs all, parallel processing)
3. Sorts pairs by similarity (descending)
4. Assigns matches optimally (each file can only match once)
5. Marks remaining files as unmatched

## Important Notes

- Files with the same name are not automatically considered matches
- Comparison is based solely on content
- High precision (10 digits) detects even very small differences
- Non-valid XML files are compared line by line
- **Parallel processing** significantly reduces execution time on multi-core systems

## Performance Tips

- Set `"generate_diff": false` in conf.json if you only need to find matches (much faster)
- The program automatically uses all available CPU cores
- Performance scales with the number of cores available
- Large files may require more memory

## Dependencies

- Python 3.7+
- openpyxl (for Excel generation)
- Standard modules: xml.etree.ElementTree, difflib, pathlib, shutil, time, concurrent.futures

## Troubleshooting

Make sure that:
- Python is correctly installed
- input/DEV68 and input/DEV78 folders contain XML files
- XML files are well-formed
- Sufficient memory available for large file sets
