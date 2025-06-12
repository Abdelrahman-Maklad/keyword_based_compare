# Magic Document Processing Tools

A collection of tools for automated document processing and analysis.

## Tools Included

1. Magic.py - Main document processing tool
2. Magic_Special.py - Specialized processing for specific document types

## Features

- Automated document analysis
- Pattern matching and extraction
- Batch processing capabilities
- Excel input/output support
- Custom keyword processing
- Results aggregation

## Requirements

- Python 3.7+
- Required packages:
  - pandas
  - openpyxl
  - PyMuPDF (fitz)

## Installation

```bash
pip install pandas openpyxl pymupdf
```

## Usage

1. Prepare input files:
   - Place documents in input directory
   - Configure input.xlsx with processing parameters
   - Set up fixed_path.txt with path configurations

2. Run the tool:
```bash
python Magic.py
```

3. Results will be saved in:
   - VC_output.xlsx for detailed results

## Configuration

- input.xlsx: Define processing parameters
- fixed_path.txt: Configure file paths
- DataBase/: Store reference data

## License

MIT License - See LICENSE file for details
