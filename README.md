# Beam Structural Analysis Report Generator

A professional Python application for generating comprehensive engineering reports for simply supported beam analysis using PyLaTeX. This tool reads force data from Excel files and produces publication-quality PDF reports with vector graphics diagrams.

## Overview

This application automates the creation of structural analysis reports for simply supported beams. It processes beam loading data from Excel spreadsheets and generates detailed PDF reports that include shear force diagrams (SFD), bending moment diagrams (BMD), and comprehensive analysis documentation.

### Key Features

- Professional PDF report generation with LaTeX typesetting
- TikZ/pgfplots vector graphics for SFD and BMD diagrams
- LaTeX-based tables with selectable text (not embedded images)
- Automated structural analysis calculations
- Command-line argument support for flexible usage
- Comprehensive error handling and validation

## Requirements

### System Requirements

- Python 3.7 or higher
- LaTeX distribution (pdflatex compiler)
  - Windows: MiKTeX or TeX Live
  - macOS: MacTeX
  - Linux: TeX Live

### Python Dependencies

```
pylatex>=1.4.1
pandas>=1.3.0
numpy>=1.21.0
openpyxl>=3.0.9
```

## Installation

### Step 1: Install LaTeX Distribution

**Windows (MiKTeX):**
1. Download from https://miktex.org/download
2. Run installer with administrator privileges
3. During installation, set "Install missing packages on-the-fly" to "Yes"
4. After installation, open MiKTeX Console and update all packages

**macOS (MacTeX):**
1. Download from https://www.tug.org/mactex/
2. Install the package (approximately 4GB download)
3. Restart terminal after installation

**Linux (Ubuntu/Debian):**
```bash
sudo apt-get update
sudo apt-get install texlive-full
```

### Step 2: Install Python Dependencies

```bash
pip install -r requirements.txt
```

Or install packages individually:
```bash
pip install pylatex pandas numpy openpyxl
```

### Step 3: Verify Installation

```bash
python --version        # Should show Python 3.7 or higher
pdflatex --version      # Should show LaTeX version information
```

## Usage

### Basic Usage

```bash
python docGenerator.py
```

This command will:
1. Read force data from `force_table.xlsx`
2. Embed the beam diagram image from `ssbeam.png`
3. Generate a PDF report named `output.pdf`

### Command-Line Arguments

```bash
python docGenerator.py -e force_table.xlsx -i ssbeam.png -o output.pdf
```

**Available options:**
- `-e, --excel`: Path to Excel file with force data (default: force_table.xlsx)
- `-i, --image`: Path to beam configuration image (default: ssbeam.png)
- `-o, --output`: Output PDF filename (default: output.pdf)
- `-v, --verbose`: Enable verbose output

### Display Help Information

```bash
python docGenerator.py --help
```

## Input Data Format

### Excel File Structure

The Excel file must contain three columns with the following headers:

| Column Header | Description | Units |
|---------------|-------------|-------|
| x | Position along beam | meters (m) |
| Shear force | Shear force at position | kilonewtons (kN) |
| Bending Moment | Bending moment at position | kilonewton-meters (kN·m) |

### Example Data

```
x       Shear force     Bending Moment
0       45              0
1.5     36              60.75
3       27              108
4.5     18              141.75
6       9               162
```

### Image Requirements

- Supported formats: JPG, PNG, PDF
- Recommended resolution: 1200x600 pixels or higher
- Clear depiction of beam configuration with supports and loading

## Output Structure

The generated PDF report contains the following sections:

### 1. Title Page
- Report title
- Author information
- Generation date
- Abstract summarizing the analysis

### 2. Table of Contents
- Automatically generated section listing
- Clickable hyperlinks to each section
- Page number references

### 3. Introduction
**3.1 Beam Description**
- Overview of simply supported beam configuration
- Embedded beam diagram image
- Description of support conditions

**3.2 Analysis Methodology**
- Description of structural analysis approach
- Step-by-step methodology

**3.3 Data Source**
- Excel file information
- Number of data points
- Beam length specification

### 4. Input Data
- Complete force and moment data table
- LaTeX-formatted tabular environment (selectable text)
- Professional table styling with horizontal rules

**4.1 Data Summary**
- Statistical summary of extracted data
- Shear force and bending moment statistics

### 5. Structural Analysis
**5.1 Shear Force Diagram (SFD)**
- Definition and engineering significance
- Key observations (maximum positive/negative values, zero crossing)
- TikZ/pgfplots vector graphic plot
- Grid lines, axis labels, and legend

**5.2 Bending Moment Diagram (BMD)**
- Definition and engineering significance
- Key observations (maximum moment location)
- TikZ/pgfplots vector graphic plot
- Filled area representation
- Grid lines, axis labels, and legend

**5.3 Analysis Summary**
- Interpretation of results
- Structural behavior characteristics
- Design considerations

### 6. Engineering Interpretation
**6.1 Critical Sections**
- Location and magnitude of maximum bending moment
- Location and magnitude of maximum shear force

**6.2 Design Implications**
- Flexural capacity requirements
- Shear resistance considerations
- Support design requirements

### 7. Conclusion
- Summary of analysis findings
- Engineering implications
- Design recommendations
- Report generation timestamp

## Troubleshooting

### LaTeX Package Missing Error

**Problem:** `LaTeX Error: File 'package.sty' not found`

**Solution for MiKTeX:**
1. Open MiKTeX Console as Administrator
2. Go to Settings tab
3. Enable "Install missing packages on-the-fly"
4. Go to Updates tab and update all packages

**Solution for TeX Live:**
```bash
sudo tlmgr update --self
sudo tlmgr install <package-name>
```

### Python Module Not Found

**Problem:** `ModuleNotFoundError: No module named 'pylatex'`

**Solution:**
```bash
pip install pylatex pandas numpy openpyxl
```

### PDF Compilation Fails

**Problem:** pdflatex compilation errors

**Solution:**
1. Verify LaTeX distribution is properly installed
2. Check that pdflatex is in system PATH
3. Review the `.log` file for specific error messages
4. Ensure all required LaTeX packages are installed

### Excel File Reading Error

**Problem:** Cannot open or read Excel file

**Solution:**
1. Verify file path is correct
2. Ensure file has .xlsx extension (not .xls)
3. Check file is not corrupted
4. Confirm file is not open in another application

### Image Not Embedded

**Problem:** Image does not appear in generated PDF

**Solution:**
1. Verify image file path is correct
2. Use supported formats (JPG, PNG, PDF)
3. Check image file is not corrupted
4. Ensure proper file permissions

## Technical Documentation

### Architecture

The application consists of the following main components:

**Data Processing Module**
- Excel file reading and validation
- Column and data type verification
- Input file existence checks

**Table Generation Module**
- LaTeX Tabular environment creation
- Data formatting and alignment
- Professional styling implementation

**Plot Generation Module**
- TikZ coordinate generation
- Automatic axis scaling
- Plot customization and styling

**Document Assembly Module**
- PyLaTeX document structure
- Section and subsection organization
- Package management
- PDF compilation

### Technologies Used

- **Python 3.7+**: Core programming language
- **PyLaTeX**: LaTeX document generation library
- **pandas**: Data manipulation and Excel reading
- **numpy**: Numerical computations
- **openpyxl**: Excel file processing engine
- **TikZ/pgfplots**: Vector graphics generation
- **pdflatex**: PDF compilation from LaTeX source

### Code Quality Standards

- PEP 8 compliant Python code
- Comprehensive function documentation with docstrings
- Command-line argument support (argparse)
- Enhanced error handling and validation
- Input file existence checks
- Modular function design
- Clear variable naming conventions

## License

This project is licensed under the Creative Commons Attribution-ShareAlike 4.0 International License (CC BY-SA 4.0).

You are free to:
- Share: copy and redistribute the material in any medium or format
- Adapt: remix, transform, and build upon the material for any purpose

Under the following terms:
- Attribution: You must give appropriate credit, provide a link to the license, and indicate if changes were made
- ShareAlike: If you remix, transform, or build upon the material, you must distribute your contributions under the same license

For full license details, see LICENSE file or visit:
https://creativecommons.org/licenses/by-sa/4.0/

## Acknowledgments

Developed as part of the FOSSEE Osdag screening task initiative. This project demonstrates the integration of Python automation with LaTeX document generation for engineering applications.

## Contributing

This project was created for educational and assessment purposes. For questions or issues related to the FOSSEE Osdag screening task, please refer to the official communication channels.

## Version History

**Version 2.0.0** (January 2026)
- Added command-line argument support
- Enhanced error handling and validation
- Added engineering interpretation section
- Improved documentation with comprehensive docstrings
- Added data summary statistics
- Enhanced report structure with methodology section

**Version 1.0.0** (January 2026)
- Initial release
- Core functionality implementation
- TikZ/pgfplots diagram generation
- LaTeX table creation
- Documentation and examples