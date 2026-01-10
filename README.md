# Panther Reference Verification

A tool to detect potentially fabricated references in student papers. Developed by Darby Proctor, Ph.D.

## Features

- **Reference Verification**: Checks references against CrossRef, PubMed (journals), Open Library, and Google Books (books)
- **Citation-Reference Matching**: Verifies that in-text citations match reference list entries
- **Detailed Reports**: Generates Word documents with color-coded verification status
- **Batch Processing**: Process multiple student papers at once

## Download

### Pre-built Executables (Recommended for most users)

Download the latest release for your operating system:

- **Windows**: [ReferenceChecker.exe](../../releases/latest)
- **macOS**: [ReferenceChecker-macOS.dmg](../../releases/latest)

### From Source

If you prefer to run from source:

```bash
# Clone the repository
git clone https://github.com/DarbyP/Panther_Reference_Verification.git
cd reference-checker

# Install dependencies
pip install -r requirements.txt

# Run the application
python reference_checker.py
```

## Usage

1. **Select Papers Folder**: Choose a folder containing student papers (.docx or .pdf files)
2. **Choose Output Location**: Select where to save the verification report
3. **Adjust Settings** (optional):
   - Matching thresholds for verified/partial matches
   - Ignore books option (book verification accuracy is limited)
   - Skip citation-reference matching if not needed
4. **Run Verification**: Click the button and wait for processing
5. **Review Report**: Open the generated Word document to see results

## Report Status Codes

| Status | Meaning |
|--------|---------|
| ✅ Verified | Reference found with high confidence match |
| ⚠️ Partial Match | Similar reference found (may need manual review) |
| ❌ No Match | Reference not found - may be fabricated |
| 🔗 DOI Mismatch | DOI links to a different paper |
| 📚 Book (Manual) | Book/chapter - verify manually |
| 🌐 Website (Manual) | Website source - verify manually |

## Building from Source

### Requirements

- Python 3.9+
- Dependencies in `requirements.txt`

### Building Executables

#### Windows

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "ReferenceChecker" --icon "assets/panther_icon.ico" --add-data "assets;assets" --add-data "docs;docs" reference_checker.py
```

#### macOS

```bash
pip install pyinstaller
pyinstaller --onedir --windowed --name "ReferenceChecker" --icon "assets/panther_icon.icns" --add-data "assets:assets" --add-data "docs:docs" reference_checker.py
```

## License

MIT License - see LICENSE file for details.

## Acknowledgments

- CrossRef API for journal article verification
- PubMed/NCBI for biomedical reference checking
- Open Library and Google Books APIs for book verification
