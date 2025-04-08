# File Compressor with 7-Zip

A Python GUI application that compresses files and folders using 7-Zip, with support for reading file paths from Excel and OpenOffice spreadsheets.

## Features

- Read file paths from Excel (.xlsx, .xls) and OpenOffice (.ods) spreadsheets
- Compress files and folders using 7-Zip with maximum compression settings
- Skip already compressed files (zip, rar, 7z, etc.)
- Show compression progress and status updates
- Modern GUI interface with progress tracking
- Sound notification when compression is complete

## Requirements

- Windows operating system
- [7-Zip](https://www.7-zip.org/) installed on your system
- Python 3.6+ (if running from source)

## Installation

### Using the Executable

1. Download the latest release
2. Make sure 7-Zip is installed on your system
3. Run `FileCompressor7z.exe`

### From Source

1. Clone this repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the application:
   ```bash
   python compressapppy/file_compressor_7z.py
   ```

## Building the Executable

To build the executable yourself:

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
2. Run the build script:
   ```bash
   build.bat
   ```
The executable will be created in the `dist` folder.

## Usage

1. Launch the application
2. Click "Browse" to select an Excel or OpenOffice spreadsheet containing file paths
   - Paths should be in the first column
   - First row is considered a header and will be skipped
3. The application will validate the paths and display them in the list
4. Click "Compress Files/Folders" to start compression
5. Monitor progress through the progress bar and status messages
6. A sound will play when compression is complete

## File Format

Your Excel or OpenOffice spreadsheet should have file/folder paths in the first column, starting from the second row. For example:

| Path |
|------|
| C:\Documents\file1.txt |
| C:\Documents\folder1 |
| D:\Pictures |

## License

MIT License 