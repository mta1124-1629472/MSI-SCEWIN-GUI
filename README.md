# MSI Hidden BIOS GUI

## Overview
MSI Hidden BIOS GUI is a powerful tool designed to manage hidden BIOS settings on MSI systems. It provides an intuitive graphical interface for importing, exporting, and editing BIOS configurations, with advanced validation and performance optimizations.

## Features
- **Integrated Import/Export**: Easily save and load BIOS settings directly to/from your system.
- **Lazy Loading**: Optimized performance for large NVRAM files with infinite scrolling.
- **Validation**: Advanced validation ensures safe changes to BIOS settings.
- **Undo/Redo**: Track and revert changes with a robust undo/redo system.
- **Search & Filter**: Fuzzy search and category-based filtering for quick navigation.
- **Backup & Restore**: Automatic backups and easy restoration of previous configurations.
- **Progress Tracking**: Visual progress indicators for long-running operations.
- **Keyboard Navigation**: Navigate search results and settings efficiently using arrow keys and Enter.

## Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/mta1124-1629472/MSI-Hidden-Bios-GUI.git
   ```
2. Install dependencies:
   ```bash
   pip install rapidfuzz
   ```
3. Run the application:
   ```bash
   python msi-bios-editor.py
   ```

## Usage
1. **Export BIOS Settings**: Use the "Export BIOS & Load" button to extract BIOS settings into a file.
2. **Edit Settings**: Modify settings using the intuitive GUI, with validation to ensure safe changes.
3. **Import Changes**: Save and import changes back to the BIOS using the "Save & Import to BIOS" button.
4. **Backup & Restore**: Automatically create backups before importing changes, and restore previous configurations if needed.

## Known Issues
- Some settings may not display correctly due to malformed NVRAM files.
- Ensure administrative privileges are granted for BIOS operations.

## Contributing
Contributions are welcome! Please fork the repository and submit a pull request with your changes.

## License
This project is licensed under the MIT License. See the LICENSE file for details.
