# MSI SCEWIN GUI

![image](https://github.com/user-attachments/assets/f0bfe5d4-47f6-446c-ac9f-9a9dc309b4dd)

## Overview

MSI SCEWIN GUI is a powerful tool designed to manage hidden BIOS settings on MSI systems using the SCEWIN utility. It provides an intuitive graphical interface for importing, exporting, and editing BIOS configurations, with advanced validation and performance optimizations.

## Features

- **🔄 Integrated Import/Export**: Easily save and load BIOS settings directly to/from your system using SCEWIN
- **⚡ Performance Optimized**: Lazy loading and pagination for handling large NVRAM files with hundreds of settings
- **✅ Advanced Validation**: Multi-layer validation ensures safe changes with range checking and option validation
- **↩️ Undo/Redo System**: Track and revert changes with a robust 30-level undo/redo system
- **🏷️ Smart Categories**: Auto-generated category filters based on setting content for quick navigation
- **🛡️ Backup & Restore**: Automatic timestamped backups before each import with easy restoration
- **📊 Progress Tracking**: Visual progress indicators for long-running BIOS operations
- **⌨️ Keyboard Navigation**: Navigate search results and settings efficiently using arrow keys and Enter
- **🎯 Real-time Search**: Debounced search with instant results as you type
- **🔧 SCEWIN Integration**: Dynamic detection of MSI Center installations with comprehensive path searching

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/mta1124-1629472/MSI-SCEWIN-GUI.git
   ```

2. Ensure MSI Center is installed on your system:
   - Download and install MSI Center from the [official MSI website](https://www.msi.com/Landing/MSI-Center)
   - The application will automatically detect SCEWIN tools in standard MSI Center installation paths
   - If SCEWIN is not found, use the "Check SCEWIN Status" button for troubleshooting

3. Run the application with administrative privileges (recommended for BIOS operations):

   ```bash
   python msi-bios-editor.py
   ```

## Key Features in Detail

### 🛡️ Safety Features

- **Change Review Dialog**: Preview all modifications before applying to BIOS
- **Advanced Validation**: Range checking, option validation, and consistency checks
- **Automatic Backups**: Timestamped backups created before each BIOS import
- **Invalid Field Detection**: Real-time validation with visual feedback for problematic values
- **Admin Rights Detection**: Warns when administrator privileges are needed

### ⚡ Performance

- **Paginated Display**: Handle hundreds of settings without UI lag (25 settings per page)
- **Memory Optimization**: Lazy loading and efficient widget management
- **Debounced Search**: Intelligent delays prevent excessive processing during typing
- **Background Operations**: Non-blocking BIOS operations with progress tracking

## Usage

### Quick Start

1. **Export BIOS Settings**: Click "🔽 Export BIOS & Load" to extract current BIOS settings
2. **Search & Edit**: Use the fuzzy search to find settings, then modify values using dropdowns or text fields
3. **Validate Changes**: Check for any red-highlighted invalid fields before proceeding
4. **Review & Import**: Click "🔼 Save & Import to BIOS" to review changes and apply them to your system

### Safety Best Practices

- **Always review changes** in the confirmation dialog before importing
- **Keep backups** - the app automatically creates them, but manual backups are wise
- **Test gradually** - make small changes first to ensure system stability
- **Use "Check SCEWIN Status"** to verify SCEWIN availability before operations

## Requirements

- **Windows OS** with MSI motherboard
- **MSI Center** installed (for SCEWIN tools)
- **Python 3.7+** with tkinter support
- **Administrator privileges** (recommended for BIOS operations)

## Troubleshooting

### SCEWIN Not Found

1. Ensure MSI Center is properly installed
2. Check the installation paths in "Check SCEWIN Status"
3. Try reinstalling MSI Center from the official website
4. Manually copy SCEWIN files to the application directory if needed

### Permission Issues

- Run the application as Administrator for full BIOS access
- Some BIOS operations require elevated privileges
- The app will warn you if admin rights are needed

### Invalid Field Errors

- Red-highlighted fields indicate invalid values
- Check ranges and allowed options for each setting
- Correct all invalid fields before importing to BIOS

## Known Issues

- Some settings may not display correctly due to malformed NVRAM files
- Very large NVRAM files (1000+ settings) may experience slower initial loading

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request with your changes.

## License

This project is licensed under the MIT License. See the LICENSE file for details.
