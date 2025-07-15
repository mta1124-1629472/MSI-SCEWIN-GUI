#!/usr/bin/env python3
"""
MSI SCEWIN GUI - Enhanced BIOS Settings Editor
- Integrated import/export functionality with SCEWIN
- Performance optimizations with lazy loading
- Progress bars for long operations
- Threading for non-blocking operations
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import re
import os
import subprocess
import ctypes
import sys
import shutil
import threading
import queue
import gc
import weakref
import glob
from datetime import datetime
import time
from collections import defaultdict
from rapidfuzz import fuzz, process
import webbrowser
import tempfile
from typing import Optional, List, Tuple, Union


def run_as_admin():
    if ctypes.windll.shell32.IsUserAnAdmin():
        return True
    # Relaunch as admin
    params = ' '.join([f'"{x}"' for x in sys.argv])
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, params, None, 1)
    sys.exit()


# Call this at the start of your script
run_as_admin()


class BIOSSetting:
    """Represents a single BIOS setting with optimized structure"""
    __slots__ = ("setup_question", "help_string", "token", "offset", "width", "bios_default", "options", "current_value", "is_numeric", "original_value", "original_has_options", "original_block_lines", "range_min", "range_max")
    
    def __init__(self, setup_question: str = "", help_string: str = "", token: str = "", offset: str = "", 
                 width: str = "", bios_default: str = "", options: Optional[List[Tuple[str, str, bool]]] = None, 
                 current_value: Optional[str] = None, is_numeric: bool = False):
        self.setup_question: str = setup_question
        self.help_string: str = help_string
        self.token: str = token
        self.offset: str = offset
        self.width: str = width
        self.bios_default: str = bios_default
        self.options: List[Tuple[str, str, bool]] = options or []
        self.current_value: str = current_value or ""
        self.is_numeric: bool = is_numeric
        self.original_value: str = ""
        self.original_has_options: bool = False
        self.original_block_lines: List[str] = []
        self.range_min: Optional[int] = None
        self.range_max: Optional[int] = None


class ProgressDialog:
    """Enhanced progress dialog with better user feedback"""
    def __init__(self, parent, title="Processing", message="Please wait..."):
        self.parent = parent
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("450x180")  # Increased height for better button spacing
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center the dialog
        self.dialog.geometry(f"+{parent.winfo_rootx() + 50}+{parent.winfo_rooty() + 50}")
        
        # Create widgets
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Message label with text wrapping
        self.message_label = ttk.Label(main_frame, text=message, font=("Arial", 10), wraplength=400)
        self.message_label.pack(pady=(0, 15))
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='determinate', length=380)
        self.progress.pack(pady=(0, 15))
        
        # Status label with text wrapping
        self.status_label = ttk.Label(main_frame, text="Starting...", font=("Arial", 9), wraplength=400)
        self.status_label.pack(pady=(0, 15))
        
        # Cancel button with proper sizing
        self.cancel_button = ttk.Button(main_frame, text="Cancel", command=self.cancel, width=12)
        self.cancel_button.pack(pady=(5, 0))
        
        self.cancelled = False
        self.dialog.protocol("WM_DELETE_WINDOW", self.cancel)
    
    def update_progress(self, value, status=""):
        """Update progress bar and status"""
        if not self.cancelled:
            self.progress['value'] = value
            if status:
                self.status_label.config(text=status)
            self.dialog.update()
    
    def cancel(self):
        """Handle cancellation"""
        self.cancelled = True
        self.close()
    
    def close(self):
        """Close the dialog"""
        try:
            self.dialog.destroy()
        except tk.TclError:
            pass


class OptimizedNVRAMParser:
    """High-performance NVRAM parser with aggressive optimizations"""
    
    def __init__(self):
        self.settings = []
        self.header_info = {}
        self.raw_header = ""
        self.categories = defaultdict(list)
        self._cancelled = False
        
        # Pre-compile regex patterns for better performance
        self._setup_question_pattern = re.compile(r'Setup Question\s*=\s*(.+)', re.IGNORECASE)
        self._token_pattern = re.compile(r'Token\s*=\s*([A-Fa-f0-9]+)', re.IGNORECASE)
        self._help_pattern = re.compile(r'Help\s*=\s*(.+)', re.IGNORECASE)
        self._range_pattern = re.compile(r'range[:=]?\s*(\d+)\s*[~\-]\s*(\d+)', re.IGNORECASE)
        self._block_split_pattern = re.compile(r'\n(?=Setup Question\s*=)', re.IGNORECASE)
        
    def parse_file(self, file_path, progress_callback=None, cancel_flag=None):
        """Parse NVRAM file with optimized processing and progress updates"""
        self.settings = []
        self.raw_header = ""
        self.categories.clear()
        self._cancelled = False
        
        try:
            if progress_callback:
                progress_callback(5, "Reading file...")
            
            # Memory-efficient file reading with buffering
            with open(file_path, 'r', encoding='utf-8', errors='replace', buffering=8192) as file:
                content = file.read()
            
            if progress_callback:
                progress_callback(15, "Parsing header...")
            
            # Extract and parse header more efficiently
            header_end_pos = content.find("Setup Question")
            if header_end_pos > 0:
                self.raw_header = content[:header_end_pos]
                self._parse_header(self.raw_header)
            
            if cancel_flag and cancel_flag():
                return []
            
            if progress_callback:
                progress_callback(25, "Splitting setting blocks...")
            
            # Optimized block splitting using pre-compiled regex
            setting_content = content[header_end_pos:] if header_end_pos > 0 else content
            setting_blocks = self._block_split_pattern.split(setting_content)
            
            # Filter out empty blocks and strip whitespace in one pass
            setting_blocks = [block.strip() for block in setting_blocks if block.strip() and len(block.strip()) > 10]
            total_blocks = len(setting_blocks)
            
            if progress_callback:
                progress_callback(35, f"Processing {total_blocks} settings...")
            
            # Process blocks with optimized batch updates
            batch_size = max(10, total_blocks // 20)  # Larger batches for better performance
            processed_settings = []
            
            for i, block in enumerate(setting_blocks):
                if cancel_flag and cancel_flag():
                    break
                
                setting = self._parse_setting_block_optimized(block)
                if setting:
                    processed_settings.append(setting)
                    
                    # Batch categorization for better performance
                    if len(processed_settings) % 100 == 0:  # Categorize in batches of 100
                        self._batch_categorize(processed_settings[-100:])
                
                # Update progress less frequently for better performance
                if progress_callback and (i % batch_size == 0 or i == total_blocks - 1):
                    progress_value = 35 + int((i + 1) / total_blocks * 55)
                    status = f"Processed {i + 1}/{total_blocks} settings"
                    progress_callback(progress_value, status)
            
            # Final categorization of remaining settings
            remaining = len(processed_settings) % 100
            if remaining > 0:
                self._batch_categorize(processed_settings[-remaining:])
            
            self.settings = processed_settings
            
            if progress_callback:
                progress_callback(95, "Optimization complete...")
            
            # Force garbage collection to free memory
            import gc
            gc.collect()
            
            return self.settings
            
        except MemoryError:
            messagebox.showerror("Memory Error", "Not enough memory to process this file. Try closing other applications.")
            return []
        except Exception as e:
            messagebox.showerror("Parsing Error", f"Failed to parse BIOS file:\n{e}")
            return []
    
    def _parse_setting_block_optimized(self, block):
        """Highly optimized setting block parser with pre-compiled patterns"""
        if not block or block.strip().startswith('//'):
            return None
        
        # Save original block lines
        block_lines = [line for line in block.split('\n') if line.strip()]
        lines = [line.strip() for line in block_lines if not line.strip().startswith('//')]
        
        if not lines:
            return None
            
        # Fast pattern matching using pre-compiled regex
        setup_match = self._setup_question_pattern.search(lines[0])
        if not setup_match:
            return None
            
        setting = BIOSSetting()
        setting.original_block_lines = block_lines
        setting.setup_question = setup_match.group(1).strip()
        
        try:
            # Optimized parsing using compiled patterns
            for line in lines[1:]:
                if line.startswith('Help String'):
                    setting.help_string = self._extract_value_fast(line)
                elif line.startswith('Token'):
                    token_match = self._token_pattern.search(line)
                    if token_match:
                        setting.token = token_match.group(1)
                elif line.startswith('Offset'):
                    setting.offset = self._extract_value_fast(line)
                elif line.startswith('Width'):
                    setting.width = self._extract_value_fast(line)
                elif line.startswith('BIOS Default'):
                    setting.bios_default = self._extract_value_fast(line).strip('<>')
                elif line.startswith('Value'):
                    value_match = re.search(r'<([^>]+)>', line)
                    setting.current_value = value_match.group(1) if value_match else self._extract_value_fast(line)
                elif line.startswith('Options') or '[' in line:
                    if not setting.options:
                        setting.options = []
                    self._process_option_line_fast(setting, line)
            
            # Fast numeric validation and range extraction
            if setting.current_value and setting.current_value.isdigit():
                setting.is_numeric = True
                if setting.help_string:
                    range_match = self._range_pattern.search(setting.help_string)
                    if range_match:
                        setting.range_min = int(range_match.group(1))
                        setting.range_max = int(range_match.group(2))
            
            # Auto-generate Enabled/Disabled options for 0/1 values
            if (setting.current_value in ('0', '1') and not setting.options and 
                ('enable' in setting.setup_question.lower() or 'disable' in setting.setup_question.lower())):
                setting.options = [('1', 'Enabled', setting.current_value == '1'), 
                                 ('0', 'Disabled', setting.current_value == '0')]
                setting.is_numeric = False
            
            return setting if setting.setup_question and setting.token else None
            
        except Exception:
            return None
    
    def _extract_value_fast(self, line):
        """Fast value extraction using string operations"""
        if '=' not in line:
            return ""
        value = line.split('=', 1)[1].strip()
        return value[1:-1] if value.startswith('<') and value.endswith('>') else value
    
    def _process_option_line_fast(self, setting, line):
        """Fast option processing with minimal regex"""
        try:
            is_current = '*' in line
            clean_line = line.replace('*', '', 1) if is_current else line
            
            # Fast bracket extraction
            start = clean_line.find('[')
            end = clean_line.find(']', start)
            if start != -1 and end != -1:
                value = clean_line[start+1:end].strip()
                desc_start = end + 1
                comment_pos = clean_line.find('//', desc_start)
                description = clean_line[desc_start:comment_pos if comment_pos != -1 else len(clean_line)].strip()
                
                setting.options.append((value, description, is_current))
                if is_current:
                    setting.current_value = value
        except Exception:
            pass
    
    def _batch_categorize(self, settings_batch):
        """Batch categorization for better performance"""
        for setting in settings_batch:
            category = self._extract_category_fast(setting.setup_question)
            self.categories[category].append(len(self.settings) + len(settings_batch) - len(settings_batch) + settings_batch.index(setting))
    
    def _extract_category_fast(self, question):
        """Fast category extraction using string operations"""
        if not question:
            return "Other"
        # Get first word up to space or special character
        for i, char in enumerate(question):
            if char in ' :-_()[]':
                return question[:i] if i > 0 else "Other"
        return question[:15] if len(question) > 15 else question  # Limit category length
    
    def _parse_header(self, header_content):
        """Extract header information with error handling"""
        try:
            for line in header_content.split('\n'):
                line = line.strip()
                if 'Script File Name' in line and ':' in line:
                    self.header_info['filename'] = line.split(':', 1)[1].strip()
                elif 'Created on' in line and ':' in line:
                    parts = line.split(':', 2)
                    if len(parts) > 2:
                        self.header_info['created'] = parts[2].strip()
                elif 'AMISCE Utility' in line:
                    self.header_info['utility'] = line.strip()
                elif 'HIICrc32' in line and '=' in line:
                    self.header_info['crc32'] = line.split('=', 1)[1].strip()
        except Exception:
            pass  # Skip header parsing errors
    
    def _parse_setting_block(self, block):
        """Optimized setting block parser, skips commented-out (//) settings and lines"""
        if not block:
            return None
        # Save original block lines (including comments and formatting)
        block_lines = [line for line in block.split('\n') if line.strip()]
        # Remove lines that start with // (commented out) for parsing, but keep for original_block_lines
        lines = [line.strip() for line in block_lines if not line.strip().startswith('//')]
        if not lines:
            return None
        # If the first line is not a Setup Question, skip this block
        if not lines[0].startswith('Setup Question'):
            return None
        setting = BIOSSetting()
        setting.original_block_lines = block_lines
        try:
            # Use regex patterns for faster parsing
            setup_match = re.search(r'Setup Question\s*=\s*(.+)', lines[0])
            if setup_match:
                setting.setup_question = setup_match.group(1).strip()
            # Track if we see a Value line and its value
            value_line = None
            value_val = None
            help_string = ""
            has_options = False
            bios_default_found = False
            for line in lines[1:]:
                if line.startswith('Help String'):
                    help_string = self._extract_value(line)
                    setting.help_string = help_string
                elif line.startswith('Token'):
                    token_value = self._extract_value(line)
                    setting.token = token_value.split('//')[0].strip() if token_value else ""
                elif line.startswith('Offset'):
                    setting.offset = self._extract_value(line)
                elif line.startswith('Width'):
                    setting.width = self._extract_value(line)
                elif line.startswith('BIOS Default'):
                    bios_default_value = self._extract_value(line)
                    # Handle different formats: "BIOS Default = <value>" or "BIOS Default = value"
                    if bios_default_value:
                        # Remove angle brackets if present
                        if bios_default_value.startswith('<') and bios_default_value.endswith('>'):
                            bios_default_value = bios_default_value[1:-1]
                        setting.bios_default = bios_default_value.strip()
                        bios_default_found = True
                elif line.startswith('Options') or (setting.options is not None and '[' in line):
                    if not setting.options:
                        setting.options = []
                    self._process_option_line(setting, line)
                    has_options = True
                elif line.startswith('Value'):
                    value_line = line
                    value_match = re.search(r'<([^>]+)>', line)
                    if value_match:
                        value_val = value_match.group(1)
                        setting.current_value = value_val
                    else:
                        value_val = self._extract_value(line)
                        setting.current_value = value_val
            setting.original_has_options = has_options
            # Special handling: if numeric and value is 0 or 1, and help string indicates Enabled/Disabled, treat as options
            if value_line is not None and value_val is not None:
                try:
                    v = int(value_val)
                    if v in (0, 1):
                        # Look for 'Enabled' and 'Disabled' in help string or comment
                        if (("enabled" in help_string.lower() and "disabled" in help_string.lower()) or
                            ("enabled" in setting.setup_question.lower() and "disabled" in setting.setup_question.lower())):
                            setting.options = [('1', 'Enabled', v == 1), ('0', 'Disabled', v == 0)]
                            setting.is_numeric = False
                            # Set current_value as string for consistency
                            setting.current_value = str(value_val)
                        else:
                            setting.is_numeric = True
                    else:
                        setting.is_numeric = True
                except Exception:
                    setting.is_numeric = True
            # Set default current value if not set
            if not setting.current_value and setting.options:
                # Find the option marked with * or use first option
                for value, desc, is_current in setting.options:
                    if is_current:
                        setting.current_value = value
                        break
                else:
                    setting.current_value = setting.options[0][0] if setting.options else ""


            # --- Set bios_default by checking explicit BIOS Default line first, then help string ---
            # Priority 1: Explicit "BIOS Default" line (if found)
            if not bios_default_found and setting.help_string:
                # Priority 2: Look for patterns in help string like "Default is X", "default: X", etc.
                help_lower = setting.help_string.lower()
                
                # Try various patterns for default value specification
                # Only match explicit default value specifications, not descriptive text
                patterns = [
                    r'default\s+is\s+([^\s,.;]+)',           # "Default is X"
                    r'default\s*:\s*([^\s,.;]+)',            # "default: X" or "default:X"
                    r'default\s*=\s*([^\s,.;]+)',            # "default = X" or "default=X"
                    r'by\s+default\s*:\s*([^\s,.;]+)',       # "by default: X"
                    r'by\s+default\s*=\s*([^\s,.;]+)',       # "by default = X"
                    r'defaults\s+to\s+([^\s,.;]+)',          # "defaults to X"
                ]
                
                for pattern in patterns:
                    match = re.search(pattern, help_lower)
                    if match:
                        default_value = match.group(1).strip()
                        # Clean up common suffixes that might be captured
                        default_value = re.sub(r'[.,;]$', '', default_value)
                        setting.bios_default = default_value
                        break
            
            # If no explicit default found in either place, leave blank
            if not hasattr(setting, 'bios_default') or not setting.bios_default:
                setting.bios_default = ""

            # Extract range information from help string if available
            if setting.help_string:
                # Try multiple patterns for range detection
                range_patterns = [
                    r'range[:=]?\s*(\d+)\s*[~\-]\s*(\d+)',                    # "range: 0 - 255" or "range 0~255"
                    r'\(\s*(\d+)\s*[~\-]\s*(\d+)\s*\)',                       # "(0 ~ 255)" or "(0 - 255)"
                    r'range\s+from\s+(\d+)[a-zA-Z]*\s*[~\-]\s*(\d+)[a-zA-Z]*', # "range from 100MHz ~ 140MHz"
                    r'(\d+)\s*[~\-]\s*(\d+)',                                 # "0 ~ 255" or "0 - 255" (more general)
                ]
                
                for pattern in range_patterns:
                    range_match = re.search(pattern, setting.help_string.lower())
                    if range_match:
                        setting.range_min = int(range_match.group(1))
                        setting.range_max = int(range_match.group(2))
                        break

            # Validate required fields
            if not setting.setup_question or not setting.token:
                return None
            return setting
        except Exception:
            return None  # Skip malformed settings
    
    def _extract_value(self, line):
        """Extract value after = sign, handling various formats"""
        if '=' in line:
            value = line.split('=', 1)[1].strip()
            # Handle angle bracket format: <value>
            if value.startswith('<') and value.endswith('>'):
                return value[1:-1].strip()
            return value
        return ""
    
    def _process_option_line(self, setting, line):
        """Process option line with regex optimization"""
        try:
            is_current = '*' in line
            clean_line = line.replace('*', '', 1) if is_current else line
            
            # Extract option with optimized regex
            match = re.search(r'\[([^\]]+)\]([^//\n]*)', clean_line)
            if match:
                value = match.group(1).strip()
                description = match.group(2).strip()
                setting.options.append((value, description, is_current))
                
                if is_current:
                    setting.current_value = value
        except Exception:
            pass  # Skip malformed options


class LazyLoadTreeview:
    """Optimized Treeview with lazy loading for better performance"""
    
    def __init__(self, parent, settings_callback):
        self.parent = parent
        self.settings_callback = settings_callback
        self.visible_items = {}
        self.item_cache = {}
        
        # Create treeview with virtual scrolling
        self.tree = ttk.Treeview(parent, show='tree')
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # Bind events for lazy loading
        self.tree.bind('<ButtonRelease-1>', self._on_selection)
        self.tree.bind('<Double-1>', self._on_double_click)
        
    def populate(self, categories):
        """Populate tree with category nodes only"""
        self.tree.delete(*self.tree.get_children())
        self.visible_items.clear()
        self.item_cache.clear()
        
        for category, setting_indices in categories.items():
            count = len(setting_indices)
            category_text = f"{category} ({count} settings)"
            category_item = self.tree.insert('', 'end', text=category_text, values=(category,))
            self.visible_items[category_item] = ('category', category, setting_indices)
    
    def _on_selection(self, event):
        """Handle tree selection with lazy loading"""
        selection = self.tree.selection()
        if selection:
            item = selection[0]
            if item in self.visible_items:
                item_type, data, extra = self.visible_items[item]
                if item_type == 'category':
                    self._load_category_children(item, data, extra)
    
    def _load_category_children(self, category_item, category_name, setting_indices):
        """Lazy load category children when expanded"""
        # Check if already loaded
        if self.tree.get_children(category_item):
            return
        
        # Load first 50 settings to avoid UI freeze
        batch_size = 50
        for i, setting_index in enumerate(setting_indices[:batch_size]):
            setting = self.settings_callback(setting_index)
            if setting:
                setting_text = setting.setup_question[:60] + ("..." if len(setting.setup_question) > 60 else "")
                setting_item = self.tree.insert(category_item, 'end', text=setting_text)
                self.visible_items[setting_item] = ('setting', setting, setting_index)
        
        # Add "Load More" if there are more settings
        if len(setting_indices) > batch_size:
            remaining = len(setting_indices) - batch_size
            load_more_item = self.tree.insert(category_item, 'end', text=f"Load {remaining} more settings...")
            self.visible_items[load_more_item] = ('load_more', category_name, setting_indices[batch_size:])
    
    def _on_double_click(self, event):
        """Handle double-click for edit action"""
        selection = self.tree.selection()
        if selection:
            item = selection[0]
            if item in self.visible_items:
                item_type, data, extra = self.visible_items[item]
                if item_type == 'setting':
                    # Trigger setting edit
                    return data


class EnhancedBIOSSettingsGUI:
    def validate_settings_against_original(self):
        """Advanced validation: check all settings against original export for option validity and value range."""
        errors = []
        for setting in self.settings:
            # Validate options
            if getattr(setting, 'original_has_options', False) and setting.options:
                # Collect all possible values
                allowed_values = [str(v) for v, _, _ in setting.options]
                current_value = str(setting.current_value)
                # Check that current_value is in allowed options
                if current_value not in allowed_values:
                    errors.append(f"Setting '{setting.setup_question}': Current value '{current_value}' is not a valid option. Allowed: {allowed_values}")
                # Check that exactly one option is marked as current
                star_count = 0
                for v, _, _ in setting.options:
                    if str(v) == current_value:
                        star_count += 1
                if star_count != 1:
                    errors.append(f"Setting '{setting.setup_question}': Exactly one option must be selected, found {star_count}.")
            # Validate value lines
            elif not getattr(setting, 'original_has_options', False):
                # Use new range fields if available, otherwise fall back to help string parsing
                value = str(setting.current_value)
                if hasattr(setting, 'range_min') and setting.range_min is not None and hasattr(setting, 'range_max') and setting.range_max is not None:
                    try:
                        v = int(value, 0)
                        if not (setting.range_min <= v <= setting.range_max):
                            errors.append(f"Setting '{setting.setup_question}': Value {v} is out of allowed range {setting.range_min}~{setting.range_max}.")
                    except ValueError:
                        errors.append(f"Setting '{setting.setup_question}': Value '{value}' is not a valid integer.")
                elif hasattr(setting, 'help_string') and setting.help_string:
                    # Fallback to old method for settings without parsed range
                    import re
                    m = re.search(r'range[:=]?\s*(\d+)\s*[~\-]\s*(\d+)', setting.help_string)
                    if m:
                        minv, maxv = int(m.group(1)), int(m.group(2))
                        try:
                            v = int(value, 0)
                            if not (minv <= v <= maxv):
                                errors.append(f"Setting '{setting.setup_question}': Value {v} is out of allowed range {minv}~{maxv}.")
                        except Exception:
                            errors.append(f"Setting '{setting.setup_question}': Value '{value}' is not a valid integer.")
        return errors
    
    def scroll_and_highlight_setting(self, setting_index):
        """Scroll to and highlight the widget for the selected setting on the standard view, ensuring it is fully visible."""
        setting = self.settings[setting_index]
        widget = self.setting_widgets.get(setting.token)
        if widget:
            widget.focus_set()
            try:
                self.canvas.update_idletasks()
                # Get widget's position and size relative to the scrollable_frame
                widget_y = widget.winfo_y()
                widget_h = widget.winfo_height()
                frame_h = self.scrollable_frame.winfo_height()
                canvas_h = self.canvas.winfo_height()
                # Get the current top of the visible area in the scrollable_frame
                y0 = int(self.canvas.canvasy(0))
                y1 = y0 + canvas_h
                # If widget is above the visible area, scroll so its top is at the top
                if widget_y < y0:
                    # Scroll so the widget's top is at the top of the visible area
                    scroll_fraction = widget_y / max(1, frame_h - canvas_h)
                    self.canvas.yview_moveto(max(0, min(1, scroll_fraction)))
                # If widget is below the visible area, scroll so its bottom is at the bottom
                elif widget_y + widget_h > y1:
                    # Scroll so the widget's bottom is at the bottom of the visible area
                    scroll_fraction = (widget_y + widget_h - canvas_h) / max(1, frame_h - canvas_h)
                    self.canvas.yview_moveto(max(0, min(1, scroll_fraction)))
                # If widget is partially visible (top above, bottom below), scroll so it's fully visible
                elif widget_y < y0 and widget_y + widget_h > y1:
                    scroll_fraction = widget_y / max(1, frame_h - canvas_h)
                    self.canvas.yview_moveto(max(0, min(1, scroll_fraction)))
                # Otherwise, already fully visible; do nothing
            except Exception:
                # As a fallback, use bbox to scroll the widget into view
                try:
                    bbox = self.canvas.bbox(widget)
                    frame_h = self.scrollable_frame.winfo_height()
                    canvas_h = self.canvas.winfo_height()
                    if bbox and frame_h and canvas_h:
                        self.canvas.yview_moveto(bbox[1] / max(1, frame_h - canvas_h))
                except Exception:
                    pass
            # Optionally, add a temporary highlight effect
            orig_bg = widget.cget('background') if 'background' in widget.keys() else None
            try:
                widget.configure(background='#ffff99')
            except Exception:
                pass
            def remove_highlight():
                try:
                    if orig_bg is not None:
                        widget.configure(background=orig_bg)
                    else:
                        widget.configure(background='white')
                except Exception:
                    pass
            self.root.after(1200, remove_highlight)
    def show_search_results_view(self, matched_settings):
        """Show a dedicated search results popup window with highlighted matches."""
        import re
        # If a previous popup exists, destroy it
        if hasattr(self, '_search_popup') and self._search_popup:
            try:
                self._search_popup.destroy()
            except Exception:
                pass
        self._search_popup = tk.Toplevel(self.root)
        self._search_popup.title("Search Results")
        self._search_popup.geometry("600x650")
        self._search_popup.transient(self.root)
        self._search_popup.grab_set()
        self._search_popup.focus_set()
        # Frame for results
        frame = ttk.Frame(self._search_popup, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text="Search Results", font=("Arial", 14, "bold")).pack(anchor=tk.W, pady=(0, 10))
        if not matched_settings:
            ttk.Label(frame, text="No matches found.", font=("Arial", 12), foreground="gray").pack(pady=40)
            ttk.Button(frame, text="Close", command=self.hide_search_results_view).pack(pady=10)
            return
        # Scrollable results list
        results_canvas = tk.Canvas(frame, borderwidth=0, highlightthickness=0)
        results_scroll = ttk.Scrollbar(frame, orient="vertical", command=results_canvas.yview)
        results_frame = ttk.Frame(results_canvas)
        results_canvas.create_window((0, 0), window=results_frame, anchor="nw")
        results_canvas.configure(yscrollcommand=results_scroll.set, height=350)
        results_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        results_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        def on_frame_configure(event):
            results_canvas.configure(scrollregion=results_canvas.bbox("all"))
        results_frame.bind("<Configure>", on_frame_configure)
        # Highlight all search terms in results
        search_text = self.inline_search_var.get().strip().lower()
        pattern = re.compile(re.escape(search_text), re.IGNORECASE) if search_text else None
        # --- Keyboard navigation state ---
        self._search_result_btns = []
        self._search_result_btn_frames = []
        self._search_result_selected = 0
        def select_result(idx):
            # Remove highlight from all
            for f in self._search_result_btn_frames:
                try:
                    f.config(style="TFrame")
                except Exception:
                    pass
            # Highlight selected
            if 0 <= idx < len(self._search_result_btn_frames):
                try:
                    self._search_result_btn_frames[idx].config(style="Selected.TFrame")
                except Exception:
                    pass
                try:
                    self._search_result_btns[idx].focus_set()
                except Exception:
                    pass
            self._search_result_selected = idx
        style = ttk.Style()
        style.configure("Selected.TFrame", background="#cce6ff")
        for idx, (i, setting) in enumerate(matched_settings):
            sq = str(setting.setup_question)
            if pattern and search_text:
                parts = pattern.split(sq)
                matches = pattern.findall(sq)
                btn_frame = ttk.Frame(results_frame)
                btn_frame.pack(fill=tk.X, padx=10, pady=3)
                for j, part in enumerate(parts):
                    if part:
                        lbl = ttk.Label(btn_frame, text=part, font=("Arial", 10))
                        lbl.pack(side=tk.LEFT, anchor=tk.W)
                    if j < len(matches):
                        hl = ttk.Label(btn_frame, text=matches[j], font=("Arial", 10, "bold"), background="#ffe066")
                        hl.pack(side=tk.LEFT, anchor=tk.W)
                btn = ttk.Button(btn_frame, text="Go", style="Action.TButton", cursor="hand2", command=lambda idx=i: self.on_search_result_selected(idx))
                btn.pack(side=tk.RIGHT, padx=5)
                self._search_result_btns.append(btn)
                self._search_result_btn_frames.append(btn_frame)
            else:
                btn = ttk.Button(results_frame, text=f"{sq[:60]}" + ("..." if len(sq) > 60 else ""),
                                style="Action.TButton", cursor="hand2",
                                command=lambda idx=i: self.on_search_result_selected(idx))
                btn.pack(fill=tk.X, padx=10, pady=3)
                self._search_result_btns.append(btn)
                self._search_result_btn_frames.append(btn)
            # Highlight matches in help_string
            if setting.help_string:
                hs = str(setting.help_string)
                if pattern and search_text:
                    parts = pattern.split(hs)
                    matches = pattern.findall(hs)
                    help_frame = ttk.Frame(results_frame)
                    help_frame.pack(fill=tk.X, padx=30, anchor=tk.W)
                    for j, part in enumerate(parts):
                        if part:
                            lbl = ttk.Label(help_frame, text=part, font=("Arial", 8), foreground="gray")
                            lbl.pack(side=tk.LEFT, anchor=tk.W)
                        if j < len(matches):
                            hl = ttk.Label(help_frame, text=matches[j], font=("Arial", 8, "bold"), background="#ffe066", foreground="gray")
                            hl.pack(side=tk.LEFT, anchor=tk.W)
                else:
                    lbl = ttk.Label(results_frame, text=hs[:80] + ("..." if len(hs) > 80 else ""), font=("Arial", 8), foreground="gray")
                    lbl.pack(fill=tk.X, padx=30, anchor=tk.W)
        # Initial highlight
        if self._search_result_btn_frames:
            select_result(0)
        def on_key(event):
            if not self._search_result_btns:
                return
            idx = self._search_result_selected
            # Handle Enter key to open dropdown or trigger "Go" button
            if event.keysym in ("Return", "KP_Enter"):
                if self._search_result_btns[idx].cget("text") == "Go":
                    self._search_result_btns[idx].invoke()
                else:
                    self._search_result_btns[idx].event_generate("<Button-1>")
                return "break"
            elif event.keysym in ("Down", "Tab"):
                idx = (idx + 1) % len(self._search_result_btns)
                select_result(idx)
                return "break"
            elif event.keysym == "Up":
                idx = (idx - 1) % len(self._search_result_btns)
                select_result(idx)
                return "break"

        # Bind to all widgets in popup so Enter always works, including Entry
        widgets_to_bind = [self._search_popup]
        for child in self._search_popup.winfo_children():
            if isinstance(child, tk.Toplevel):
                widgets_to_bind.append(child)
            try:
                widgets_to_bind.extend(c for c in child.winfo_children() if isinstance(c, tk.Toplevel))
            except Exception:
                pass
        for w in widgets_to_bind:
            w.bind("<KeyPress>", on_key, add="+")
        ttk.Button(frame, text="Cancel Search", command=self.hide_search_results_view).pack(pady=10)

    def hide_search_results_view(self):
        """Hide the search results popup and return to normal view."""
        if hasattr(self, '_search_popup') and self._search_popup:
            try:
                self._search_popup.grab_release()
            except Exception:
                pass
            self._search_popup.destroy()
            self._search_popup = None

    def on_search_result_selected(self, setting_index):
        """Handle user selecting a search result: go to page, highlight field, exit search mode."""
        self.hide_search_results_view()
        page_size = 20
        page = setting_index // page_size
        self.load_page_settings(page)
        self.root.after(150, lambda: self.scroll_and_highlight_setting(setting_index))

    def show_change_review_dialog(self):
        """Show a dialog summarizing all changes before import, with user-friendly option descriptions."""
        changes = []
        for setting in self.settings:
            orig_value = getattr(setting, 'original_value', setting.current_value)
            curr_value = setting.current_value
            if str(orig_value) != str(curr_value):
                # If options exist, show the description instead of just the value
                def get_desc(val):
                    if hasattr(setting, 'options') and setting.options:
                        for v, desc, *_ in setting.options:
                            if str(v) == str(val):
                                return f"{desc} ({val})"
                    return str(val)
                changes.append((setting.setup_question, get_desc(orig_value), get_desc(curr_value)))
        if not changes:
            messagebox.showinfo("No Changes", "No changes to review.")
            return False
        # Build review text
        review_text = "The following changes will be applied to the BIOS:\n\n"
        for q, old, new in changes:
            review_text += f"- {q}\n    Old: {old}\n    New: {new}\n\n"
        return messagebox.askyesno("Review Changes", review_text + "\nProceed with import?")

    def push_undo(self):
        """Push current state to the undo stack and clear redo stack. Limit stack size for memory efficiency."""
        snapshot = [(s.token, s.current_value) for s in self.settings]
        self.undo_stack.append(snapshot)
        # Limit undo stack size
        max_undo = 30
        if len(self.undo_stack) > max_undo:
            self.undo_stack = self.undo_stack[-max_undo:]
        if hasattr(self, 'redo_stack'):
            self.redo_stack.clear()
        else:
            self.redo_stack = []

    def undo(self):
        """Undo the last change and push to redo stack."""
        if not self.undo_stack:
            messagebox.showinfo("Undo", "Nothing to undo.")
            return
            
        # Check for invalid fields first
        if not self.check_can_proceed("undo"):
            return
            
        # Save current state to redo stack
        if not hasattr(self, 'redo_stack'):
            self.redo_stack = []
        current_snapshot = [(s.token, s.current_value) for s in self.settings]
        self.redo_stack.append(current_snapshot)
        # Restore previous state from undo stack
        last_state = self.undo_stack.pop()
        token_to_value = dict(last_state)
        for s in self.settings:
            if s.token in token_to_value:
                s.current_value = token_to_value[s.token]
        self.on_inline_search_changed()

    def redo(self):
        """Redo the last undone change. Limit redo stack size for memory efficiency."""
        if not hasattr(self, 'redo_stack') or not self.redo_stack:
            messagebox.showinfo("Redo", "Nothing to redo.")
            return
            
        # Check for invalid fields first
        if not self.check_can_proceed("redo"):
            return
            
        # Save current state to undo stack
        current_snapshot = [(s.token, s.current_value) for s in self.settings]
        self.undo_stack.append(current_snapshot)
        # Limit undo stack size
        max_undo = 30
        if len(self.undo_stack) > max_undo:
            self.undo_stack = self.undo_stack[-max_undo:]
        # Restore state from redo stack
        redo_state = self.redo_stack.pop()
        token_to_value = dict(redo_state)
        for s in self.settings:
            if s.token in token_to_value:
                s.current_value = token_to_value[s.token]
        self.on_inline_search_changed()

    def restore_last_backup(self):
        """Restore the most recent NVRAM backup from the temp directory."""
        import glob
        temp_dir = os.path.join(tempfile.gettempdir(), "scewin_temp")
        backup_files = sorted(glob.glob(os.path.join(temp_dir, "nvram_backup_*.txt")), reverse=True)
        if not backup_files:
            messagebox.showinfo("Restore Backup", "No backup files found.")
            return
        last_backup = backup_files[0]
        if not messagebox.askyesno("Restore Backup", f"Restore the most recent backup?\n\n{last_backup}\n\nThis will overwrite your current NVRAM file in the temp directory."):
            return
        nvram_txt = os.path.join(temp_dir, "nvram.txt")
        try:
            shutil.copy2(last_backup, nvram_txt)
            self.load_file(nvram_txt)
            messagebox.showinfo("Restore Complete", f"Restored backup:\n{last_backup}")
        except Exception as e:
            messagebox.showerror("Restore Error", f"Failed to restore backup:\n{e}")

    """Main application with integrated import/export and performance optimizations"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("MSI SCEWIN GUI")
        self.root.geometry("1400x750")
        self.root.minsize(1000, 600)
        
        # Initialize components
        self.parser = OptimizedNVRAMParser()
        self.settings = []  # Only store current settings in memory
        self.setting_widgets = {}  # Only keep widgets for currently visible settings
        self.original_file_path = ""
        self.operation_queue = queue.Queue()
        self.current_progress = None
        self.undo_stack = []
        self.redo_stack = []
        
        # Simple navigation state
        self.current_page = 0
        self.page_size = 25
        self.filtered_settings = []  # Current filtered/searched settings
        self.search_active = False
        self.selected_category = ""

        # Track the most recently changed token for highlight in raw view
        self._last_changed_token = None
        
        # Track invalid fields to prevent other operations
        self._invalid_fields = set()

        # Reduce memory footprint by limiting thread stack size (if needed)
        import threading
        try:
            threading.stack_size(2 * 1024 * 1024)  # 2MB per thread, adjust as needed
        except Exception:
            pass

        # Style configuration
        self.setup_styles()
        
        # GUI setup
        self.setup_gui()
        
        # Check admin rights on startup
        self.check_admin_rights()
        
        # Apply performance optimizations
        self.finalize_optimizations()
    
    def setup_styles(self):
        """Configure modern styling"""
        style = ttk.Style()
        
        # Configure treeview for better appearance
        style.configure("Treeview", rowheight=25)
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
        
        # Button styling
        style.configure("Action.TButton", font=("Arial", 9, "bold"))
    
    def check_admin_rights(self):
        """Check and prompt for admin rights and SCEWIN availability"""
        try:
            is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0
        except:
            is_admin = False
        
        # Check SCEWIN availability
        scewin_available, _ = self.check_scewin_availability()
        
        # Set appropriate status message
        if not is_admin and not scewin_available:
            self.status_var.set("⚠️ Admin rights recommended + SCEWIN not found")
        elif not is_admin:
            self.status_var.set("⚠️ Admin rights recommended for BIOS operations")
        elif not scewin_available:
            self.status_var.set("⚠️ SCEWIN not found - Click 'Check SCEWIN Status' for details")
        else:
            self.status_var.set("✅ Ready - SCEWIN available and admin rights detected")
    
    def setup_gui(self):
        """Initialize the enhanced GUI with better layout"""
        # Main container with paned window
        main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Left panel for file operations and navigation
        left_frame = ttk.Frame(main_paned, width=350)
        main_paned.add(left_frame, weight=0)
        
        # Right panel for settings display
        right_frame = ttk.Frame(main_paned)
        main_paned.add(right_frame, weight=1)
        
        self.setup_left_panel(left_frame)
        self.setup_right_panel(right_frame)
        
        # Status bar
        self.setup_status_bar()
    
    def setup_left_panel(self, parent):
        """Setup left panel with file operations, file info, search/filter, and navigation"""
        # File operations section
        file_frame = ttk.LabelFrame(parent, text="File Operations", padding=10)
        file_frame.pack(fill=tk.X, padx=5, pady=5)

        # Export button (primary action)
        export_btn = ttk.Button(file_frame, text="🔽 Export BIOS & Load", 
                               command=self.export_bios_and_load, style="Action.TButton")
        export_btn.pack(fill=tk.X, pady=(0, 5))

        # Load from file button
        load_btn = ttk.Button(file_frame, text="📁 Load From File...", 
                             command=self.load_file_dialog)
        load_btn.pack(fill=tk.X, pady=(0, 5))

        # Save and import button (primary action)
        save_btn = ttk.Button(file_frame, text="🔼 Save & Import to BIOS", 
                             command=self.save_and_import_bios_with_review, style="Action.TButton")
        save_btn.pack(fill=tk.X, pady=(0, 10))

        # Save to file only
        save_file_btn = ttk.Button(file_frame, text="💾 Save to File Only", 
                                  command=self.save_file_only)
        save_file_btn.pack(fill=tk.X)

        # Restore last backup button
        restore_btn = ttk.Button(file_frame, text="🛡️ Restore Last Backup", command=self.restore_last_backup)
        restore_btn.pack(fill=tk.X, pady=(10, 0))

        # SCEWIN status check button
        status_btn = ttk.Button(file_frame, text="🔍 Check SCEWIN Status", command=self.show_scewin_status)
        status_btn.pack(fill=tk.X, pady=(5, 0))

        # Undo/Redo buttons
        undo_btn = ttk.Button(file_frame, text="↩️ Undo", command=self.undo)
        undo_btn.pack(fill=tk.X, pady=(10, 0))
        redo_btn = ttk.Button(file_frame, text="↪️ Redo", command=self.redo)
        redo_btn.pack(fill=tk.X)

        # File info section
        info_frame = ttk.LabelFrame(parent, text="File Information", padding=10)
        info_frame.pack(fill=tk.X, padx=5, pady=5)
        self.file_info_label = ttk.Label(info_frame, text="No file loaded", font=("Arial", 9), foreground="gray")
        self.file_info_label.pack(anchor=tk.W)

        # Search and filter section
        search_frame = ttk.LabelFrame(parent, text="Search & Filter", padding=10)
        search_frame.pack(fill=tk.X, padx=5, pady=5)

        # --- Expandable category menu ---
        self.category_menu_frame = ttk.Frame(search_frame)
        self.category_menu_frame.pack(fill=tk.X, pady=(0, 5))
        self._category_menu_state = {'shown': 3, 'words': []}  # Track how many categories are shown

        # --- Inline search box ---
        self.inline_search_var = tk.StringVar()
        self.inline_search_var.trace_add('write', self.on_inline_search_changed)
        search_entry = ttk.Entry(search_frame, textvariable=self.inline_search_var, font=("Arial", 10), width=28)
        search_entry.pack(fill=tk.X, pady=(0, 5))
        search_entry.insert(0, "Search BIOS settings...")
        def clear_placeholder(event):
            if search_entry.get() == "Search BIOS settings...":
                search_entry.delete(0, tk.END)
        def restore_placeholder(event):
            if not search_entry.get():
                search_entry.insert(0, "Search BIOS settings...")
        search_entry.bind("<FocusIn>", clear_placeholder)
        search_entry.bind("<FocusOut>", restore_placeholder)

        # --- Clear search button ---
        clear_btn = ttk.Button(search_frame, text="Clear", command=lambda: self.inline_search_var.set(""))
        clear_btn.pack(fill=tk.X, pady=(0, 2))

        # --- Fuzzy search toggle ---
        fuzzy_frame = ttk.Frame(search_frame)
        fuzzy_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.fuzzy_search_var = tk.BooleanVar()
        self.fuzzy_search_var.set(False)  # Default to exact search
        fuzzy_check = ttk.Checkbutton(fuzzy_frame, text="🔍 Fuzzy Search (finds similar matches)", 
                                     variable=self.fuzzy_search_var,
                                     command=self.on_fuzzy_search_toggled)
        fuzzy_check.pack(anchor=tk.W)
        
        # Add tooltip-like help text
        help_text = ttk.Label(fuzzy_frame, 
                             text="💡 Fuzzy search finds settings even with typos or partial words", 
                             font=("Arial", 8), foreground="gray")
        help_text.pack(anchor=tk.W, pady=(2, 0))
        
        # Fuzzy search threshold slider
        self.fuzzy_threshold_frame = ttk.Frame(search_frame)
        self.fuzzy_threshold_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Label(self.fuzzy_threshold_frame, text="Match Sensitivity:", font=("Arial", 8)).pack(anchor=tk.W)
        self.fuzzy_threshold_var = tk.DoubleVar()
        self.fuzzy_threshold_var.set(70.0)  # Default threshold
        
        threshold_scale = ttk.Scale(self.fuzzy_threshold_frame, from_=50.0, to=90.0, 
                                   variable=self.fuzzy_threshold_var, orient=tk.HORIZONTAL,
                                   command=self.on_fuzzy_threshold_changed)
        threshold_scale.pack(fill=tk.X, pady=(2, 0))
        
        self.threshold_label = ttk.Label(self.fuzzy_threshold_frame, text="70% (Balanced)", 
                                        font=("Arial", 8), foreground="blue")
        self.threshold_label.pack(anchor=tk.W)
        
        # Initially hide fuzzy controls
        self.fuzzy_threshold_frame.pack_forget()

    def save_and_import_bios_with_review(self):
        """Show change review dialog and validate before importing to BIOS."""
        # Check for invalid fields first
        if not self.check_can_proceed("BIOS import"):
            return
            
        # Save original values for review if not already present
        for s in self.settings:
            if not hasattr(s, 'original_value'):
                s.original_value = s.current_value
        if not self.show_change_review_dialog():
            return
        # Advanced validation step
        errors = self.validate_settings_against_original()
        if errors:
            messagebox.showerror("Validation Error", "The following issues were found:\n\n" + "\n".join(errors))
            return
        self.save_and_import_bios()
    
    def setup_right_panel(self, parent):
        """Setup right panel for settings display and editing"""
        # Settings display frame
        settings_frame = ttk.LabelFrame(parent, text="BIOS Settings", padding=5)
        settings_frame.pack(fill=tk.BOTH, expand=True)

        # Create notebook for different views
        self.notebook = ttk.Notebook(settings_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Settings list tab
        self.setup_settings_tab()

    
    def setup_settings_tab(self):
        """Setup the main settings editing tab"""
        self.settings_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_tab, text="Settings Editor")
        
        # Create scrollable frame for settings
        self.setup_scrollable_settings(self.settings_tab)
    
    def setup_scrollable_settings(self, parent):
        """Setup optimized scrollable frame for settings"""
        # Create canvas and scrollbar
        canvas = tk.Canvas(parent, highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        
        # Configure scrolling
        self.scrollable_frame.bind("<Configure>", 
                                  lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack components
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Mouse wheel support
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
        canvas.bind("<MouseWheel>", _on_mousewheel)
        self.canvas = canvas
        
        # Performance optimization: only show placeholder initially
        placeholder = ttk.Label(self.scrollable_frame, 
                               text="Load a NVRAM file to view settings", 
                               font=("Arial", 12), foreground="gray")
        placeholder.pack(expand=True, fill=tk.BOTH, pady=50)
    
    def setup_status_bar(self):
        """Setup enhanced status bar"""
        status_frame = ttk.Frame(self.root)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        self.status_var = tk.StringVar()
        self.status_var.set("Ready - Load NVRAM file to begin")
        
        status_label = ttk.Label(status_frame, textvariable=self.status_var, 
                                relief=tk.SUNKEN, font=("Arial", 9))
        status_label.pack(fill=tk.X, padx=2, pady=2)
    
    # File operation methods with threading
    def export_bios_and_load(self):
        """Export BIOS settings and load them with progress tracking (user-writable temp dir)"""
        # Search for SCEWIN installation dynamically
        scewin_dir = self._find_scewin_installation()
        if not scewin_dir:
            # Show more helpful error with search locations
            search_paths = self._get_scewin_search_paths()
            error_msg = ("Could not find SCEWIN_64.exe in any MSI Center installation.\n\n"
                        "Searched locations:\n" + "\n".join(f"• {path}" for path in search_paths) + 
                        "\n\nPlease ensure MSI Center is properly installed, or manually copy "
                        "SCEWIN_64.exe and related files to the application directory.")
            messagebox.showerror("SCEWIN Not Found", error_msg)
            return
        temp_dir = os.path.join(tempfile.gettempdir(), "scewin_temp")
        os.makedirs(temp_dir, exist_ok=True)
        files_to_copy = ["SCEWIN_64.exe", "amifldrv64.sys", "amigendrv64.sys"]
        for fname in files_to_copy:
            src = os.path.join(scewin_dir, fname)
            dst = os.path.join(temp_dir, fname)
            if not os.path.isfile(src):
                messagebox.showerror("Missing File", f"Required file not found: {src}")
                return
            shutil.copy2(src, dst)
        scewin_exe = os.path.join(temp_dir, "SCEWIN_64.exe")
        nvram_txt = os.path.join(temp_dir, "nvram.txt")
        log_file = os.path.join(temp_dir, "log-file.txt")
        def export_worker():
            progress = None
            try:
                progress = ProgressDialog(self.root, "Exporting BIOS", "Exporting BIOS settings to nvram.txt...")
                progress.update_progress(10, "Cleaning up old files...")
                for f in [nvram_txt, log_file]:
                    if os.path.isfile(f):
                        try:
                            os.remove(f)
                        except Exception:
                            pass
                progress.update_progress(30, "Running SCEWIN_64.exe export...")
                cmd = [scewin_exe, "/o", "/s", "nvram.txt"]
                result = subprocess.run(cmd, cwd=temp_dir, capture_output=True, text=True)
                log_content = ""
                if os.path.isfile(log_file):
                    with open(log_file, "r", encoding="utf-8", errors="ignore") as f:
                        log_content = f.read()
                if result.returncode != 0:
                    if progress: progress.close()
                    self.root.after(0, lambda: messagebox.showerror(
                        "Export Failed", f"STDOUT:\n{result.stdout}\n\nSTDERR:\n{result.stderr}\n\nLog:\n{log_content}"))
                    return
                if not os.path.isfile(nvram_txt):
                    if progress: progress.close()
                    self.root.after(0, lambda: messagebox.showerror(
                        "Export Failed", "nvram.txt was not created. See log for details:\n" + log_content))
                    return
                progress.update_progress(90, "Loading exported file...")
                # Load the file directly - no popup needed since it auto-loads
                self.root.after(0, lambda: self.load_file(nvram_txt, progress))
            except Exception as e:
                if progress: progress.close()
                self.root.after(0, lambda: messagebox.showerror("Export Error", str(e)))
        thread = threading.Thread(target=export_worker)
        thread.daemon = True
        thread.start()
    
    def save_and_import_bios(self):
        """Save settings and import to BIOS with progress tracking (user-writable temp dir)"""
        if not self.settings:
            messagebox.showwarning("No Settings", "No settings loaded to save.")
            return
        if not messagebox.askyesno(
            "Confirm BIOS Import",
            "⚠️ WARNING: This will modify your BIOS settings.\n\n"
            "This operation can be risky. Make sure you understand "
            "the changes you're making.\n\n"
            "Do you want to proceed?"
        ):
            return
        
        # Ensure SCEWIN files are available in temp directory
        temp_dir = os.path.join(tempfile.gettempdir(), "scewin_temp")
        os.makedirs(temp_dir, exist_ok=True)
        scewin_exe = os.path.join(temp_dir, "SCEWIN_64.exe")
        
        # Check if SCEWIN files are already in temp directory, if not, copy them
        if not os.path.isfile(scewin_exe):
            scewin_dir = self._find_scewin_installation()
            if not scewin_dir:
                messagebox.showerror("SCEWIN Not Found", 
                                   "SCEWIN files not found. Please run 'Export BIOS & Load' first "
                                   "to copy the required files, or ensure MSI Center is properly installed.")
                return
            
            # Copy required files to temp directory
            files_to_copy = ["SCEWIN_64.exe", "amifldrv64.sys", "amigendrv64.sys"]
            for fname in files_to_copy:
                src = os.path.join(scewin_dir, fname)
                dst = os.path.join(temp_dir, fname)
                if not os.path.isfile(src):
                    messagebox.showerror("Missing File", f"Required file not found: {src}")
                    return
                shutil.copy2(src, dst)
        nvram_txt = os.path.join(temp_dir, "nvram.txt")
        import_file = os.path.join(temp_dir, "nvram_import.txt")
        log_file = os.path.join(temp_dir, "log-file.txt")
        def import_worker():
            progress = None
            try:
                progress = ProgressDialog(self.root, "Importing to BIOS", "Preparing settings for import...")
                progress.update_progress(10, "Generating NVRAM file...")
                # Generate import file from current settings
                if not self.generate_nvram_file(import_file, lambda p, s: progress.update_progress(10 + p * 0.4, s)):
                    if progress: progress.close()
                    return
                progress.update_progress(50, "Creating backup of exported NVRAM...")
                backup_file = os.path.join(temp_dir, f"nvram_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
                if os.path.exists(nvram_txt):
                    shutil.copy2(nvram_txt, backup_file)
                progress.update_progress(70, "Importing to BIOS with SCEWIN_64.exe...")
                cmd = [scewin_exe, "/i", "/s", "nvram_import.txt"]
                result = subprocess.run(cmd, cwd=temp_dir, capture_output=True, text=True)
                log_content = ""
                if os.path.isfile(log_file):
                    with open(log_file, "r", encoding="utf-8", errors="ignore") as f:
                        log_content = f.read()
                if progress: progress.close()
                if result.returncode == 0:
                    def show_success_with_restart():
                        win = tk.Toplevel(self.root)
                        win.title("Import Successful")
                        win.geometry("500x350")
                        win.transient(self.root)
                        win.grab_set()
                        msg = ("Settings imported successfully!\n\nA system reboot is recommended for changes to take effect.\n\nLog:\n" + log_content)
                        label = tk.Text(win, wrap=tk.WORD, height=12, width=60)
                        label.insert(tk.END, msg)
                        label.config(state=tk.DISABLED)
                        label.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
                        def do_restart():
                            if messagebox.askyesno("Restart Computer", "Are you sure you want to restart now?"):
                                import subprocess
                                subprocess.Popen(["shutdown", "/r", "/t", "0"])
                                win.destroy()
                        def do_reboot_uefi():
                            if messagebox.askyesno("Reboot to UEFI", "Are you sure you want to reboot to UEFI firmware settings now?"):
                                import subprocess
                                subprocess.Popen(["shutdown", "/r", "/fw", "/t", "0"])
                                win.destroy()
                        btn_frame = ttk.Frame(win)
                        btn_frame.pack(pady=10)
                        ttk.Button(btn_frame, text="Restart Now", command=do_restart).pack(side=tk.LEFT, padx=10)
                        ttk.Button(btn_frame, text="Reboot to UEFI", command=do_reboot_uefi).pack(side=tk.LEFT, padx=10)
                        ttk.Button(btn_frame, text="Close", command=win.destroy).pack(side=tk.LEFT, padx=10)
                    self.root.after(0, show_success_with_restart)
                    self.status_var.set("✅ Settings imported successfully - Reboot recommended")                  
                else:
                    self.root.after(0, lambda: messagebox.showerror(
                        "Import Failed", f"Failed to import settings.\n\nSTDOUT:\n{result.stdout}\n\nSTDERR:\n{result.stderr}\n\nLog:\n{log_content}"))
            except Exception as e:
                if progress: progress.close()
                self.root.after(0, lambda: messagebox.showerror("Import Error", str(e)))
        thread = threading.Thread(target=import_worker)
        thread.daemon = True
        thread.start()
    
    def load_file_dialog(self):
        """Load NVRAM file with dialog"""
        file_path = filedialog.askopenfilename(
            title="Select NVRAM file",
            filetypes=[
                ("Text files", "*.txt"),
                ("NVRAM files", "*.nvram"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self.load_file(file_path)
    
    def load_file(self, file_path, existing_progress=None):
        """Load NVRAM file with optimized parsing and progress tracking"""
        def load_worker():
            progress = None  # Ensure progress is always defined
            try:
                # Create or use existing progress dialog
                if existing_progress:
                    progress = existing_progress
                    progress.update_progress(85, "Parsing NVRAM file...")
                else:
                    progress = ProgressDialog(self.root, "Loading File", 
                                            "Loading and parsing NVRAM file...")
                
                # Parse file with progress tracking
                cancel_flag = lambda: getattr(progress, 'cancelled', False)
                
                def progress_callback(value, status=""):
                    if not cancel_flag():
                        base_progress = 85 if existing_progress else 0
                        final_progress = base_progress + (value * (15 if existing_progress else 95) / 100)
                        progress.update_progress(final_progress, status)
                
                self.settings = self.parser.parse_file(file_path, progress_callback, cancel_flag)
                
                if cancel_flag():
                    return
                
                self.original_file_path = file_path
    
                # Reset setting widgets to avoid stale references
                self.setting_widgets = {}
    
                # Update UI on main thread
                self.root.after(0, lambda: self.finalize_load(file_path, progress))
                
            except Exception as e:
                if existing_progress:
                    existing_progress.close()
                elif progress is not None:
                    progress.close()
                self.root.after(0, lambda: messagebox.showerror("Load Error", 
                                f"Failed to load file: {str(e)}"))
        
        # Run in separate thread
        thread = threading.Thread(target=load_worker)
        thread.daemon = True
        thread.start()
    
    def finalize_load(self, file_path, progress):
        """Finalize file loading on main thread"""
        try:
            progress.update_progress(100, "Updating interface...")
            # Update file info
            filename = os.path.basename(file_path)
            self.file_info_label.config(text=f"📄 {filename}\n🔢 {len(self.settings)} settings loaded")
            # Remove page navigation, just update category menu
            self.update_category_menu()
            self.status_var.set(f"✅ Loaded {len(self.settings)} settings from {filename}")
            # Show first page of settings
            self.current_page = 0
            self.search_active = False
            self.filtered_settings = []
            self.display_current_page()
            progress.close()
        except Exception as e:
            progress.close()
            messagebox.showerror("Interface Error", f"Failed to update interface: {str(e)}")

    def update_category_menu(self):
        """Create simple category filter using top keywords"""
        if not hasattr(self, 'category_menu_frame') or not self.settings:
            return
        for widget in self.category_menu_frame.winfo_children():
            widget.destroy()

        # Extract common keywords from settings
        from collections import Counter
        word_counter = Counter()
        for s in self.settings:
            # Get words from question and help text
            text = f"{s.setup_question} {s.help_string}".lower()
            words = re.findall(r'\b[a-z]{3,}\b', text)  # 3+ letter words only
            for word in words:
                if word not in {'the', 'and', 'for', 'with', 'this', 'that', 'setting', 'option', 'value', 'default', 'enable', 'disable'}:
                    word_counter[word] += 1
        
        # Get top 15 categories
        top_categories = [word.title() for word, count in word_counter.most_common(15) if count >= 2]
        
        # Category selection
        self.category_var = tk.StringVar()
        category_frame = ttk.Frame(self.category_menu_frame)
        category_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(category_frame, text="Category:", font=("Arial", 9, "bold")).pack(side=tk.LEFT)
        category_combo = ttk.Combobox(category_frame, textvariable=self.category_var, 
                                     values=["All"] + top_categories, state="readonly", width=20)
        category_combo.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        category_combo.set("All")
        category_combo.bind("<<ComboboxSelected>>", self.on_category_filter_changed)

    def on_category_filter_changed(self, event=None):
        """Apply category filter"""
        category = self.category_var.get()
        if category == "All":
            self.selected_category = ""
        else:
            self.selected_category = category.lower()
        self.apply_search_and_filters()

    def on_inline_search_changed(self, *args):
        """Handle search text changes with debouncing"""
        # Cancel any pending search
        if hasattr(self, '_search_timer'):
            self.root.after_cancel(self._search_timer)
        
        # Schedule new search after short delay
        self._search_timer = self.root.after(300, self.apply_search_and_filters)

    def apply_search_and_filters(self):
        """Apply search and category filters, then display results"""
        search_text = self.inline_search_var.get().strip()
        
        # Clear placeholder text
        if search_text == "Search BIOS settings...":
            search_text = ""
        
        # Start with all settings
        if not search_text and not self.selected_category:
            # No filters - show all settings
            self.search_active = False
            self.filtered_settings = []
            self.current_page = 0
            self.display_current_page()
            return
        
        # Apply filters
        self.search_active = True
        
        if search_text and self.fuzzy_search_var.get():
            # Use fuzzy search with scoring for better ranking
            self.filtered_settings = self._get_fuzzy_search_results_with_scores(search_text)
        else:
            # Use exact search or category-only filtering
            self.filtered_settings = []
            
            for i, setting in enumerate(self.settings):
                matches = True
                
                # Category filter
                if self.selected_category:
                    setting_text = f"{setting.setup_question} {setting.help_string}".lower()
                    if self.selected_category not in setting_text:
                        matches = False
                
                # Search filter
                if matches and search_text:
                    search_lower = search_text.lower()
                    searchable_text = f"{setting.setup_question} {setting.help_string} {setting.token}".lower()
                    if search_lower not in searchable_text:
                        matches = False
                
                if matches:
                    self.filtered_settings.append((i, setting))
        
        # Reset to first page and display
        self.current_page = 0
        self.display_current_page()
        
        # Update status
        if self.search_active:
            search_mode = "fuzzy" if (search_text and self.fuzzy_search_var.get()) else "exact"
            threshold_info = f" ({int(self.fuzzy_threshold_var.get())}% threshold)" if search_mode == "fuzzy" else ""
            self.status_var.set(f"🔍 Found {len(self.filtered_settings)} matching settings ({search_mode} search{threshold_info})")
        else:
            filename = os.path.basename(getattr(self, 'original_file_path', 'Unknown'))
            self.status_var.set(f"✅ Loaded {len(self.settings)} settings from {filename}")

    def on_fuzzy_search_toggled(self):
        """Toggle visibility of fuzzy search threshold controls"""
        if self.fuzzy_search_var.get():
            # Show fuzzy threshold controls
            self.fuzzy_threshold_frame.pack(fill=tk.X, pady=(5, 0))
        else:
            # Hide fuzzy threshold controls
            self.fuzzy_threshold_frame.pack_forget()

        # Update search results if search is active
        search_text = self.inline_search_var.get().strip()
        if search_text and search_text != "Search BIOS settings...":
            self.apply_search_and_filters()

    def on_fuzzy_threshold_changed(self, value):
        """Update threshold label and refresh search results if needed"""
        threshold = float(value)

        # Determine description based on threshold range
        if threshold < 60:
            description = "Very Loose"
        elif threshold < 70:
            description = "Loose"
        elif threshold < 80:
            description = "Balanced"
        elif threshold < 85:
            description = "Strict"
        else:
            description = "Very Strict"

        # Update label
        self.threshold_label.config(text=f"{int(threshold)}% ({description})")

        # Update search results if search is active
        search_text = self.inline_search_var.get().strip()
        if search_text and search_text != "Search BIOS settings..." and self.fuzzy_search_var.get():
            # Use debouncing to avoid too many updates
            if hasattr(self, '_threshold_timer'):
                self.root.after_cancel(self._threshold_timer)
            self._threshold_timer = self.root.after(300, self.apply_search_and_filters)

    def _get_fuzzy_search_results_with_scores(self, search_text):
        """Perform fuzzy search on settings with RapidFuzz and return scored results"""
        results = []
        threshold = self.fuzzy_threshold_var.get()
        search_lower = search_text.lower()

        for i, setting in enumerate(self.settings):
            # Skip if category filter doesn't match
            if self.selected_category:
                setting_text = f"{setting.setup_question} {setting.help_string}".lower()
                if self.selected_category not in setting_text:
                    continue

            # Prepare searchable text
            searchable_text = f"{setting.setup_question} {setting.help_string} {setting.token}".lower()

            # Calculate fuzzy match score
            score = fuzz.partial_ratio(search_lower, searchable_text)

            # Add to results if score meets threshold
            if score >= threshold:
                results.append((i, setting, score))

        # Sort by score (highest first)
        results.sort(key=lambda x: x[2], reverse=True)

        # Return in the format expected by the rest of the code (index, setting)
        return [(i, setting) for i, setting, _ in results]

    def add_pagination_controls(self, total_settings):
        """Add simple, clear pagination controls"""
        total_pages = (total_settings + self.page_size - 1) // self.page_size
        
        if total_pages <= 1:
            return  # No pagination needed
            
        # Pagination frame
        pagination_frame = ttk.Frame(self.scrollable_frame)
        pagination_frame.pack(fill=tk.X, pady=20)
        
        # Center the pagination controls
        center_frame = ttk.Frame(pagination_frame)
        center_frame.pack(anchor=tk.CENTER)
        
        # Previous button
        if self.current_page > 0:
            prev_btn = ttk.Button(center_frame, text="← Previous", 
                                 command=lambda: self.go_to_page(self.current_page - 1))
            prev_btn.pack(side=tk.LEFT, padx=5)
        
        # Page info
        page_label = ttk.Label(center_frame, 
                              text=f"Page {self.current_page + 1} of {total_pages} ({total_settings} settings)",
                              font=("Arial", 10))
        page_label.pack(side=tk.LEFT, padx=15)
        
        # Next button
        if self.current_page < total_pages - 1:
            next_btn = ttk.Button(center_frame, text="Next →", 
                                 command=lambda: self.go_to_page(self.current_page + 1))
            next_btn.pack(side=tk.LEFT, padx=5)
        
        # Quick jump to first/last for large result sets
        if total_pages > 5:
            ttk.Separator(center_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, padx=10, fill=tk.Y)
            
            if self.current_page > 2:
                first_btn = ttk.Button(center_frame, text="First", 
                                      command=lambda: self.go_to_page(0))
                first_btn.pack(side=tk.LEFT, padx=5)
            
            if self.current_page < total_pages - 3:
                last_btn = ttk.Button(center_frame, text="Last", 
                                     command=lambda: self.go_to_page(total_pages - 1))
                last_btn.pack(side=tk.LEFT, padx=5)

    def go_to_page(self, page_num):
        """Navigate to a specific page"""
        self.current_page = page_num
        self.display_current_page()
        # Scroll to top of page
        self.canvas.yview_moveto(0)

    def check_can_proceed(self, operation_name="operation"):
        """Check if operations can proceed (no invalid fields)"""
        if self.has_invalid_fields():
            messagebox.showwarning("Invalid Values", 
                                 f"Cannot proceed with {operation_name}.\n\n"
                                 "Please correct all invalid field values first.\n"
                                 "Look for fields highlighted in red.")
            return False
        return True

    def has_invalid_fields(self):
        """Check if there are any fields with invalid values"""
        invalid_count = len(getattr(self, '_invalid_fields', set()))
        if invalid_count > 0:
            # Update status bar to show invalid field count
            self.status_var.set(f"⚠️ {invalid_count} invalid field(s) - Please correct before saving")
        return invalid_count > 0

    def save_file_only(self):
        """Save to file without importing to BIOS"""
        if not self.settings:
            messagebox.showwarning("No Settings", "No settings loaded to save.")
            return
            
        # Check for invalid fields first
        if not self.check_can_proceed("file save"):
            return
            
        file_path = filedialog.asksaveasfilename(
            title="Save NVRAM file",
            defaultextension=".txt",
            filetypes=[
                ("Text files", "*.txt"),
                ("NVRAM files", "*.nvram"),
                ("All files", "*.*")
            ]
        )
        if file_path:
            def save_worker():
                progress = ProgressDialog(self.root, "Saving File", 
                                        "Generating NVRAM file...")
                def progress_callback(value, status=""):
                    progress.update_progress(value, status)
                try:
                    self.generate_nvram_file(file_path, progress_callback)
                    progress.close()
                except Exception as e:
                    progress.close()
                    messagebox.showerror("File Generation Error", f"Failed to generate NVRAM file: {str(e)}")
            thread = threading.Thread(target=save_worker)
            thread.daemon = True
            thread.start()

    def generate_nvram_file(self, file_path, progress_callback=None):
        """Generate NVRAM file from current settings, preserving original block and only updating * or value as needed."""
        try:
            with open(file_path, 'w', encoding='utf-8') as file:
                # Write header
                if progress_callback:
                    progress_callback(10, "Writing header...")
                if self.parser.raw_header:
                    file.write(self.parser.raw_header)
                else:
                    # Default header
                    file.write("// Script File Name : nvram_import.txt\n")
                    file.write(f"// Created on {datetime.now().strftime('%a %b %d %H:%M:%S %Y')}\n")
                    file.write("// AMISCE Utility. Ver 5.05.01.0002\n")
                    file.write("// Copyright (c) 2021 AMI. All rights reserved.\n")
                    file.write(f"HIICrc32= {self.parser.header_info.get('crc32', '67B9B44E')}\n\n")
                # Write settings
                total_settings = len(self.settings)
                for i, setting in enumerate(self.settings):
                    if progress_callback and i % 50 == 0:
                        progress_callback(10 + int((i / total_settings) * 80), f"Writing setting {i+1} of {total_settings}...")
                    # Copy original block and only update * or value as needed
                    block_lines = list(setting.original_block_lines)
                    # Update * marker for options
                    if getattr(setting, 'original_has_options', False) and setting.options:
                        new_lines = []
                        for line in block_lines:
                            updated_line = line
                            # Remove existing * first
                            if '*' in line:
                                updated_line = line.replace('*', '', 1)
                            # Check if this line matches current selection
                            for value, desc, is_current in setting.options:
                                line_lower = updated_line.lower()
                                if f'[{value}]' in line_lower or (desc and desc.lower() in line_lower):
                                    if str(value) == str(setting.current_value):
                                        updated_line = '*' + updated_line
                                    break
                            new_lines.append(updated_line)
                        block_lines = new_lines
                    # Update value for value lines
                    elif not getattr(setting, 'original_has_options', False):
                        new_lines = []
                        for line in block_lines:
                            if line.strip().startswith('Value'):
                                new_lines.append(f"Value         = <{setting.current_value}>")
                            else:
                                new_lines.append(line)
                        block_lines = new_lines
                    # Write the block
                    for line in block_lines:
                        file.write(line.rstrip() + '\n')
                if progress_callback:
                    progress_callback(100, "Done writing NVRAM file.")
                return True
        except Exception as e:
            if progress_callback:
                progress_callback(0, f"Failed to generate NVRAM file: {str(e)}")
            messagebox.showerror("File Generation Error", 
                               f"Failed to generate NVRAM file: {str(e)}")
            return False
    
    def populate_navigation(self):
        pass  # No longer needed for infinite scroll
        
    def display_current_page(self):
        """Display the current page of settings"""
        # Clear current display
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.setting_widgets.clear()
        
        # Get settings to display
        if self.search_active and self.filtered_settings:
            settings_to_show = self.filtered_settings
        else:
            settings_to_show = [(i, setting) for i, setting in enumerate(self.settings)]
        
        # Calculate page boundaries
        start_idx = self.current_page * self.page_size
        end_idx = min(start_idx + self.page_size, len(settings_to_show))
        
        if not settings_to_show:
            # Show "no results" message
            no_results_label = ttk.Label(
                self.scrollable_frame,                text="No settings found matching your search criteria.",
                font=("Arial", 12), 
                foreground="gray"
            )
            no_results_label.pack(expand=True, pady=50)
            return
        
        # Display settings for current page
        for i in range(start_idx, end_idx):
            setting_idx, setting = settings_to_show[i]
            self.create_setting_widget(setting, setting_idx)
        
        # Add pagination controls at the bottom
        self.add_pagination_controls(len(settings_to_show))
        
        # Add quick overview at top if filtering/searching
        if self.search_active:
            self.add_search_overview()

    def add_search_overview(self):
        """Add overview info for current search/filter results"""
        overview_frame = ttk.Frame(self.scrollable_frame)
        overview_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Create overview text
        total_filtered = len(self.filtered_settings)
        search_text = self.inline_search_var.get().strip()
        category_text = self.selected_category
        
        overview_parts = []
        if search_text and search_text != "Search BIOS settings...":
            overview_parts.append(f"Search: '{search_text}'")
        if category_text:
            overview_parts.append(f"Category: {category_text.title()}")
        
        if overview_parts:
            filter_text = " | ".join(overview_parts)
            overview_label = ttk.Label(overview_frame, 
                                     text=f"🔍 Filtering by: {filter_text} | {total_filtered} results",
                                     font=("Arial", 9, "italic"), 
                                     foreground="blue")
            overview_label.pack(side=tk.LEFT)
            
            # Clear filters button
            clear_btn = ttk.Button(overview_frame, text="Clear Filters", 
                                  command=self.clear_all_filters)
            clear_btn.pack(side=tk.RIGHT)

    def clear_all_filters(self):
        """Clear all search and category filters"""
        self.inline_search_var.set("")
        if hasattr(self, 'category_var'):
            self.category_var.set("All")
        self.selected_category = ""
        self.search_active = False
        self.filtered_settings = []
        self.current_page = 0
        self.display_current_page()

    def on_category_changed(self, event=None):
        """Category filter removed: do nothing."""
        pass
    
    def load_category_settings(self, category):
        """Category navigation removed: do nothing."""
        pass

    def load_page_settings(self, page):
        """Navigate to a specific page - replaced by go_to_page"""
        self.go_to_page(page)
    
    def create_setting_widget(self, setting, index):
        """Create optimized widget for individual setting (memory and UI optimized)"""
        # Use the new optimized widget creation method
        return self._create_widget_with_validation(setting, self.scrollable_frame)
    
    def create_setting_widget_legacy(self, setting, index):
        """Legacy widget creation method - kept for compatibility"""
        # Ignore/commented-out settings (leading //): do not allow configuration
        if str(setting.setup_question).strip().startswith("//") or str(setting.token).strip().startswith("//"):
            # Show as grayed-out, not editable
            setting_frame = ttk.LabelFrame(
                self.scrollable_frame, 
                text="[IGNORED] " + setting.setup_question[:80] + ("..." if len(setting.setup_question) > 80 else ""), 
                padding=8
            )
            setting_frame.pack(fill=tk.X, padx=5, pady=3)
            
            # Help string with proper wrapping
            if setting.help_string and setting.help_string.strip():
                help_label = tk.Text(setting_frame, wrap=tk.WORD, height=3, font=("Arial", 8), 
                                   foreground="gray", background="SystemButtonFace", 
                                   borderwidth=0, highlightthickness=0, state=tk.DISABLED)
                help_label.insert(tk.END, setting.help_string)
                help_label.pack(fill=tk.X, pady=(0, 5))
            
            # Only show default value, no token/offset
            if setting.bios_default:
                default_info = f"Default: {setting.bios_default}"
                ttk.Label(setting_frame, text=default_info, font=("Arial", 8), foreground="gray").pack(anchor=tk.W, pady=(0, 5))
            
            ttk.Label(setting_frame, text="This setting is ignored (commented out in BIOS file)", 
                     font=("Arial", 9, "italic"), foreground="gray").pack(anchor=tk.W, pady=(0, 5))
            return

        # Use a lightweight frame for each setting
        setting_frame = ttk.LabelFrame(
            self.scrollable_frame, 
            text=setting.setup_question[:80] + ("..." if len(setting.setup_question) > 80 else ""), 
            padding=8
        )
        setting_frame.pack(fill=tk.X, padx=5, pady=3)

        # Help string with proper wrapping - use Text widget for multiline display
        if setting.help_string and setting.help_string.strip():
            # Calculate height based on text length (rough estimate)
            text_lines = len(setting.help_string) // 80 + setting.help_string.count('\n') + 1
            height = min(max(2, text_lines), 6)  # Between 2 and 6 lines
            
            help_text = tk.Text(setting_frame, wrap=tk.WORD, height=height, font=("Arial", 8), 
                               foreground="gray", background="SystemButtonFace", 
                               borderwidth=0, highlightthickness=0, state=tk.NORMAL, cursor="arrow")
            help_text.insert(tk.END, setting.help_string)
            help_text.config(state=tk.DISABLED)  # Make read-only
            help_text.pack(fill=tk.X, pady=(0, 5))

        # Show default info only (no token/offset or range - range shown next to input)
        info_parts = []
        if setting.bios_default:
            info_parts.append(f"Default: {setting.bios_default}")
        
        if info_parts:
            info_text = " | ".join(info_parts)
            ttk.Label(setting_frame, text=info_text, font=("Arial", 8), foreground="darkblue").pack(anchor=tk.W, pady=(0, 5))

        # Value input section
        value_frame = ttk.Frame(setting_frame)
        value_frame.pack(fill=tk.X, pady=(5, 0))

        # Use a local function to minimize closure memory
        def push_and_refresh(s, val):
            self.push_undo()
            s.current_value = val
            # Ensure only one option is marked as current
            if s.options:
                new_options = []
                for value, desc, is_current in s.options:
                    new_is_current = (str(value) == str(val))
                    new_options.append((value, desc, new_is_current))
                s.options = new_options
            # Track the most recently changed token for highlight
            self._last_changed_token = s.token

        # Validation function for numeric inputs with range checking
        def validate_numeric_input(s, val, entry_widget):
            """Validate numeric input against range constraints and enforce correction"""
            field_id = id(entry_widget)
            
            if not val.strip():
                # Allow empty values - reset style and remove from invalid set
                entry_widget.config(style="TEntry")
                self._invalid_fields.discard(field_id)
                self.clear_invalid_field_status()
                return True
                
            try:
                num_val = int(val, 0)  # Support hex (0x) and decimal
                if hasattr(s, 'range_min') and s.range_min is not None and hasattr(s, 'range_max') and s.range_max is not None:
                    if not (s.range_min <= num_val <= s.range_max):
                        # Mark field as invalid
                        self._invalid_fields.add(field_id)
                        self.update_invalid_field_status()
                        # Show error and force correction
                        messagebox.showerror("Invalid Value", 
                                           f"Value {num_val} is out of range.\nAllowed range: {s.range_min} ~ {s.range_max}\n\nPlease enter a valid value before making other changes.")
                        # Reset to previous valid value and keep focus
                        entry_widget.delete(0, tk.END)
                        entry_widget.insert(0, str(s.current_value))
                        entry_widget.focus_set()
                        entry_widget.selection_range(0, tk.END)
                        # Briefly highlight the field in red
                        entry_widget.config(background='#ffcccc')
                        self.root.after(2000, lambda: entry_widget.config(background='white'))
                        return False
                # Valid value - reset style and remove from invalid set
                entry_widget.config(style="TEntry")
                self._invalid_fields.discard(field_id)
                self.clear_invalid_field_status()
                return True
            except ValueError:
                # Mark field as invalid
                self._invalid_fields.add(field_id)
                self.update_invalid_field_status()
                # Show error and force correction
                messagebox.showerror("Invalid Value", f"'{val}' is not a valid number.\n\nPlease enter a valid numeric value before making other changes.")
                # Reset to previous valid value and keep focus
                entry_widget.delete(0, tk.END)
                entry_widget.insert(0, str(s.current_value))
                entry_widget.focus_set()
                entry_widget.selection_range(0, tk.END)
                # Briefly highlight the field in red
                entry_widget.config(background='#ffcccc')
                self.root.after(2000, lambda: entry_widget.config(background='white'))
                return False

        # --- Dropdown logic for numeric fields with options ---
        # If the setting has options (even if is_numeric), always use a dropdown if there are 2-10 options
        use_dropdown = False
        if setting.options and 2 <= len(setting.options) <= 10:
            use_dropdown = True
        # If the setting is numeric but has a small set of valid values, use dropdown
        if setting.is_numeric and setting.options and 2 <= len(setting.options) <= 10:
            use_dropdown = True

        if use_dropdown:
            # Use a dropdown (Combobox) for small option sets
            values = [f"{value} : {desc[:50]}" + ("..." if len(desc) > 50 else "") for value, desc, _ in setting.options]
            combo = ttk.Combobox(value_frame, values=values, state="readonly", width=60)
            # Set current value
            current_value = str(setting.current_value)
            found = False
            for i, (value, desc, _) in enumerate(setting.options):
                if str(value) == current_value:
                    combo.current(i)
                    found = True
                    break
            if not found:
                # fallback: set to current_value as string
                combo.set(current_value)
            combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
            self.setting_widgets[setting.token] = combo

            # Create option values list for event handler
            option_vals = [value for value, _, _ in setting.options]

            # Optimized event handler with proper scope
            def on_combo_select_legacy(event=None):
                try:
                    idx = combo.current()
                    if 0 <= idx < len(option_vals):
                        old_value = setting.current_value
                        setting.current_value = option_vals[idx]
                        
                        # Use the legacy push_and_refresh function if available
                        if 'push_and_refresh' in locals():
                            push_and_refresh(setting, option_vals[idx])
                        else:
                            # Fallback to undo system
                            self.push_undo()
                except Exception:
                    pass
            
            combo.bind("<<ComboboxSelected>>", on_combo_select_legacy)
        elif setting.is_numeric:
            # Check if this should be read-only (hex values without clear ranges/options)
            current_val = str(setting.current_value).strip()
            is_readonly_hex = False
            
            # Detect read-only hex values: all hex digits, no range, no options, no clear guidance
            if (len(current_val) >= 4 and 
                all(c in '0123456789ABCDEFabcdef' for c in current_val) and
                not hasattr(setting, 'range_min') and
                not setting.options and
                not setting.help_string):
                is_readonly_hex = True
            
            # Also check for values that are clearly system-managed (all same character)
            if len(current_val) >= 4 and len(set(current_val.upper())) == 1 and current_val.upper()[0] in 'F0':
                is_readonly_hex = True
                
            if is_readonly_hex:
                # Create read-only field for hex values that shouldn't be changed
                readonly_frame = ttk.Frame(value_frame)
                readonly_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
                
                readonly_entry = ttk.Entry(readonly_frame, width=30, state="readonly")
                readonly_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
                readonly_entry.config(state="normal")
                readonly_entry.insert(0, current_val)
                readonly_entry.config(state="readonly")
                self.setting_widgets[setting.token] = readonly_entry
                
                ttk.Label(value_frame, text="(Read-only hex value)", font=("Arial", 8), foreground="gray").pack(side=tk.LEFT, padx=(5, 0))
            else:
                # Only use free-entry numeric field if no options or too many options
                entry = ttk.Entry(value_frame, width=30)
                entry.insert(0, str(setting.current_value))
                entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
                self.setting_widgets[setting.token] = entry
                
                # Show range info in label if available
                range_text = "Numeric"
                if hasattr(setting, 'range_min') and setting.range_min is not None and hasattr(setting, 'range_max') and setting.range_max is not None:
                    range_text = f"Range: {setting.range_min}-{setting.range_max}"
                ttk.Label(value_frame, text=f"({range_text})", font=("Arial", 8), foreground="blue").pack(side=tk.LEFT, padx=(5, 0))
                
                def on_entry_change(event=None, s=setting, e=entry):
                    val = e.get().strip()
                    if not val:
                        # Allow empty values
                        push_and_refresh(s, val)
                    elif validate_numeric_input(s, val, e):
                        # Only update if validation passes
                        push_and_refresh(s, val)
                    # If validation fails, the validate_numeric_input function handles resetting the field
                entry.bind('<FocusOut>', on_entry_change)
                entry.bind('<Return>', on_entry_change)
                
                # Prevent Tab/Escape from leaving invalid fields
                def on_key_press(event, s=setting, e=entry):
                    if event.keysym in ('Tab', 'ISO_Left_Tab'):
                        val = e.get().strip()
                        if val and not validate_numeric_input(s, val, e):
                            return "break"  # Prevent tab navigation
                    return None
                entry.bind('<KeyPress>', on_key_press)
        elif setting.options:
            # Non-numeric, but has options: use dropdown
            values = [f"{value} : {desc[:50]}" + ("..." if len(desc) > 50 else "") if desc else str(value) for value, desc, _ in setting.options]
            combo = ttk.Combobox(value_frame, values=values, state="readonly", width=60)
            current_value = str(setting.current_value)
            found = False
            for i, (value, desc, _) in enumerate(setting.options):
                if str(value) == current_value:
                    combo.current(i)
                    found = True
                    break
            if not found:
                combo.set(current_value)
            combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
            self.setting_widgets[setting.token] = combo

            def on_combo_change_text(event=None, s=setting, c=combo):
                sel = c.get()
                if ' : ' in sel:
                    val = sel.split(' : ')[0].strip()
                elif ' - ' in sel:
                    val = sel.split(' - ')[0].strip()
                else:
                    val = sel.strip()
                push_and_refresh(s, val)
            combo.bind("<<ComboboxSelected>>", on_combo_change_text)
        else:
            # Fallback: free-entry field
            entry = ttk.Entry(value_frame, width=30)
            entry.insert(0, str(setting.current_value))
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
            self.setting_widgets[setting.token] = entry
            def on_entry_change(event=None, s=setting, e=entry):
                push_and_refresh(s, e.get())
            entry.bind('<FocusOut>', on_entry_change)
            entry.bind('<Return>', on_entry_change)

    def clear_invalid_field_status(self):
        """Clear invalid field status and update status bar if no invalid fields remain"""
        if not self.has_invalid_fields():
            # Restore normal status if no invalid fields
            if hasattr(self, 'settings') and self.settings:
                filename = os.path.basename(getattr(self, 'original_file_path', 'Unknown'))
                self.status_var.set(f"✅ Loaded {len(self.settings)} settings from {filename}")
    
    def update_invalid_field_status(self):
        """Update status bar to reflect current invalid field count"""
        self.has_invalid_fields()  # This will update the status bar

    def _optimize_memory_usage(self):
        """Optimize memory usage by cleaning up unused references"""
        # Clean up old widget references
        self._cleanup_widget_cache()
        
        # Limit undo stack size
        max_undo = 20
        if len(self.undo_stack) > max_undo:
            self.undo_stack = self.undo_stack[-max_undo:]
        
        if hasattr(self, 'redo_stack') and len(self.redo_stack) > max_undo:
            self.redo_stack = self.redo_stack[-max_undo:]
        
        # Force garbage collection
        gc.collect()
    
    def _cleanup_widget_cache(self):
        """Clean up widgets that are no longer visible"""
        # Get currently visible setting tokens
        visible_tokens = set()
        if self.search_active and self.filtered_settings:
            start_idx = self.current_page * self.page_size
            end_idx = min(start_idx + self.page_size, len(self.filtered_settings))
            for i in range(start_idx, end_idx):
                _, setting = self.filtered_settings[i]
                visible_tokens.add(setting.token)
        else:
            start_idx = self.current_page * self.page_size
            end_idx = min(start_idx + self.page_size, len(self.settings))
            for i in range(start_idx, end_idx):
                setting = self.settings[i]
                visible_tokens.add(setting.token)
        
        # Remove widgets for settings not currently visible
        tokens_to_remove = []
        for token in self.setting_widgets:
            if token not in visible_tokens:
                tokens_to_remove.append(token)
        
        for token in tokens_to_remove:
            widget = self.setting_widgets.pop(token, None)
            if widget:
                try:
                    widget.destroy()
                except:
                    pass
    
    def _create_widget_with_validation(self, setting, parent_frame):
        """Create widget with optimized validation and memory management"""
        widget_frame = ttk.Frame(parent_frame)
        widget_frame.pack(fill=tk.X, padx=10, pady=2)
        
        # Setting label with truncation for performance
        label_text = setting.setup_question
        if len(label_text) > 80:
            label_text = label_text[:77] + "..."
        
        setting_label = ttk.Label(widget_frame, text=label_text, font=("Arial", 9, "bold"))
        setting_label.pack(anchor=tk.W, pady=(0, 2))
        
        # Help text with truncation
        if setting.help_string:
            help_text = setting.help_string
            if len(help_text) > 200:
                help_text = help_text[:197] + "..."
            help_label = ttk.Label(widget_frame, text=help_text, font=("Arial", 8), 
                                 foreground="gray", wraplength=400)
            help_label.pack(anchor=tk.W, pady=(0, 5))
        
        # Create appropriate input widget
        if setting.options:
            widget = self._create_combobox_widget(setting, widget_frame)
        else:
            widget = self._create_entry_widget(setting, widget_frame)
        
        # Store weak reference to avoid memory leaks
        self.setting_widgets[setting.token] = widget
        return widget
    def _create_combobox_widget(self, setting, parent):
        """Create optimized combobox widget"""
        combo_frame = ttk.Frame(parent)
        combo_frame.pack(fill=tk.X, pady=2)
        
        # Optimize option display
        option_values = []
        option_display = []
        current_idx = 0
        
        for i, (value, desc, is_current) in enumerate(setting.options):
            option_values.append(value)
            # Truncate long descriptions for performance
            display_text = f"{desc} ({value})" if desc else value
            if len(display_text) > 50:
                display_text = display_text[:47] + "..."
            option_display.append(display_text)
            
            if is_current or value == setting.current_value:
                current_idx = i
        
        combo = ttk.Combobox(combo_frame, values=option_display, state="readonly", width=60)
        combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        if current_idx < len(option_display):
            combo.set(option_display[current_idx])
        
        # Optimized event handler with closure for option_values
        def on_combo_select(event=None):
            try:
                idx = combo.current()
                if 0 <= idx < len(option_values):
                    old_value = setting.current_value
                    setting.current_value = option_values[idx]
                    
                    # Update validation state
                    self._update_validation_state(setting.token, True)
                    
                    # Push to undo stack only if value actually changed
                    if old_value != setting.current_value:
                        self.push_undo()
            except Exception:
                pass
        
        combo.bind("<<ComboboxSelected>>", on_combo_select)
        return combo
    
    def _create_entry_widget(self, setting, parent):
        """Create optimized entry widget with validation"""
        entry_frame = ttk.Frame(parent)
        entry_frame.pack(fill=tk.X, pady=2)
        
        entry = ttk.Entry(entry_frame, width=30)
        entry.pack(side=tk.LEFT, padx=(0, 10))
        entry.insert(0, str(setting.current_value))
        
        # Add range info if available
        if hasattr(setting, 'range_min') and setting.range_min is not None:
            range_label = ttk.Label(entry_frame, text=f"Range: {setting.range_min}-{setting.range_max}", 
                                  font=("Arial", 8), foreground="blue")
            range_label.pack(side=tk.LEFT)
        
        # Optimized validation with debouncing
        validation_timer = None
        
        def validate_delayed():
            nonlocal validation_timer
            if validation_timer:
                entry.after_cancel(validation_timer)
            validation_timer = entry.after(500, lambda: self._validate_entry_value(setting, entry))
        
        def on_entry_change(event=None):
            validate_delayed()
            # Update current value immediately for responsiveness
            setting.current_value = entry.get()
        
        entry.bind('<KeyRelease>', on_entry_change)
        entry.bind('<FocusOut>', lambda e: self._validate_entry_value(setting, entry))
        
        return entry
    
    def _validate_entry_value(self, setting, entry):
        """Fast entry validation with visual feedback"""
        try:
            value = entry.get().strip()
            is_valid = True
            
            # Fast numeric validation
            if setting.is_numeric or (hasattr(setting, 'range_min') and setting.range_min is not None):
                try:
                    num_value = int(value, 0)  # Support hex with 0x prefix
                    if hasattr(setting, 'range_min') and setting.range_min is not None:
                        is_valid = setting.range_min <= num_value <= setting.range_max
                except ValueError:
                    is_valid = False
            
            # Update visual state
            if is_valid:
                entry.configure(style="TEntry")
                self._update_validation_state(setting.token, True)
            else:
                entry.configure(style="Invalid.TEntry")
                self._update_validation_state(setting.token, False)
            
            setting.current_value = value
            
        except Exception:
            self._update_validation_state(setting.token, False)
    
    def _update_validation_state(self, token, is_valid):
        """Track validation state for memory efficiency"""
        if not hasattr(self, '_invalid_fields'):
            self._invalid_fields = set()
        
        if is_valid:
            self._invalid_fields.discard(token)
        else:
            self._invalid_fields.add(token)
        
        # Update status bar efficiently
        if len(self._invalid_fields) > 0:
            self.status_var.set(f"⚠️ {len(self._invalid_fields)} invalid field(s)")
        else:
            # Restore normal status
            if hasattr(self, 'original_file_path'):
                filename = os.path.basename(self.original_file_path)
                self.status_var.set(f"✅ Loaded {len(self.settings)} settings from {filename}")
    
    def finalize_optimizations(self):
        """Apply final performance optimizations"""
        # Set up TTK styles for validation
        style = ttk.Style()
        style.configure("Invalid.TEntry", fieldbackground="#ffcccc")
        style.configure("Valid.TEntry", fieldbackground="#ffffff")
        
        # Configure memory management
        self._setup_memory_management()
        
        # Optimize search patterns
        self._setup_search_optimization()
    
    def _setup_memory_management(self):
        """Setup automatic memory management"""
        # Schedule periodic memory cleanup
        def periodic_cleanup():
            self._optimize_memory_usage()
            # Schedule next cleanup in 30 seconds
            self.root.after(30000, periodic_cleanup)
        
        # Start the cleanup cycle
        self.root.after(30000, periodic_cleanup)
    
    def _setup_search_optimization(self):
        """Setup search pattern compilation for better performance"""
        if not hasattr(self, '_compiled_search_patterns'):
            self._compiled_search_patterns = {}
        
        # Pre-compile common search patterns
        common_terms = ['enable', 'disable', 'mode', 'speed', 'power', 'clock', 'memory', 'cpu', 'gpu']
        for term in common_terms:
            try:
                self._compiled_search_patterns[term] = re.compile(re.escape(term), re.IGNORECASE)
            except:
                pass
    
    def optimize_display_performance(self):
        """Call this method to apply all display optimizations"""
        # Optimize memory usage
        self._optimize_memory_usage()
        
        # Update validation states efficiently
        if hasattr(self, '_invalid_fields'):
            invalid_count = len(self._invalid_fields)
            if invalid_count > 0:
                self.status_var.set(f"⚠️ {invalid_count} invalid field(s) - Please correct before saving")
        
        # Force widget update
        self.root.update_idletasks()

    def check_scewin_availability(self):
        """Check if SCEWIN is available and show status to user"""
        scewin_dir = self._find_scewin_installation()
        
        if scewin_dir:
            return True, f"SCEWIN found at: {scewin_dir}"
        else:
            search_paths = self._get_scewin_search_paths()
            searched_info = "Searched locations:\n" + "\n".join(f"• {path}" for path in search_paths[:10])
            if len(search_paths) > 10:
                searched_info += f"\n• ... and {len(search_paths) - 10} more locations"
            
            return False, (
                "SCEWIN not found in any standard MSI Center installation.\n\n"
                f"{searched_info}\n\n"
                "Solutions:\n"
                "• Install or reinstall MSI Center\n"
                "• Manually copy SCEWIN_64.exe and driver files to the application directory\n"
                "• Check if MSI Center is installed in a custom location"
            )

    def show_scewin_status(self):
        """Show SCEWIN availability status to user"""
        available, message = self.check_scewin_availability()
        
        if available:
            messagebox.showinfo("SCEWIN Status", f"✅ SCEWIN Available\n\n{message}")
        else:
            messagebox.showwarning("SCEWIN Status", f"⚠️ SCEWIN Not Available\n\n{message}")

    def _get_scewin_search_paths(self):
        """Get list of potential SCEWIN installation paths to search"""
        base_paths = [
            r"C:\Program Files\MSI\MSI Center\Lib\SCEWIN",
            r"C:\Program Files (x86)\MSI\MSI Center\Lib\SCEWIN",
            r"C:\Program Files\MSI\MSI Center\SCEWIN",
            r"C:\Program Files (x86)\MSI\MSI Center\SCEWIN",
            # Legacy Dragon Center paths
            r"C:\Program Files\MSI\Dragon Center\Lib\SCEWIN",
            r"C:\Program Files (x86)\MSI\Dragon Center\Lib\SCEWIN",
            # Alternative installation paths
            r"C:\MSI\MSI Center\Lib\SCEWIN",
            r"C:\MSI\Dragon Center\Lib\SCEWIN"
        ]
        
        search_paths = []
        
        # Add base paths directly
        search_paths.extend(base_paths)
        
        # For each base path, also search for version subdirectories
        for base_path in base_paths:
            if os.path.isdir(base_path):
                try:
                    # Look for version directories (e.g., 5.05.01.0002, 5.06.*, etc.)
                    for item in os.listdir(base_path):
                        item_path = os.path.join(base_path, item)
                        if (os.path.isdir(item_path) and 
                            re.match(r'^\d+\.\d+\.\d+\.\d+$', item)):  # Version pattern
                            search_paths.append(item_path)
                except (OSError, PermissionError):
                    continue
        
        return search_paths

    def _find_scewin_installation(self):
        """Dynamically find SCEWIN installation directory"""
        search_paths = self._get_scewin_search_paths()
        
        for path in search_paths:
            if os.path.isfile(os.path.join(path, "SCEWIN_64.exe")):
                # Verify all required files are present
                required_files = ["SCEWIN_64.exe", "amifldrv64.sys", "amigendrv64.sys"]
                all_files_present = True
                
                for filename in required_files:
                    if not os.path.isfile(os.path.join(path, filename)):
                        all_files_present = False
                        break
                
                if all_files_present:
                    return path
                else:
                    # Log missing files for debugging
                    missing_files = [f for f in required_files 
                                   if not os.path.isfile(os.path.join(path, f))]
                    print(f"Found SCEWIN_64.exe at {path} but missing: {missing_files}")
        
        # If not found in standard locations, check current directory and common locations
        fallback_paths = [
            os.path.dirname(os.path.abspath(__file__)),  # Script directory
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "SCEWIN"),
            r"C:\SCEWIN",
            r"C:\Tools\SCEWIN"
        ]
        
        for path in fallback_paths:
            if os.path.isfile(os.path.join(path, "SCEWIN_64.exe")):
                return path
        
        return None


if __name__ == "__main__":
    # Create main window and start application
    root = tk.Tk()
    app = EnhancedBIOSSettingsGUI(root)
    root.mainloop()
