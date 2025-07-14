#!/usr/bin/env python3
"""
Enhanced MSI BIOS Hidden Settings Editor
- Integrated import/export functionality
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
from datetime import datetime
import time
from collections import defaultdict
import webbrowser
import tempfile


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
    __slots__ = ("setup_question", "help_string", "token", "offset", "width", "bios_default", "options", "current_value", "is_numeric", "original_value", "original_has_options", "original_block_lines")
    def __init__(self, setup_question="", help_string="", token="", offset="", 
                 width="", bios_default="", options=None, current_value=None, is_numeric=False):
        self.setup_question = setup_question
        self.help_string = help_string
        self.token = token
        self.offset = offset
        self.width = width
        self.bios_default = bios_default
        self.options = options or []
        self.current_value = current_value or ""
        self.is_numeric = is_numeric
        self.original_has_options = False
        self.original_block_lines = []


class ProgressDialog:
    """Enhanced progress dialog with better user feedback"""
    def __init__(self, parent, title="Processing", message="Please wait..."):
        self.parent = parent
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("400x150")
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center the dialog
        self.dialog.geometry(f"+{parent.winfo_rootx() + 50}+{parent.winfo_rooty() + 50}")
        
        # Create widgets
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Message label
        self.message_label = ttk.Label(main_frame, text=message, font=("Arial", 10))
        self.message_label.pack(pady=(0, 15))
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='determinate', length=350)
        self.progress.pack(pady=(0, 10))
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="Starting...", font=("Arial", 9))
        self.status_label.pack()
        
        # Cancel button
        self.cancel_button = ttk.Button(main_frame, text="Cancel", command=self.cancel)
        self.cancel_button.pack(pady=(10, 0))
        
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
    """High-performance NVRAM parser with progress tracking"""
    
    def __init__(self):
        self.settings = []
        self.header_info = {}
        self.raw_header = ""
        self.categories = defaultdict(list)
        self._cancelled = False
        
    def parse_file(self, file_path, progress_callback=None, cancel_flag=None):
        """Parse NVRAM file with optimized processing and progress updates"""
        self.settings = []
        self.raw_header = ""
        self.categories.clear()
        self._cancelled = False
        
        try:
            # Read file with proper encoding handling
            with open(file_path, 'r', encoding='utf-8', errors='replace') as file:
                content = file.read()
            
            if progress_callback:
                progress_callback(10, "File loaded, parsing header...")
            
            # Extract and parse header
            header_end_pos = content.find("Setup Question")
            if header_end_pos > 0:
                self.raw_header = content[:header_end_pos]
                self._parse_header(self.raw_header)
            
            if cancel_flag and cancel_flag():
                return []
            
            if progress_callback:
                progress_callback(20, "Splitting into setting blocks...")
            
            # Optimized block splitting using regex
            setting_content = content[header_end_pos:] if header_end_pos > 0 else content
            setting_blocks = re.split(r'\n(?=Setup Question\s*=)', setting_content)
            
            # Filter out empty blocks
            setting_blocks = [block.strip() for block in setting_blocks if block.strip()]
            total_blocks = len(setting_blocks)
            
            if progress_callback:
                progress_callback(30, f"Processing {total_blocks} settings...")
            
            # Process blocks with batch progress updates
            batch_size = max(1, total_blocks // 50)  # Update progress every 2%
            
            for i, block in enumerate(setting_blocks):
                if cancel_flag and cancel_flag():
                    break
                
                setting = self._parse_setting_block(block)
                if setting:
                    self.settings.append(setting)
                    
                    # Categorize settings for faster filtering
                    category = self._extract_category(setting.setup_question)
                    self.categories[category].append(len(self.settings) - 1)
                
                # Update progress in batches
                if progress_callback and (i % batch_size == 0 or i == total_blocks - 1):
                    progress_value = 30 + int((i + 1) / total_blocks * 60)
                    status = f"Processed {i + 1} of {total_blocks} settings..."
                    progress_callback(progress_value, status)
            
            if progress_callback:
                progress_callback(95, "Finalizing...")
            
            return self.settings
            
        except Exception as e:
            messagebox.showerror("Parsing Error", f"Failed to parse BIOS file:\n{e}")
            return False
    
    def _extract_category(self, question):
        """Extract category from setting question for organization"""
        if not question:
            return "Other"
        
        # Extract first meaningful word as category
        words = question.split()
        if words:
            return words[0]
        return "Other"
    
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
                    setting.bios_default = self._extract_value(line)
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
                        if (('enabled' in help_string.lower() and 'disabled' in help_string.lower()) or
                            ('enabled' in setting.setup_question.lower() and 'disabled' in setting.setup_question.lower())):
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
            # Validate required fields
            if not setting.setup_question or not setting.token:
                return None
            return setting
        except Exception:
            return None  # Skip malformed settings
    
    def _extract_value(self, line):
        """Extract value after = sign"""
        if '=' in line:
            return line.split('=', 1)[1].strip()
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
                # Try to infer allowed range from help string (e.g., 'range:0 ~ 31')
                value = str(setting.current_value)
                if hasattr(setting, 'help_string') and setting.help_string:
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
    def on_inline_search_changed(self, *args):
        """Handler for inline search box. Filters settings and displays results in main area. Now uses fuzzy matching if available."""
        # Prevent error if scrollable_frame is not yet created
        if not hasattr(self, 'scrollable_frame'):
            return
        search_text = self.inline_search_var.get().strip().lower()
        # If search is empty, show normal paged view
        if not search_text:
            page = getattr(self, '_current_page', 0)
            self.load_page_settings(page)
            return
        # Try to use rapidfuzz for fuzzy matching, else fallback to substring
        try:
            from rapidfuzz import fuzz
            def fuzzy_score(s):
                return max(
                    fuzz.partial_ratio(search_text, str(s.setup_question).lower()),
                    fuzz.partial_ratio(search_text, str(s.help_string).lower()),
                    fuzz.partial_ratio(search_text, str(s.token).lower())
                )
            scored = [(i, setting, fuzzy_score(setting)) for i, setting in enumerate(self.settings)]
            # Only keep those with score >= 60 (tune as needed)
            filtered = [(i, s) for i, s, score in scored if score >= 60]
            # Sort by best match
            matched_settings = sorted(filtered, key=lambda x: -fuzzy_score(x[1]))
        except ImportError:
            # Fallback: substring match
            matched_settings = []
            for i, setting in enumerate(self.settings):
                if (search_text in str(setting.setup_question).lower() or
                    search_text in str(setting.help_string).lower() or
                    search_text in str(setting.token).lower()):
                    matched_settings.append((i, setting))
        if not matched_settings:
            # Clear display
            for widget in self.scrollable_frame.winfo_children():
                widget.destroy()
            label = ttk.Label(self.scrollable_frame, text="No matches found.", font=("Arial", 12), foreground="gray")
            label.pack(pady=40)
            return
        # Show first batch (20) with lazy loading
        self.display_filtered_settings(matched_settings)

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
        self.root.title("Enhanced MSI BIOS Hidden Settings Editor")
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
        # Advanced optimization: cache for settings batches (for lazy loading)
        self._settings_batch_cache = {}
        # Advanced optimization: track last access for LRU purging
        self._batch_access_times = {}
        self._max_batches_in_memory = 5  # Tune as needed for memory/performance

        # Track the most recently changed token for highlight in raw view
        self._last_changed_token = None

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
    
    def setup_styles(self):
        """Configure modern styling"""
        style = ttk.Style()
        
        # Configure treeview for better appearance
        style.configure("Treeview", rowheight=25)
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
        
        # Button styling
        style.configure("Action.TButton", font=("Arial", 9, "bold"))
    
    def check_admin_rights(self):
        """Check and prompt for admin rights"""
        try:
            is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0
        except:
            is_admin = False
        
        if not is_admin:
            self.status_var.set("⚠️ Admin rights recommended for BIOS operations")
    
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

    def save_and_import_bios_with_review(self):
        """Show change review dialog and validate before importing to BIOS."""
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
        possible_dirs = [
            r"C:\\Program Files (x86)\\MSI\\MSI Center\\Lib\\SCEWIN\\5.05.01.0002",
            r"C:\\Program Files\\MSI\\MSI Center\\Lib\\SCEWIN\\5.05.01.0002"
        ]
        scewin_dir = None
        for d in possible_dirs:
            if os.path.isfile(os.path.join(d, "SCEWIN_64.exe")):
                scewin_dir = d
                break
        if not scewin_dir:
            messagebox.showerror("Missing SCEWIN_64.exe", "Could not find SCEWIN_64.exe in known MSI Center locations.")
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
                self.root.after(0, lambda: messagebox.showinfo(
                    "Export Successful", "BIOS settings exported to nvram.txt.\n\nLog:\n" + log_content))
                self.root.after(0, lambda: self.load_file(nvram_txt))
                if progress: progress.close()
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
        temp_dir = os.path.join(tempfile.gettempdir(), "scewin_temp")
        os.makedirs(temp_dir, exist_ok=True)
        scewin_exe = os.path.join(temp_dir, "SCEWIN_64.exe")
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
            # Show first batch of settings with lazy loading
            self.display_lazy_loaded_settings()
            progress.close()
        except Exception as e:
            progress.close()
            messagebox.showerror("Interface Error", f"Failed to update interface: {str(e)}")
    
    def populate_navigation(self):
        pass  # No longer needed for infinite scroll
    def display_lazy_loaded_settings(self, batch_size=20, keep_widgets=60):
        """Display settings with infinite/lazy loading as user scrolls. Only keep a limited number of widgets in memory."""
        self._lazy_settings_offset = 0
        self._lazy_batch_size = batch_size
        self._lazy_keep_widgets = keep_widgets
        self._lazy_loaded_indices = set()
        # Clear current display
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.setting_widgets.clear()
        self._lazy_settings_total = len(self.settings)
        self._lazy_settings_widgets = []
        self._lazy_settings_frame = self.scrollable_frame
        # Use self.canvas as the scrollable Canvas
        self._lazy_settings_canvas = self.canvas
        self._lazy_settings_canvas.bind('<Configure>', self._on_lazy_scroll)
        self._lazy_settings_canvas.bind_all('<MouseWheel>', self._on_lazy_scroll)
        self._lazy_settings_canvas.bind_all('<Button-4>', self._on_lazy_scroll)  # Linux scroll up
        self._lazy_settings_canvas.bind_all('<Button-5>', self._on_lazy_scroll)  # Linux scroll down
        self._lazy_settings_last_y = 0
        self._lazy_settings_last_max = 0
        self._lazy_settings_loading = False
        self._lazy_settings_load_batch(0)

    def _on_lazy_scroll(self, event=None):
        # Check if near bottom, then load more
        canvas = self._lazy_settings_canvas
        try:
            yview = canvas.yview()
            if yview[1] > 0.95 and not self._lazy_settings_loading:
                # Near bottom, load next batch
                self._lazy_settings_loading = True
                self._lazy_settings_load_batch(len(self._lazy_loaded_indices))
                self._lazy_settings_loading = False
        except Exception:
            pass

    def _lazy_settings_load_batch(self, start_index):
        end_index = min(start_index + self._lazy_batch_size, self._lazy_settings_total)
        for i in range(start_index, end_index):
            if i not in self._lazy_loaded_indices:
                self.create_setting_widget(self.settings[i], i)
                self._lazy_loaded_indices.add(i)
        # Remove widgets far above current scroll for memory
        if len(self._lazy_loaded_indices) > self._lazy_keep_widgets:
            min_index = min(self._lazy_loaded_indices)
            max_index = max(self._lazy_loaded_indices)
            to_remove = [idx for idx in self._lazy_loaded_indices if idx < max_index - self._lazy_keep_widgets]
            for idx in to_remove:
                widget = self.setting_widgets.get(self.settings[idx].token)
                if widget:
                    widget.master.destroy()
                self._lazy_loaded_indices.remove(idx)
                self.setting_widgets.pop(self.settings[idx].token, None)

    def update_category_menu(self):
        """Replace category pills with a searchable dropdown (combobox) for categories, sorted by frequency."""
        if not hasattr(self, 'category_menu_frame') or not self.settings:
            return
        for widget in self.category_menu_frame.winfo_children():
            widget.destroy()

        from collections import Counter
        stopwords = set([
            'the', 'and', 'or', 'to', 'of', 'in', 'for', 'on', 'with', 'by', 'is', 'at', 'as', 'an', 'be', 'are',
            'from', 'this', 'that', 'it', 'if', 'not', 'can', 'will', 'a', 'but', 'was', 'has', 'have', 'may', 'all',
            'help', 'string', 'text', 'value', 'option', 'set', 'setting', 'settings', 'default', 'enable', 'disable',
            'yes', 'no', 'auto', 'user', 'system', 'mode', 'type', 'select', 'use', 'change', 'current', 'bios', 'token',
            'offset', 'width', 'page', 'number', 'data', 'field', 'bit', 'bits', 'description', 'desc', 'info', 'information',
            'boot', 'save', 'import', 'export', 'file', 'load', 'backup', 'restore', 'undo', 'redo', 'option', 'options',
            'nvram', 'msi', 'ami', 'center', 'utility', 'ver', 'copyright', 'reserved', 'crc32', 'script', 'name', 'created',
            'do', 'not', 'change', 'line', 'move', 'desired', 'move', 'move', 'move', 'move', 'move', 'move', 'move', 'move',
        ])
        word_counter = Counter()
        for s in self.settings:
            for text in (str(s.setup_question), str(s.help_string), str(s.token)):
                words = re.findall(r'\b\w+\b', text.lower())
                for w in words:
                    if w not in stopwords and len(w) > 2:
                        word_counter[w] += 1
        sorted_words = [w for w, _ in word_counter.most_common()]
        self._category_menu_state['words'] = sorted_words

        # Create a combobox for categories
        self.category_var = tk.StringVar()
        category_combo = ttk.Combobox(self.category_menu_frame, textvariable=self.category_var, values=[w.capitalize() for w in sorted_words], font=("Arial", 10))
        category_combo.pack(fill=tk.X, padx=2, pady=2)
        category_combo.set("")
        category_combo.bind("<KeyRelease>", self._on_category_combo_typed)
        category_combo.bind("<<ComboboxSelected>>", self._on_category_combo_selected)

    def _on_category_combo_selected(self, event=None):
        val = self.category_var.get().strip()
        if val:
            self.inline_search_var.set(val)

    def _on_category_combo_typed(self, event=None):
        # As user types, filter the dropdown list
        val = self.category_var.get().strip().lower()
        all_words = self._category_menu_state.get('words', [])
        filtered = [w.capitalize() for w in all_words if val in w]
        if event is not None and getattr(event, 'widget', None) is not None:
            event.widget['values'] = filtered
    
    def validate_setting(self, setting, value):
        """Validate setting value before saving"""
        if setting.is_numeric:
            try:
                int(value)
            except ValueError:
                messagebox.showerror("Validation Error", f"Value for {setting.setup_question} must be numeric.")
                return False
        if setting.options and value not in setting.options:
            messagebox.showerror("Validation Error", f"Value '{value}' not in allowed options for {setting.setup_question}.")
            return False
        return True
    
    # Utility methods
    def find_scetool_path(self):
        """Locate MSI SCEWIN tools directory"""
        # Try PATH first
        scetool = shutil.which("SCEWIN.exe")
        if scetool:
            return scetool
        # Try common MSI Center install location
        possible_dirs = [
            r"C:\Program Files (x86)\MSI\MSI Center\Lib\SCEWIN\5.05.01.0002",
            r"C:\Program Files\MSI\MSI Center\Lib\SCEWIN\5.05.01.0002"
        ]
        for d in possible_dirs:
            exe_path = os.path.join(d, "SCEWIN.exe")
            if os.path.isfile(exe_path):
                return exe_path
        return None
    
    def run_scetool_with_progress(self, scetool_path, command, progress):
        """Run SCEWIN tool with progress tracking"""
        try:
            # Check admin privileges
            try:
                is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0
            except:
                is_admin = False
            
            if not is_admin:
                progress.update_progress(50, "Requesting administrator privileges...")
                # Request elevation
                result = ctypes.windll.shell32.ShellExecuteW(
                    None, "runas", "cmd.exe", 
                    f'/c cd /d "{scetool_path}" && {command}', 
                    None, 1
                )
                return result > 32  # Success if > 32
            else:
                # Run directly
                full_command = f'cd /d "{scetool_path}" && {command}'
                result = subprocess.run(['cmd.exe', '/c', full_command], 
                                      capture_output=True, text=True, timeout=60)
                return result.returncode == 0
                
        except Exception as e:
            progress.update_progress(0, f"Error: {str(e)}")
            return False

    def save_file_only(self):
        """Save to file without importing to BIOS"""
        if not self.settings:
            messagebox.showwarning("No Settings", "No settings loaded to save.")
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
                        progress_callback(int(10 + 90 * i / total_settings), f"Writing setting {i+1}/{total_settings}")
                    # Copy original block and only update * or value as needed
                    block_lines = list(setting.original_block_lines)
                    # Update * marker for options
                    if getattr(setting, 'original_has_options', False) and setting.options:
                        # Remove all * markers
                        new_lines = []
                        for line in block_lines:
                            # Remove * from options lines
                            if re.match(r'\s*\*?\[', line.strip()):
                                new_lines.append(re.sub(r'\*', '', line, count=1))
                            elif re.match(r'Options\s*=\s*\*?\[', line.strip()):
                                new_lines.append(re.sub(r'\*', '', line, count=1))
                            else:
                                new_lines.append(line)
                        # Add * to the correct option
                        for idx, (value, desc, _) in enumerate(setting.options):
                            # Find the line for this option
                            val_str = f'[{value}]'
                            for j, l in enumerate(new_lines):
                                if val_str in l:
                                    # Add * if this is the current value
                                    if str(value) == str(setting.current_value):
                                        # Insert * at the right place
                                        new_lines[j] = re.sub(r'(Options\s*=\s*)?(\s*)', r'\1\2*', new_lines[j], count=1)
                        block_lines = new_lines
                    # Update value for value lines
                    elif not getattr(setting, 'original_has_options', False):
                        new_lines = []
                        for line in block_lines:
                            if line.strip().startswith('Value'):
                                new_lines.append(re.sub(r'<[^>]*>', f'<{setting.current_value}>', line))
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

    def _display_settings_batch(self, matched_settings, start, batch_size=5):
        end = min(start + batch_size, len(matched_settings), 20)
        for i in range(start, end):
            index, setting = matched_settings[i]
            self.create_setting_widget(setting, index)
        if end < min(len(matched_settings), 20):
            # Schedule next batch
            self.root.after(10, lambda: self._display_settings_batch(matched_settings, end, batch_size))
        elif len(matched_settings) > 20:
            # Add Load More button after first 20
            load_more_btn = ttk.Button(
                self.scrollable_frame,
                text=f"Load {len(matched_settings) - 20} more matches...",
                command=lambda: self.load_more_search_results(matched_settings[20:])
            )
            load_more_btn.pack(pady=10)

    def _get_settings_batch(self, batch_key, settings_list, cache=True):
        """Return a batch of settings, using LRU cache for large sets."""
        import time
        if not cache:
            return settings_list
        # LRU cache logic
        now = time.time()
        self._batch_access_times[batch_key] = now
        # Purge old batches if over limit
        if len(self._settings_batch_cache) > self._max_batches_in_memory:
            # Remove least recently used
            sorted_batches = sorted(self._batch_access_times.items(), key=lambda x: x[1])
            for k, _ in sorted_batches[:-self._max_batches_in_memory]:
                self._settings_batch_cache.pop(k, None)
                self._batch_access_times.pop(k, None)
        # Return from cache or store
        if batch_key in self._settings_batch_cache:
            return self._settings_batch_cache[batch_key]
        self._settings_batch_cache[batch_key] = settings_list
        return settings_list
    
    def on_category_changed(self, event=None):
        """Category filter removed: do nothing."""
        pass
    
    def load_category_settings(self, category):
        """Category navigation removed: do nothing."""
        pass

    def load_page_settings(self, page):
        """Display a page of settings in the scrollable frame."""
        page_size = 20
        start = page * page_size
        end = min(start + page_size, len(self.settings))
        # Clear current display
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.setting_widgets.clear()
        for i in range(start, end):
            self.create_setting_widget(self.settings[i], i)
        # Track current page
        self._current_page = page
    
    def load_more_settings(self, category, start_index):
        """Load more settings for a category, clearing unused widgets. Uses batch cache."""
        if category in self.parser.categories:
            batch = self._get_settings_batch(category, [(i, self.settings[i]) for i in self.parser.categories[category]], cache=True)

            # Remove the "Load More" button
            for widget in self.scrollable_frame.winfo_children():
                if isinstance(widget, ttk.Button) and "Load" in widget.cget("text"):
                    widget.destroy()
                    break

            # Load next batch
            end_index = min(start_index + 20, len(batch))
            for i, (setting_index, setting) in enumerate(batch[start_index:end_index]):
                self.create_setting_widget(setting, setting_index)

            # Add "Load More" button if there are still more settings
            if end_index < len(batch):
                load_more_btn = ttk.Button(
                    self.scrollable_frame,
                    text=f"Load {len(batch) - end_index} more settings...",
                    command=lambda: self.load_more_settings(category, end_index)
                )
                load_more_btn.pack(pady=10)
    
    def load_more_search_results(self, remaining_settings, show_goto=False):
        """Load more search results, clearing unused widgets. Supports Go to button."""
        # Remove the "Load More" button
        for widget in self.scrollable_frame.winfo_children():
            if isinstance(widget, ttk.Button) and "Load" in widget.cget("text"):
                widget.destroy()
                break

        # Load next batch of settings
        for i, (index, setting) in enumerate(remaining_settings[:20]):
            self.create_setting_widget(setting, index, show_goto=show_goto)

        # Add "Load More" button if there are more settings
        if len(remaining_settings) > 20:
            load_more_btn = ttk.Button(
                self.scrollable_frame,
                text=f"Load {len(remaining_settings) - 20} more matches...",
                command=lambda: self.load_more_search_results(remaining_settings[20:], show_goto=show_goto)
            )
            load_more_btn.pack(pady=10)
    
    def create_setting_widget(self, setting, index, show_goto=False):
        """Create optimized widget for individual setting (memory and UI optimized)"""
        # Ignore/commented-out settings (leading //): do not allow configuration
        if str(setting.setup_question).strip().startswith("//") or str(setting.token).strip().startswith("//"):
            # Show as grayed-out, not editable
            setting_frame = ttk.LabelFrame(
                self.scrollable_frame, 
                text="[IGNORED] " + setting.setup_question[:80] + ("..." if len(setting.setup_question) > 80 else ""), 
                padding=8
            )
            setting_frame.pack(fill=tk.X, padx=5, pady=3)
            help_text = setting.help_string[:200] + ("..." if len(setting.help_string) > 200 else "") if setting.help_string else ""
            if help_text.strip():
                help_label = ttk.Label(setting_frame, text=help_text, foreground="gray", font=("Arial", 8))
                help_label.pack(anchor=tk.W, pady=(0, 5))
            tech_info = f"Token: {setting.token} | Offset: {setting.offset} | Default: {setting.bios_default}"
            ttk.Label(setting_frame, text=tech_info, font=("Consolas", 7), foreground="gray").pack(anchor=tk.W, pady=(0, 5))
            ttk.Label(setting_frame, text="This setting is ignored (commented out in BIOS file)", font=("Arial", 9, "italic"), foreground="gray").pack(anchor=tk.W, pady=(0, 5))
            return

        # Use a lightweight frame for each setting
        setting_frame = ttk.LabelFrame(
            self.scrollable_frame, 
            text=setting.setup_question[:80] + ("..." if len(setting.setup_question) > 80 else ""), 
            padding=8
        )
        setting_frame.pack(fill=tk.X, padx=5, pady=3)

        # Only show help if not empty and not too long
        if setting.help_string:
            help_text = setting.help_string[:200] + ("..." if len(setting.help_string) > 200 else "")
            if help_text.strip():
                help_label = ttk.Label(setting_frame, text=help_text, 
                                      foreground="gray", font=("Arial", 8))
                help_label.pack(anchor=tk.W, pady=(0, 5))

        # Technical info in compact format
        tech_info = f"Token: {setting.token} | Offset: {setting.offset} | Default: {setting.bios_default}"
        ttk.Label(setting_frame, text=tech_info, font=("Consolas", 7), foreground="darkblue").pack(anchor=tk.W, pady=(0, 5))

        # Value input section
        value_frame = ttk.Frame(setting_frame)
        value_frame.pack(fill=tk.X, pady=(5, 0))

        # Add Go to button if requested
        if show_goto:
            ttk.Button(value_frame, text="Go to", command=lambda idx=index: self.goto_setting_in_main(idx)).pack(side=tk.RIGHT, padx=(8, 0))

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
            values = [f"{value} : {desc[:50]}" + ("..." if len(desc) > 50 else "") if desc else str(value) for value, desc, _ in setting.options]
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

            def on_combo_change(event=None, s=setting, c=combo):
                sel = c.get()
                # Extract value from "value : desc"
                if ' : ' in sel:
                    val = sel.split(' : ')[0].strip()
                elif ' - ' in sel:
                    val = sel.split(' - ')[0].strip()
                else:
                    val = sel.strip()
                push_and_refresh(s, val)
            combo.bind("<<ComboboxSelected>>", on_combo_change)
        elif setting.is_numeric:
            # Only use free-entry numeric field if no options or too many options
            entry = ttk.Entry(value_frame, width=30)
            entry.insert(0, str(setting.current_value))
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
            self.setting_widgets[setting.token] = entry
            ttk.Label(value_frame, text="(Numeric)", font=("Arial", 8), foreground="blue").pack(side=tk.LEFT, padx=(5, 0))
            def on_entry_change(event=None, s=setting, e=entry):
                push_and_refresh(s, e.get())
            entry.bind('<FocusOut>', on_entry_change)
            entry.bind('<Return>', on_entry_change)
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

            def on_combo_change(event=None, s=setting, c=combo):
                sel = c.get()
                if ' : ' in sel:
                    val = sel.split(' : ')[0].strip()
                elif ' - ' in sel:
                    val = sel.split(' - ')[0].strip()
                else:
                    val = sel.strip()
                push_and_refresh(s, val)
            combo.bind("<<ComboboxSelected>>", on_combo_change)
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

    def goto_setting_in_main(self, idx):
        """Scroll to and highlight the setting at the given index in the main view."""
        self.load_page_settings(idx // 20)
        self.root.after(150, lambda: self.scroll_and_highlight_setting(idx))

    def display_filtered_settings(self, settings_list):
        """Display filtered settings with lazy loading and Go to buttons"""
        # Clear current display
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        # Display first batch with Go to buttons
        for i, (index, setting) in enumerate(settings_list[:20]):
            self.create_setting_widget(setting, index, show_goto=True)
        # Add "Load More" button if needed
        if len(settings_list) > 20:
            load_more_btn = ttk.Button(
                self.scrollable_frame,
                text=f"Load {len(settings_list) - 20} more settings...",
                command=lambda: self.load_more_search_results(settings_list[20:], show_goto=True)
            )
            load_more_btn.pack(pady=10)

if __name__ == "__main__":
    root = tk.Tk()
    app = EnhancedBIOSSettingsGUI(root)
    root.mainloop()

