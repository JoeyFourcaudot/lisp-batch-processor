#!/usr/bin/env python
"""
Revision: 9 - Updated status messages for each Lisp execution with the Lisp name.
Created by Jiraiya78 | Version 1.0.3

Changes:
- After executing each Lisp script for a DWG file, the status is updated to indicate
  which Lisp (by name and order) completed for that file.
  For example: "myscript.lsp completed for file X (Lisp 1 of 3)"
- Other functionalities remain unchanged.
"""

# Import necessary modules for system operations, GUI, file handling, and AutoCAD integration
import sys  # For system-specific parameters and functions
import os  # For operating system dependent functionality like file paths
import tkinter as tk  # Main GUI library for creating the interface
from tkinter import filedialog, ttk, messagebox  # Additional Tkinter components for dialogs and widgets
from tkinterdnd2 import TkinterDnD, DND_FILES  # For drag-and-drop functionality
import pythoncom  # For COM initialization in multithreaded environments
import win32com.client  # For interacting with AutoCAD via COM
import win32gui  # For Windows GUI operations like hiding windows
import win32con  # Windows constants for GUI operations
import threading  # For running processing in a separate thread to avoid freezing the GUI
import json  # For loading and saving settings in JSON format
from PIL import Image, ImageTk  # For handling and displaying images like logos
import time  # For adding delays in processing

# Function to get the absolute path to resources, handling both development and bundled executable environments
def resource_path(relative_path):
    """
    Get absolute path to resource, works for PyInstaller.
    If running as a bundled executable, sys._MEIPASS points to the temporary folder
    where resources are stored. Otherwise, return the path relative to current directory.
    """
    try:
        base_path = sys._MEIPASS  # Path used by PyInstaller for bundled apps
    except Exception:
        base_path = os.path.abspath(".")  # Current directory for development
    return os.path.join(base_path, relative_path)

# Set the TKDND_LIBRARY environment variable to the path of the tkdnd library for drag-and-drop support
os.environ["TKDND_LIBRARY"] = resource_path("tkdnd2.8")

# Function to hide AutoCAD windows to prevent them from appearing during batch processing
def hide_autocad_window():
    """
    Enumerate all top-level windows and hide those whose title contains 'AutoCAD'.
    """
    def enum_handler(hwnd, lparam):
        # Check if the window is visible
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)  # Get the window title
            if "AutoCAD" in title:  # If 'AutoCAD' is in the title, hide the window
                win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
    win32gui.EnumWindows(enum_handler, None)  # Enumerate all windows and apply the handler

# Function to recursively find all .lsp files in a directory and its subdirectories
def get_lisp_files(directory):
    """Recursively scan for .lsp files in a given directory."""
    lisp_files = []  # List to store found Lisp files
    for root, dirs, files in os.walk(directory):  # Walk through directory tree
        for file in files:  # Check each file
            if file.lower().endswith(".lsp"):  # If it's a .lsp file (case insensitive)
                lisp_files.append(os.path.join(root, file))  # Add full path to list
    return lisp_files  # Return the list of Lisp files

# Main application class for the Lisp Batch Processor GUI
class LispBatchProcessorApp:
    # Initialize the application with GUI setup and default configurations
    def __init__(self, root):
        self.root = root  # Store reference to the root Tkinter window
        self.root.title("Lisp Batch Processor")  # Set window title
        self.root.resizable(True, True)  # Allow window resizing
        self.file_list = []  # List to hold selected DWG files
        # Change lisp_files to a list of dicts with keys "path" and "var" for maintaining order.
        self.lisp_files = []  # List of Lisp files as dictionaries with path and checkbox variable

        # Determine the base path for the application (handles both script and executable modes)
        if getattr(sys, "frozen", False):  # Check if running as PyInstaller bundle
            base_path = os.path.dirname(sys.executable)  # Use executable directory
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))  # Use script directory
        self.default_lisp_dir = os.path.join(base_path, "lisp")  # Default directory for Lisp files

        self.settings_file = "settings.json"  # File to store user settings
        self.load_settings()  # Load settings from file
        self.load_default_lisps()  # Load default Lisp files from directory
        self.create_widgets()  # Create GUI widgets
        self.style_widgets()  # Apply styling to widgets
        self.update_process_button_state()  # Update button states based on selections
        self.options_window = None  # Reference to options window (initially None)
        self.success_count = 0  # Counter for successfully processed files

    # Load application settings from JSON file, or create default settings if file doesn't exist
    def load_settings(self):
        if os.path.exists(self.settings_file):  # Check if settings file exists
            with open(self.settings_file, "r") as f:  # Open file for reading
                self.settings = json.load(f)  # Load JSON data into settings dictionary
        else:
            self.settings = {"autocad_location": ""}  # Default settings with empty AutoCAD location
            self.save_settings()  # Save default settings to file

    # Save current settings to JSON file
    def save_settings(self):
        with open(self.settings_file, "w") as f:  # Open file for writing
            json.dump(self.settings, f, indent=4)  # Write settings as formatted JSON

    # Load default Lisp files from the default Lisp directory
    def load_default_lisps(self):
        # Scan the default Lisp directory (and subdirectories) for .lsp files.
        if os.path.isdir(self.default_lisp_dir):  # Check if default directory exists
            default_lisps = get_lisp_files(self.default_lisp_dir)  # Get all Lisp files in directory
            for lisp in default_lisps:  # For each found Lisp file
                self.lisp_files.append({"path": lisp, "var": tk.BooleanVar(value=True)})  # Add as dict with path and enabled checkbox
        else:
            os.makedirs(self.default_lisp_dir, exist_ok=True)  # Create directory if it doesn't exist

    # Create and arrange all GUI widgets for the application
    def create_widgets(self):
        frame = tk.Frame(self.root, padx=10, pady=10)  # Main frame with padding
        frame.pack(fill=tk.BOTH, expand=True)  # Pack to fill window and expand

        file_frame = tk.LabelFrame(frame, text="DWG Files", padx=10, pady=10)  # Frame for file selection
        file_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")  # Grid layout

        lisp_frame = tk.LabelFrame(frame, text="Lisp Scripts", padx=10, pady=10)  # Frame for Lisp selection
        lisp_frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")  # Grid layout

        self.file_listbox = tk.Listbox(file_frame, selectmode=tk.EXTENDED, width=50, height=15, font=("Helvetica", 12), activestyle="none")  # Listbox for DWG files
        self.file_listbox.grid(row=0, column=0, sticky="nsew")  # Grid in file frame

        file_scrollbar = tk.Scrollbar(file_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)  # Vertical scrollbar for listbox
        file_scrollbar.grid(row=0, column=1, sticky="ns", padx=(0, 5))  # Position scrollbar
        self.file_listbox.config(yscrollcommand=file_scrollbar.set)  # Link scrollbar to listbox

        self.file_listbox.bind("<Delete>", lambda e: self.remove_files())  # Bind delete key to remove files
        self.file_listbox.bind("<<ListboxSelect>>", self.update_backdrop_text)  # Update backdrop on selection

        self.backdrop_text = tk.Label(self.file_listbox, text="Drag and drop to add file or use button", font=("Helvetica", 12, "italic"), fg="grey")  # Placeholder text
        self.backdrop_text.pack(side="top", fill="both", expand=True)  # Pack backdrop label

        file_buttons_frame = tk.Frame(file_frame)  # Frame for file buttons
        file_buttons_frame.grid(row=0, column=2, padx=(5, 0), pady=5, sticky="n")

        self.add_file_button = tk.Button(file_buttons_frame, text="+", command=self.add_files, font=("Helvetica", 24, "bold"), fg="green", width=2, height=1)  # Add file button
        self.add_file_button.pack(pady=(10, 5))

        self.remove_file_button = tk.Button(file_buttons_frame, text="-", command=self.remove_files, font=("Helvetica", 24, "bold"), fg="red", width=2, height=1)  # Remove file button
        self.remove_file_button.pack(pady=(5, 10))

        self.root.drop_target_register(DND_FILES)  # Register root for drag-and-drop of files
        self.root.dnd_bind("<<Drop>>", self.drop_files)  # Bind drop event to handler

        # Create a frame to hold the list of Lisp entries with reordering buttons.
        self.lisp_listbox_frame = tk.Frame(lisp_frame)  # Frame for Lisp list with buttons
        self.lisp_listbox_frame.pack(pady=5, fill=tk.BOTH, expand=True)
        self.refresh_lisp_list()  # Populate the Lisp list

        lisp_buttons_frame = tk.Frame(lisp_frame)  # Frame for Lisp management buttons
        lisp_buttons_frame.pack(pady=5)
        self.add_lisp_button = ttk.Button(lisp_buttons_frame, text="Add Lisp", command=self.add_lisp)  # Button to add Lisp files
        self.add_lisp_button.pack(side=tk.LEFT, padx=5)
        self.remove_lisp_button = ttk.Button(lisp_buttons_frame, text="Remove Lisp", command=self.remove_lisp)  # Button to remove unchecked Lisps
        self.remove_lisp_button.pack(side=tk.LEFT, padx=5)

        self.process_button = ttk.Button(self.root, text="Process Files", command=self.start_processing, width=20)  # Button to start processing
        self.process_button.pack(pady=10)
        self.process_button.config(state=tk.DISABLED)  # Initially disabled

        self.progress = ttk.Progressbar(self.root, length=600, mode="determinate")  # Progress bar for processing
        self.progress.pack(pady=5, fill=tk.X, padx=10)

        status_frame = tk.Frame(self.root)  # Frame for status text area
        status_frame.pack(pady=5, fill=tk.X, padx=10)
        self.status_text = tk.Text(status_frame, height=7, font=("Helvetica", 12, "italic"), state="disabled")  # Text widget for status messages
        self.status_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        status_scrollbar = tk.Scrollbar(status_frame, orient=tk.VERTICAL, command=self.status_text.yview)  # Scrollbar for status text
        status_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.status_text.config(yscrollcommand=status_scrollbar.set)

        # Use resource_path to correctly load the Autocad logo.
        gear_img_path = resource_path("logo\\Autocad.png")  # Path to gear/options icon
        try:
            gear_img = Image.open(gear_img_path)  # Open image
            gear_img = gear_img.resize((24, 24), Image.Resampling.LANCZOS)  # Resize to 24x24
            self.gear_photo = ImageTk.PhotoImage(gear_img)  # Convert to PhotoImage
        except Exception as e:
            print("Failed to load gear image:", e)  # Print error if loading fails
            self.gear_photo = None
        self.options_button = ttk.Button(self.root, image=self.gear_photo, command=self.open_options)  # Options button with icon
        self.options_button.pack(side=tk.RIGHT, padx=10, pady=10)
        self.credit_label = ttk.Label(self.root, text="Created by Jiraiya78 | Version 1.0.3", font=("Helvetica", 10, "italic"))  # Credit label
        self.credit_label.pack(side=tk.BOTTOM, pady=5)

        # Configure grid weights to expand proportionally when resizing.
        #root.grid_rowconfigure(0, weight=1)
        #root.grid_columnconfigure(0, weight=1)
        #root.grid_columnconfigure(1, weight=1)
        #root.grid_columnconfigure(2, weight=1)


# Refresh the display of Lisp scripts with checkboxes and reorder buttons
    def refresh_lisp_list(self):
        """Clear and redraw the list of Lisp scripts with up/down buttons."""
        # Remove existing widgets in lisp_listbox_frame.
        for widget in self.lisp_listbox_frame.winfo_children():  # Destroy all child widgets
            widget.destroy()
        # For each Lisp entry, create a frame with a checkbutton and arrow buttons.
        for index, item in enumerate(self.lisp_files):  # Iterate through Lisp files
            row_frame = tk.Frame(self.lisp_listbox_frame)  # Frame for each row
            row_frame.pack(fill=tk.X, pady=2)  # Pack row frame
            
            chk = tk.Checkbutton(row_frame, text=os.path.basename(item["path"]), variable=item["var"], font=("Helvetica", 12))  # Checkbox with filename
            chk.pack(side=tk.LEFT, padx=5)
            
            btn_frame = tk.Frame(row_frame)  # Frame for buttons
            btn_frame.pack(side=tk.RIGHT)
            
            # Up button (disable if first item)
            up_state = tk.NORMAL if index > 0 else tk.DISABLED  # Enable only if not first
            up_btn = ttk.Button(btn_frame, text="▲", width=2, state=up_state, command=lambda idx=index: self.move_lisp_up(idx))  # Up button
            up_btn.pack(side=tk.LEFT, padx=2)
            # Down button (disable if last item)
            down_state = tk.NORMAL if index < len(self.lisp_files) - 1 else tk.DISABLED  # Enable only if not last
            down_btn = ttk.Button(btn_frame, text="▼", width=2, state=down_state, command=lambda idx=index: self.move_lisp_down(idx))  # Down button
            down_btn.pack(side=tk.LEFT, padx=2)

    # Move the Lisp script at the given index up in the list
    def move_lisp_up(self, index):
        """Move the Lisp script at the given index up in the list."""
        if index > 0:  # Only move if not already at top
            self.lisp_files[index - 1], self.lisp_files[index] = self.lisp_files[index], self.lisp_files[index - 1]  # Swap with previous
            self.refresh_lisp_list()  # Refresh display

    # Move the Lisp script at the given index down in the list
    def move_lisp_down(self, index):
        """Move the Lisp script at the given index down in the list."""
        if index < len(self.lisp_files) - 1:  # Only move if not already at bottom
            self.lisp_files[index + 1], self.lisp_files[index] = self.lisp_files[index], self.lisp_files[index + 1]  # Swap with next
            self.refresh_lisp_list()  # Refresh display

    # Apply custom styling to Tkinter widgets
    def style_widgets(self):
        style = ttk.Style()  # Create style object
        style.configure("TButton", font=("Helvetica", 12, "bold"), padding=6)  # Style for buttons
        style.configure("TLabel", font=("Helvetica", 12))  # Style for labels
        style.configure("TListbox", font=("Courier", 12))  # Style for listboxes (though not ttk)
        style.configure("TProgressbar", thickness=20)  # Style for progress bar

    # Open file dialog to add DWG files to the list
    def add_files(self):
        files = filedialog.askopenfilenames(filetypes=[("DWG Files", "*.dwg")])  # Open file dialog for DWG files
        for file in files:  # For each selected file
            if file not in self.file_list:  # Avoid duplicates
                self.file_list.append(file)  # Add to file list
                self.file_listbox.insert(tk.END, os.path.basename(file))  # Add filename to listbox
        self.update_process_button_state()  # Update button states
        self.update_backdrop_text()  # Update backdrop visibility

    # Handle files dropped onto the application window
    def drop_files(self, event):
        files = self.root.tk.splitlist(event.data)  # Parse dropped file paths
        for file in files:  # For each dropped file
            if file.endswith(".dwg") and file not in self.file_list:  # Check if DWG and not duplicate
                self.file_list.append(file)  # Add to file list
                self.file_listbox.insert(tk.END, os.path.basename(file))  # Add to listbox
        self.update_process_button_state()  # Update button states
        self.update_backdrop_text()  # Update backdrop

    # Remove selected files from the list
    def remove_files(self):
        selected_files = self.file_listbox.curselection()  # Get indices of selected items
        for index in reversed(selected_files):  # Iterate in reverse to maintain indices
            self.file_listbox.delete(index)  # Remove from listbox
            del self.file_list[index]  # Remove from file list
        self.update_process_button_state()  # Update button states
        self.update_backdrop_text()  # Update backdrop

    # Show or hide the backdrop text based on whether files are present
    def update_backdrop_text(self, event=None):
        if not self.file_list:  # If no files in list
            self.backdrop_text.pack(side="top", fill="both", expand=True)  # Show backdrop
        else:
            self.backdrop_text.pack_forget()  # Hide backdrop

    # Open file dialog to add Lisp files to the list
    def add_lisp(self):
        lisp_files = filedialog.askopenfilenames(filetypes=[("Lisp Files", "*.lsp")])  # Open dialog for Lisp files
        for lisp in lisp_files:  # For each selected Lisp
            lisp_path = os.path.abspath(lisp)  # Get absolute path
            # Only add new Lisp if not already in the list.
            if not any(item["path"] == lisp_path for item in self.lisp_files):  # Check for duplicates
                self.lisp_files.append({"path": lisp_path, "var": tk.BooleanVar(value=True)})  # Add with enabled checkbox
        self.refresh_lisp_list()  # Refresh display
        self.update_process_button_state()  # Update button states

    # Remove Lisp files that are unchecked
    def remove_lisp(self):
        # Remove entries where the checkbutton is unchecked.
        self.lisp_files = [item for item in self.lisp_files if item["var"].get()]  # Keep only checked items
        self.refresh_lisp_list()  # Refresh display
        self.update_process_button_state()  # Update button states

    # Enable or disable the process button based on selections
    def update_process_button_state(self):
        if self.file_list and any(item["var"].get() for item in self.lisp_files):  # If files and at least one Lisp selected
            self.process_button.config(state=tk.NORMAL)  # Enable button
        else:
            self.process_button.config(state=tk.DISABLED)  # Disable button

    # Start the file processing in a separate thread
    def start_processing(self):
        self.disable_buttons()  # Disable UI buttons during processing
        self.success_count = 0  # Reset success counter
        threading.Thread(target=self.process_files).start()  # Start processing thread

    # Main processing function that handles batch processing of DWG files with Lisp scripts
    def process_files(self):
        pythoncom.CoInitialize()  # Initialize COM for this thread
        total_files = len(self.file_list)  # Total number of files to process
        self.update_status("Initializing AutoCAD...", "blue")  # Update status
        try:
            acad_location = self.settings["autocad_location"]  # Get AutoCAD executable path
            if not os.path.exists(acad_location):  # Check if path exists
                self.update_status(f"AutoCAD.exe not found at {acad_location}", "red")  # Error message
                self.enable_buttons()  # Re-enable buttons
                return

            acad = win32com.client.Dispatch("AutoCAD.Application")  # Launch AutoCAD
            acad.Visible = False  # Hide AutoCAD window
            acad.WindowState = 1  # Minimize window
            hide_autocad_window()  # Ensure window is hidden

            # Use the order defined in lisp_files.
            selected_lisps = [item["path"] for item in self.lisp_files if item["var"].get()]  # Get selected Lisp paths

            for index, file in enumerate(self.file_list):  # Process each file
                self.update_status(f"Processing file: {os.path.basename(file)} ({index+1}/{total_files})", "blue")  # Update status
                self.update_progress(index+1, total_files)  # Update progress bar
                try:
                    self.run_lisp_process(acad, file, selected_lisps)  # Run Lisp on file
                    self.update_status(f"Process successful for file {file}", "green")  # Success message
                    self.success_count += 1  # Increment success count
                except Exception as e:  # Handle errors
                    error_str = str(e)
                    if "Open.Close" in error_str:  # Specific error handling
                        self.update_status(f"Error processing file {file}: The file could not be opened or closed.", "red")
                    elif "disconnected" in error_str:
                        self.update_status(f"Error processing file {file}: AutoCAD may have crashed.", "red")
                    else:
                        self.update_status(f"Error processing file {file}: {e}", "red")

            try:
                acad.Quit()  # Quit AutoCAD
            except Exception as quit_exception:
                self.update_status(f"Error quitting AutoCAD: {quit_exception}", "red")  # Handle quit error
        except Exception as e:
            self.update_status(f"Error initializing AutoCAD: {e}", "red")  # Initialization error
        finally:
            self.update_status(f"Processing complete: {self.success_count} of {total_files} files processed successfully.", "blue")  # Final status
            self.update_progress(total_files, total_files)  # Complete progress
            pythoncom.CoUninitialize()  # Uninitialize COM
            self.enable_buttons()  # Re-enable buttons

    # Run Lisp scripts on a single DWG file using AutoCAD
    def run_lisp_process(self, acad, file, selected_lisps):
        try:
            doc = self.safe_open_document(acad, file)  # Open the DWG document safely
            if doc:  # If document opened successfully
                for i, lisp in enumerate(selected_lisps):  # For each selected Lisp
                    lisp_path_fixed = lisp.replace("\\", "/")  # Fix path separators for Lisp
                    self.send_command_with_retry(acad, f'(load "{lisp_path_fixed}")\n')  # Load Lisp script
                    time.sleep(1)  # Wait for loading
                    self.send_command_with_retry(acad, f'(c:MyLispFunction)\n')  # Execute Lisp function
                    time.sleep(1)  # Wait for execution
                    # Extract the Lisp name for the status message.
                    lisp_name = os.path.basename(lisp)  # Get Lisp filename
                    self.update_status(f'{lisp_name} completed for file {os.path.basename(file)} (Lisp {i+1} of {len(selected_lisps)})', "blue")  # Status update
                self.send_command_with_retry(acad, '(command "_.QSAVE")\n')  # Save document
                time.sleep(2)  # Wait for save
                self.send_command_with_retry(acad, '(command "_.CLOSE")\n')  # Close document
                time.sleep(3)  # Wait for close
                if self.is_document_open(acad, file):  # Check if still open
                    self.update_status(f"Warning: Document did not close properly on first attempt for {file}", "orange")  # Warning
                    self.send_command_with_retry(acad, '(command "_.CLOSE")\n')  # Try close again
                    time.sleep(3)  # Wait
                if not self.is_document_open(acad, file):  # If closed
                    try:
                        doc.Close(SaveChanges=True)  # Ensure closed
                    except Exception as close_exception:
                        self.update_status(f"Suppressed final close error for {file}: {close_exception}", "orange")  # Suppress error
                else:
                    self.update_status(f"Warning: Document still appears open for {file}", "orange")  # Warning
            else:
                raise Exception("Failed to open document after multiple attempts")  # Raise error if couldn't open
        except Exception as e:
            raise e  # Re-raise exception

    # Check if a specific DWG file is currently open in AutoCAD
    def is_document_open(self, acad, file_path):
        try:
            for doc in acad.Documents:  # Iterate through open documents
                if os.path.normcase(doc.FullName) == os.path.normcase(file_path):  # Compare paths case-insensitively
                    return True  # Document is open
            return False  # Document not found
        except Exception:
            return False  # Assume not open on error

    # Attempt to open a DWG document with retries and delays
    def safe_open_document(self, acad, file, retries=5, delay=4):
        for attempt in range(retries):  # Try multiple times
            try:
                return acad.Documents.Open(file)  # Attempt to open document
            except Exception as e:
                if attempt < retries - 1:  # If not last attempt
                    self.update_status(f"Retrying to open file {file}... (Attempt {attempt+1}/{retries})", "orange")  # Status update
                    time.sleep(delay)  # Wait before retry
                else:
                    raise e  # Raise exception on final failure

    # Send a command to AutoCAD with retry logic
    def send_command_with_retry(self, acad, command, retries=3):
        for attempt in range(retries):  # Try multiple times
            try:
                acad.ActiveDocument.SendCommand(command)  # Send command to active document
                time.sleep(1)  # Wait for command to process
                return  # Success, exit
            except Exception as e:
                if attempt < retries - 1:  # If not last attempt
                    time.sleep(2)  # Wait before retry
                else:
                    raise e  # Raise exception on final failure

    # Update the status text area with a new message and color
    def update_status(self, status, color="blue"):
        self.root.after(0, self._set_status_text, status, color)  # Schedule update on main thread

    # Internal method to set status text with color tagging
    def _set_status_text(self, status, color):
        self.status_text.config(state="normal")  # Enable text editing
        self.status_text.insert("end", f"{status}\n")  # Insert new status line
        if color == "red":  # Apply color tags
            self.status_text.tag_configure("error", foreground="red")
            self.status_text.tag_add("error", "end-2l", "end-1c")
        elif color == "green":
            self.status_text.tag_configure("success", foreground="green")
            self.status_text.tag_add("success", "end-2l", "end-1c")
        elif color == "orange":
            self.status_text.tag_configure("warning", foreground="orange")
            self.status_text.tag_add("warning", "end-2l", "end-1c")
        else:
            self.status_text.tag_configure("info", foreground="blue")
            self.status_text.tag_add("info", "end-2l", "end-1c")
        self.status_text.config(state="disabled")  # Disable editing
        self.status_text.see("end")  # Scroll to end

    # Update the progress bar value
    def update_progress(self, current, total):
        progress_value = (current / total) * 100  # Calculate percentage
        self.root.after(0, self._set_progress, progress_value)  # Schedule update

    # Internal method to set progress bar value
    def _set_progress(self, value):
        self.progress["value"] = value  # Set progress value

    # Disable all interactive buttons during processing
    def disable_buttons(self):
        self.root.after(0, lambda: self._set_buttons_state(tk.DISABLED))  # Schedule disable

    # Enable all interactive buttons after processing
    def enable_buttons(self):
        self.root.after(0, lambda: self._set_buttons_state(tk.NORMAL))  # Schedule enable

    # Set the state of all buttons
    def _set_buttons_state(self, state):
        self.add_file_button.config(state=state)  # File buttons
        self.remove_file_button.config(state=state)
        self.add_lisp_button.config(state=state)  # Lisp buttons
        self.remove_lisp_button.config(state=state)
        self.process_button.config(state=state)  # Process button
        self.options_button.config(state=state)  # Options button

    # Open the options window for configuring AutoCAD location
    def open_options(self):
        if self.options_window and self.options_window.winfo_exists():  # Check if window already open
            return  # Do nothing if already open
        self.options_window = tk.Toplevel(self.root)  # Create new window
        self.options_window.title("Options")  # Set title
        self.options_window.geometry("400x200")  # Set size
        self.options_window.transient(self.root)  # Set as child window
        self.options_window.grab_set()  # Modal window
        self.root.update_idletasks()  # Update geometry
        x = self.root.winfo_x()  # Position relative to main window
        y = self.root.winfo_y()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        self.options_window.geometry(f"+{x + width//2 - 200}+{y + height//2 - 100}")  # Center window
        options_frame = tk.Frame(self.options_window, padx=10, pady=10)  # Frame for content
        options_frame.pack(fill=tk.BOTH, expand=True)
        autocad_label = ttk.Label(options_frame, text="AutoCAD Location:", font=("Helvetica", 12))  # Label
        autocad_label.pack(anchor="w", pady=5)
        self.autocad_entry = ttk.Entry(options_frame, font=("Helvetica", 12))  # Entry field
        self.autocad_entry.pack(fill=tk.X, pady=5, padx=10)
        if self.settings["autocad_location"]:  # If location set
            self.autocad_entry.insert(0, self.settings["autocad_location"])  # Insert current value
        else:
            self.autocad_entry.insert(0, self.find_autocad_location())  # Insert found location
        autocad_browse_button = ttk.Button(options_frame, text="Browse...", command=self.browse_autocad)  # Browse button
        autocad_browse_button.pack(pady=5)
        save_button = ttk.Button(options_frame, text="Save", command=self.save_options)  # Save button
        save_button.pack(pady=10)

    # Attempt to automatically find the AutoCAD executable location
    def find_autocad_location(self):
        possible_paths = [  # Common installation directories
            "C:\\Program Files\\Autodesk",
            "C:\\Program Files (x86)\\Autodesk"
        ]
        for path in possible_paths:  # Search each path
            for root_dir, dirs, files in os.walk(path):  # Walk directory tree
                if "acad.exe" in files:  # If acad.exe found
                    return os.path.join(root_dir, "acad.exe")  # Return full path
        return ""  # Return empty if not found

    # Open file dialog to browse for AutoCAD executable
    def browse_autocad(self):
        initial_dir = os.path.dirname(self.settings.get("autocad_location", ""))  # Start in current location directory
        if not initial_dir:  # If no current location
            initial_dir = "C:\\Program Files\\Autodesk\\"  # Default start directory
        filepath = filedialog.askopenfilename(initialdir=initial_dir, filetypes=[("AutoCAD Executable", "acad.exe")])  # Open dialog
        if filepath:  # If file selected
            self.autocad_entry.delete(0, tk.END)  # Clear entry
            self.autocad_entry.insert(0, filepath)  # Insert selected path

    # Save the AutoCAD location setting
    def save_options(self):
        autocad_location = self.autocad_entry.get()  # Get entered path
        if os.path.basename(autocad_location) == "acad.exe" and os.path.exists(autocad_location):  # Validate path
            self.settings["autocad_location"] = autocad_location  # Save to settings
            self.save_settings()  # Persist settings
            messagebox.showinfo("Settings Saved", "AutoCAD location has been updated.")  # Success message
        else:
            messagebox.showerror("Invalid Path", "The specified AutoCAD location is invalid.")  # Error message

# Main entry point of the application
if __name__ == "__main__":
    root = TkinterDnD.Tk()  # Create drag-and-drop enabled root window
    # Load Nordmin Logo 
    ico = Image.open(resource_path("logo\\Logo-Blue.png"))  # Load logo image
    
    # Resize the logo to fit the title bar (optional, adjust size as needed)
    resized_logo = ico.resize((500, 500), Image.Resampling.BICUBIC)  # Resize logo
    
    # Converting PNG logo to ICO format
    photo = ImageTk.PhotoImage(resized_logo)  # Convert to PhotoImage for icon


    root.wm_iconphoto(False, photo)  # Set window icon
    app = LispBatchProcessorApp(root)  # Create application instance
    root.mainloop()  # Start GUI event loop