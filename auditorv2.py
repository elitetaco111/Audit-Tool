import os
import sys
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
import pandas as pd
import download_helper
import csv
from ttkthemes import ThemedTk
import threading
import datetime
import tempfile

"""
Developed by Dave Nissly
Rally House Product Audit Tool
This tool is designed to help audit product images and metadata.
It allows users to load a CSV file, view product images, and mark them as correct or incorrect.
It also provides functionality to handle missing or incorrect data by allowing users to select the correct values from predefined lists.

github.com/elitetaco111/audit-tool

To Package: pyinstaller --onefile --noconsole --hidden-import=tkinter --add-data "ColorList.csv;." --add-data "LogoList.csv;." --add-data "TeamList.csv;." --add-data "ClassMappingList.csv;." --add-data "choose.png;." --add-data "back.png;." --add-data "background.png;." --add-data "Logos;Logos" --add-data "Colors;Colors" auditorv2.py
"""

TEMP_FOLDER = "TEMP"
LOGOS_FOLDER = "Logos"
COLORS_FOLDER = "Colors"

def resource_path(relative_path):
    """
    Get absolute path to resource, works for dev and for PyInstaller.
    """
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def find_image(folder, base_name):
    """
    Looks for an image file (.jpg or .png, case-insensitive) in the given folder matching base_name.
    Returns the full path if found, else None.
    """
    if not base_name or not isinstance(base_name, str):
        return None
    for ext in ['.jpg', '.png']:
        for file in os.listdir(resource_path(folder)):
            if file.lower() == f"{base_name.lower()}{ext}":
                return os.path.join(resource_path(folder), file)
    return None

class AuditApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Product Audit Tool --- Press Left arrow to mark wrong, Right arrow to mark right")
        self.root.state('zoomed')
        self.data = None
        self.index = 0
        self.choices = []
        self.images = []
        self.logo_imgs = []
        self.color_imgs = []
        self.missing_rows = []
        self.missing_index = 0
        self.data_missing = None
        self.in_missing_loop = False
        self.setup_ui()
        self.root.bind('<Left>', self.mark_wrong)
        self.root.bind('<Right>', self.mark_right)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)  # Handle window close

    def setup_ui(self):
        self.frame = tk.Frame(self.root)
        self.frame.pack(fill=tk.BOTH, expand=True)
        choose_img_path = resource_path("choose.png")
        if os.path.exists(choose_img_path):
            choose_img = Image.open(choose_img_path)
            choose_img = choose_img.resize((200, 200))
            self.tk_choose_img = ImageTk.PhotoImage(choose_img)
            self.btn_load = tk.Button(self.frame, image=self.tk_choose_img, command=self.load_csv, borderwidth=0)
        else:
            self.btn_load = tk.Button(self.frame, text="Load CSV", command=self.load_csv)
        self.btn_load.pack(pady=10)

        style = ttk.Style(self.root)
        style.theme_use("arc")
        style.configure("big.Horizontal.TProgressbar", thickness=30, troughcolor="#e0e0e0", background="#4a90e2")

        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(
            self.frame,
            variable=self.progress_var,
            maximum=100,
            length=600,
            style="big.Horizontal.TProgressbar"
        )
        self.progress_bar.pack(pady=30, side=tk.TOP)  # <-- Explicitly pack at the top

        self.canvas = tk.Canvas(self.frame, width=1920, height=1080, bg="white")
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas_font = ("Roboto", 24)

        # Add back button (hidden until images are shown)
        back_img_path = resource_path("back.png")
        if os.path.exists(back_img_path):
            back_img = Image.open(back_img_path)
            back_img = back_img.resize((148, 148))
            self.tk_back_img = ImageTk.PhotoImage(back_img)
            self.btn_back = tk.Button(self.frame, image=self.tk_back_img, command=self.undo_last, borderwidth=0)
        else:
            self.btn_back = tk.Button(self.frame, text="Back", command=self.undo_last)
        self.btn_back.place_forget()

    def load_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not file_path:
            return
        # Ensure TEMP folder exists
        if not os.path.exists(TEMP_FOLDER):
            os.makedirs(TEMP_FOLDER)
        self.btn_load.pack_forget()
        self.progress_var.set(0)
        self.progress_bar.pack(pady=30)
        self.progress_bar.lift()
        self.progress_bar.update_idletasks()

        self.data = pd.read_csv(file_path, dtype=str)

        # --- Preprocess: Separate parent and child records ---
        self.child_records = {}  # {parent_name: [child_rows]}
        parent_rows = []
        for idx, row in self.data.iterrows():
            name = str(row['Name']) if 'Name' in row else ""
            if " :" in name:
                parent_name = name.split(" :")[0]
                self.child_records.setdefault(parent_name, []).append(row.copy())
            else:
                parent_rows.append(row.copy())
        self.data = pd.DataFrame(parent_rows).reset_index(drop=True)
        total_images = len(self.data)

        # --- Create a temporary CSV for parent records only ---
        self.temp_parent_csv = tempfile.NamedTemporaryFile(delete=False, suffix=".csv")
        self.data.to_csv(self.temp_parent_csv.name, index=False)
        self.temp_parent_csv.close()

        # Start download in a background thread (using parent-only CSV)
        threading.Thread(
            target=self.download_images_thread,
            args=(self.temp_parent_csv.name, TEMP_FOLDER),
            daemon=True
        ).start()

        # Start polling for progress in the main thread
        self.poll_progress(total_images)

    def download_images_thread(self, parent_csv_path, temp_folder):
        download_helper.download_images(parent_csv_path, temp_folder, item_col='Name', picture_id_col='Picture ID')

    def poll_progress(self, total_images):
        downloaded = len([f for f in os.listdir(TEMP_FOLDER) if f.lower().endswith('.jpg') or f.lower().endswith('.png')])
        percent = (downloaded / total_images) * 100 if total_images else 0
        self.progress_var.set(percent)
        self.progress_bar.update_idletasks()
        self.progress_bar.lift()
        self.root.update_idletasks()
        if downloaded >= total_images:
            self.progress_bar.pack_forget()
            # Continue with rest of setup
            self.data.reset_index(drop=True, inplace=True)
            self.index = 0
            self.choices = []
            self.missing_rows = []
            self.missing_index = 0
            self.data_missing = None
            self.btn_load.pack_forget()
            bg_path = resource_path("background.png")
            if os.path.exists(bg_path):
                bg_img = Image.open(bg_path)
                bg_img = bg_img.resize((1920, 1080))
                self.tk_bg_img = ImageTk.PhotoImage(bg_img)
                # Draw background image as the first (bottom) item on the canvas
                self.bg_image_id = self.canvas.create_image(0, 0, anchor='nw', image=self.tk_bg_img)
                self.canvas.tag_lower(self.bg_image_id)  # Ensure it's at the very back

            self.show_image()
        else:
            self.root.after(100, lambda: self.poll_progress(total_images))

    def show_image(self):
        if self.data is None or self.index >= len(self.data):
            # Start missing info fix loop if needed
            if self.missing_rows:
                # --- Add warning popup before fix_missing_loop ---
                def show_missing_warning():
                    popup = tk.Toplevel(self.root)
                    popup.title("Missing Fields Detected")
                    popup.grab_set()
                    popup.transient(self.root)
                    popup.lift()
                    popup.focus_force()
                    self.root.unbind('<Left>')
                    self.root.unbind('<Right>')
                    self.root.attributes('-disabled', True)
                    frame = ttk.Frame(popup, padding=40)
                    frame.pack(fill=tk.BOTH, expand=True)
                    ttk.Label(frame, text="You are about to audit products with missing fields.\nYou must fix these products; the missing field will be autoselected for you.", font=(self.canvas_font)).pack(pady=20)
                    def on_ok():
                        popup.destroy()
                    ttk.Button(frame, text="OK", command=on_ok).pack(pady=10)
                    popup.protocol("WM_DELETE_WINDOW", lambda: None)  # Prevent closing
                    popup.wait_window()
                    self.root.attributes('-disabled', False)
                    self.root.bind('<Left>', self.mark_wrong)
                    self.root.bind('<Right>', self.mark_right)
                    self.root.focus_force()
                show_missing_warning()
                # -------------------------------------------------
                indices, rows = zip(*self.missing_rows)
                self.data_missing = pd.DataFrame(list(rows), index=list(indices)).astype('object')
                self.missing_index = 0
                self.in_missing_loop = True  # NEW: start missing-loop mode
                self.fix_missing_loop()
                return
            self.finish()
            return
        row = self.data.iloc[self.index]
        # Check for missing/invalid fields
        logo_id = row['Logo ID'] if pd.notna(row['Logo ID']) else ""
        class_mapping = row['Class Mapping'] if pd.notna(row['Class Mapping']) else ""
        color_id = row['Parent Color Primary'] if pd.notna(row['Parent Color Primary']) else ""
        team_league = row['Team League Data'] if pd.notna(row['Team League Data']) else ""

        # Track missing info rows, but show them normally
        logo_id_missing = (
            not logo_id or
            "-tbd" in logo_id.lower() or
            logo_id.strip() == "- None -"
        )
        if (
            logo_id_missing or
            not class_mapping or
            not color_id or
            not team_league
        ):
            # Only add if not already in missing_rows AND not already fixed in choices
            already_fixed = any(
                entry[1].name == self.index and entry[0] in ('accepted', 'to_audit')
                for entry in self.choices
            )
            if not any(idx == self.index for idx, _ in self.missing_rows) and not already_fixed:
                self.missing_rows.append((self.index, row.copy()))
            self.index += 1
            self.show_image()
            return

        self.display_row(row)
        #self.btn_back.place(x=205, y=750)

    def display_row(self, row):
        logo_id = row['Logo ID'] if pd.notna(row['Logo ID']) else ""
        class_mapping = row['Class Mapping'] if pd.notna(row['Class Mapping']) else ""
        color_id = row['Parent Color Primary'] if pd.notna(row['Parent Color Primary']) else ""
        team_league = row['Team League Data'] if pd.notna(row['Team League Data']) else ""
        img_path = os.path.join(TEMP_FOLDER, f"{row['Name']}.jpg")
        logo_path = find_image(LOGOS_FOLDER, logo_id)
        color_path = find_image(COLORS_FOLDER, color_id)

        # Remove all items except the background image
        items = self.canvas.find_all()
        for item in items:
            if not hasattr(self, 'bg_image_id') or item != self.bg_image_id:
                self.canvas.delete(item)

        # Main product image at natural resolution (511x730)
        img_x, img_y = 0, 0
        if os.path.exists(img_path):
            img = Image.open(img_path)
            img = img.resize((511, 730))  # Ensure natural resolution
            self.tk_img = ImageTk.PhotoImage(img)
            self.canvas.create_image(img_x, img_y, anchor='nw', image=self.tk_img)
        else:
            self.canvas.create_text(img_x + 100, img_y + 100, text="Image not found", anchor='nw', font=self.canvas_font)

        # Remove any previous button window from the canvas
        if hasattr(self, 'btn_back_canvas_id'):
            self.canvas.delete(self.btn_back_canvas_id)

        # Create the button directly on the canvas, matching the image size
        btn_x = img_x + (511 // 2)  # Center of product image
        btn_y = img_y + 730
        self.btn_back.config(bg="white", activebackground="white", highlightthickness=0, bd=0)
        self.btn_back_canvas_id = self.canvas.create_window(
            btn_x, btn_y, anchor='n', window=self.btn_back, width=100, height=44
        )

        # Info boxes, at least 50px to the right of the product image
        x_offset = 511 + 50
        y_offset = 10
        box_height = 90

        # Logo ID
        self.canvas.create_text(x_offset, y_offset, anchor='nw', text=f"Logo ID: {logo_id}", font=self.canvas_font)
        if logo_path and os.path.exists(logo_path):
            logo_img = Image.open(logo_path).resize((200, 200))
            self.tk_logo = ImageTk.PhotoImage(logo_img)
            self.canvas.create_image(x_offset+100, y_offset+40, anchor='nw', image=self.tk_logo)
        y_offset += box_height + 180

        # Class Mapping
        self.canvas.create_text(x_offset, y_offset, anchor='nw', text=f"Class Mapping: {class_mapping}", font=self.canvas_font)
        y_offset += box_height

        # Parent Color Primary
        self.canvas.create_text(x_offset, y_offset, anchor='nw', text=f"Parent Color Primary: {color_id}", font=self.canvas_font)
        if color_path and os.path.exists(color_path):
            color_img = Image.open(color_path).resize((200, 200))
            self.tk_color = ImageTk.PhotoImage(color_img)
            self.canvas.create_image(x_offset+100, y_offset+40, anchor='nw', image=self.tk_color)
        y_offset += box_height + 180
        self.canvas.create_text(x_offset, y_offset, anchor='nw', text=f"Team League Data: {team_league}", font=self.canvas_font)
        y_offset += box_height

        # Add Style Number label on canvas
        self.canvas.create_text(x_offset, y_offset, anchor='nw', text="Style Number:", font=self.canvas_font)

        # Add Style Number (shows Name column value) as a copyable Entry
        style_number = row['Name'] if pd.notna(row['Name']) else ""
        if hasattr(self, 'style_entry') and self.style_entry.winfo_exists():
            self.style_entry.destroy()
        self.style_entry = ttk.Entry(self.frame, font=self.canvas_font, width=30)
        self.style_entry.insert(0, style_number)
        self.style_entry.config(state='readonly')
        # Place the entry box just to the right of the label
        self.style_entry.place(x=x_offset + 220, y=y_offset)

    def fix_missing_loop(self):
        if self.missing_index >= len(self.data_missing):
            self.in_missing_loop = False  # NEW: exit missing-loop mode
            self.finish()
            return
        row = self.data_missing.iloc[self.missing_index]
        # Detect missing/invalid fields for autoselection
        missing_fields = []
        if not row['Logo ID'] or pd.isna(row['Logo ID']):
            missing_fields.append("Logo ID")
        if not row['Class Mapping'] or pd.isna(row['Class Mapping']):
            missing_fields.append("Class Mapping")
        if not row['Parent Color Primary'] or pd.isna(row['Parent Color Primary']):
            missing_fields.append("Parent Color Primary")
        if not row['Team League Data'] or pd.isna(row['Team League Data']):
            missing_fields.append("Team League Data")
        #add more checks if needed here

        self.display_row(row)
        # self.btn_back.place_forget()  # REMOVE: keep Back visible during missing-loop

        self._popup_open = True
        wrong_info = self.ask_wrong_fields(row, preselected_fields=missing_fields)
        self._popup_open = False

        # NEW: allow "Back" from popup to go to previous missing product
        if isinstance(wrong_info, dict) and wrong_info.get("back"):
            if self.missing_index > 0:
                self.missing_index -= 1
            # Optional: remove any previous choice for this row to avoid duplicates
            target_idx = self.data_missing.iloc[self.missing_index].name
            self.choices = [c for c in self.choices if getattr(c[1], "name", None) != target_idx]
            self.fix_missing_loop()
            return

        wrong_fields = wrong_info["fields"] if isinstance(wrong_info, dict) else []
        wrong_details = wrong_info["details"] if isinstance(wrong_info, dict) else {}
        if not wrong_fields or (isinstance(wrong_fields, list) and all(f.strip() == "" for f in wrong_fields)):
            messagebox.showwarning("Input required", "You must select at least one field that is wrong before continuing.")
            return
        # Update the row in self.data_missing and self.data with corrected values
        row_idx = row.name  # This is now the original index in self.data
        for field, value in wrong_details.items():
            self.data_missing.at[row_idx, field] = str(value)
            self.data.at[row_idx, field] = str(value)
        self.choices.append(('to_audit', row, False, wrong_fields, wrong_details))
        self.missing_index += 1
        self.fix_missing_loop()

    def mark_right(self, event=None):
        # Prevent advancing if popup is open
        if getattr(self, "_popup_open", False):
            return
        self.choices.append(('accepted', self.data.iloc[self.index], False))
        self.index += 1
        self.show_image()

    def mark_wrong(self, event=None):
        # Prevent multiple popups or input while popup is open
        if getattr(self, "_popup_open", False):
            return
        self._popup_open = True
        row = self.data.iloc[self.index]
        wrong_info = self.ask_wrong_fields(row)
        self._popup_open = False
        wrong_fields = wrong_info["fields"] if isinstance(wrong_info, dict) else []
        wrong_details = wrong_info["details"] if isinstance(wrong_info, dict) else {}
        if not wrong_fields or (isinstance(wrong_fields, list) and all(f.strip() == "" for f in wrong_fields)):
            messagebox.showwarning("Input required", "You must select at least one field that is wrong before continuing.")
            return
        # Update the row in self.data with corrected values
        for field, value in wrong_details.items():
            self.data.at[row.name, field] = value
        self.choices.append(('to_audit', row, False, wrong_fields, wrong_details))
        self.index += 1
        self.show_image()

    def ask_wrong_fields(self, row, preselected_fields=None, force_fix=False):
        fields = [
            "Logo ID",
            "Class Mapping",
            "Parent Color Primary",
            "Team League Data",
            "Other"
        ]
        # include a 'back' flag
        result = {"value": None, "details": {}, "back": False}

        # Style number to expose in popups
        style_number = row['Name'] if pd.notna(row['Name']) else ""

        def copy_to_clip(text):
            try:
                self.root.clipboard_clear()
                self.root.clipboard_append(text)
                # Optional: brief visual feedback can be added here if wanted
            except Exception:
                pass

        def show_popup():
            popup = tk.Toplevel(self.root)
            style = ttk.Style(popup)
            style.theme_use("arc")
            popup.configure(bg="#f7f7f7")
            popup.title("Select the field(s) that are wrong")
            popup.grab_set()
            popup.transient(self.root)
            popup.lift()
            # Position the popup in the top right corner
            self.root.update_idletasks()
            root_x = self.root.winfo_rootx()
            root_y = self.root.winfo_rooty()
            root_width = self.root.winfo_width()
            popup_width = 500  # Adjust as needed
            popup_height = 420 # Adjust as needed
            x = root_x + root_width - popup_width - 40
            y = root_y + 40
            popup.geometry(f"{popup_width}x{popup_height}+{x}+{y}")
            popup.focus_force()
            self.root.unbind('<Left>')
            self.root.unbind('<Right>')
            self.root.attributes('-disabled', True)

            main_frame = ttk.Frame(popup, padding=60)
            main_frame.pack(fill=tk.BOTH, expand=True)

            # NEW: Style Number copy bar (double-click to copy)
            copy_bar = ttk.Frame(main_frame)
            copy_bar.pack(fill=tk.X, pady=(0, 12))
            ttk.Label(copy_bar, text="Style Number:").pack(side=tk.LEFT)
            sn_var = tk.StringVar(value=style_number)
            sn_entry = ttk.Entry(copy_bar, textvariable=sn_var, state='readonly', width=32)
            sn_entry.pack(side=tk.LEFT, padx=(6, 6))
            sn_entry.bind('<Double-Button-1>', lambda e: copy_to_clip(sn_var.get()))
            ttk.Button(copy_bar, text="Copy", command=lambda: copy_to_clip(sn_var.get())).pack(side=tk.LEFT)

            ttk.Label(main_frame, text="Which field(s) are wrong?").pack(pady=(0, 10), anchor='w')
            vars = {field: tk.BooleanVar(value=False) for field in fields}
            if preselected_fields:
                for field in preselected_fields:
                    if field in vars:
                        vars[field].set(True)
            for field in fields:
                ttk.Checkbutton(main_frame, text=field, variable=vars[field]).pack(anchor='w', pady=2)

            custom_var = tk.StringVar()
            entry = ttk.Entry(main_frame, textvariable=custom_var)

            def on_other_checked(*args):
                if vars["Other"].get():
                    entry.pack(pady=5, anchor='w')
                else:
                    entry.pack_forget()
            vars["Other"].trace_add("write", on_other_checked)

            def submit():
                selected = [field for field in fields if vars[field].get() and field != "Other"]
                if vars["Other"].get():
                    other_text = custom_var.get().strip()
                    if other_text:
                        selected.append(other_text)
                    else:
                        selected.append("Other")
                # --- ENFORCE LOGIC HERE ---
                if "Team League Data" in selected:
                    if "Parent Color Primary" not in selected:
                        selected.append("Parent Color Primary")
                    if "Logo ID" not in selected:
                        selected.append("Logo ID")
                if not selected or (len(selected) == 1 and selected[0] == "Other"):
                    messagebox.showwarning("Input required", "Please select at least one field or enter a value for 'Other'.", parent=popup)
                    return
                result["value"] = selected
                popup.destroy()

            def go_back():
                result["back"] = True
                popup.destroy()

            # prevent closing by X: reopen the dialog
            def on_close():
                popup.destroy()
                show_popup()

            popup.protocol("WM_DELETE_WINDOW", on_close)

            # Buttons: show Back only during the missing loop
            btn_bar = ttk.Frame(main_frame)
            btn_bar.pack(pady=10, anchor='e', fill=tk.X)
            if getattr(self, "in_missing_loop", False):
                ttk.Button(btn_bar, text="Back", command=go_back).pack(side=tk.LEFT)
            ttk.Button(btn_bar, text="OK", command=submit).pack(side=tk.RIGHT)

            popup.bind('<Return>', lambda event: submit())
            popup.wait_window()
            self.root.attributes('-disabled', False)
            self.root.bind('<Left>', self.mark_wrong)
            self.root.bind('<Right>', self.mark_right)
            self.root.focus_force()

        show_popup()

        if result.get("back"):
            return {"back": True}

        wrong_fields = result["value"]
        wrong_details = {}

        def select_from_list(title, label, options, show_images=False, image_folder=None):
            sel_popup = tk.Toplevel(self.root)
            style = ttk.Style(sel_popup)
            style.theme_use("arc")
            sel_popup.configure(bg="#f7f7f7")
            sel_popup.title(title)
            sel_popup.grab_set()
            sel_popup.focus_force()
            self.root.attributes('-disabled', True)

            main_frame = ttk.Frame(sel_popup, padding=20)
            main_frame.pack(fill=tk.BOTH, expand=True)

            # NEW: Style Number copy bar in selector popups too
            copy_bar = ttk.Frame(main_frame)
            copy_bar.pack(fill=tk.X, pady=(0, 8))
            ttk.Label(copy_bar, text="Style Number:").pack(side=tk.LEFT)
            sn_entry = ttk.Entry(copy_bar, width=32)
            sn_entry.insert(0, style_number)
            sn_entry.config(state='readonly')
            sn_entry.pack(side=tk.LEFT, padx=(6, 6))
            sn_entry.bind('<Double-Button-1>', lambda e: copy_to_clip(style_number))
            ttk.Button(copy_bar, text="Copy", command=lambda: copy_to_clip(style_number)).pack(side=tk.LEFT)

            ttk.Label(main_frame, text=label).pack(pady=(0, 10), anchor='w')
            search_var = tk.StringVar()
            search_entry = ttk.Entry(main_frame, textvariable=search_var, width=60)
            search_entry.pack(pady=(0, 10), anchor='w')
            search_entry.focus_set()

            filtered_options = options.copy()
            listbox_var = tk.StringVar(value=filtered_options)
            listbox = tk.Listbox(
                main_frame,
                listvariable=listbox_var,
                width=80,
                height=min(15, len(filtered_options)),
                exportselection=False,
                bg="#f7f7f7",
                relief=tk.FLAT,
                highlightthickness=0,
                borderwidth=0
            )
            listbox.pack(side=tk.LEFT, pady=10, fill=tk.Y)

            image_label = None
            img_cache = {}
            def show_logo_img(event):
                sel = listbox.curselection()
                if sel and show_images and image_folder:
                    logo_id = filtered_options[sel[0]]
                    img_path = find_image(image_folder, logo_id)
                    if img_path and os.path.exists(img_path):
                        img = Image.open(img_path)
                        img = img.resize((150, 150))
                        tk_img = ImageTk.PhotoImage(img)
                        img_cache["img"] = tk_img
                        image_label.config(image=tk_img, text="")
                    else:
                        image_label.config(image="", text="No image found")
                elif image_label:
                    image_label.config(image="", text="")
            if show_images and image_folder:
                image_label = ttk.Label(main_frame)
                image_label.pack(side=tk.LEFT, padx=(10, 0), pady=10, fill=tk.BOTH, expand=True)
                listbox.bind("<<ListboxSelect>>", show_logo_img)

            def filter_options(*args):
                nonlocal filtered_options
                search = search_var.get().lower()
                filtered_options = [opt for opt in options if search in opt.lower()]
                listbox_var.set(filtered_options)
                if filtered_options:
                    listbox.selection_clear(0, tk.END)
                    listbox.selection_set(0)
                    show_logo_img(None)
            search_var.trace_add("write", filter_options)

            local = {"value": None}
            def on_select(event=None):
                sel = listbox.curselection()
                if sel:
                    local["value"] = filtered_options[sel[0]]
                    sel_popup.destroy()

            btn_frame = ttk.Frame(main_frame)
            btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))
            ttk.Button(btn_frame, text="OK", command=on_select).pack()
            sel_popup.bind('<Return>', on_select)
            sel_popup.wait_window()
            self.root.attributes('-disabled', False)
            self.root.focus_force()
            return local["value"]

        # CSV loaders
        def load_csv_column(filename, colname, filter_col=None, filter_val=None):
            path = resource_path(filename)
            values = []
            if not os.path.exists(path):
                return values
            with open(path, newline='', encoding='utf-8') as csvfile:
                reader = csv.DictReader(csvfile)
                for r in reader:
                    if filter_col and filter_val and r.get(filter_col) != filter_val:
                        continue
                    values.append(r[colname])
            return sorted(set(values))

        # Always handle Team League Data first if selected
        if "Team League Data" in wrong_fields:
            team_options = load_csv_column("TeamList.csv", "Team League Data")
            new_team = select_from_list("Select Team", "Select the correct Team League Data:", team_options)
            if new_team:
                wrong_details["Team League Data"] = new_team
                row = row.copy()
                row['Team League Data'] = new_team
        if "Logo ID" in wrong_fields:
            logo_options = load_csv_column("LogoList.csv", "Logo ID", filter_col="Team League Data", filter_val=row['Team League Data'])
            new_logo = select_from_list(
                "Select Logo ID",
                f"Select the correct Logo ID for team '{row['Team League Data']}':",
                logo_options,
                show_images=True,
                image_folder=LOGOS_FOLDER
            )
            if new_logo:
                wrong_details["Logo ID"] = new_logo
        if "Parent Color Primary" in wrong_fields:
            color_options = load_csv_column("ColorList.csv", "Parent Color Primary", filter_col="Team League Data", filter_val=row['Team League Data'])
            new_color = select_from_list("Select Parent Color Primary", f"Select the correct Parent Color Primary for team '{row['Team League Data']}':", color_options)
            if new_color:
                wrong_details["Parent Color Primary"] = new_color
        if "Class Mapping" in wrong_fields:
            class_options = load_csv_column("ClassMappingList.csv", "Name")
            new_class = select_from_list("Select Class Mapping", "Select the correct Class Mapping:", class_options)
            if new_class:
                wrong_details["Class Mapping"] = new_class

        return {"fields": wrong_fields, "details": wrong_details}

    def download_images_with_progress(self, file_path, temp_folder, total_images):
        # Start download in main thread (for simplicity)
        download_helper.download_images(file_path, temp_folder, item_col='Name', picture_id_col='Picture ID')
        # Poll for progress
        while True:
            downloaded = len([f for f in os.listdir(temp_folder) if f.lower().endswith('.jpg') or f.lower().endswith('.png')])
            percent = (downloaded / total_images) * 100 if total_images else 0
            self.progress_var.set(percent)
            self.progress_bar.update_idletasks()
            self.progress_bar.lift()
            self.root.update_idletasks()
            if downloaded >= total_images:
                break
            self.root.after(10)  # Wait a bit before checking again

    def undo_last(self):
        # If we are fixing missing rows, go back within that list
        if getattr(self, "in_missing_loop", False) and not getattr(self, "_popup_open", False):
            if self.missing_index > 0:
                self.missing_index -= 1
            self.fix_missing_loop()
            return

        # Undo the last user action (not auto-rejected)
        for i in range(len(self.choices) - 1, -1, -1):
            entry = self.choices[i]
            status, row, auto_rejected = entry[:3]
            if not auto_rejected:
                self.choices.pop(i)
                self.index = row.name
                self.missing_rows = [(idx, r) for idx, r in self.missing_rows if idx != self.index]
                self.show_image()
                return
        # If nothing to undo, do nothing

    def finish(self):
        self.save_outputs()
        messagebox.showinfo("Done", "Audit complete!\nFile saved as to_audit.csv.")
        self.root.quit()

    def save_outputs(self):
        if self.data is None or self.data.empty:
            return
        exclude_cols = {"Picture ID", "Image Assignment"}
        output_rows = []
        for _, parent_row in self.data.iterrows():
            output_rows.append(parent_row.copy())
            parent_name = str(parent_row['Name']) if 'Name' in parent_row else ""
            if hasattr(self, 'child_records') and parent_name in self.child_records:
                for child_row in self.child_records[parent_name]:
                    # Copy all parent values to child, except Name
                    new_child = parent_row.copy()
                    new_child['Name'] = child_row['Name']
                    output_rows.append(new_child)
        output_df = pd.DataFrame(output_rows)
        output_df = output_df[[col for col in output_df.columns if col not in exclude_cols]].copy()
        today_str = datetime.datetime.now().strftime("%m/%d/%Y")
        output_df["Flash Sale Date"] = today_str
        output_df.to_csv("to_audit.csv", index=False)
        # Delete the entire TEMP folder
        if os.path.exists(TEMP_FOLDER):
            try:
                shutil.rmtree(TEMP_FOLDER)
            except Exception as e:
                print(f"Failed to delete TEMP folder: {e}")

    def on_close(self):
        self.save_outputs()
        self.root.destroy()

if __name__ == "__main__":
    root = ThemedTk(theme="arc")
    app = AuditApp(root)
    try:
        root.mainloop()
    finally:
        app.save_outputs()