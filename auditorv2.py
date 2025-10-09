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
import json 

"""
Developed by Dave Nissly
Rally House Product Audit Tool
This tool is designed to help audit product images and metadata.
It allows users to load a CSV file, view product images, and mark them as correct or incorrect.
It also provides functionality to handle missing or incorrect data by allowing users to select the correct values from predefined lists.

github.com/elitetaco111/audit-tool

To Package: pyinstaller --onefile --noconsole --hidden-import=tkinter --add-data "ColorList.csv;." --add-data "LogoList.csv;." --add-data "TeamList.csv;." --add-data "ClassMappingList.csv;." --add-data "choose.png;." --add-data "back.png;." --add-data "background.png;." --add-data "Logos;Logos" --add-data "Colors;Colors" auditorv2.py

Note: Setting wrong image makes it disappear from the audit flow and changes will not be applied, can change if needed
"""
#TODO:

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
        self.progress_label = None
        self._app_quitting = False
        # ---- session/resume additions ----
        self.original_csv_path = None
        self.session_manifest_path = None
        self.session_data_csv_path = None
        self.temp_folder = TEMP_FOLDER  # per-session subfolder will override this
        self.completed = False
        # Track products marked as wrong image (parent Name values)
        self.wrong_image_names = set()
        # NEW: track download lifecycle and expected names
        self.download_done = False
        self.expected_names = []
        # Map each Name (parent or child) to its Internal ID from the original CSV
        self.name_to_internal_id = {}
        # Background image state
        self.bg_original = None
        self.bg_image_id = None
        self.tk_bg_img = None

    # Helper: place a popup on the same screen as the root
    def _place_popup(self, popup, width, height, align="center", margin=40):
        try:
            self.root.update_idletasks()
            rx = self.root.winfo_rootx()
            ry = self.root.winfo_rooty()
            rw = self.root.winfo_width()
            rh = self.root.winfo_height()
            if align == "top-right":
                x = rx + max(0, rw - width - margin)
                y = ry + margin
            elif align == "top-center":
                x = rx + (rw - width) // 2
                y = ry + margin
            else:  # center
                x = rx + (rw - width) // 2
                y = ry + (rh - height) // 2
            popup.geometry(f"{int(width)}x{int(height)}+{int(x)}+{int(y)}")
        except Exception:
            # Fallback: just center on default screen
            popup.geometry(f"{int(width)}x{int(height)}")

    def _get_missing_fields(self, row):
        fields = []
        def _val(key):
            v = row.get(key) if hasattr(row, "get") else row[key]
            return v if pd.notna(v) else ""
        logo_val = str(_val("Logo ID")).strip()
        if (not logo_val) or logo_val.lower() == "- none -" or "-tbd" in logo_val.lower():
            fields.append("Logo ID")
        if not str(_val("Class Mapping")).strip():
            fields.append("Class Mapping")
        if not str(_val("Parent Color Primary")).strip():
            fields.append("Parent Color Primary")
        if not str(_val("Team League Data")).strip():
            fields.append("Team League Data")
        return fields

    def _filter_missing_rows_after_resume(self):
        try:
            audited_indices = {getattr(c[1], "name", None) for c in self.choices if c and c[0] in ("accepted", "to_audit", "wrong_image")}
            filtered = []
            for idx, _row in (self.missing_rows or []):
                if idx is None or idx < 0 or idx >= len(self.data):
                    continue
                current = self.data.iloc[idx]
                if self._get_missing_fields(current) and idx not in audited_indices:
                    filtered.append((idx, current.copy()))
            self.missing_rows = filtered
        except Exception:
            # Leave as-is if anything goes wrong
            pass

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
        self.canvas_font = ("Roboto", 18)
        # Resize background with the canvas
        self.canvas.bind("<Configure>", self._on_canvas_resize)

        # Add a Save & Quit button pinned to the top-right
        self.btn_save_quit = ttk.Button(self.frame, text="Save and Quit", command=self.save_and_quit)
        self.btn_save_quit.place(relx=1.0, x=-20, y=10, anchor='ne')
        self.btn_save_quit.lift()

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
        self.original_csv_path = file_path

        # Build per-session paths
        def _session_paths(csv_path):
            # Store all session artifacts inside the session TEMP subfolder
            base_name = os.path.splitext(os.path.basename(csv_path))[0]
            temp_sub = os.path.join(TEMP_FOLDER, base_name)
            os.makedirs(temp_sub, exist_ok=True)
            session_manifest = os.path.join(temp_sub, "audit_session.json")
            session_data_csv = os.path.join(temp_sub, "audit_session.data.csv")
            return session_manifest, session_data_csv, temp_sub

        self.session_manifest_path, self.session_data_csv_path, self.temp_folder = _session_paths(file_path)

        # Ask to resume if a previous session exists
        resumed = False
        if os.path.exists(self.session_manifest_path):
            if messagebox.askyesno("Resume audit?", "A previous session was found for this CSV. Resume where you left off?", parent=self.root):
                try:
                    with open(self.session_manifest_path, "r", encoding="utf-8") as f:
                        m = json.load(f)
                    # restore paths (in case they moved)
                    self.original_csv_path = m.get("original_csv_path", file_path)
                    self.session_data_csv_path = m.get("data_csv_path", self.session_data_csv_path)
                    self.temp_folder = m.get("temp_folder", self.temp_folder)

                    # load working parent rows
                    self.data = pd.read_csv(self.session_data_csv_path, dtype=str)

                    # rebuild child mapping from original CSV
                    self.child_records = self._build_child_records(self.original_csv_path)

                    # restore progress
                    self.index = int(m.get("index", 0))
                    self.in_missing_loop = m.get("phase", "main") == "missing"
                    self.missing_index = int(m.get("missing_index", 0))
                    # Force indices to Python ints
                    mr_idx = [int(i) for i in m.get("missing_rows_indices", [])]
                    self.missing_rows = [(idx, self.data.iloc[idx].copy()) for idx in mr_idx if 0 <= idx < len(self.data)]

                    # restore choices
                    self.choices = []
                    for rec in m.get("choices", []):
                        ridx = int(rec.get("row_index", -1))
                        if 0 <= ridx < len(self.data):
                            row = self.data.iloc[ridx]
                            tup = [rec.get("status"), row, bool(rec.get("auto", False))]
                            if "wrong_fields" in rec:
                                tup.append(rec["wrong_fields"])
                            if "wrong_details" in rec:
                                tup.append(rec["wrong_details"])
                            self.choices.append(tuple(tup))
                    # restore wrong image selections
                    self.wrong_image_names = set(m.get("wrong_images", []))
                    # NEW: ensure saved missing list only includes rows still missing and not audited
                    self._filter_missing_rows_after_resume()
                    # Build Name -> Internal ID map from original CSV for reliable output
                    self.name_to_internal_id = self._build_name_to_id(self.original_csv_path)
                    resumed = True
                except Exception as e:
                    messagebox.showwarning("Resume failed", f"Could not resume session. Starting a new one.\n\n{e}", parent=self.root)
        # Ensure per-session TEMP folder exists
        os.makedirs(self.temp_folder, exist_ok=True)

        self.btn_load.pack_forget()
        self.progress_var.set(0)
        self.progress_bar.pack(pady=30)
        self.progress_bar.lift()
        self.progress_bar.update_idletasks()

        if not resumed:
            self.data = pd.read_csv(file_path, dtype=str)

            # Preprocess: Separate parent and child records
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
            # NEW: expected names list for download verification
            self.expected_names = [str(n) if pd.notna(n) else "" for n in self.data.get('Name', [])]

            # Create a persistent CSV for parent records only
            self.data.to_csv(self.session_data_csv_path, index=False)

            parent_csv_for_dl = self.session_data_csv_path

            # NEW: Build Name -> Internal ID map (parents and children) from the original CSV
            self.name_to_internal_id = self._build_name_to_id(self.original_csv_path)
        else:
            total_images = len(self.data)
            # NEW: expected names list for download verification
            self.expected_names = [str(n) if pd.notna(n) else "" for n in self.data.get('Name', [])]
            parent_csv_for_dl = self.session_data_csv_path
            # Ensure Name -> Internal ID map exists on resume
            if not getattr(self, "name_to_internal_id", None):
                self.name_to_internal_id = self._build_name_to_id(self.original_csv_path)
        # Start download in a background thread (idempotent; will skip existing images)
        threading.Thread(
            target=self.download_images_thread,
            args=(parent_csv_for_dl, self.temp_folder),
            daemon=True
        ).start()

        # Start polling for progress in the main thread
        self.poll_progress(total_images)

    def download_images_thread(self, parent_csv_path, temp_folder):
        try:
            download_helper.download_images(parent_csv_path, temp_folder, item_col='Name', picture_id_col='Picture ID')
        finally:
            # NEW: mark download complete so we can reconcile failures
            self.download_done = True

    def poll_progress(self, total_images):
        # NEW: use expected count (parents) for progress display
        total_expected = len(self.expected_names) if self.expected_names else total_images
        downloaded = len([f for f in os.listdir(self.temp_folder) if f.lower().endswith('.jpg') or f.lower().endswith('.png')])

        if not self.download_done:
            percent = (downloaded / total_expected) * 100 if total_expected else 0
            # Keep UI responsive and show progress until thread completes
            self.progress_var.set(min(percent, 99.0))
            self.progress_bar.update_idletasks()
            self.progress_bar.lift()
            self.root.update_idletasks()
            self.root.after(100, lambda: self.poll_progress(total_images))
            return

        # Download thread finished: reconcile failures (missing files)
        present_basenames = {os.path.splitext(f)[0] for f in os.listdir(self.temp_folder)
                             if f.lower().endswith('.jpg') or f.lower().endswith('.png')}
        expected_set = set(self.expected_names)
        failed_names = sorted(expected_set - present_basenames)

        # Add failed downloads to "wrong images" and exclude from audit flow
        if failed_names:
            self.wrong_image_names.update(failed_names)

        # Wrap up progress UI
        self.progress_var.set(100)
        self.progress_bar.update_idletasks()
        self.progress_bar.pack_forget()

        # Continue with rest of setup
        self.data.reset_index(drop=True, inplace=True)
        self.missing_rows = self.missing_rows or []
        self.data_missing = None
        self.btn_load.pack_forget()
        # Load and draw background stretched to canvas size
        self._load_bg_image()
        self._update_bg_image()

        self.show_image()

    def show_image(self):
        if self.data is None or self.index >= len(self.data):
            if self.missing_rows:
                # Use native messagebox so OK button isn't tiny
                messagebox.showinfo(
                    "Missing Fields Detected",
                    "You are about to audit products with missing fields.\nYou must fix these products; the missing field will be autoselected for you.",
                    parent=self.root
                )
                if getattr(self, "_app_quitting", False):
                    return
                # -------------------------------------------------
                indices, rows = zip(*self.missing_rows)
                self.data_missing = pd.DataFrame(list(rows), index=list(indices)).astype('object')
                self.missing_index = 0
                self.in_missing_loop = True
                self.fix_missing_loop()
                return
            self.finish()
            return
        row = self.data.iloc[self.index]

        # NEW: skip rows whose image failed to download (removed from audit flow)
        try:
            name_val = str(row['Name']) if 'Name' in row else ""
        except Exception:
            name_val = ""
        if name_val in self.wrong_image_names:
            self.index += 1
            self.show_image()
            return

        # Use unified missing detection
        missing_fields = self._get_missing_fields(row)
        if missing_fields:
            # Only add if not already in missing_rows AND not already fixed in choices
            already_fixed = any(
                entry[1].name == self.index and entry[0] in ('accepted', 'to_audit', 'wrong_image')
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
        display_name = row['Web Display Name'] if pd.notna(row['Web Display Name']) else ""
        marketing_event = row['Marketing Event'] if 'Marketing Event' in row and pd.notna(row['Marketing Event']) else ""
        silhouette = row['Silhouette'] if pd.notna(row['Silhouette']) else ""
        web_style = row['Web Style'] if pd.notna(row['Web Style']) else ""
        img_path = os.path.join(self.temp_folder, f"{row['Name']}.jpg")  # <-- was TEMP_FOLDER
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
        x_offset = 511 + 25
        y_offset = 10
        box_height = 70

        # Logo ID
        self.canvas.create_text(x_offset, y_offset, anchor='nw', text=f"Logo ID: {logo_id}", font=self.canvas_font)
        if logo_path and os.path.exists(logo_path):
            logo_img = Image.open(logo_path).resize((200, 200))
            self.tk_logo = ImageTk.PhotoImage(logo_img)
            self.canvas.create_image(x_offset+100, y_offset+40, anchor='nw', image=self.tk_logo)
        y_offset += box_height + 180

        # Class Mapping
        self.canvas.create_text(x_offset, y_offset, anchor='nw', text=f"Class Mapping: {class_mapping}", font=self.canvas_font)
        y_offset += 35

        # Parent Color Primary
        self.canvas.create_text(x_offset, y_offset, anchor='nw', text=f"Parent Color Primary: {color_id}", font=self.canvas_font)
        if color_path and os.path.exists(color_path):
            color_img = Image.open(color_path).resize((200, 200))
            self.tk_color = ImageTk.PhotoImage(color_img)
            self.canvas.create_image(x_offset+100, y_offset+40, anchor='nw', image=self.tk_color)
        y_offset += box_height + 180
        # Team League Data
        self.canvas.create_text(x_offset, y_offset, anchor='nw', text=f"Team League Data: {team_league}", font=self.canvas_font)
        y_offset += 35

        # Silhouette Name
        self.canvas.create_text(x_offset, y_offset, anchor='nw', text=f"Silhouette: {silhouette}", font=self.canvas_font)
        y_offset += 35

        # Web Style Name
        self.canvas.create_text(x_offset, y_offset, anchor='nw', text=f"Web Style: {web_style}", font=self.canvas_font)
        y_offset += 35

        # Web Display Name (wrap to two lines if > 25 chars)
        def _wrap_two_lines(text, max_chars=25):
            text = str(text) if text is not None else ""
            if len(text) <= max_chars:
                return text
            cut = text.rfind(" ", 0, max_chars + 1)
            if cut == -1 or cut < max_chars // 2:
                cut = max_chars
            return text[:cut].rstrip() + "\n" + text[cut:].lstrip()

        display_name_wrapped = _wrap_two_lines(display_name, 45)
        self.canvas.create_text(
            x_offset,
            y_offset,
            anchor='nw',
            text=f"Web Display Name: {display_name_wrapped}",
            font=self.canvas_font
        )
        y_offset += 60 if "\n" in display_name_wrapped else 30

        # Marketing Event (NEW, placed after Web Display Name)
        self.canvas.create_text(x_offset, y_offset, anchor='nw', text=f"Marketing Event: {marketing_event}", font=self.canvas_font)
        y_offset += 35

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
        self.style_entry.place(x=x_offset + 200, y=y_offset)

        # progress counter to the right of the Style Number entry
        audited = self._audited_count()
        # NEW: effective total excludes rows with failed downloads or user-marked wrong images
        try:
            data_names_iter = self.data['Name'] if 'Name' in self.data.columns else []
            effective_total = sum(1 for n in data_names_iter if str(n) not in self.wrong_image_names)
        except Exception:
            effective_total = len(self.data) if self.data is not None else 0
        progress_text = f"Audited: {audited} / {effective_total}"
        if not self.progress_label or not self.progress_label.winfo_exists():
            self.progress_label = ttk.Label(self.frame, text=progress_text, font=self.canvas_font)
        else:
            self.progress_label.config(text=progress_text)
        # Position to the right of the entry
        self.frame.update_idletasks()
        entry_width_px = self.style_entry.winfo_width()
        self.progress_label.place(x=(x_offset + 200 + entry_width_px + 20), y=y_offset)

    def fix_missing_loop(self):
        if self.missing_index >= len(self.data_missing):
            self.in_missing_loop = False  # exit missing-loop mode
            self.finish()
            return
        row = self.data_missing.iloc[self.missing_index]

        # Use unified missing detection for preselection
        missing_fields = self._get_missing_fields(row)

        self.display_row(row)
        # self.btn_back.place_forget()  # REMOVE: keep Back visible during missing-loop

        self._popup_open = True
        wrong_info = self.ask_wrong_fields(row, preselected_fields=missing_fields)
        self._popup_open = False

        # allow "Back" from popup to go to previous missing product
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

        # If user marked as Wrong Image, record and move on (do not include in to_audit)
        if "Wrong Image" in wrong_fields:
            try:
                name_val = str(row['Name']) if 'Name' in row else ""
            except Exception:
                name_val = ""
            if name_val:
                self.wrong_image_names.add(name_val)
            self.choices.append(('wrong_image', row, False))
            self.missing_index += 1
            self.fix_missing_loop()
            return

        if not wrong_fields or (isinstance(wrong_fields, list) and all(f.strip() == "" for f in wrong_fields)):
            messagebox.showwarning("Input required", "You must select at least one field that is wrong before continuing.", parent=self.root)
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

        # If user marked as Wrong Image, record and move on (do not include in to_audit)
        if "Wrong Image" in wrong_fields:
            try:
                name_val = str(row['Name']) if 'Name' in row else ""
            except Exception:
                name_val = ""
            if name_val:
                self.wrong_image_names.add(name_val)
            self.choices.append(('wrong_image', row, False))
            self.index += 1
            self.show_image()
            return

        if not wrong_fields or (isinstance(wrong_fields, list) and all(f.strip() == "" for f in wrong_fields)):
            messagebox.showwarning("Input required", "You must select at least one field that is wrong before continuing.", parent=self.root)
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
            "Silhouette",
            "Web Style",
            "Marketing Event",   # <-- Add here
            "Wrong Image"
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
            popup.transient(self.root)
            popup.lift()
            # Ensure this popup has focus and captures all events
            popup.grab_set()
            popup.focus_force()
            # Optionally keep on top so focus isn’t stolen on Windows
            try:
                popup.attributes("-topmost", True)
                popup.after(50, lambda: popup.attributes("-topmost", False))
            except Exception:
                pass

            # Place this popup at the top-right of the app window
            self._place_popup(popup, width=500, height=520, align="top-right")

            main_frame = ttk.Frame(popup, padding=60)
            main_frame.pack(fill=tk.BOTH, expand=True)

            # Style Number copy bar (double-click to copy)
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

            # Checkboxes container
            checks_frame = ttk.Frame(main_frame)
            checks_frame.pack(fill=tk.X, anchor='w')

            silhouette_var = tk.StringVar(value=(row['Silhouette'] if 'Silhouette' in row and pd.notna(row['Silhouette']) else ""))
            web_style_var = tk.StringVar(value=(row['Web Style'] if 'Web Style' in row and pd.notna(row['Web Style']) else ""))
            marketing_event_var = tk.StringVar(value=(row['Marketing Event'] if 'Marketing Event' in row and pd.notna(row['Marketing Event']) else ""))

            silhouette_frame = ttk.Frame(main_frame)
            ttk.Label(silhouette_frame, text="Silhouette should be:").pack(side=tk.LEFT, padx=(0, 8))
            ttk.Entry(silhouette_frame, textvariable=silhouette_var, width=36).pack(side=tk.LEFT, fill=tk.X, expand=True)

            web_style_frame = ttk.Frame(main_frame)
            ttk.Label(web_style_frame, text="Web Style should be:").pack(side=tk.LEFT, padx=(0, 8))
            ttk.Entry(web_style_frame, textvariable=web_style_var, width=36).pack(side=tk.LEFT, fill=tk.X, expand=True)

            marketing_event_frame = ttk.Frame(main_frame)
            ttk.Label(marketing_event_frame, text="Marketing Event should be:").pack(side=tk.LEFT, padx=(0, 8))
            ttk.Entry(marketing_event_frame, textvariable=marketing_event_var, width=36).pack(side=tk.LEFT, fill=tk.X, expand=True)

            # Toggle helpers
            def toggle_silhouette():
                if vars["Silhouette"].get():
                    silhouette_frame.pack(fill=tk.X, pady=(6, 0), anchor='w')
                else:
                    silhouette_frame.pack_forget()

            def toggle_web_style():
                if vars["Web Style"].get():
                    web_style_frame.pack(fill=tk.X, pady=(6, 0), anchor='w')
                else:
                    web_style_frame.pack_forget()

            def toggle_marketing_event():
                if vars["Marketing Event"].get():
                    marketing_event_frame.pack(fill=tk.X, pady=(6, 0), anchor='w')
                else:
                    marketing_event_frame.pack_forget()

            # Build checkboxes (attach toggle command for the new field)
            for field in fields:
                if field == "Silhouette":
                    ttk.Checkbutton(checks_frame, text=field, variable=vars[field], command=toggle_silhouette).pack(anchor='w', pady=2)
                elif field == "Web Style":
                    ttk.Checkbutton(checks_frame, text=field, variable=vars[field], command=toggle_web_style).pack(anchor='w', pady=2)
                elif field == "Marketing Event":
                    ttk.Checkbutton(checks_frame, text=field, variable=vars[field], command=toggle_marketing_event).pack(anchor='w', pady=2)
                else:
                    ttk.Checkbutton(checks_frame, text=field, variable=vars[field]).pack(anchor='w', pady=2)

            # If preselected includes these, show their inputs now
            toggle_silhouette()
            toggle_web_style()
            toggle_marketing_event()

            def submit():
                selected = [field for field in fields if vars[field].get()]
                # Auto-include dependencies if Team League Data is selected
                if "Team League Data" in selected:
                    if "Parent Color Primary" not in selected:
                        selected.append("Parent Color Primary")
                    if "Logo ID" not in selected:
                        selected.append("Logo ID")

                # Require values for free-form fields if selected
                if "Silhouette" in selected and not silhouette_var.get().strip():
                    messagebox.showwarning("Input required", "Please enter a value for Silhouette.", parent=popup)
                    return
                if "Web Style" in selected and not web_style_var.get().strip():
                    messagebox.showwarning("Input required", "Please enter a value for Web Style.", parent=popup)
                    return
                if "Marketing Event" in selected and not marketing_event_var.get().strip():
                    messagebox.showwarning("Input required", "Please enter a value for Marketing Event.", parent=popup)
                    return

                result["value"] = selected
                details = {}
                if "Silhouette" in selected:
                    details["Silhouette"] = silhouette_var.get().strip()
                if "Web Style" in selected:
                    details["Web Style"] = web_style_var.get().strip()
                if "Marketing Event" in selected:
                    details["Marketing Event"] = marketing_event_var.get().strip()
                result["details"] = details
                popup.destroy()

            def go_back():
                result["back"] = True
                popup.destroy()

            # If user clicks X, quit the whole app
            popup.protocol("WM_DELETE_WINDOW", self.quit_app)

            # Buttons: show Back only during the missing loop
            btn_bar = ttk.Frame(main_frame)
            btn_bar.pack(pady=10, anchor='e', fill=tk.X)
            if getattr(self, "in_missing_loop", False):
                ttk.Button(btn_bar, text="Back", command=go_back).pack(side=tk.LEFT)
            ok_btn = ttk.Button(btn_bar, text="OK", command=submit)
            ok_btn.pack(side=tk.RIGHT)

            # Bind Enter to submit and focus OK by default so user can press Enter
            popup.bind('<Return>', lambda event: submit())
            popup.after(0, ok_btn.focus_set)
            popup.wait_window()

        show_popup()
        if getattr(self, "_app_quitting", False):
            return {"fields": [], "details": {}}

        if result.get("back"):
            return {"back": True}

        wrong_fields = result["value"]
        wrong_details = dict(result.get("details", {}))  # may already contain Silhouette/Web Style

        def select_from_list(title, label, options, show_images=False, image_folder=None):
            sel_popup = tk.Toplevel(self.root)
            style = ttk.Style(sel_popup)
            style.theme_use("arc")
            sel_popup.configure(bg="#f7f7f7")
            sel_popup.title(title)
            sel_popup.transient(self.root)
            sel_popup.lift()

            # Place selector popups centered on the same screen
            self._place_popup(sel_popup, width=720, height=520, align="center")

            main_frame = ttk.Frame(sel_popup, padding=20)
            main_frame.pack(fill=tk.BOTH, expand=True)

            # Style Number copy bar
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

            # NEW: wrap horizontal content in its own frame so the bottom buttons don't get squeezed
            content_frame = ttk.Frame(main_frame)
            content_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

            filtered_options = options.copy()
            listbox_var = tk.StringVar(value=filtered_options)

            # List with scrollbar (keeps a stable width/height)
            list_frame = ttk.Frame(content_frame)
            list_frame.pack(side=tk.LEFT, pady=10, fill=tk.Y)
            listbox = tk.Listbox(
                list_frame,
                listvariable=listbox_var,
                width=80,
                height=min(15, len(filtered_options)),
                exportselection=False,
                bg="#f7f7f7",
                relief=tk.FLAT,
                highlightthickness=0,
                borderwidth=0
            )
            listbox.pack(side=tk.LEFT, fill=tk.Y)
            scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=listbox.yview)
            scrollbar.pack(side=tk.LEFT, fill=tk.Y)
            listbox.config(yscrollcommand=scrollbar.set)

            # Image preview in a fixed-size container to avoid resizing the dialog/buttons
            image_label = None
            img_cache = {}
            if show_images and image_folder:
                image_container = ttk.Frame(content_frame, width=180, height=180)
                image_container.pack_propagate(False)
                image_container.pack(side=tk.LEFT, padx=(10, 0), pady=10)
                image_label = ttk.Label(image_container, anchor="center")
                image_label.pack(expand=True, fill=tk.BOTH)

            def show_logo_img(event):
                if not (image_label and show_images and image_folder):
                    return
                sel = listbox.curselection()
                if sel:
                    logo_id = filtered_options[sel[0]]
                    img_path = find_image(image_folder, logo_id)
                    if img_path and os.path.exists(img_path):
                        img = Image.open(img_path).resize((150, 150))
                        tk_img = ImageTk.PhotoImage(img)
                        img_cache["img"] = tk_img
                        image_label.config(image=tk_img, text="")
                    else:
                        image_label.config(image="", text="No image found")
                else:
                    image_label.config(image="", text="")

            if show_images and image_folder:
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

            # Bottom button row stays full size
            btn_frame = ttk.Frame(main_frame)
            btn_frame.pack(side=tk.BOTTOM, fill=tk.X)
            ttk.Button(btn_frame, text="OK", command=on_select).pack(anchor='e')

            sel_popup.protocol("WM_DELETE_WINDOW", self.quit_app)
            sel_popup.bind('<Return>', on_select)
            sel_popup.wait_window()
            if getattr(self, "_app_quitting", False):
                return None
            return local["value"]

        # CSV loaders
        def load_csv_column(filename, colname, filter_col=None, filter_val=None):
            path = resource_path(filename)
            values = []
            if not os.path.exists(path):
                return values

            # Allow column aliases (Logo ID -> Name in LogoList.csv, case-insensitive)
            alias = {
                "logo id": "name",
            }
            try:
                with open(path, newline='', encoding='utf-8') as csvfile:
                    reader = csv.DictReader(csvfile)
                    # Build a case-insensitive header map
                    header_map = {h.lower().strip(): h for h in (reader.fieldnames or [])}

                    # Resolve target and filter columns against aliases and actual headers
                    want_col_key = alias.get(colname.lower().strip(), colname.lower().strip())
                    target_col = header_map.get(want_col_key)
                    if target_col is None:
                        # Nothing we can do; return empty to avoid KeyError
                        return values

                    filter_col_key = None
                    filter_col_name = None
                    if filter_col:
                        filter_col_key = alias.get(filter_col.lower().strip(), filter_col.lower().strip())
                        filter_col_name = header_map.get(filter_col_key)

                    fval = str(filter_val).strip() if filter_val is not None and str(filter_val).strip() != "" else None

                    for r in reader:
                        if filter_col_name and fval:
                            rv = r.get(filter_col_name, "")
                            if str(rv).strip() != fval:
                                continue
                        v = r.get(target_col, "")
                        if v is not None and str(v).strip() != "":
                            values.append(str(v).strip())

                # If filter yielded no options but we had a filter, fall back to all values in target_col
                if not values and filter_col_name:
                    with open(path, newline='', encoding='utf-8') as csvfile:
                        reader = csv.DictReader(csvfile)
                        # Rebuild header map in case DictReader changed
                        header_map = {h.lower().strip(): h for h in (reader.fieldnames or [])}
                        target_col = header_map.get(want_col_key)
                        for r in reader:
                            v = r.get(target_col, "")
                            if v is not None and str(v).strip() != "":
                                values.append(str(v).strip())

                return sorted(set(values))
            except Exception:
                # Be defensive: never raise from here
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
            new_color = select_from_list(
                "Select Parent Color Primary",
                f"Select the correct Parent Color Primary for team '{row['Team League Data']}':",
                color_options,
                show_images=True,
                image_folder=COLORS_FOLDER
            )
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
                # If we undid a wrong_image, remove it from the set
                if status == 'wrong_image':
                    try:
                        name_val = str(row['Name']) if 'Name' in row else ""
                        if name_val in self.wrong_image_names:
                            self.wrong_image_names.remove(name_val)
                    except Exception:
                        pass
                self.index = row.name
                self.missing_rows = [(idx, r) for idx, r in self.missing_rows if idx != self.index]
                self.show_image()
                return
        # If nothing to undo, do nothing

    def finish(self):
        self.save_outputs()
        # Add date to the filenames in the message
        date_suffix = datetime.datetime.now().strftime("%Y-%m-%d")
        to_audit_filename = f"to_audit_{date_suffix}.csv"
        wrong_images_filename = f"wrong_images_{date_suffix}.csv"
        messagebox.showinfo("Done", f"Audit complete!\nFiles saved as {to_audit_filename} and {wrong_images_filename}.", parent=self.root)
        self.completed = True
        self._cleanup_session_files()
        # Clean up this session's TEMP folder
        if os.path.exists(self.temp_folder):
            try:
                shutil.rmtree(self.temp_folder)
            except Exception as e:
                print(f"Failed to delete TEMP folder: {e}")
        self.root.quit()

    def save_outputs(self):
        # Build dated filenames
        date_suffix = datetime.datetime.now().strftime("%Y-%m-%d")
        to_audit_filename = f"to_audit_{date_suffix}.csv"
        wrong_images_filename = f"wrong_images_{date_suffix}.csv"

        # Helper to collect wrong-image parent+children rows
        def _collect_wrong_rows():
            rows = []
            wrong_set_local = set(self.wrong_image_names)
            if not wrong_set_local:
                return rows

            # Candidate keys for fallback internal id read (if needed)
            id_keys = ('Internal ID', 'Internal ID.1', 'Internal ID 0', 'InternalID0', 'InternalID', 'Internal ID0')

            def _id_from_series(series):
                if series is None:
                    return ""
                for k in id_keys:
                    if k in series and pd.notna(series[k]) and str(series[k]).strip():
                        return str(series[k]).strip()
                return ""

            # Map parent Name -> parent row for fallback
            name_to_parent = {}
            if self.data is not None and not self.data.empty and 'Name' in self.data.columns:
                for _, prow in self.data.iterrows():
                    name_to_parent[str(prow.get('Name', '') or '')] = prow

            for parent_name in sorted(wrong_set_local):
                # Parent row id via map, then fallback
                prow = name_to_parent.get(parent_name)
                pid = self.name_to_internal_id.get(parent_name, "") or _id_from_series(prow)
                rows.append({"Internal ID": pid, "Name": parent_name})

                # Children
                if hasattr(self, 'child_records') and parent_name in getattr(self, 'child_records', {}):
                    for crow in self.child_records[parent_name]:
                        child_name = str(crow.get('Name', '') or '').strip()
                        cid = self.name_to_internal_id.get(child_name, "") or _id_from_series(crow)
                        rows.append({"Internal ID": cid, "Name": child_name})
            return rows

        if self.data is None or self.data.empty:
            # Still write a wrong_images file (with new policy columns)
            try:
                wrong_rows = _collect_wrong_rows()
                wrong_df = pd.DataFrame(wrong_rows if wrong_rows else [], columns=["Internal ID", "Name"])
                wrong_df["Did you make a POL"] = ""
                wrong_df["Do Not Display in Web Store"] = "Yes"
                wrong_df["Do Not Display Reason"] = "Bad Image"
                wrong_df["Display in Web Store"] = "No"

                wrong_df = wrong_df.drop_duplicates(subset=["Internal ID", "Name"], keep="first")
                wrong_df = wrong_df.reindex(columns=[
                    "Internal ID", "Name",
                    "Did you make a POL",
                    "Do Not Display in Web Store",
                    "Do Not Display Reason",
                    "Display in Web Store"
                ])

                wrong_df.to_csv(wrong_images_filename, index=False)
            except Exception as e:
                print(f"Failed to write {wrong_images_filename}: {e}")
            return

        exclude_cols = {"Picture ID", "Image Assignment"}
        output_rows = []
        wrong_set = set(self.wrong_image_names)

        for _, parent_row in self.data.iterrows():
            parent_name = str(parent_row['Name']) if 'Name' in parent_row else ""
            # Skip parents marked as wrong image
            if parent_name in wrong_set:
                continue

            output_rows.append(parent_row.copy())

            # Include children only for included parents
            if hasattr(self, 'child_records') and parent_name in self.child_records:
                for child_row in self.child_records[parent_name]:
                    new_child = parent_row.copy()
                    new_child['Name'] = child_row['Name']
                    internal_id = None
                    for key in ('Internal ID', 'Internal ID.1', 'Internal ID 0', 'InternalID0', 'InternalID', 'Internal ID0'):
                        if key in child_row and pd.notna(child_row[key]):
                            internal_id = child_row[key]
                            break
                    new_child['Internal ID'] = internal_id if internal_id is not None else ""
                    output_rows.append(new_child)

        output_df = pd.DataFrame(output_rows)
        output_df = output_df[[col for col in output_df.columns if col not in exclude_cols]].copy()

        # NEW: Ensure Internal ID is present and populated using the Name -> ID map
        if 'Name' in output_df.columns:
            mapped_ids = output_df['Name'].map(lambda x: self.name_to_internal_id.get(str(x), ""))
            if 'Internal ID' in output_df.columns:
                mask = output_df['Internal ID'].isna() | (output_df['Internal ID'].astype(str).str.strip() == "")
                output_df.loc[mask, 'Internal ID'] = mapped_ids[mask]
            else:
                output_df['Internal ID'] = mapped_ids

        today_str = datetime.datetime.now().strftime("%m/%d/%Y")
        output_df["Flash Sale Date"] = today_str
        output_df.to_csv(to_audit_filename, index=False)

        # Write wrong_images (parents + children) with required policy columns
        try:
            wrong_rows = _collect_wrong_rows()
            wrong_df = pd.DataFrame(wrong_rows, columns=["Internal ID", "Name"])
            wrong_df["Did you make a POL"] = ""
            wrong_df["Do Not Display in Web Store"] = "Yes"
            wrong_df["Do Not Display Reason"] = "Bad Image"
            wrong_df["Display in Web Store"] = "No"

            wrong_df = wrong_df.drop_duplicates(subset=["Internal ID", "Name"], keep="first")
            wrong_df = wrong_df.reindex(columns=[
                "Internal ID", "Name",
                "Did you make a POL",
                "Do Not Display in Web Store",
                "Do Not Display Reason",
                "Display in Web Store"
            ])

            wrong_df.to_csv(wrong_images_filename, index=False)
        except Exception as e:
            print(f"Failed to write {wrong_images_filename}: {e}")

        # Delete the entire TEMP folder
        if os.path.exists(TEMP_FOLDER):
            try:
                shutil.rmtree(TEMP_FOLDER)
            except Exception as e:
                print(f"Failed to delete TEMP folder: {e}")

    def on_close(self):
        # Save session state (resume later) and close
        try:
            if not self.completed:
                self.save_session()
        finally:
            self.root.destroy()

    def _audited_count(self):
        try:
            return len({getattr(c[1], "name", None) for c in self.choices if len(c) >= 3 and not c[2]})
        except Exception:
            return len([c for c in self.choices if len(c) >= 3 and not c[2]])

    def quit_app(self):
        # Gracefully save and close everything
        if getattr(self, "_app_quitting", False):
            return
        self._app_quitting = True
        try:
            self.on_close()
        except Exception:
            try:
                self.root.destroy()
            except Exception:
                pass

    # explicit Save & Quit handler
    def save_and_quit(self):
        if messagebox.askyesno("Save and Quit", "Save your current progress and quit now?", parent=self.root):
            self.quit_app()

    # ---- session helpers ----
    def _build_child_records(self, csv_path):
        try:
            df_full = pd.read_csv(csv_path, dtype=str)
        except Exception:
            return {}
        child_records = {}
        for _, row in df_full.iterrows():
            name = str(row['Name']) if 'Name' in row else ""
            if " :" in name:
                parent_name = name.split(" :")[0]
                child_records.setdefault(parent_name, []).append(row.copy())
        return child_records

    # NEW: Build a reliable Name -> Internal ID map from any CSV (includes parents and children)
    def _build_name_to_id(self, csv_path):
        mapping = {}
        try:
            df_full = pd.read_csv(csv_path, dtype=str)
            id_keys = ['Internal ID', 'Internal ID.1', 'Internal ID 0', 'InternalID0', 'InternalID', 'Internal ID0']
            for _, r in df_full.iterrows():
                name = str(r.get('Name', '') or '').strip()
                if not name:
                    continue
                iid = ""
                for k in id_keys:
                    if k in r and pd.notna(r[k]) and str(r[k]).strip():
                        iid = str(r[k]).strip()
                        break
                if name and iid:
                    mapping[name] = iid
        except Exception:
            pass
        return mapping

    def save_session(self):
        if self.data is None:
            return

        # persist current parent rows (with any corrections)
        try:
            self.data.to_csv(self.session_data_csv_path, index=False)
        except Exception as e:
            print(f"Failed to write session data CSV: {e}")

        # helper: safe int conversion (handles numpy int64)
        def _to_int(v, default=None):
            try:
                return int(v)
            except Exception:
                return default

        # serialize minimal state (ensure pure Python types)
        missing_indices = []
        for idx, _ in (self.missing_rows or []):
            i = _to_int(idx)
            if i is not None:
                missing_indices.append(i)

        serial_choices = []
        for entry in self.choices:
            try:
                row_index = _to_int(getattr(entry[1], "name", -1), default=-1)
                rec = {
                    "status": entry[0],
                    "row_index": row_index,
                    "auto": bool(entry[2]),
                }
                if len(entry) > 3:
                    # ensure wrong_fields is list of strings
                    rec["wrong_fields"] = [str(x) for x in entry[3]] if isinstance(entry[3], (list, tuple)) else [str(entry[3])]
                if len(entry) > 4 and isinstance(entry[4], dict):
                    # ensure wrong_details is a dict of strings
                    rec["wrong_details"] = {str(k): str(v) for k, v in entry[4].items()}
                serial_choices.append(rec)
            except Exception:
                pass

        manifest = {
            "created": datetime.datetime.now().isoformat(),
            "original_csv_path": str(self.original_csv_path) if self.original_csv_path else "",
            "data_csv_path": str(self.session_data_csv_path) if self.session_data_csv_path else "",
            "index": _to_int(self.index, 0),
            "phase": "missing" if self.in_missing_loop else "main",
            "missing_index": _to_int(self.missing_index, 0),
            "missing_rows_indices": missing_indices,
            "choices": serial_choices,
            "temp_folder": str(self.temp_folder) if self.temp_folder else TEMP_FOLDER,
            "wrong_images": sorted(list(self.wrong_image_names)),
        }
        try:
            with open(self.session_manifest_path, "w", encoding="utf-8") as f:
                json.dump(manifest, f, indent=2)
        except Exception as e:
            print(f"Failed to write session manifest: {e}")

    def _cleanup_session_files(self):
        for p in [self.session_manifest_path, self.session_data_csv_path]:
            try:
                if p and os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass

    def handle_app_exit(self):
        if not self.completed:
            self.save_session()

    # Background helpers
    def _load_bg_image(self):
        if self.bg_original is None:
            try:
                bg_path = resource_path("background.png")
                if os.path.exists(bg_path):
                    self.bg_original = Image.open(bg_path)
            except Exception:
                self.bg_original = None

    def _update_bg_image(self, width=None, height=None):
        if not self.bg_original or not hasattr(self, "canvas"):
            return
        # Determine target size (canvas size)
        w = int(width) if width else int(self.canvas.winfo_width())
        h = int(height) if height else int(self.canvas.winfo_height())
        if w <= 1 or h <= 1:
            # Canvas not yet laid out; try again shortly
            self.root.after(50, self._update_bg_image)
            return
        try:
            # Stretch to fill canvas
            resized = self.bg_original.resize((w, h), Image.LANCZOS)
            self.tk_bg_img = ImageTk.PhotoImage(resized)
            if self.bg_image_id:
                self.canvas.itemconfig(self.bg_image_id, image=self.tk_bg_img)
            else:
                self.bg_image_id = self.canvas.create_image(0, 0, anchor='nw', image=self.tk_bg_img)
            self.canvas.tag_lower(self.bg_image_id)
        except Exception:
            pass

    def _on_canvas_resize(self, event):
        # Keep background stretched to new size
        self._update_bg_image(event.width, event.height)

if __name__ == "__main__":
    try:
        root = ThemedTk(theme="arc")
    except Exception:
        # Fallback if ttkthemes isn't available
        root = tk.Tk()
    app = AuditApp(root)
    try:
        root.mainloop()
    finally:
        # Save session on unexpected exit
        app.handle_app_exit()