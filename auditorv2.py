import os
import sys
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import pandas as pd
import download_helper
import csv

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
        self.setup_ui()
        self.root.bind('<Left>', self.mark_wrong)
        self.root.bind('<Right>', self.mark_right)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)  # Handle window close

    def setup_ui(self):
        self.frame = tk.Frame(self.root)
        self.frame.pack(fill=tk.BOTH, expand=True)
        # Replace button with image button
        choose_img_path = resource_path("choose.png")
        if os.path.exists(choose_img_path):
            choose_img = Image.open(choose_img_path)
            choose_img = choose_img.resize((200, 200))  # Resize as needed
            self.tk_choose_img = ImageTk.PhotoImage(choose_img)
            self.btn_load = tk.Button(self.frame, image=self.tk_choose_img, command=self.load_csv, borderwidth=0)
        else:
            self.btn_load = tk.Button(self.frame, text="Load CSV", command=self.load_csv)
        self.btn_load.pack(pady=10)
        # Make canvas large enough for fullscreen and info
        self.canvas = tk.Canvas(self.frame, width=1920, height=1080, bg="white")
        self.canvas.pack(fill=tk.BOTH, expand=True)
        # Set default font for canvas text
        self.canvas_font = ("Roboto", 24)

        # Add back button (hidden until images are shown)
        back_img_path = resource_path("back.png")
        if os.path.exists(back_img_path):
            back_img = Image.open(back_img_path)
            back_img = back_img.resize((100, 100))
            self.tk_back_img = ImageTk.PhotoImage(back_img)
            self.btn_back = tk.Button(self.frame, image=self.tk_back_img, command=self.undo_last, borderwidth=0)
        else:
            self.btn_back = tk.Button(self.frame, text="Back", command=self.undo_last)
        self.btn_back.place_forget()  # Hide initially

    def load_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not file_path:
            return
        self.data = pd.read_csv(file_path, dtype=str)
        download_helper.download_images(file_path, TEMP_FOLDER, item_col='Name', picture_id_col='Picture ID')
        self.data.reset_index(drop=True, inplace=True)
        self.index = 0
        self.choices = []
        self.btn_load.pack_forget()  # Hide the choose file button after loading

        # Set background to background.png if it exists
        bg_path = resource_path("background.png")
        if os.path.exists(bg_path):
            bg_img = Image.open(bg_path)
            bg_img = bg_img.resize((1920, 1080))  # Resize as needed for your canvas
            self.tk_bg_img = ImageTk.PhotoImage(bg_img)
            # Draw background image as the first (bottom) item on the canvas
            self.bg_image_id = self.canvas.create_image(0, 0, anchor='nw', image=self.tk_bg_img)
            self.canvas.tag_lower(self.bg_image_id)  # Ensure it's at the very back

        self.show_image()

    def show_image(self):
        if self.data is None or self.index >= len(self.data):
            self.finish()
            return
        row = self.data.iloc[self.index]
        # Check for missing/invalid fields
        logo_id = row['Logo ID'] if pd.notna(row['Logo ID']) else ""
        class_mapping = row['Class Mapping'] if pd.notna(row['Class Mapping']) else ""
        color_id = row['Parent Color Primary'] if pd.notna(row['Parent Color Primary']) else ""
        team_league = row['Team League Data'] if pd.notna(row['Team League Data']) else ""

        # Reject if any required field is blank or Logo ID contains "-TBD"
        if (
            not logo_id or
            not class_mapping or
            not color_id or
            not team_league or
            "-tbd" in logo_id.lower()
        ):
            self.choices.append(('to_audit', row, True))  # Mark as auto-rejected
            self.index += 1
            self.show_image()
            return

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

        # Place the back button under the product image
        self.btn_back.place(x=img_x + 205, y=img_y + 750)  # Centered under image

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

        # Team League Data
        self.canvas.create_text(x_offset, y_offset, anchor='nw', text=f"Team League Data: {team_league}", font=self.canvas_font)

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
        self.choices.append(('to_audit', row, False, wrong_fields, wrong_details))
        self.index += 1
        self.show_image()

    def ask_wrong_fields(self, row):
        fields = [
            "Logo ID",
            "Class Mapping",
            "Parent Color Primary",
            "Team League Data",
            "Other"
        ]
        popup = tk.Toplevel(self.root)
        popup.title("Select the field(s) that are wrong")
        popup.grab_set()
        popup.transient(self.root)
        self.root.unbind('<Left>')
        self.root.unbind('<Right>')
        self.root.attributes('-disabled', True)

        tk.Label(popup, text="Which field(s) are wrong?").pack(padx=20, pady=10)

        vars = {field: tk.BooleanVar(value=False) for field in fields}
        for field in fields:
            tk.Checkbutton(popup, text=field, variable=vars[field]).pack(anchor='w', padx=20)

        custom_var = tk.StringVar()
        entry = tk.Entry(popup, textvariable=custom_var)
        def on_other_checked(*args):
            if vars["Other"].get():
                entry.pack(padx=20, pady=5)
            else:
                entry.pack_forget()
        vars["Other"].trace_add("write", on_other_checked)

        result = {"value": None, "details": {}}
        def submit():
            selected = [field for field in fields if vars[field].get() and field != "Other"]
            if vars["Other"].get():
                other_text = custom_var.get().strip()
                if other_text:
                    selected.append(other_text)
                else:
                    selected.append("Other")
            if not selected:
                messagebox.showwarning("Input required", "Please select at least one field or enter a value for 'Other'.", parent=popup)
                return
            popup.destroy()
            result["value"] = selected

        popup.protocol("WM_DELETE_WINDOW", lambda: popup.destroy())
        tk.Button(popup, text="OK", command=submit).pack(pady=10)
        popup.wait_window()
        self.root.attributes('-disabled', False)
        self.root.bind('<Left>', self.mark_wrong)
        self.root.bind('<Right>', self.mark_right)
        self.root.focus_force()

        # Now, for each selected field, prompt for the new value as needed
        wrong_fields = result["value"]
        wrong_details = {}

        # Helper to select from a list
        def select_from_list(title, label, options, show_images=False, image_folder=None):
            sel_popup = tk.Toplevel(self.root)
            sel_popup.title(title)
            sel_popup.grab_set()
            sel_popup.focus_force()  # Focus the popup window
            self.root.attributes('-disabled', True)
            tk.Label(sel_popup, text=label).pack(padx=50, pady=10)

            # Make the listbox much wider
            listbox_width = 80  # Wider for long values

            var = tk.StringVar(value=options[0] if options else "")
            listbox = tk.Listbox(
                sel_popup,
                listvariable=tk.StringVar(value=options),
                width=listbox_width,
                height=min(15, len(options)),
                exportselection=False
            )
            listbox.pack(side=tk.LEFT, padx=(50, 10), pady=10, fill=tk.Y)
            listbox.focus_set()  # Focus the listbox for keyboard navigation

            # For showing images next to Logo ID options
            image_label = None
            img_cache = {}

            if show_images and image_folder:
                image_label = tk.Label(sel_popup)
                image_label.pack(side=tk.LEFT, padx=(10, 50), pady=10, fill=tk.BOTH, expand=True)

                def show_logo_img(event):
                    sel = listbox.curselection()
                    if sel:
                        logo_id = options[sel[0]]
                        img_path = find_image(image_folder, logo_id)
                        if img_path and os.path.exists(img_path):
                            img = Image.open(img_path)
                            img = img.resize((150, 150))  # Adjust size as needed
                            tk_img = ImageTk.PhotoImage(img)
                            img_cache["img"] = tk_img  # Prevent garbage collection
                            image_label.config(image=tk_img, text="")
                        else:
                            image_label.config(image="", text="No image found")
                    else:
                        image_label.config(image="", text="")
                listbox.bind("<<ListboxSelect>>", show_logo_img)
                # Show image for first item by default
                if options:
                    listbox.selection_set(0)
                    show_logo_img(None)

            result = {"value": None}
            def on_select(event=None):
                sel = listbox.curselection()
                if sel:
                    result["value"] = options[sel[0]]
                    sel_popup.destroy()

            # OK button at the bottom
            btn_frame = tk.Frame(sel_popup)
            btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 20))
            ok_btn = tk.Button(btn_frame, text="OK", command=on_select)
            ok_btn.pack()
            sel_popup.bind('<Return>', on_select)  # Allow Enter key to confirm

            sel_popup.wait_window()
            self.root.attributes('-disabled', False)
            self.root.focus_force()
            return result["value"]

        # Load CSVs as needed
        def load_csv_column(filename, colname, filter_col=None, filter_val=None):
            path = resource_path(filename)
            values = []
            if not os.path.exists(path):
                return values
            with open(path, newline='', encoding='utf-8') as csvfile:
                reader = csv.DictReader(csvfile)
                for row in reader:
                    if filter_col and filter_val and row.get(filter_col) != filter_val:
                        continue
                    values.append(row[colname])
            return sorted(set(values))

        # Always handle Team League Data first if selected
        if "Team League Data" in wrong_fields:
            team_options = load_csv_column("TeamList.csv", "Team League Data")
            new_team = select_from_list("Select Team", "Select the correct Team League Data:", team_options)
            if new_team:
                wrong_details["Team League Data"] = new_team
                # Update row for filtering other fields
                row = row.copy()
                row['Team League Data'] = new_team

        # Now handle other fields, filtering by team if needed
        team_val = row['Team League Data'] if pd.notna(row['Team League Data']) else ""
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

        # Return both the wrong fields and the new values chosen
        return {"fields": wrong_fields, "details": wrong_details}

    def undo_last(self):
        # Undo the last user action (not auto-rejected)
        for i in range(len(self.choices) - 1, -1, -1):
            status, row, auto_rejected = self.choices[i]
            if not auto_rejected:
                self.choices.pop(i)
                self.index = row.name  # row.name is the original index in DataFrame
                self.show_image()
                return
        # If nothing to undo, do nothing

    def finish(self):
        self.save_outputs()
        messagebox.showinfo("Done", "Audit complete!\nFile saved as to_audit.csv.")
        self.root.quit()

    def save_outputs(self):
        if not self.choices:
            return
        has_wrong_field = any(len(choice) > 3 for choice in self.choices)
        accepted = [row for status, row, *rest in self.choices if status == 'accepted']
        to_audit = []
        wrong_fields = []
        wrong_details_list = []
        for choice in self.choices:
            if choice[0] == 'to_audit':
                to_audit.append(choice[1])
                if len(choice) > 3:
                    wrong_fields.append("; ".join(choice[3]) if isinstance(choice[3], list) else str(choice[3]))
                else:
                    wrong_fields.append("")
                if len(choice) > 4:
                    # Save details as a stringified dict
                    wrong_details_list.append(str(choice[4]))
                else:
                    wrong_details_list.append("")
        exclude_cols = {"Picture ID", "Image Assignment"}
        if to_audit:
            to_audit_df = pd.DataFrame(to_audit)
            to_audit_df = to_audit_df[[col for col in to_audit_df.columns if col not in exclude_cols]]
            if has_wrong_field:
                to_audit_df["Wrong Field"] = wrong_fields
                to_audit_df["Wrong Field Details"] = wrong_details_list
            to_audit_df.to_csv("to_audit.csv", index=False)
        # Delete TEMP folder contents
        if os.path.exists(TEMP_FOLDER):
            for filename in os.listdir(TEMP_FOLDER):
                file_path = os.path.join(TEMP_FOLDER, filename)
                try:
                    if os.path.isfile(file_path) or os.path.islink(file_path):
                        os.unlink(file_path)
                    elif os.path.isdir(file_path):
                        shutil.rmtree(file_path)
                except Exception as e:
                    print(f"Failed to delete {file_path}: {e}")

    def on_close(self):
        self.save_outputs()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = AuditApp(root)
    try:
        root.mainloop()
    finally:
        app.save_outputs()