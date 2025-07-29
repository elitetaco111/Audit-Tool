import os
import sys
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import pandas as pd
import download_helper

TEMP_FOLDER = "TEMP"
LOGOS_FOLDER = "Logos"
COLORS_FOLDER = "Colors"

# DEPRECATED USE auditorv2.py now
# This script is a GUI application for auditing product images and metadata.

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
        self.choices.append(('accepted', self.data.iloc[self.index], False))
        self.index += 1
        self.show_image()

    def mark_wrong(self, event=None):
        self.choices.append(('to_audit', self.data.iloc[self.index], False))
        self.index += 1
        self.show_image()

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
        accepted = [row for status, row, auto_rejected in self.choices if status == 'accepted']
        to_audit = [row for status, row, auto_rejected in self.choices if status == 'to_audit']

        # Columns to exclude
        exclude_cols = {"Picture ID", "Image Assignment"}

        if to_audit:
            to_audit_df = pd.DataFrame(to_audit)
            to_audit_df = to_audit_df[[col for col in to_audit_df.columns if col not in exclude_cols]]
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