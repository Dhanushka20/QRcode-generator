import tkinter as tk
from tkinter import messagebox, filedialog, colorchooser
import qrcode
from PIL import Image, ImageTk
import qrcode.image.svg
import openpyxl

# Function to generate QR code
def generate_qr(*args):
    text = entry.get()
    if text.strip() == "":
        qr_label.config(image='')
        return

    global qr_img, qr_svg
    fg_color = fg_color_var.get()
    bg_color = bg_color_var.get()

    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(text)
    qr.make(fit=True)

    qr_img = qr.make_image(fill_color=fg_color, back_color=bg_color)
    img = qr_img.resize((200, 200), Image.Resampling.LANCZOS)  # Preview size fixed at 200x200
    img_tk = ImageTk.PhotoImage(img)

    qr_label.config(image=img_tk)
    qr_label.image = img_tk

    qr_svg = qr.make_image(image_factory=qrcode.image.svg.SvgImage)

# Function to save the QR code
def save_qr():
    if qr_img is None or qr_svg is None:
        messagebox.showerror("Save Error", "No QR code to save. Generate a QR code first.")
        return

    file_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG files", "*.png"), ("SVG files", "*.svg"), ("All files", "*.*")])
    if file_path:
        if file_path.lower().endswith('.png'):
            try:
                size = int(size_entry.get())
                if size <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Input Error", "Size must be a positive integer.")
                return

            qr_img_resized = qr_img.resize((size, size), Image.Resampling.LANCZOS)
            qr_img_resized.save(file_path)
            messagebox.showinfo("Save Successful", f"QR code saved as {file_path}")
        elif file_path.lower().endswith('.svg'):
            with open(file_path, 'w') as f:
                qr_svg.save(f)
            messagebox.showinfo("Save Successful", f"QR code saved as {file_path}")
        else:
            messagebox.showerror("Save Error", "Unsupported file format. Please use .png or .svg.")

# Function to choose foreground color
def choose_fg_color():
    color_code = colorchooser.askcolor(title="Choose foreground color")[1]
    if color_code:
        fg_color_var.set(color_code)

# Function to choose background color
def choose_bg_color():
    color_code = colorchooser.askcolor(title="Choose background color")[1]
    if color_code:
        bg_color_var.set(color_code)

# Function to handle Excel file upload and generate QR codes
# Function to handle Excel file upload and generate QR codes
def handle_excel_upload():
    global qr_img_list, qr_svg_list

    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if file_path:
        try:
            wb = openpyxl.load_workbook(file_path)
            ws = wb.active

            save_dir = filedialog.askdirectory()
            if not save_dir:
                return

            qr_img_list = []
            qr_svg_list = []

            for row in ws.iter_rows(values_only=True):
                for data in row:
                    if data:
                        qr = qrcode.QRCode(
                            version=1,
                            error_correction=qrcode.constants.ERROR_CORRECT_L,
                            box_size=10,
                            border=4,
                        )
                        qr.add_data(str(data))
                        qr.make(fit=True)

                        qr_img = qr.make_image(fill_color=fg_color_var.get(), back_color=bg_color_var.get())
                        qr_svg = qr.make_image(image_factory=qrcode.image.svg.SvgImage)

                        qr_img_list.append(qr_img)
                        qr_svg_list.append(qr_svg)

            messagebox.showinfo("QR Code Generation", "QR codes generated successfully.")

            for i, (qr_img, qr_svg) in enumerate(zip(qr_img_list, qr_svg_list), start=1):
                qr_img_path = f"{save_dir}/qr_code_{i}.png"
                qr_svg_path = f"{save_dir}/qr_code_{i}.svg"

                qr_img.save(qr_img_path)
                with open(qr_svg_path, 'w') as f:
                    qr_svg.save(f)

            messagebox.showinfo("Save Successful", f"QR codes saved to {save_dir}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while processing the Excel file: {str(e)}")


# Create the main window
root = tk.Tk()
root.title("QR Code Generator")

# Initialize global variables
qr_img = None
qr_svg = None
qr_img_list = []
qr_svg_list = []

# Create and place the widgets
frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

entry_label = tk.Label(frame, text="Enter text:")
entry_label.grid(row=0, column=0, padx=5, pady=5, sticky='e')

entry = tk.Entry(frame, width=30)
entry.grid(row=0, column=1, padx=5, pady=5)
entry.bind("<KeyRelease>", generate_qr)

size_label = tk.Label(frame, text="Save Size:")
size_label.grid(row=1, column=0, padx=5, pady=5, sticky='e')

size_entry = tk.Entry(frame, width=10)
size_entry.insert(0, "200")  # Default size for saving
size_entry.grid(row=1, column=1, padx=5, pady=5)

fg_color_var = tk.StringVar(value="black")
bg_color_var = tk.StringVar(value="white")

fg_color_var.trace("w", generate_qr)
bg_color_var.trace("w", generate_qr)

fg_color_button = tk.Button(frame, text="Foreground Color", command=choose_fg_color)
fg_color_button.grid(row=2, column=0, columnspan=2, padx=5, pady=5)

bg_color_button = tk.Button(frame, text="Background Color", command=choose_bg_color)
bg_color_button.grid(row=3, column=0, columnspan=2, padx=5, pady=5)

qr_label = tk.Label(frame)
qr_label.grid(row=4, column=0, columnspan=2, padx=5, pady=10)

# Create a frame for the buttons
button_frame = tk.Frame(root)
button_frame.pack(pady=10)

upload_button = tk.Button(button_frame, text="Upload Excel File", command=handle_excel_upload)
upload_button.grid(row=0, column=0, padx=5)

generate_button = tk.Button(button_frame, text="Generate QR Codes", command=generate_qr)
generate_button.grid(row=0, column=1, padx=5)

save_button = tk.Button(button_frame, text="Save QR Codes", command=save_qr)
save_button.grid(row=0, column=2, padx=5)

# Run the application
root.mainloop()
