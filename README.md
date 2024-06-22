# QR Code Generator

This Python-based application generates QR codes from user-input text or from data in an Excel file. The generated QR codes can be customized with different foreground and background colors, previewed within the application, and saved as PNG or SVG files.

## Features

- Generate QR codes from user-input text.
- Customize QR code colors (foreground and background).
- Preview QR codes before saving.
- Save QR codes in PNG or SVG format.
- Generate multiple QR codes from an Excel file and save them in bulk.

## Requirements

- Python 3.x
- `tkinter` (for GUI)
- `qrcode` (for generating QR codes)
- `Pillow` (for image processing)
- `openpyxl` (for handling Excel files)

## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/yourusername/qr-code-generator.git
    cd qr-code-generator
    ```

2. Install the required Python packages:
    ```sh
    pip install qrcode[pil] pillow openpyxl
    ```

## Usage

1. Run the application:
    ```sh
    python main.py
    ```

2. Enter text in the input field to generate a QR code. The QR code will be displayed in the preview area.

3. Customize the QR code by choosing the foreground and background colors.

4. To save the QR code:
    - Specify the size of the QR code in the "Save Size" field.
    - Click the "Save QR Codes" button and choose the file format (PNG or SVG).

5. To generate multiple QR codes from an Excel file:
    - Click the "Upload Excel File" button.
    - Select an Excel file containing the data for QR codes.
    - Choose a directory to save the generated QR codes.
