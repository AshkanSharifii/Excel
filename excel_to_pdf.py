import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
import threading
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.enums import TA_RIGHT, TA_CENTER
from reportlab.pdfbase.pdfmetrics import stringWidth
from datetime import datetime
import sys

# Try to import Arabic reshaper and BiDi support
try:
    import arabic_reshaper
    from bidi.algorithm import get_display

    BIDI_SUPPORT = True
    print("✓ BiDi support loaded")
except ImportError:
    BIDI_SUPPORT = False
    print("Warning: arabic_reshaper and python-bidi not installed.")
    print("Install them for better Persian/Arabic text support:")
    print("  pip install arabic-reshaper python-bidi")


# Register Persian/Arabic font support
def setup_persian_font():
    """Try to find and register Persian font from multiple locations"""
    font_names = ['Vazir.ttf', 'Vazir-Regular.ttf', 'Vazir-Medium.ttf', 'Vazir-Bold.ttf', 'Sahel.ttf', 'B-Nazanin.ttf',
                  'IRANSans.ttf']

    # Possible font locations
    font_paths = [
        # Same directory as script
        '',
        # Fonts subdirectory
        'fonts/',
        # Common system paths
        '/usr/share/fonts/truetype/',
        'C:/Windows/Fonts/',
        os.path.expanduser('~/.fonts/'),
        '/System/Library/Fonts/',
    ]

    for font_name in font_names:
        for path in font_paths:
            try:
                font_file = os.path.join(path, font_name)
                if os.path.exists(font_file):
                    pdfmetrics.registerFont(TTFont('Persian', font_file))
                    print(f"✓ Persian font loaded: {font_file}")
                    return 'Persian'
            except Exception as e:
                continue

    print("Warning: Persian font not found. Using default font.")
    return 'Helvetica'


PERSIAN_FONT = setup_persian_font()


def fix_persian_text(text):
    """Fix Persian text for proper display in PDF"""
    if text is None:
        return ''

    text = str(text)

    if not BIDI_SUPPORT:
        return text

    try:
        # Configure reshaper for Persian
        configuration = {
            'delete_harakat': False,
            'support_ligatures': True,
            'RIAL SIGN': True,  # Support for ریال sign
        }

        reshaper = arabic_reshaper.ArabicReshaper(configuration=configuration)
        reshaped_text = reshaper.reshape(text)
        bidi_text = get_display(reshaped_text)

        return bidi_text
    except Exception as e:
        print(f"Error reshaping text: {e}")
        # Fallback to simple reshaping
        try:
            reshaped_text = arabic_reshaper.reshape(text)
            bidi_text = get_display(reshaped_text)
            return bidi_text
        except:
            return text


def read_excel_file(file_path):
    """Read Excel file and return DataFrame"""
    try:
        if file_path.endswith(('.xlsx', '.xlsm', '.xlsb')):
            df = pd.read_excel(file_path, engine='openpyxl')
        elif file_path.endswith('.xls'):
            df = pd.read_excel(file_path, engine='xlrd')
        else:
            df = pd.read_excel(file_path)
        return df
    except Exception as e:
        raise Exception(f"Error reading Excel file: {e}")


def create_pdf_for_person(person_data, person_name, output_dir, columns):
    """Create a PDF file for a single person with their data"""

    # Create filename (remove invalid characters)
    safe_filename = "".join(c for c in person_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
    pdf_filename = os.path.join(output_dir, f"{safe_filename}.pdf")

    # Create PDF document
    doc = SimpleDocTemplate(
        pdf_filename,
        pagesize=landscape(A4),
        rightMargin=30,
        leftMargin=30,
        topMargin=30,
        bottomMargin=30
    )

    elements = []
    styles = getSampleStyleSheet()

    # Create custom style for Persian text
    persian_title_style = ParagraphStyle(
        'PersianTitle',
        parent=styles['Title'],
        fontName=PERSIAN_FONT,
        fontSize=18,
        alignment=TA_CENTER,
        leading=24,
        spaceAfter=30,
        wordWrap='RTL'
    )

    persian_normal_style = ParagraphStyle(
        'PersianNormal',
        parent=styles['Normal'],
        fontName=PERSIAN_FONT,
        fontSize=10,
        alignment=TA_RIGHT,
        leading=14,
        wordWrap='RTL'
    )

    # Add title with fixed Persian text
    title_persian = fix_persian_text('گزارش اطلاعات')
    name_persian = fix_persian_text(person_name)
    title_text = f"{title_persian}: {name_persian}"

    title = Paragraph(title_text, persian_title_style)
    elements.append(title)
    elements.append(Spacer(1, 0.5 * inch))

    # Add timestamp with fixed Persian text
    date_label = fix_persian_text('تاریخ تولید گزارش')
    timestamp_text = f"{date_label}: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    timestamp = Paragraph(timestamp_text, persian_normal_style)
    elements.append(timestamp)
    elements.append(Spacer(1, 0.3 * inch))

    # Prepare data for table
    table_data = []

    # Process columns and values
    if len(columns) > 10:
        # For many columns, create a vertical layout
        for col in columns:
            value = person_data[col].iloc[0] if not person_data[col].empty else ''
            if pd.isna(value):
                value = '-'

            # Fix Persian text for both column name and value
            col_fixed = fix_persian_text(col)
            value_fixed = fix_persian_text(str(value))

            table_data.append([value_fixed, col_fixed])  # Reversed order for RTL
    else:
        # For fewer columns, use traditional horizontal layout
        headers = []
        values = []

        for col in columns:
            # Fix column header
            headers.append(fix_persian_text(col))

            # Get and fix value
            value = person_data[col].iloc[0] if not person_data[col].empty else ''
            if pd.isna(value):
                value = '-'
            values.append(fix_persian_text(str(value)))

        # Reverse headers for RTL display
        headers.reverse()
        values.reverse()

        table_data = [headers, values]

    # Create table
    table = Table(table_data, repeatRows=1)

    # Table style
    table_style = TableStyle([
        # Header style
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, -1), PERSIAN_FONT),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        # Alternate row colors
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
        # RTL alignment
        ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
    ])

    table.setStyle(table_style)
    elements.append(table)

    # Build PDF
    try:
        doc.build(elements)
        return pdf_filename
    except Exception as e:
        print(f"Error building PDF: {e}")
        raise


class ExcelToPDFConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to PDF Converter")
        self.root.geometry("650x600")

        # Variables
        self.file_path = None
        self.output_dir = os.path.join(os.getcwd(), "output_pdfs")
        self.name_column = 'نام'

        self.setup_ui()

    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Title
        title_label = ttk.Label(main_frame, text="Excel to PDF Converter",
                                font=('Arial', 18, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=10)

        # BiDi support status
        status_text = "✓ Persian text support enabled" if BIDI_SUPPORT else "⚠ Install arabic-reshaper and python-bidi for better Persian support"
        status_color = "green" if BIDI_SUPPORT else "orange"
        status_label = ttk.Label(main_frame, text=status_text, foreground=status_color,
                                 font=('Arial', 9))
        status_label.grid(row=1, column=0, columnspan=2, pady=5)

        # Drop zone
        self.drop_frame = tk.Frame(main_frame, bg='lightgray', relief=tk.RIDGE, bd=5)
        self.drop_frame.grid(row=2, column=0, columnspan=2, pady=20, padx=20, sticky='nsew')
        self.drop_frame.config(height=150, width=450)

        self.drop_label = tk.Label(self.drop_frame,
                                   text="Drag and Drop Excel File Here\nor Click Browse",
                                   bg='lightgray', font=('Arial', 12))
        self.drop_label.pack(expand=True)

        # Enable drag and drop
        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind('<<Drop>>', self.drop_file)

        # Browse button
        browse_btn = ttk.Button(main_frame, text="📁 Browse File", command=self.browse_file)
        browse_btn.grid(row=3, column=0, columnspan=2, pady=10)

        # File info
        self.file_info = ttk.Label(main_frame, text="No file selected", foreground="gray")
        self.file_info.grid(row=4, column=0, columnspan=2, pady=5)

        # Settings frame
        settings_frame = ttk.LabelFrame(main_frame, text="Settings", padding="15")
        settings_frame.grid(row=5, column=0, columnspan=2, pady=10, sticky='ew')

        # Name column
        ttk.Label(settings_frame, text="Name Column:").grid(row=0, column=0, sticky='w', pady=5)
        self.name_column_var = tk.StringVar(value=self.name_column)
        name_entry = ttk.Entry(settings_frame, textvariable=self.name_column_var, width=25)
        name_entry.grid(row=0, column=1, padx=10, pady=5)

        # Output directory
        ttk.Label(settings_frame, text="Output Directory:").grid(row=1, column=0, sticky='w', pady=5)

        output_frame = ttk.Frame(settings_frame)
        output_frame.grid(row=1, column=1, sticky='w', pady=5)

        self.output_label = ttk.Label(output_frame, text=self.shorten_path(self.output_dir),
                                      foreground="blue")
        self.output_label.pack(side='left', padx=(10, 5))

        ttk.Button(output_frame, text="Change",
                   command=self.change_output_dir).pack(side='left')

        # Convert button
        self.convert_btn = ttk.Button(main_frame, text="🔄 Convert to PDFs",
                                      command=self.convert, state='disabled',
                                      style='Accent.TButton')
        self.convert_btn.grid(row=6, column=0, columnspan=2, pady=20)

        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=7, column=0, columnspan=2, sticky='ew', padx=20)

        # Status label
        self.status_label = ttk.Label(main_frame, text="", foreground="blue")
        self.status_label.grid(row=8, column=0, columnspan=2, pady=10)

        # Configure grid weights
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Style
        style = ttk.Style()
        style.configure('Accent.TButton', foreground='blue')

    def shorten_path(self, path, max_length=40):
        """Shorten path for display"""
        if len(path) <= max_length:
            return path
        return "..." + path[-(max_length - 3):]

    def drop_file(self, event):
        # Get the dropped file path
        file_path = event.data
        # Remove curly braces if present (Windows)
        file_path = file_path.strip('{}')
        # Check if it's an Excel file
        if file_path.endswith(('.xls', '.xlsx', '.xlsm', '.xlsb')):
            self.load_file(file_path)
        else:
            messagebox.showerror("Error", "Please drop an Excel file (.xls, .xlsx, .xlsm, .xlsb)")

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[
                ("Excel Files", "*.xlsx *.xls *.xlsm *.xlsb"),
                ("All Files", "*.*")
            ]
        )
        if file_path:
            self.load_file(file_path)

    def load_file(self, file_path):
        self.file_path = file_path
        filename = os.path.basename(file_path)
        self.file_info.config(text=f"Selected: {filename}", foreground="green")
        self.drop_label.config(text=f"✅ File loaded:\n{filename}")
        self.drop_frame.config(bg='lightgreen')
        self.drop_label.config(bg='lightgreen')
        self.convert_btn.config(state='normal')
        self.status_label.config(text="")

    def change_output_dir(self):
        directory = filedialog.askdirectory(title="Select Output Directory")
        if directory:
            self.output_dir = directory
            self.output_label.config(text=self.shorten_path(directory))

    def convert(self):
        if not self.file_path:
            messagebox.showerror("Error", "Please select an Excel file first")
            return

        # Start conversion in a separate thread
        self.convert_btn.config(state='disabled')
        self.progress.start()
        self.status_label.config(text="Processing...", foreground="blue")

        thread = threading.Thread(target=self.process_conversion)
        thread.start()

    def process_conversion(self):
        try:
            # Create output directory if it doesn't exist
            if not os.path.exists(self.output_dir):
                os.makedirs(self.output_dir)

            # Read Excel file
            df = read_excel_file(self.file_path)

            # Check if name column exists
            name_column = self.name_column_var.get()
            if name_column not in df.columns:
                self.root.after(0, lambda: messagebox.showerror(
                    "Error",
                    f"Column '{name_column}' not found.\nAvailable columns: {', '.join(df.columns)}"
                ))
                return

            # Get unique names
            unique_names = df[name_column].dropna().unique()
            total_names = len(unique_names)

            # Process each person
            success_count = 0
            error_count = 0

            for i, name in enumerate(unique_names):
                # Update status
                self.root.after(0, lambda n=name, i=i, t=total_names:
                self.status_label.config(text=f"Processing {i + 1}/{t}: {n}"))

                # Filter data for this person
                person_data = df[df[name_column] == name]

                # Create PDF
                try:
                    create_pdf_for_person(person_data, name, self.output_dir, df.columns.tolist())
                    success_count += 1
                except Exception as e:
                    error_count += 1
                    print(f"Error creating PDF for {name}: {e}")

            # Show completion message
            message = f"Conversion complete!\n✅ Success: {success_count} PDFs created"
            if error_count > 0:
                message += f"\n❌ Errors: {error_count} PDFs failed"
            message += f"\n\nOutput directory:\n{self.output_dir}"

            self.root.after(0, lambda: messagebox.showinfo("Conversion Complete", message))

            # Update status
            self.root.after(0, lambda: self.status_label.config(
                text=f"Completed: {success_count} PDFs created" + (
                    f", {error_count} errors" if error_count > 0 else ""),
                foreground="green" if error_count == 0 else "orange"
            ))

            # Open output directory
            if success_count > 0:
                if sys.platform == "win32":
                    os.startfile(self.output_dir)
                elif sys.platform == "darwin":
                    os.system(f"open '{self.output_dir}'")
                else:
                    os.system(f"xdg-open '{self.output_dir}'")

        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
            self.root.after(0, lambda: self.status_label.config(
                text=f"Error: {str(e)[:50]}...",
                foreground="red"
            ))
        finally:
            self.root.after(0, self.progress.stop)
            self.root.after(0, lambda: self.convert_btn.config(state='normal'))


def main():
    # Check if tkinterdnd2 is installed
    try:
        root = TkinterDnD.Tk()
    except Exception as e:
        print(f"Error initializing TkinterDnD: {e}")
        print("Make sure tkinterdnd2 is properly installed")
        return

    app = ExcelToPDFConverter(root)
    root.mainloop()


if __name__ == "__main__":
    # Check and install required packages
    required_packages = ['pandas', 'openpyxl', 'xlrd', 'reportlab', 'tkinterdnd2']
    optional_packages = ['arabic-reshaper', 'python-bidi']

    print("Checking packages...")

    for package in required_packages:
        try:
            __import__(package.replace('-', '_'))
        except ImportError:
            print(f"Installing {package}...")
            os.system(f"pip install {package}")

    # Check for optional RTL support packages
    if not BIDI_SUPPORT:
        print("\n" + "=" * 50)
        print("IMPORTANT: For proper Persian/Arabic text display:")
        print("  pip install arabic-reshaper python-bidi")
        print("=" * 50 + "\n")

    main()