# CertiGen 🎓

Automatic certificate generator with OCR-based placeholder detection. Generate hundreds of personalized certificates in seconds!

## Features

- 🔍 **Auto-detect placeholder** - Uses OCR to find "John Doe" or any placeholder text
- 🎨 **Auto-detect colors** - Automatically extracts font and background colors
- 📏 **Smart text sizing** - Automatically resizes text for long names
- 📄 **Multiple exports** - PNG, PDF, or ZIP
- ☁️ **Cloud upload** - S3 and Google Drive support
- 📧 **Email** - Send certificates directly via email

## Installation

```bash
pip install certigen
```

For OCR support (recommended):
```bash
pip install certigen[ocr]
```

For all features:
```bash
pip install certigen[all]
```

### Tesseract OCR (Required for auto-detection)

- **Windows**: Download from [UB-Mannheim](https://github.com/UB-Mannheim/tesseract/wiki)
- **macOS**: `brew install tesseract`
- **Linux**: `sudo apt install tesseract-ocr`

## Quick Start

### Python API

```python
from certigen import CertificateGenerator

# Basic usage - auto-detects everything
gen = CertificateGenerator(
    template_path="template.png",
    excel_path="names.xlsx",
    name_column="Name",
    font_path="arial.ttf",
    placeholder="John Doe"
)
gen.generate_all()

# Export options
gen.export_as_pdf()
gen.zip_certificates()
```

### Command Line

```bash
# Basic usage
certigen -t template.png -e names.xlsx -f arial.ttf

# With options
certigen -t template.png -e names.xlsx -f arial.ttf -p "John Doe" --pdf --zip

# Find coordinates interactively
certigen -t template.png --find-coords
```

## Advanced Usage

### Manual Position Override

If OCR doesn't detect the placeholder correctly:

```python
gen = CertificateGenerator(
    template_path="template.png",
    excel_path="names.xlsx",
    name_column="Name",
    font_path="arial.ttf",
    manual_position=(800, 600),  # (x, y) center point
    font_color=(0, 0, 0),        # RGB black
    bg_color=(255, 255, 255),    # RGB white
)
```

### Cloud Upload

```python
# AWS S3
gen.upload_to_s3(
    bucket="my-bucket",
    access_key="...",
    secret_key="...",
)

# Google Drive
gen.upload_to_drive(
    credentials_path="service_account.json",
    folder_id="..."
)
```

### Email Certificates

```python
gen.email_certificates(
    smtp_server="smtp.gmail.com",
    smtp_port=587,
    sender_email="you@gmail.com",
    sender_password="app_password",
    recipient_emails=["recipient@example.com"]
)
```

## CLI Options

| Option | Description |
|--------|-------------|
| `-t, --template` | Template image path (required) |
| `-e, --excel` | Excel/CSV file with names (required) |
| `-f, --font` | Font file path (required) |
| `-c, --column` | Column name for names (default: "Name") |
| `-o, --output` | Output directory (default: "output") |
| `-p, --placeholder` | Placeholder text (default: "John Doe") |
| `--font-color` | Font color as R,G,B |
| `--bg-color` | Background color as R,G,B |
| `--position` | Manual position as X,Y |
| `--pdf` | Create combined PDF |
| `--zip` | Create ZIP archive |
| `--find-coords` | Interactive coordinate finder |

## License

MIT License
