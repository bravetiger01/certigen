⭐ If you find CertiGen useful, consider starring the repo!

# CertiGen 🎓

Automatic certificate generator with OCR-based placeholder detection. Generate hundreds of personalized certificates in seconds!

## Features

- 🔍 **Auto-detect placeholder** - Uses OCR to find "John Doe" or any placeholder text
- 🎨 **Auto-detect colors** - Automatically extracts font and background colors
- 📏 **Smart text sizing** - Automatically resizes text for long names
- 📄 **Multiple exports** - PNG, PDF, or ZIP
- ☁️ **Cloud upload** - S3 and Google Drive support
- 📧 **Email** - Send certificates directly via email
- 🚀 **No external dependencies** - OCR works out of the box, no Tesseract needed!

## Installation

```bash
pip install certigen
```

That's it! OCR is included and works offline. No external software required.

For cloud upload features:
```bash
pip install certigen[cloud]
```

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

## How It Works

1. **OCR Detection** - Scans your template to find the placeholder text (e.g., "John Doe")
2. **Color Extraction** - Automatically detects text color and background color
3. **Font Matching** - Estimates the font size from the placeholder
4. **Generation** - Replaces placeholder with each name, auto-scaling for long names

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

## Parameters

| Parameter | Description | Default |
|-----------|-------------|---------|
| `template_path` | Path to certificate template image | Required |
| `excel_path` | Path to Excel/CSV file with names | Required |
| `name_column` | Column name containing names | Required |
| `font_path` | Path to .ttf font file | Required |
| `placeholder` | Text to find and replace | "John Doe" |
| `output_dir` | Output directory | "output" |
| `font_color` | RGB tuple for text color | Auto-detected |
| `bg_color` | RGB tuple for background | Auto-detected |
| `manual_position` | (x, y) to override OCR | None |
| `base_font_size` | Starting font size | 180 |
| `min_font_size` | Minimum font size | 60 |
| `verbose` | Print progress | True |

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

## Changelog

### v1.1.0
- Replaced Tesseract with RapidOCR (no external dependencies!)
- OCR now works out of the box with just `pip install`
- Removed `tesseract_path` parameter

### v1.0.0
- Initial release

## License

MIT License
