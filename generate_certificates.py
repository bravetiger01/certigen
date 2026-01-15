"""
Certificate Generator - OCR-based placeholder detection with exact position and color extraction
"""

from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import numpy as np
import os
import zipfile
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from dataclasses import dataclass
from typing import Optional, Tuple, List
import re
from pathlib import Path

try:
    import pytesseract
    # Set Tesseract path for Windows (adjust if installed elsewhere)
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    HAS_OCR = True
except ImportError:
    HAS_OCR = False
    print("⚠️ pytesseract not installed. Install with: pip install pytesseract")
    print("   Also install Tesseract OCR: https://github.com/UB-Mannheim/tesseract/wiki")


@dataclass
class TextRegion:
    """Detected text region with position and color info"""
    x: int  # center x
    y: int  # center y
    width: int  # max width for text
    height: int
    text_color: Tuple[int, int, int]
    bg_color: Tuple[int, int, int]
    detected_font_size: Optional[int] = None
    # Original placeholder bounding box (to clear it completely)
    placeholder_box: Optional[Tuple[int, int, int, int]] = None  # (x1, y1, x2, y2)


class CertificateGenerator:
    def __init__(
        self,
        template_path: str,
        excel_path: str,
        name_column: str,
        font_path: str,
        output_dir: str = "output",
        placeholder: str = "John Doe",
        font_color: Optional[Tuple[int, int, int]] = None,
        bg_color: Optional[Tuple[int, int, int]] = None,
        base_font_size: int = 180,
        min_font_size: int = 60,
        # Manual position override (use if OCR fails)
        manual_position: Optional[Tuple[int, int]] = None,  # (center_x, center_y)
        max_text_width: Optional[int] = None,
    ):
        self.template_path = template_path
        self.excel_path = excel_path
        self.name_column = name_column
        self.font_path = font_path
        self.output_dir = output_dir
        self.placeholder = placeholder
        self.user_font_color = font_color
        self.user_bg_color = bg_color
        self.base_font_size = base_font_size
        self.min_font_size = min_font_size
        self.manual_position = manual_position
        self.max_text_width = max_text_width
        
        # Load template and names
        self.template = Image.open(template_path).convert("RGB")
        self.names = self._load_names()
        
        # Detect placeholder position and colors
        self.text_region = self._detect_placeholder()
        
        os.makedirs(output_dir, exist_ok=True)
        
        print(f"\n📍 Detected position: ({self.text_region.x}, {self.text_region.y})")
        print(f"🎨 Text color: {self.text_region.text_color}")
        print(f"🖼️ Background color: {self.text_region.bg_color}")
        print(f"📏 Max width: {self.text_region.width}")
        print(f"🔤 Base font size: {self.text_region.detected_font_size or self.base_font_size}px")
        if self.text_region.placeholder_box:
            print(f"📦 Placeholder box to clear: {self.text_region.placeholder_box}")
        print()

    def _load_names(self) -> List[str]:
        """Load names from Excel file"""
        df = pd.read_excel(self.excel_path)
        return [str(name).strip() for name in df[self.name_column] if pd.notna(name)]

    def _detect_placeholder(self) -> TextRegion:
        """
        Use OCR to find the exact position of placeholder text in template.
        Falls back to manual position or center if OCR fails.
        """
        img_array = np.array(self.template)
        height, width = img_array.shape[:2]
        
        # Default values
        detected_x, detected_y = width // 2, height // 2
        detected_width = int(width * 0.6)
        detected_height = 100
        detected_font_size = None
        placeholder_box = None
        found_placeholder = False
        
        # Try OCR detection
        if HAS_OCR and self.manual_position is None:
            try:
                # Get detailed OCR data with bounding boxes
                ocr_data = pytesseract.image_to_data(
                    self.template, 
                    output_type=pytesseract.Output.DICT
                )
                
                # Search for placeholder text
                placeholder_lower = self.placeholder.lower()
                placeholder_words = placeholder_lower.split()
                
                # Find all matching word indices
                matching_indices = []
                for i, text in enumerate(ocr_data['text']):
                    if not text.strip():
                        continue
                    text_lower = text.strip().lower()
                    if text_lower in placeholder_words:
                        matching_indices.append(i)
                
                if matching_indices:
                    # Get bounding box that covers all matching words
                    x_min = min(ocr_data['left'][i] for i in matching_indices)
                    y_min = min(ocr_data['top'][i] for i in matching_indices)
                    x_max = max(ocr_data['left'][i] + ocr_data['width'][i] for i in matching_indices)
                    y_max = max(ocr_data['top'][i] + ocr_data['height'][i] for i in matching_indices)
                    
                    full_w = x_max - x_min
                    full_h = y_max - y_min
                    
                    # Store the exact placeholder box for clearing
                    placeholder_box = (x_min, y_min, x_max, y_max)
                    
                    # Calculate center of detected region
                    detected_x = x_min + full_w // 2
                    detected_y = y_min + full_h // 2
                    detected_width = max(full_w, int(width * 0.4))
                    detected_height = full_h
                    
                    # Estimate font size from detected text
                    detected_font_size = self._estimate_font_size_from_placeholder(
                        self.placeholder, full_w, full_h
                    )
                    
                    found_placeholder = True
                    print(f"✅ Found placeholder '{self.placeholder}' via OCR")
                    print(f"   Bounding box: ({x_min}, {y_min}) to ({x_max}, {y_max}) = {full_w}x{full_h}px")
                
                if not found_placeholder:
                    print(f"⚠️ Placeholder '{self.placeholder}' not found via OCR, using center")
                    
            except Exception as e:
                print(f"⚠️ OCR failed: {e}")
        
        # Use manual position if provided
        if self.manual_position:
            detected_x, detected_y = self.manual_position
            print(f"📍 Using manual position: ({detected_x}, {detected_y})")
        
        if self.max_text_width:
            detected_width = self.max_text_width
        
        # Detect colors from the region
        text_color, bg_color = self._extract_colors(
            detected_x, detected_y, detected_width, detected_height
        )
        
        # Override with user colors if provided
        if self.user_font_color:
            text_color = self.user_font_color
        if self.user_bg_color:
            bg_color = self.user_bg_color
        
        return TextRegion(
            x=detected_x,
            y=detected_y,
            width=detected_width,
            height=detected_height,
            text_color=text_color,
            bg_color=bg_color,
            detected_font_size=detected_font_size,
            placeholder_box=placeholder_box
        )

    def _estimate_font_size_from_placeholder(self, placeholder: str, box_width: int, box_height: int) -> int:
        """
        Estimate the font size used for placeholder text by finding the size
        that produces similar dimensions. Prioritize width matching.
        """
        best_size = 50
        best_width_diff = float('inf')
        
        for size in range(10, 400, 1):
            try:
                font = ImageFont.truetype(self.font_path, size)
                bbox = font.getbbox(placeholder)
                text_width = bbox[2] - bbox[0]
                
                width_diff = abs(text_width - box_width)
                
                if width_diff < best_width_diff:
                    best_width_diff = width_diff
                    best_size = size
                    
                # Stop if we overshoot
                if text_width > box_width + 20:
                    break
                    
            except Exception:
                continue
        
        print(f"   Estimated font size: {best_size}px (placeholder width: {box_width}px)")
        return best_size

    def _extract_colors(
        self, 
        center_x: int, 
        center_y: int, 
        width: int, 
        height: int
    ) -> Tuple[Tuple[int, int, int], Tuple[int, int, int]]:
        """
        Extract text and background colors from the region around detected text.
        Uses color clustering to find dominant (background) and secondary (text) colors.
        """
        img_array = np.array(self.template)
        img_h, img_w = img_array.shape[:2]
        
        # Define sampling region
        x1 = max(0, center_x - width // 2)
        x2 = min(img_w, center_x + width // 2)
        y1 = max(0, center_y - height)
        y2 = min(img_h, center_y + height)
        
        region = img_array[y1:y2, x1:x2]
        pixels = region.reshape(-1, 3)
        
        # Find unique colors and their counts
        unique_colors, counts = np.unique(pixels, axis=0, return_counts=True)
        
        # Sort by frequency
        sorted_indices = np.argsort(-counts)
        
        # Background is most common color
        bg_color = tuple(int(x) for x in unique_colors[sorted_indices[0]])
        
        # Text color is the most different from background among top colors
        bg_array = np.array(bg_color)
        best_text_color = (0, 0, 0)  # Default to black
        best_distance = 0
        
        # Check top 20 most common colors for text color
        for idx in sorted_indices[:20]:
            color = unique_colors[idx]
            distance = np.sqrt(np.sum((color - bg_array) ** 2))
            
            # Must be significantly different from background
            if distance > best_distance and distance > 30:
                best_distance = distance
                best_text_color = tuple(int(x) for x in color)
        
        # If no good text color found, use black or white based on background brightness
        if best_distance < 30:
            brightness = sum(bg_color) / 3
            best_text_color = (0, 0, 0) if brightness > 128 else (255, 255, 255)
        
        return best_text_color, bg_color

    def _calculate_font_size(self, name: str, max_width: int) -> int:
        """Calculate optimal font size so the name fits within the same width as placeholder"""
        base_size = self.text_region.detected_font_size or self.base_font_size
        
        # Find the size that makes this name fit in the same width as placeholder
        # Start from base size and adjust
        for size in range(base_size + 20, self.min_font_size - 1, -1):
            font = ImageFont.truetype(self.font_path, size)
            bbox = font.getbbox(name)
            text_width = bbox[2] - bbox[0]
            if text_width <= max_width:
                return size
        
        return self.min_font_size

    def _generate_single(self, name: str, index: int) -> str:
        """Generate a single certificate"""
        img = self.template.copy()
        draw = ImageDraw.Draw(img)
        img_width, img_height = img.size
        
        # Calculate font size for this name
        max_width = self.text_region.width
        font_size = self._calculate_font_size(name, max_width)
        font = ImageFont.truetype(self.font_path, font_size)
        
        # Get text dimensions for clearing
        text_bbox = draw.textbbox((0, 0), name, font=font, anchor="lt")
        text_width = text_bbox[2] - text_bbox[0]
        text_height = text_bbox[3] - text_bbox[1]
        
        # Get the original placeholder position
        if self.text_region.placeholder_box:
            px1, py1, px2, py2 = self.text_region.placeholder_box
            # Use the LEFT edge of placeholder (preserve left-alignment)
            # and vertical center for baseline alignment
            start_x = px1
            center_y = (py1 + py2) // 2
        else:
            # Fallback to detected center
            start_x = self.text_region.x - (text_width // 2)
            center_y = self.text_region.y
        
        # FIRST: Clear the original placeholder area completely
        if self.text_region.placeholder_box:
            px1, py1, px2, py2 = self.text_region.placeholder_box
            padding = 25
            draw.rectangle(
                [px1 - padding, py1 - padding, px2 + padding, py2 + padding],
                fill=self.text_region.bg_color
            )
        
        # Also clear the area where new text will go
        half_h = text_height // 2
        padding = 10
        clear_x1 = max(0, start_x - padding)
        clear_y1 = max(0, center_y - half_h - padding)
        clear_x2 = min(img_width, start_x + text_width + padding)
        clear_y2 = min(img_height, center_y + half_h + padding)
        
        draw.rectangle(
            [clear_x1, clear_y1, clear_x2, clear_y2],
            fill=self.text_region.bg_color
        )
        
        # Boundary checks
        margin = 10
        if start_x < margin:
            start_x = margin
        if start_x + text_width > img_width - margin:
            start_x = img_width - margin - text_width
        
        # Draw name: left-aligned at start_x, vertically centered at center_y
        # Using anchor="lm" = left edge, middle vertically
        draw.text(
            (start_x, center_y), 
            name, 
            fill=self.text_region.text_color, 
            font=font,
            anchor="lm"  # left-middle: left edge at x, vertical center at y
        )
        
        # Save
        safe_name = re.sub(r'[^\w\s-]', '', name).replace(" ", "_")
        output_path = os.path.join(self.output_dir, f"{safe_name}_certificate.png")
        img.save(output_path, "PNG", quality=95)
        
        print(f"[{index}/{len(self.names)}] Generated: {safe_name}_certificate.png (font: {font_size}px, x: {start_x})")
        return output_path

    def generate_all(self) -> List[str]:
        """Generate certificates for all names"""
        paths = []
        for idx, name in enumerate(self.names, 1):
            path = self._generate_single(name, idx)
            paths.append(path)
        print(f"\n✅ Generated {len(paths)} certificates in '{self.output_dir}'")
        return paths

    def export_as_pdf(self, output_name: str = "certificates.pdf") -> str:
        """Convert all PNGs to a single PDF"""
        png_files = sorted(Path(self.output_dir).glob("*_certificate.png"))
        if not png_files:
            raise ValueError("No certificates found. Run generate_all() first.")
        
        images = [Image.open(f).convert("RGB") for f in png_files]
        pdf_path = os.path.join(self.output_dir, output_name)
        
        images[0].save(pdf_path, "PDF", save_all=True, append_images=images[1:])
        print(f"📄 PDF created: {pdf_path}")
        return pdf_path

    def zip_certificates(self, output_name: str = "certificates.zip") -> str:
        """Zip all generated certificates"""
        zip_path = os.path.join(self.output_dir, output_name)
        
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for file in Path(self.output_dir).glob("*_certificate.png"):
                zf.write(file, file.name)
        
        print(f"📦 Zipped certificates: {zip_path}")
        return zip_path

    def email_certificates(
        self,
        smtp_server: str,
        smtp_port: int,
        sender_email: str,
        sender_password: str,
        recipient_emails: List[str],
        subject: str = "Your Certificate",
        body: str = "Please find your certificate attached."
    ):
        """Email certificates to recipients"""
        zip_path = self.zip_certificates()
        
        for recipient in recipient_emails:
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = recipient
            msg['Subject'] = subject
            
            with open(zip_path, 'rb') as f:
                part = MIMEBase('application', 'zip')
                part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', 'attachment; filename="certificates.zip"')
                msg.attach(part)
            
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(sender_email, sender_password)
                server.send_message(msg)
            
            print(f"📧 Sent to: {recipient}")

    def upload_to_s3(
        self,
        bucket_name: str,
        aws_access_key: str,
        aws_secret_key: str,
        region: str = "us-east-1",
        prefix: str = "certificates/"
    ):
        """Upload certificates to AWS S3"""
        import boto3
        
        s3 = boto3.client(
            's3',
            aws_access_key_id=aws_access_key,
            aws_secret_access_key=aws_secret_key,
            region_name=region
        )
        
        for file in Path(self.output_dir).glob("*_certificate.png"):
            key = f"{prefix}{file.name}"
            s3.upload_file(str(file), bucket_name, key)
            print(f"☁️ Uploaded: {key}")
        
        print(f"✅ All certificates uploaded to s3://{bucket_name}/{prefix}")

    def upload_to_drive(self, credentials_path: str, folder_id: Optional[str] = None):
        """Upload certificates to Google Drive"""
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaFileUpload
        
        creds = service_account.Credentials.from_service_account_file(
            credentials_path,
            scopes=['https://www.googleapis.com/auth/drive.file']
        )
        service = build('drive', 'v3', credentials=creds)
        
        for file in Path(self.output_dir).glob("*_certificate.png"):
            metadata = {'name': file.name}
            if folder_id:
                metadata['parents'] = [folder_id]
            
            media = MediaFileUpload(str(file), mimetype='image/png')
            service.files().create(body=metadata, media_body=media).execute()
            print(f"📁 Uploaded to Drive: {file.name}")
        
        print("✅ All certificates uploaded to Google Drive")


# ---------------- HELPER: Find coordinates interactively ---------------- #

def find_coordinates(template_path: str):
    """
    Interactive tool to find the exact coordinates for text placement.
    Opens the image and lets you click to get coordinates.
    """
    import matplotlib.pyplot as plt
    
    img = Image.open(template_path)
    fig, ax = plt.subplots(figsize=(12, 8))
    ax.imshow(img)
    ax.set_title("Click on the CENTER of where the name should go\nClose window when done")
    
    coords = []
    
    def onclick(event):
        if event.xdata and event.ydata:
            x, y = int(event.xdata), int(event.ydata)
            coords.append((x, y))
            print(f"Clicked: ({x}, {y})")
            ax.plot(x, y, 'r+', markersize=20, markeredgewidth=3)
            fig.canvas.draw()
    
    fig.canvas.mpl_connect('button_press_event', onclick)
    plt.show()
    
    if coords:
        print(f"\n📍 Use this position: manual_position=({coords[-1][0]}, {coords[-1][1]})")
    return coords[-1] if coords else None


# ---------------- CLI INTERFACE ---------------- #

def main():
    """Command-line interface"""
    import argparse
    
    parser = argparse.ArgumentParser(description="Certificate Generator")
    parser.add_argument("--template", "-t", required=True, help="Template image path")
    parser.add_argument("--excel", "-e", required=True, help="Excel file with names")
    parser.add_argument("--column", "-c", default="Name", help="Column name for names")
    parser.add_argument("--font", "-f", required=True, help="Font file path (.ttf)")
    parser.add_argument("--output", "-o", default="output", help="Output directory")
    parser.add_argument("--placeholder", "-p", default="John Doe", help="Placeholder text to find")
    parser.add_argument("--font-color", help="Font color as R,G,B (e.g., 0,0,0)")
    parser.add_argument("--bg-color", help="Background color as R,G,B (e.g., 255,255,255)")
    parser.add_argument("--position", help="Manual position as X,Y (e.g., 500,300)")
    parser.add_argument("--max-width", type=int, help="Max text width in pixels")
    parser.add_argument("--font-size", type=int, default=180, help="Base font size")
    parser.add_argument("--min-font-size", type=int, default=60, help="Minimum font size")
    parser.add_argument("--zip", action="store_true", help="Create zip file")
    parser.add_argument("--pdf", action="store_true", help="Create PDF")
    parser.add_argument("--find-coords", action="store_true", help="Interactive coordinate finder")
    
    args = parser.parse_args()
    
    # Interactive coordinate finder
    if args.find_coords:
        find_coordinates(args.template)
        return
    
    # Parse colors and position if provided
    font_color = tuple(map(int, args.font_color.split(','))) if args.font_color else None
    bg_color = tuple(map(int, args.bg_color.split(','))) if args.bg_color else None
    position = tuple(map(int, args.position.split(','))) if args.position else None
    
    generator = CertificateGenerator(
        template_path=args.template,
        excel_path=args.excel,
        name_column=args.column,
        font_path=args.font,
        output_dir=args.output,
        placeholder=args.placeholder,
        font_color=font_color,
        bg_color=bg_color,
        manual_position=position,
        max_text_width=args.max_width,
        base_font_size=args.font_size,
        min_font_size=args.min_font_size,
    )
    
    generator.generate_all()
    
    if args.zip:
        generator.zip_certificates()
    if args.pdf:
        generator.export_as_pdf()


# ---------------- EXAMPLE USAGE ---------------- #

if __name__ == "__main__":
    # Let OCR detect placeholder position automatically
    generator = CertificateGenerator(
        template_path="hgfxdjh.png",
        excel_path="Names for Certi.xlsx",
        name_column="Name",
        font_path="OpenSans-VariableFont_wdth,wght.ttf",
        output_dir="output",
        placeholder="John Doe",  # Text to find and replace in template
    )
    
    generator.generate_all()
