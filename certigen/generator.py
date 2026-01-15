"""
Certificate Generator - Core module with OCR-based placeholder detection
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
from typing import Optional, Tuple, List, Union
import re
from pathlib import Path

try:
    import pytesseract
    HAS_OCR = True
except ImportError:
    HAS_OCR = False


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
    placeholder_box: Optional[Tuple[int, int, int, int]] = None  # (x1, y1, x2, y2)


class CertificateGenerator:
    """
    Generate certificates by replacing placeholder text with names from Excel/CSV.
    
    Features:
    - OCR-based automatic placeholder detection
    - Auto-detect font color and background color
    - Auto-resize text for long names
    - Export to PNG, PDF, or ZIP
    - Upload to S3 or Google Drive
    
    Example:
        >>> from certigen import CertificateGenerator
        >>> gen = CertificateGenerator(
        ...     template_path="template.png",
        ...     excel_path="names.xlsx",
        ...     name_column="Name",
        ...     font_path="arial.ttf",
        ...     placeholder="John Doe"
        ... )
        >>> gen.generate_all()
    """
    
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
        manual_position: Optional[Tuple[int, int]] = None,
        max_text_width: Optional[int] = None,
        tesseract_path: Optional[str] = None,
        verbose: bool = True,
    ):
        """
        Initialize the certificate generator.
        
        Args:
            template_path: Path to certificate template image
            excel_path: Path to Excel/CSV file with names
            name_column: Column name containing the names
            font_path: Path to .ttf font file
            output_dir: Directory for generated certificates
            placeholder: Text to find and replace (e.g., "John Doe")
            font_color: RGB tuple for text color (auto-detected if None)
            bg_color: RGB tuple for background color (auto-detected if None)
            base_font_size: Starting font size for text
            min_font_size: Minimum font size for long names
            manual_position: (x, y) tuple to override OCR detection
            max_text_width: Maximum width for text in pixels
            tesseract_path: Path to Tesseract executable (for Windows)
            verbose: Print progress messages
        """
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
        self.verbose = verbose
        
        # Configure Tesseract path
        if tesseract_path and HAS_OCR:
            pytesseract.pytesseract.tesseract_cmd = tesseract_path
        elif HAS_OCR and os.name == 'nt':
            # Default Windows path
            default_path = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
            if os.path.exists(default_path):
                pytesseract.pytesseract.tesseract_cmd = default_path
        
        # Load template and names
        self.template = Image.open(template_path).convert("RGB")
        self.names = self._load_names()
        
        # Detect placeholder position and colors
        self.text_region = self._detect_placeholder()
        
        os.makedirs(output_dir, exist_ok=True)
        
        if self.verbose:
            print(f"\n📍 Detected position: ({self.text_region.x}, {self.text_region.y})")
            print(f"🎨 Text color: {self.text_region.text_color}")
            print(f"🖼️ Background color: {self.text_region.bg_color}")
            print(f"📏 Max width: {self.text_region.width}")
            print(f"🔤 Font size: {self.text_region.detected_font_size or self.base_font_size}px")
            if self.text_region.placeholder_box:
                print(f"📦 Placeholder box: {self.text_region.placeholder_box}")
            print()

    def _load_names(self) -> List[str]:
        """Load names from Excel or CSV file"""
        path = Path(self.excel_path)
        if path.suffix.lower() == '.csv':
            df = pd.read_csv(self.excel_path)
        else:
            df = pd.read_excel(self.excel_path)
        return [str(name).strip() for name in df[self.name_column] if pd.notna(name)]

    def _detect_placeholder(self) -> TextRegion:
        """Use OCR to find placeholder text position, with fallbacks"""
        img_array = np.array(self.template)
        height, width = img_array.shape[:2]
        
        detected_x, detected_y = width // 2, height // 2
        detected_width = int(width * 0.6)
        detected_height = 100
        detected_font_size = None
        placeholder_box = None
        found_placeholder = False
        
        if HAS_OCR and self.manual_position is None:
            try:
                ocr_data = pytesseract.image_to_data(
                    self.template, output_type=pytesseract.Output.DICT
                )
                
                placeholder_words = self.placeholder.lower().split()
                matching_indices = []
                
                for i, text in enumerate(ocr_data['text']):
                    if text.strip() and text.strip().lower() in placeholder_words:
                        matching_indices.append(i)
                
                if matching_indices:
                    x_min = min(ocr_data['left'][i] for i in matching_indices)
                    y_min = min(ocr_data['top'][i] for i in matching_indices)
                    x_max = max(ocr_data['left'][i] + ocr_data['width'][i] for i in matching_indices)
                    y_max = max(ocr_data['top'][i] + ocr_data['height'][i] for i in matching_indices)
                    
                    full_w = x_max - x_min
                    full_h = y_max - y_min
                    placeholder_box = (x_min, y_min, x_max, y_max)
                    
                    detected_x = x_min + full_w // 2
                    detected_y = y_min + full_h // 2
                    detected_width = max(full_w, int(width * 0.4))
                    detected_height = full_h
                    
                    detected_font_size = self._estimate_font_size(self.placeholder, full_w)
                    found_placeholder = True
                    
                    if self.verbose:
                        print(f"✅ Found placeholder '{self.placeholder}' via OCR")
                        print(f"   Bounding box: ({x_min}, {y_min}) to ({x_max}, {y_max})")
                
                if not found_placeholder and self.verbose:
                    print(f"⚠️ Placeholder '{self.placeholder}' not found, using center")
                    
            except Exception as e:
                if self.verbose:
                    print(f"⚠️ OCR failed: {e}")
        
        if self.manual_position:
            detected_x, detected_y = self.manual_position
            if self.verbose:
                print(f"📍 Using manual position: ({detected_x}, {detected_y})")
        
        if self.max_text_width:
            detected_width = self.max_text_width
        
        text_color, bg_color = self._extract_colors(
            detected_x, detected_y, detected_width, detected_height
        )
        
        if self.user_font_color:
            text_color = self.user_font_color
        if self.user_bg_color:
            bg_color = self.user_bg_color
        
        return TextRegion(
            x=detected_x, y=detected_y, width=detected_width, height=detected_height,
            text_color=text_color, bg_color=bg_color,
            detected_font_size=detected_font_size, placeholder_box=placeholder_box
        )

    def _estimate_font_size(self, text: str, target_width: int) -> int:
        """Find font size that matches target width"""
        best_size = 50
        best_diff = float('inf')
        
        for size in range(10, 400):
            try:
                font = ImageFont.truetype(self.font_path, size)
                bbox = font.getbbox(text)
                text_width = bbox[2] - bbox[0]
                diff = abs(text_width - target_width)
                
                if diff < best_diff:
                    best_diff = diff
                    best_size = size
                if text_width > target_width + 20:
                    break
            except Exception:
                continue
        
        if self.verbose:
            print(f"   Estimated font size: {best_size}px")
        return best_size

    def _extract_colors(self, cx: int, cy: int, w: int, h: int) -> Tuple[Tuple[int, int, int], Tuple[int, int, int]]:
        """Extract text and background colors from region"""
        img_array = np.array(self.template)
        img_h, img_w = img_array.shape[:2]
        
        x1, x2 = max(0, cx - w // 2), min(img_w, cx + w // 2)
        y1, y2 = max(0, cy - h), min(img_h, cy + h)
        
        region = img_array[y1:y2, x1:x2]
        pixels = region.reshape(-1, 3)
        unique_colors, counts = np.unique(pixels, axis=0, return_counts=True)
        sorted_indices = np.argsort(-counts)
        
        bg_color = tuple(int(x) for x in unique_colors[sorted_indices[0]])
        bg_array = np.array(bg_color)
        
        best_text_color = (0, 0, 0)
        best_distance = 0
        
        for idx in sorted_indices[:20]:
            color = unique_colors[idx]
            distance = np.sqrt(np.sum((color - bg_array) ** 2))
            if distance > best_distance and distance > 30:
                best_distance = distance
                best_text_color = tuple(int(x) for x in color)
        
        if best_distance < 30:
            brightness = sum(bg_color) / 3
            best_text_color = (0, 0, 0) if brightness > 128 else (255, 255, 255)
        
        return best_text_color, bg_color

    def _calculate_font_size(self, name: str, max_width: int) -> int:
        """Calculate font size to fit name within max width"""
        base_size = self.text_region.detected_font_size or self.base_font_size
        
        for size in range(base_size + 20, self.min_font_size - 1, -1):
            font = ImageFont.truetype(self.font_path, size)
            bbox = font.getbbox(name)
            if bbox[2] - bbox[0] <= max_width:
                return size
        return self.min_font_size

    def _generate_single(self, name: str, index: int) -> str:
        """Generate a single certificate"""
        img = self.template.copy()
        draw = ImageDraw.Draw(img)
        img_width, img_height = img.size
        
        font_size = self._calculate_font_size(name, self.text_region.width)
        font = ImageFont.truetype(self.font_path, font_size)
        
        text_bbox = draw.textbbox((0, 0), name, font=font, anchor="lt")
        text_width = text_bbox[2] - text_bbox[0]
        text_height = text_bbox[3] - text_bbox[1]
        
        if self.text_region.placeholder_box:
            px1, py1, px2, py2 = self.text_region.placeholder_box
            start_x = px1
            center_y = (py1 + py2) // 2
        else:
            start_x = self.text_region.x - (text_width // 2)
            center_y = self.text_region.y
        
        # Clear placeholder area
        if self.text_region.placeholder_box:
            px1, py1, px2, py2 = self.text_region.placeholder_box
            draw.rectangle([px1 - 25, py1 - 25, px2 + 25, py2 + 25], fill=self.text_region.bg_color)
        
        # Clear new text area
        half_h = text_height // 2
        draw.rectangle(
            [max(0, start_x - 10), max(0, center_y - half_h - 10),
             min(img_width, start_x + text_width + 10), min(img_height, center_y + half_h + 10)],
            fill=self.text_region.bg_color
        )
        
        # Boundary checks
        if start_x < 10:
            start_x = 10
        if start_x + text_width > img_width - 10:
            start_x = img_width - 10 - text_width
        
        draw.text((start_x, center_y), name, fill=self.text_region.text_color, font=font, anchor="lm")
        
        safe_name = re.sub(r'[^\w\s-]', '', name).replace(" ", "_")
        output_path = os.path.join(self.output_dir, f"{safe_name}_certificate.png")
        img.save(output_path, "PNG", quality=95)
        
        if self.verbose:
            print(f"[{index}/{len(self.names)}] Generated: {safe_name}_certificate.png")
        return output_path

    def generate_all(self) -> List[str]:
        """Generate certificates for all names"""
        paths = []
        for idx, name in enumerate(self.names, 1):
            paths.append(self._generate_single(name, idx))
        if self.verbose:
            print(f"\n✅ Generated {len(paths)} certificates in '{self.output_dir}'")
        return paths

    def export_as_pdf(self, output_name: str = "certificates.pdf") -> str:
        """Combine all certificates into a single PDF"""
        png_files = sorted(Path(self.output_dir).glob("*_certificate.png"))
        if not png_files:
            raise ValueError("No certificates found. Run generate_all() first.")
        
        images = [Image.open(f).convert("RGB") for f in png_files]
        pdf_path = os.path.join(self.output_dir, output_name)
        images[0].save(pdf_path, "PDF", save_all=True, append_images=images[1:])
        
        if self.verbose:
            print(f"📄 PDF created: {pdf_path}")
        return pdf_path

    def zip_certificates(self, output_name: str = "certificates.zip") -> str:
        """Create ZIP archive of all certificates"""
        zip_path = os.path.join(self.output_dir, output_name)
        
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for file in Path(self.output_dir).glob("*_certificate.png"):
                zf.write(file, file.name)
        
        if self.verbose:
            print(f"📦 Zipped: {zip_path}")
        return zip_path

    def email_certificates(
        self, smtp_server: str, smtp_port: int, sender_email: str,
        sender_password: str, recipient_emails: List[str],
        subject: str = "Your Certificate", body: str = "Please find your certificate attached."
    ):
        """Email certificates as ZIP attachment"""
        zip_path = self.zip_certificates()
        
        for recipient in recipient_emails:
            msg = MIMEMultipart()
            msg['From'], msg['To'], msg['Subject'] = sender_email, recipient, subject
            
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
            
            if self.verbose:
                print(f"📧 Sent to: {recipient}")

    def upload_to_s3(self, bucket: str, access_key: str, secret_key: str, 
                     region: str = "us-east-1", prefix: str = "certificates/"):
        """Upload certificates to AWS S3"""
        import boto3
        s3 = boto3.client('s3', aws_access_key_id=access_key, 
                          aws_secret_access_key=secret_key, region_name=region)
        
        for file in Path(self.output_dir).glob("*_certificate.png"):
            key = f"{prefix}{file.name}"
            s3.upload_file(str(file), bucket, key)
            if self.verbose:
                print(f"☁️ Uploaded: {key}")

    def upload_to_drive(self, credentials_path: str, folder_id: Optional[str] = None):
        """Upload certificates to Google Drive"""
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaFileUpload
        
        creds = service_account.Credentials.from_service_account_file(
            credentials_path, scopes=['https://www.googleapis.com/auth/drive.file']
        )
        service = build('drive', 'v3', credentials=creds)
        
        for file in Path(self.output_dir).glob("*_certificate.png"):
            metadata = {'name': file.name}
            if folder_id:
                metadata['parents'] = [folder_id]
            media = MediaFileUpload(str(file), mimetype='image/png')
            service.files().create(body=metadata, media_body=media).execute()
            if self.verbose:
                print(f"📁 Uploaded: {file.name}")
