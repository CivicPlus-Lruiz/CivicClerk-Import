"""
OCR Handler for Meeting Import GUI
Extracts dates and meeting info from PDFs using text extraction and OCR fallback
Based on FileScanPDFRename v2.5
"""

import os
import re
from datetime import datetime, date
from typing import List, Optional, Tuple, Dict
import warnings

# Suppress torch warnings
warnings.filterwarnings('ignore', category=UserWarning, module='torch')

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

# EasyOCR will be lazy-loaded
EASYOCR_READER = None


class OCRHandler:
    """Handles PDF text extraction and OCR for date/meeting info detection."""
    
    # Date search limits
    DATE_SEARCH_LIMIT_PDF = 1500  # Increased: cover headers pushed down by boilerplate text
    DATE_SEARCH_LIMIT_OCR = 600
    
    # Time patterns
    TIME_PATTERNS = [
        r'\b(1[0-2]|0?[1-9]):[0-5][0-9]\s?(AM|PM)\b',
        r'\b(1[0-2]|0?[1-9])\s?(AM|PM)\b',
        r'\b([01]?[0-9]|2[0-3]):[0-5][0-9]\b',
    ]
    
    # Subtype patterns (ordered by priority)
    SUBTYPE_PATTERNS = [
        ("Emergency", [r"\bemergency\b"]),
        ("Executive_Session", [r"\bexecutive session\b", r"\bclosed session\b"]),
        ("Special", [r"\bspecial\b", r"\bspc\b"]),
        ("Joint", [r"\bjoint\b"]),
        ("Work_Session", [r"\bwork session\b", r"\bworksession\b", r"\bstudy session\b"]),
        ("Workshop", [r"\bworkshop\b", r"\bwork shop\b"]),
        ("Budget_Workshop", [r"\bbudget workshop\b"]),
        ("Budget_Hearing", [r"\bbudget hearing\b"]),
        ("Budget_Work_Session", [r"\bbudget work session\b"]),
        ("Budget", [r"\bbudget\b"]),
        ("Public_Hearing", [r"\bpublic hearing\b"]),
        ("Hearing", [r"\bhearing\b"]),
        ("Organizational", [r"\borganizational\b"]),
        ("Reorganizational", [r"\breorganizational\b", r"\breorg\b"]),
    ]
    
    def __init__(self):
        """Initialize OCR handler."""
        self.reader = None
    
    def suggest_filename(self, pdf_path: str, meeting_body: str = "", 
                        file_type: str = "") -> Dict[str, any]:
        """
        Analyze PDF and suggest a filename.
        
        Args:
            pdf_path: Path to PDF file
            meeting_body: Meeting body name (from folder structure)
            file_type: File type (Agenda, Minutes, etc. from folder structure)
        
        Returns:
            Dict with keys: 'suggested_name', 'confidence', 'date', 'subtype', 'error'
        """
        if not fitz:
            return {
                'suggested_name': None,
                'confidence': None,
                'date': None,
                'subtype': None,
                'error': 'PyMuPDF not installed'
            }
        
        try:
            # Open PDF
            doc = fitz.open(pdf_path)
            
            # Extract text from first few pages
            pages_text = []
            for page_num in range(min(3, len(doc))):
                pages_text.append(doc[page_num].get_text())
            
            # Try to extract date with fallbacks
            date_result = self._extract_date_with_fallbacks(
                pages_text, 
                os.path.basename(pdf_path), 
                doc
            )
            
            doc.close()
            
            if not date_result:
                return {
                    'suggested_name': "OCR Unable to parse data",
                    'confidence': None,
                    'date': None,
                    'subtype': None,
                    'error': 'No date found'
                }
            
            extracted_date, confidence, source = date_result
            
            # Detect subtype from text
            full_text = ' '.join(pages_text)
            subtype = self._detect_subtype(full_text)
            
            # Build suggested filename
            # Format: {Meeting Body}_{Subtype}_{File Type}_{Date}.pdf
            parts = []
            
            if meeting_body:
                parts.append(meeting_body.replace(" ", "_"))
            
            if subtype:
                parts.append(subtype)
            
            if file_type:
                parts.append(file_type.replace(" ", "_"))
            
            # Format date as YYYY-MM-DD
            date_str = extracted_date.strftime("%Y-%m-%d")
            parts.append(date_str)
            
            suggested_name = "_".join(parts) + ".pdf"
            
            # Map confidence to readable format
            confidence_map = {3: "HIGH", 2: "MEDIUM", 1: "LOW"}
            confidence_str = confidence_map.get(confidence, "UNKNOWN")
            
            return {
                'suggested_name': suggested_name,
                'confidence': confidence_str,
                'date': extracted_date,
                'subtype': subtype,
                'error': None,
                'source': source
            }
            
        except Exception as e:
            return {
                'suggested_name': "OCR Unable to parse data",
                'confidence': None,
                'date': None,
                'subtype': None,
                'error': str(e)
            }
    
    def _extract_date_with_fallbacks(self, pages_text: List[str], 
                                    filename: str, doc) -> Optional[Tuple[date, int, str]]:
        """
        Try multiple methods to extract date.
        Returns: (date, confidence_level, source_description)
        """
        # Fallback 1: PDF text extraction
        result = self._extract_date_from_pdf_text(pages_text)
        if result:
            return result
        
        # Fallback 2: Filename parsing
        result = self._extract_date_from_filename(filename)
        if result:
            return result
        
        # Fallback 3: PDF metadata
        result = self._extract_date_from_metadata(doc)
        if result:
            return result
        
        # Fallback 4: OCR
        result = self._extract_date_from_ocr(doc)
        if result:
            return result
        
        return None
    
    def _extract_date_from_pdf_text(self, pages: List[str]) -> Optional[Tuple[date, int, str]]:
        """Extract date from PDF text."""
        combined = ' '.join(pages)[:self.DATE_SEARCH_LIMIT_PDF]
        result = self._parse_date_from_text(combined, allow_8digit=False)
        if result:
            return (result[0], result[1], "PDF text")
        return None
    
    def _extract_date_from_filename(self, filename: str) -> Optional[Tuple[date, int, str]]:
        """Extract date from filename."""
        result = self._parse_date_from_text(filename, allow_8digit=True)
        if result:
            return (result[0], result[1], "Filename")
        return None
    
    def _extract_date_from_metadata(self, doc) -> Optional[Tuple[date, int, str]]:
        """Extract date from PDF metadata."""
        try:
            metadata = doc.metadata
            if not metadata:
                return None
            
            title = metadata.get('title', '').lower()
            subject = metadata.get('subject', '').lower()
            
            # Check if metadata contains meeting-related keywords
            meeting_keywords = ['meeting', 'agenda', 'minutes', 'session']
            combined = title + ' ' + subject
            
            if any(kw in combined for kw in meeting_keywords):
                result = self._parse_date_from_text(combined, allow_8digit=False)
                if result:
                    return (result[0], 2, "PDF metadata")  # Medium confidence
        except:
            pass
        
        return None
    
    def _extract_date_from_ocr(self, doc) -> Optional[Tuple[date, int, str]]:
        """Extract date using OCR on first page."""
        try:
            # Lazy load EasyOCR
            if not self.reader:
                import easyocr
                self.reader = easyocr.Reader(['en'], gpu=False, verbose=False)
            
            # Get first page as image
            page = doc[0]
            pix = page.get_pixmap(dpi=150)
            img_data = pix.tobytes("png")
            
            # Run OCR
            import io
            from PIL import Image
            img = Image.open(io.BytesIO(img_data))
            result = self.reader.readtext(img, detail=0)
            
            # Combine OCR text
            ocr_text = ' '.join(result)[:self.DATE_SEARCH_LIMIT_OCR]
            
            # Parse date
            date_result = self._parse_date_from_text(ocr_text, allow_8digit=False)
            if date_result:
                return (date_result[0], 1, "OCR")  # Low confidence
        
        except Exception:
            pass
        
        return None
    
    # Date patterns from FileScanPDFRename v2.5
    # NOTE: inner alternations use (?:...) so that each pattern has exactly the
    # capture groups its parsing code expects (month, day, year = groups 1,2,3).
    MONTH_RE = r'(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)'
    
    DATE_PATTERNS = [
        # Pattern 1: Full text dates like "Wednesday, April 13, 2022" or "April 13, 2022"  
        re.compile(
            rf'(?:(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)(?:day)?)?[,\s]*({MONTH_RE})\s+([0-3]?\d)(?:st|nd|rd|th)?[,\s]+([12]\d{{3}})',
            re.I
        ),
        # Pattern 2: Numeric MM/DD/YYYY with separators (/, -, _, .)
        re.compile(r'([01]?\d)[\/\-_\.]([0-3]?\d)[\/\-_\.]((?:19|20)?\d{2,4})'),
        # Pattern 3: ISO format YYYY-MM-DD with separators (/, -, _, .)
        re.compile(r'([12]\d{3})[\/\-_\.]([01]?\d)[\/\-_\.]([0-3]?\d)'),
        # Pattern 4: Day-abbreviated month-year: 24_Mar_2020, 24-Mar-2020
        re.compile(
            rf'([0-3]?\d)[\/\-_\.]+({MONTH_RE})[\/\-_\.]+([12]\d{{3}})',
            re.I
        ),
        # Pattern 5: Month_Day_Year with separators: April_13_2022, March-24-2020
        re.compile(
            rf'({MONTH_RE})[\/\-_\.]+([0-3]?\d)[\/\-_\.]+([12]\d{{3}})',
            re.I
        ),
        # Pattern 6: 8-digit no separator MMDDYYYY (filename only, strict rules)
        re.compile(r'\b([01]\d)([0-3]\d)([12]\d{3})\b'),
    ]
    
    @staticmethod
    def _is_valid_date(year: int, month: int, day: int) -> bool:
        """Validate date is reasonable (between 1800 and 2 years from today)."""
        try:
            from datetime import timedelta
            dt = date(year, month, day)
            future_limit = date.today() + timedelta(days=730)  # Allow up to 2 years ahead
            return 1800 <= year and dt <= future_limit
        except ValueError:
            return False
    
    @staticmethod
    def _is_repeated_pattern(s: str) -> bool:
        """Check if string has repeated patterns like 01010101."""
        if len(s) < 4:
            return False
        # Check for 2-char repeats
        first_two = s[:2]
        if all(s[i:i+2] == first_two for i in range(0, len(s), 2)):
            return True
        # Check for 4-char repeats
        if len(s) >= 8:
            first_four = s[:4]
            if all(s[i:i+4] == first_four for i in range(0, len(s), 4)):
                return True
        return False
    
    def _parse_date_from_text(self, text: str, max_chars: int = 300, 
                             allow_8digit: bool = False) -> Optional[Tuple[date, int]]:
        """
        Parse date from text using multiple patterns (from FileScanPDFRename v2.5).
        Returns: (date_object, confidence_level)
        """
        text = text[:max_chars]
        
        for pat_idx, pat in enumerate(self.DATE_PATTERNS):
            # Skip 8-digit pattern unless explicitly allowed
            if pat_idx == 5 and not allow_8digit:
                continue
            
            for m in pat.finditer(text):
                try:
                    dt = None
                    
                    if pat_idx == 0:
                        # Pattern 1: "Wednesday, April 13, 2022" or "April 13, 2022"
                        month_str = m.group(1)
                        day_str = m.group(2)
                        year_str = m.group(3)
                        
                        for fmt in ["%B %d %Y", "%b %d %Y"]:
                            try:
                                dt = datetime.strptime(f"{month_str} {day_str} {year_str}", fmt).date()
                                break
                            except ValueError:
                                continue
                        
                        if dt is None or not self._is_valid_date(dt.year, dt.month, dt.day):
                            continue
                    
                    elif pat_idx == 1:
                        # Pattern 2: MM/DD/YYYY, MM-DD-YYYY, MM_DD_YYYY, MM.DD.YYYY
                        mm, dd, yy = int(m.group(1)), int(m.group(2)), int(m.group(3))
                        if yy < 100:
                            yy += 2000
                        
                        if not self._is_valid_date(yy, mm, dd):
                            continue
                        dt = date(yy, mm, dd)
                    
                    elif pat_idx == 2:
                        # Pattern 3: 2020-03-24, 2020_03_24 (ISO format YYYY-MM-DD)
                        yy, mm, dd = int(m.group(1)), int(m.group(2)), int(m.group(3))
                        
                        if not self._is_valid_date(yy, mm, dd):
                            continue
                        dt = date(yy, mm, dd)
                    
                    elif pat_idx == 3:
                        # Pattern 4: 24_Mar_2020, 24-Mar-2020
                        day_str = m.group(1)
                        month_str = m.group(2)
                        year_str = m.group(3)
                        
                        for fmt in ["%d %b %Y", "%d %B %Y"]:
                            try:
                                dt = datetime.strptime(f"{day_str} {month_str} {year_str}", fmt).date()
                                break
                            except ValueError:
                                continue
                        
                        if dt is None or not self._is_valid_date(dt.year, dt.month, dt.day):
                            continue
                    
                    elif pat_idx == 4:
                        # Pattern 5: April_13_2022, March-24-2020
                        month_str = m.group(1)
                        day_str = m.group(2)
                        year_str = m.group(3)
                        
                        for fmt in ["%B %d %Y", "%b %d %Y"]:
                            try:
                                dt = datetime.strptime(f"{month_str} {day_str} {year_str}", fmt).date()
                                break
                            except ValueError:
                                continue
                        
                        if dt is None or not self._is_valid_date(dt.year, dt.month, dt.day):
                            continue
                    
                    elif pat_idx == 5:
                        # Pattern 6: MMDDYYYY (8-digit, no separator) - STRICT RULES
                        mm_str, dd_str, yy_str = m.group(1), m.group(2), m.group(3)
                        full_match = mm_str + dd_str + yy_str
                        
                        # Check for repeated patterns
                        if self._is_repeated_pattern(full_match):
                            continue
                        
                        mm, dd, yy = int(mm_str), int(dd_str), int(yy_str)
                        
                        # No 00 month or day
                        if mm == 0 or dd == 0:
                            continue
                        
                        if not self._is_valid_date(yy, mm, dd):
                            continue
                        dt = date(yy, mm, dd)
                    
                    if dt:
                        # Return confidence level: 3=HIGH, 2=MEDIUM, 1=LOW
                        confidence = 3 if pat_idx in [0, 2, 3, 4] else 2
                        return (dt, confidence)
                
                except Exception:
                    continue
        
        return None
    
    def _detect_subtype(self, text: str) -> Optional[str]:
        """Detect meeting subtype from text."""
        text_lower = text.lower()
        
        for subtype_name, patterns in self.SUBTYPE_PATTERNS:
            for pattern in patterns:
                if re.search(pattern, text_lower):
                    return subtype_name
        
        return None
