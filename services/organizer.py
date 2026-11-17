import os
import re
import json
import shutil
from dataclasses import dataclass, field
from typing import List, Dict, Tuple, Optional, Callable
from openpyxl import load_workbook
import pdfplumber


@dataclass
class ClassifiedReturnFile:
    original_path: str
    filename: str
    detected_form: str             # 'GSTR-1', 'GSTR-2A', 'GSTR-2B', 'GSTR-3B', 'UNIDENTIFIED'
    confidence: str                # 'EXACT_PORTAL_PATTERN', 'FUZZY_KEYWORD', 'DEEP_CONTENT_INSPECTION', 'LOW'
    gstin: str = ""
    period: str = ""
    destination_subfolder: str = ""
    error_or_warning: str = ""


@dataclass
class OrganizationResult:
    total_files_scanned: int = 0
    copied_counts: Dict[str, int] = field(default_factory=lambda: {
        "GSTR-1": 0, "GSTR-2A": 0, "GSTR-2B": 0, "GSTR-3B": 0, "UNIDENTIFIED": 0
    })
    unidentified_files: List[str] = field(default_factory=list)
    destination_root: str = ""
    file_mappings: List[Tuple[str, str]] = field(default_factory=list)


class GSTFileOrganizer:
    """Classifies and safely organizes GST return files into categorized folders."""

    def __init__(self, log_callback: Optional[Callable[[str], None]] = None):
        self.log_callback = log_callback

    def log(self, message: str):
        if self.log_callback:
            self.log_callback(message)

    def classify_file(self, file_path: str, deep_inspection: bool = True) -> Tuple[str, str, str, str]:
        """
        Classifies a file by filename first, and falls back to deep content inspection if enabled.
        Returns: (Form_Name, Confidence, GSTIN, Period)
        """
        filename = os.path.basename(file_path)
        form_name, conf, gstin, period = self.classify_filename(filename)

        # If name is unidentified or ambiguous, and deep inspection is enabled, inspect file content
        if (form_name == "UNIDENTIFIED" or conf == "LOW") and deep_inspection and os.path.exists(file_path):
            deep_form, deep_conf, deep_gstin, deep_period = self.inspect_file_content(file_path)
            if deep_form != "UNIDENTIFIED":
                return deep_form, deep_conf, deep_gstin or gstin, deep_period or period
            elif deep_conf != "LOW":
                return form_name, deep_conf, deep_gstin or gstin, deep_period or period

        return form_name, conf, gstin, period

    def classify_filename(self, filename: str) -> Tuple[str, str, str, str]:
        """
        Classifies a filename based on portal naming conventions and fuzzy keywords.
        """
        name_no_ext, _ = os.path.splitext(filename)
        name_upper = name_no_ext.upper()
        clean_name = name_upper.replace("-", "_").replace(" ", "_")

        # 1. Regex for standard portal formats
        # GSTR-1: GSTR1_<GSTIN>_<MMYYYY>
        m_r1 = re.search(r'\bGSTR1_([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})_([0-9]{6})(?:_|\b|$)', clean_name)
        if m_r1:
            return "GSTR-1", "EXACT_PORTAL_PATTERN", m_r1.group(1), m_r1.group(2)

        # GSTR-3B: GSTR3B_<GSTIN>_<MMYYYY>
        m_3b = re.search(r'\bGSTR3B_([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})_([0-9]{6})(?:_|\b|$)', clean_name)
        if m_3b:
            return "GSTR-3B", "EXACT_PORTAL_PATTERN", m_3b.group(1), m_3b.group(2)

        # GSTR-2B: <MMYYYY>_<GSTIN>_GSTR2B_<DDMMYYYY> or <MMYYYY>_<GSTIN>_GSTR2B
        m_2b = re.search(r'\b([0-9]{6})_([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})_GSTR2B(?:_([0-9]{8}))?(?:_|\b|$)', clean_name)
        if m_2b:
            return "GSTR-2B", "EXACT_PORTAL_PATTERN", m_2b.group(2), m_2b.group(1)

        # GSTR-2A: <GSTIN>_<MMYYYY>_R2A
        m_2a = re.search(r'\b([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})_([0-9]{6})_R2A(?:_|\b|$)', clean_name)
        if m_2a:
            return "GSTR-2A", "EXACT_PORTAL_PATTERN", m_2a.group(1), m_2a.group(2)

        # 2. Extract general GSTIN and Period if present in filename
        gstin_m = re.search(r'\b[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}\b', clean_name)
        gstin = gstin_m.group(0) if gstin_m else ""

        period_m = re.search(r'\b(0[1-9]|1[0-2])(20[2-9][0-9])\b', clean_name)
        period = period_m.group(0) if period_m else ""

        # 3. Fuzzy & Symbol-Tolerant Pattern Matching (GSTR 3B, GSTR-3B, gstr3b, 3b, GSTR3, R3B, etc.)
        # Check GSTR-3B first (handles GSTR3B, GSTR-3B, GSTR 3B, GSTR3, GSTR_3, R3B, 3B, 3-B, 3_B)
        if any(k in clean_name for k in ["GSTR3B", "GSTR_3B", "GSTR3", "GSTR_3", "R3B", "R_3B", "3B", "3_B"]) or "3B" in name_upper or "GSTR 3" in name_upper:
            return "GSTR-3B", "FUZZY_KEYWORD", gstin, period

        # Check GSTR-2B (handles GSTR2B, GSTR-2B, GSTR 2B, R2B, 2B, 2-B, 2_B)
        if any(k in clean_name for k in ["GSTR2B", "GSTR_2B", "R2B", "R_2B", "2B", "2_B"]) or " 2B" in name_upper or "GSTR 2B" in name_upper:
            return "GSTR-2B", "FUZZY_KEYWORD", gstin, period

        # Check GSTR-2A (handles GSTR2A, GSTR-2A, GSTR 2A, R2A, R-2A, 2A, 2-A, 2_A)
        if any(k in clean_name for k in ["GSTR2A", "GSTR_2A", "R2A", "R_2A", "2A", "2_A"]) or " 2A" in name_upper or "GSTR 2A" in name_upper:
            return "GSTR-2A", "FUZZY_KEYWORD", gstin, period

        # Check GSTR-1 (handles GSTR1, GSTR-1, GSTR 1, GSTR_1, R1, R-1, 1)
        if any(k in clean_name for k in ["GSTR1", "GSTR_1", "R1", "R_1", "1_B", "1_A"]) or "GSTR 1" in name_upper or "GSTR-1" in name_upper or " R1" in name_upper:
            return "GSTR-1", "FUZZY_KEYWORD", gstin, period

        # Single tokens with separators
        tokens = re.split(r'[^A-Z0-9]', clean_name)
        if "3B" in tokens or "GSTR3" in tokens or "R3B" in tokens:
            return "GSTR-3B", "FUZZY_KEYWORD", gstin, period
        if "2B" in tokens or "GSTR2B" in tokens or "R2B" in tokens:
            return "GSTR-2B", "FUZZY_KEYWORD", gstin, period
        if "2A" in tokens or "GSTR2A" in tokens or "R2A" in tokens:
            return "GSTR-2A", "FUZZY_KEYWORD", gstin, period
        if "1" in tokens and any(k in tokens for k in ["GSTR", "RETURN", "SALES", "R1"]):
            return "GSTR-1", "FUZZY_KEYWORD", gstin, period

        return "UNIDENTIFIED", "LOW", gstin, period

    def inspect_file_content(self, file_path: str) -> Tuple[str, str, str, str]:
        """
        Deep Content Sniffer: Reads the interior content of PDFs, Excel workbooks, or JSONs
        to accurately classify returns even if the filename is completely altered.
        """
        ext = os.path.splitext(file_path)[1].lower()

        # ─── A. PDF Content Inspection ─────────────────────────────────────────
        if ext == ".pdf":
            try:
                with pdfplumber.open(file_path) as pdf:
                    full_text = ""
                    for p in pdf.pages[:3]:
                        t = p.extract_text() or ""
                        full_text += " " + t

                    text_upper = full_text.upper()

                    # Extract GSTIN if present in text
                    gstin_m = re.search(r'GSTIN\s*[:\-]?\s*([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})', text_upper)
                    gstin = gstin_m.group(1) if gstin_m else ""

                    # Extract Period
                    period_m = re.search(r'(?:TAX\s*PERIOD|PERIOD|MONTH)\s*[:\-]?\s*([A-Z]+|\d{2}\s*[-_]\s*\d{4})', text_upper)
                    period = period_m.group(1).strip() if period_m else ""

                    # 1. Search specifically for GSTR-3B terms (e.g. "FORM GSTR-3B", "GSTR-3B", "GSTR 3B", "GSTR3B", " 3B ")
                    if any(term in text_upper for term in ["FORM GSTR-3B", "FORM GSTR 3B", "GSTR-3B", "GSTR 3B", "GSTR3B", "TABLE 3.1", "TABLE 4. ELIGIBLE"]):
                        return "GSTR-3B", "DEEP_CONTENT_INSPECTION", gstin, period
                    if re.search(r'\b(?:FORM\s*)?GSTR[-\s]?3B\b|\b3B\b', text_upper):
                        return "GSTR-3B", "DEEP_CONTENT_INSPECTION", gstin, period

                    # 2. Search specifically for GSTR-1 terms (e.g. "FORM GSTR-1", "GSTR-1", "GSTR 1", "GSTR1", " R1 ")
                    if any(term in text_upper for term in ["FORM GSTR-1", "FORM GSTR 1", "GSTR-1", "GSTR 1", "GSTR1"]):
                        return "GSTR-1", "DEEP_CONTENT_INSPECTION", gstin, period
                    if re.search(r'\b(?:FORM\s*)?GSTR[-\s]?1\b|\bR1\b', text_upper):
                        return "GSTR-1", "DEEP_CONTENT_INSPECTION", gstin, period

                    # 3. Search specifically for GSTR-2B terms
                    if any(term in text_upper for term in ["FORM GSTR-2B", "FORM GSTR 2B", "GSTR-2B", "GSTR 2B", "GSTR2B"]):
                        return "GSTR-2B", "DEEP_CONTENT_INSPECTION", gstin, period
                    if re.search(r'\b(?:FORM\s*)?GSTR[-\s]?2B\b|\b2B\b', text_upper):
                        return "GSTR-2B", "DEEP_CONTENT_INSPECTION", gstin, period

                    # 4. Search specifically for GSTR-2A terms
                    if any(term in text_upper for term in ["FORM GSTR-2A", "FORM GSTR 2A", "GSTR-2A", "GSTR 2A", "GSTR2A", "R2A"]):
                        return "GSTR-2A", "DEEP_CONTENT_INSPECTION", gstin, period
                    if re.search(r'\b(?:FORM\s*)?GSTR[-\s]?2A\b|\b2A\b|\bR2A\b', text_upper):
                        return "GSTR-2A", "DEEP_CONTENT_INSPECTION", gstin, period

            except Exception as e:
                self.log(f"[Deep Inspection Notice] Could not read PDF {os.path.basename(file_path)}: {e}")

        # ─── B. Excel Content Inspection ──────────────────────────────────────
        elif ext in [".xlsx", ".xls"]:
            try:
                wb = load_workbook(file_path, read_only=True, data_only=True)
                sheet_names = [s.upper().strip() for s in wb.sheetnames]
                all_sheets_str = " ".join(sheet_names)
                
                # Check specific form names in sheets
                if any(k in sheet_names for k in ["GSTR-3B", "GSTR3B", "GSTR 3B", "3B", "R3B"]):
                    wb.close()
                    return "GSTR-3B", "DEEP_CONTENT_INSPECTION", "", ""

                if any(k in sheet_names for k in ["GSTR-2B", "GSTR2B", "GSTR 2B", "2B", "R2B", "ITC AVAILABLE"]):
                    wb.close()
                    return "GSTR-2B", "DEEP_CONTENT_INSPECTION", "", ""

                if any(k in sheet_names for k in ["GSTR-2A", "GSTR2A", "GSTR 2A", "2A", "R2A", "B2BA"]):
                    wb.close()
                    return "GSTR-2A", "DEEP_CONTENT_INSPECTION", "", ""

                if any(k in sheet_names for k in ["GSTR-1", "GSTR1", "GSTR 1", "R1", "1"]):
                    wb.close()
                    return "GSTR-1", "DEEP_CONTENT_INSPECTION", "", ""

                wb.close()
            except Exception as e:
                self.log(f"[Deep Inspection Notice] Could not read Excel {os.path.basename(file_path)}: {e}")

        # ─── C. JSON Content Inspection ───────────────────────────────────────
        elif ext == ".json":
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    payload = json.load(f)

                if isinstance(payload, dict):
                    gstin = payload.get("gstin", "")
                    period = payload.get("fp", "")

                    form_val = str(payload.get("form", "") or payload.get("form_name", "") or "").upper()
                    if "3B" in form_val:
                        return "GSTR-3B", "DEEP_CONTENT_INSPECTION", gstin, period
                    if "2B" in form_val:
                        return "GSTR-2B", "DEEP_CONTENT_INSPECTION", gstin, period
                    if "2A" in form_val or "R2A" in form_val:
                        return "GSTR-2A", "DEEP_CONTENT_INSPECTION", gstin, period
                    if "1" in form_val:
                        return "GSTR-1", "DEEP_CONTENT_INSPECTION", gstin, period

                    keys = [k.lower() for k in payload.keys()]
                    if any(k in keys for k in ["sup_details", "itc_elg"]):
                        return "GSTR-3B", "DEEP_CONTENT_INSPECTION", gstin, period
                    if any(k in keys for k in ["itc_avl", "itc_unavl"]):
                        return "GSTR-2B", "DEEP_CONTENT_INSPECTION", gstin, period
                    if any(k in keys for k in ["b2b", "b2cl", "b2cs", "hsn", "exp"]):
                        return "GSTR-1", "DEEP_CONTENT_INSPECTION", gstin, period
            except Exception:
                pass

        # ─── D. ZIP Archive Content Inspection (Only R2A zips are sorted to GSTR-2A, bundled zips go to Unidentified) ─
        elif ext == ".zip":
            try:
                import zipfile
                with zipfile.ZipFile(file_path, 'r') as zf:
                    internal_names = [n for n in zf.namelist() if not n.startswith("__MACOSX") and not n.endswith("/")]
                    internal_str = " ".join(internal_names).upper()

                    # 1. Check if the zip contains multiple different return forms (Bundled ZIP)
                    forms_found = set()
                    if any(k in internal_str for k in ["GSTR3B", " 3B", "_3B", "-3B"]):
                        forms_found.add("GSTR-3B")
                    if any(k in internal_str for k in ["GSTR1", " 1", "_1", "-1", "R1"]):
                        forms_found.add("GSTR-1")
                    if any(k in internal_str for k in ["GSTR2B", " 2B", "_2B", "-2B"]):
                        forms_found.add("GSTR-2B")
                    if any(k in internal_str for k in ["GSTR2A", "R2A", " 2A", "_2A", "-2A"]):
                        forms_found.add("GSTR-2A")

                    # If it's a bundled zip containing multiple forms or non-2A forms, put in Unidentified
                    if len(forms_found) > 1 or ("GSTR-2A" not in forms_found and len(forms_found) > 0):
                        return "UNIDENTIFIED", "BUNDLED_ZIP_ARCHIVE", "", ""

                    # 2. If it is an original GSTR-2A (R2A) zip download
                    if "GSTR-2A" in forms_found or "R2A" in internal_str or "GSTR2A" in internal_str:
                        return "GSTR-2A", "PORTAL_R2A_ZIP", "", ""

            except Exception as e:
                self.log(f"[Deep Inspection Notice] Could not read ZIP {os.path.basename(file_path)}: {e}")

        return "UNIDENTIFIED", "LOW", "", ""

    def preview_source_folder(
        self,
        source_folder: str,
        deep_inspection: bool = True
    ) -> List[ClassifiedReturnFile]:
        """Scans source folder (non-destructive) and returns classified preview (1 file on disk = 1 item)."""
        classified_list: List[ClassifiedReturnFile] = []
        if not os.path.exists(source_folder):
            return classified_list

        for item in sorted(os.listdir(source_folder)):
            full_path = os.path.join(source_folder, item)
            if os.path.isfile(full_path):
                ext = os.path.splitext(item)[1].lower()
                if ext in [".pdf", ".xlsx", ".xls", ".json", ".zip", ".csv"]:
                    # Standard file classification (single file on disk = 1 item)
                    form_name, conf, gstin, period = self.classify_file(full_path, deep_inspection=deep_inspection)
                    
                    subfolder = ""
                    warning = ""
                    if form_name != "UNIDENTIFIED":
                        subfolder = form_name  # 'GSTR-1', 'GSTR-2A', 'GSTR-2B', 'GSTR-3B'
                    else:
                        subfolder = "Unidentified_Returns"
                        if conf == "BUNDLED_ZIP_ARCHIVE":
                            warning = "Bundled multi-return ZIP archive -> placed in 'Unidentified_Returns'."
                        else:
                            warning = f"Could not determine return type from filename or content: '{item}'"

                    classified_list.append(ClassifiedReturnFile(
                        original_path=full_path,
                        filename=item,
                        detected_form=form_name,
                        confidence=conf,
                        gstin=gstin,
                        period=period,
                        destination_subfolder=subfolder,
                        error_or_warning=warning
                    ))

        return classified_list

    def preview_files(
        self,
        file_paths: List[str],
        deep_inspection: bool = True
    ) -> List[ClassifiedReturnFile]:
        """Scans and classifies an explicit list of file paths."""
        classified_list: List[ClassifiedReturnFile] = []
        for full_path in file_paths:
            if os.path.isfile(full_path):
                item = os.path.basename(full_path)
                ext = os.path.splitext(item)[1].lower()
                if ext in [".pdf", ".xlsx", ".xls", ".json", ".zip", ".csv"]:
                    form_name, conf, gstin, period = self.classify_file(full_path, deep_inspection=deep_inspection)
                    
                    subfolder = ""
                    warning = ""
                    if form_name != "UNIDENTIFIED":
                        subfolder = form_name
                    else:
                        subfolder = "Unidentified_Returns"
                        if conf == "BUNDLED_ZIP_ARCHIVE":
                            warning = "Bundled multi-return ZIP archive -> placed in 'Unidentified_Returns'."
                        else:
                            warning = f"Could not determine return type from filename or content: '{item}'"

                    classified_list.append(ClassifiedReturnFile(
                        original_path=full_path,
                        filename=item,
                        detected_form=form_name,
                        confidence=conf,
                        gstin=gstin,
                        period=period,
                        destination_subfolder=subfolder,
                        error_or_warning=warning
                    ))
        return classified_list

    def organize_files(
        self,
        source_folder: str,
        destination_folder: str,
        deep_inspection: bool = True,
        progress_callback: Optional[Callable[[int, int, str], None]] = None
    ) -> OrganizationResult:
        """
        Safely organizes classified files into categorized subfolders.
        R2A ZIPs go to GSTR-2A/, bundled ZIPs go to Unidentified_Returns/.
        NEVER alters, renames, or deletes original files.
        """
        preview = self.preview_source_folder(source_folder, deep_inspection=deep_inspection)
        return self._execute_copy_organization(preview, destination_folder, deep_inspection, progress_callback)

    def organize_specific_files(
        self,
        file_paths: List[str],
        destination_folder: str,
        deep_inspection: bool = True,
        progress_callback: Optional[Callable[[int, int, str], None]] = None
    ) -> OrganizationResult:
        """
        Safely organizes a specific list of return files into categorized subfolders.
        """
        preview = self.preview_files(file_paths, deep_inspection=deep_inspection)
        return self._execute_copy_organization(preview, destination_folder, deep_inspection, progress_callback)

    def _execute_copy_organization(
        self,
        preview: List[ClassifiedReturnFile],
        destination_folder: str,
        deep_inspection: bool,
        progress_callback: Optional[Callable[[int, int, str], None]] = None
    ) -> OrganizationResult:
        """Internal helper to safely copy classified return items into their destination folders."""
        result = OrganizationResult(destination_root=destination_folder)
        result.total_files_scanned = len(preview)

        if not preview:
            self.log("[Organizer] No return files found or provided to organize.")
            return result

        os.makedirs(destination_folder, exist_ok=True)
        self.log(f"[Organizer] Starting organization of {len(preview)} return files (Deep Inspection: {'ON' if deep_inspection else 'OFF'})...")

        for idx, item in enumerate(preview, start=1):
            target_sub = os.path.join(destination_folder, item.destination_subfolder)
            os.makedirs(target_sub, exist_ok=True)

            dest_file_path = os.path.join(target_sub, item.filename)
            try:
                shutil.copy2(item.original_path, dest_file_path)
                result.file_mappings.append((item.original_path, dest_file_path))
                
                if item.detected_form in result.copied_counts:
                    result.copied_counts[item.detected_form] += 1
                else:
                    result.copied_counts["UNIDENTIFIED"] += 1

                if item.detected_form == "UNIDENTIFIED":
                    result.unidentified_files.append(item.filename)
                    self.log(f"   [Unsorted] {item.filename} -> Copied to '{item.destination_subfolder}' ({item.error_or_warning or 'Unidentified'}).")
                else:
                    self.log(f"   [Sorted ({item.confidence})] {item.filename} -> {item.destination_subfolder}/")

            except Exception as e:
                self.log(f"   [Error copying {item.filename}] {e}")

            if progress_callback:
                progress_callback(idx, len(preview), item.filename)

        self.log("[Organizer] File organization completed successfully.")
        return result

