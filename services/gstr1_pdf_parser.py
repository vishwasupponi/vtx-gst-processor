import os
import re
from typing import Dict, Any, Optional, List
import pdfplumber


class GSTR1PDFParser:
    """Parses GSTR-1 PDF documents into structured Python dictionaries."""

    MONTH_ORDER = [
        "April", "May", "June", "July", "August", "September",
        "October", "November", "December", "January", "February", "March"
    ]

    MONTH_MAP = {
        "04": "April", "05": "May", "06": "June", "07": "July",
        "08": "August", "09": "September", "10": "October", "11": "November",
        "12": "December", "01": "January", "02": "February", "03": "March",
        "APR": "April", "MAY": "May", "JUN": "June", "JUL": "July",
        "AUG": "August", "SEP": "September", "OCT": "October", "NOV": "November",
        "DEC": "December", "JAN": "January", "FEB": "February", "MAR": "March",
        "APRIL": "April", "JUNE": "June", "JULY": "July", "AUGUST": "August",
        "SEPTEMBER": "September", "OCTOBER": "October", "NOVEMBER": "November",
        "DECEMBER": "December", "JANUARY": "January", "FEBRUARY": "February", "MARCH": "March"
    }

    def __init__(self, log_callback=None):
        self.log_callback = log_callback

    def log(self, msg: str):
        if self.log_callback:
            self.log_callback(msg)

    def _parse_num(self, val_str: Any) -> float:
        """Safely parses currency/number string into float."""
        if not val_str or val_str in ["-", "NA", "N/A", ""]:
            return 0.0
        try:
            cleaned = str(val_str).replace(",", "").replace("₹", "").replace(" ", "").strip()
            return float(cleaned)
        except Exception:
            return 0.0

    def parse_pdf(self, pdf_path: str) -> Dict[str, Any]:
        """Extracts complete metadata and all 17 tables from a GSTR-1 PDF."""
        result: Dict[str, Any] = {
            "source_file": os.path.basename(pdf_path),
            "gstin": "",
            "legal_name": "",
            "trade_name": "",
            "financial_year": "",
            "tax_period_raw": "",
            "month_name": "",
            "month_index": -1,
            "arn": "",
            "arn_date": "",
            "tables": {}
        }

        if not os.path.exists(pdf_path):
            return result

        full_text = ""
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                t = page.extract_text() or ""
                full_text += "\n" + t

        lines = [l.strip() for l in full_text.split("\n") if l.strip()]

        # ─── 1. Header Metadata Extraction ────────────────────────────────────
        for idx, line in enumerate(lines[:35]):
            line_u = line.upper()
            
            # GSTIN
            m_gstin = re.search(r'GSTIN\s*[:\-]?\s*([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})', line_u)
            if m_gstin and not result["gstin"]:
                result["gstin"] = m_gstin.group(1)

            # Financial Year
            m_fy = re.search(r'FINANCIAL\s*YEAR\s*[:\-]?\s*(\d{4}\s*-\s*\d{2,4})', line_u)
            if m_fy and not result["financial_year"]:
                result["financial_year"] = m_fy.group(1).replace(" ", "")

            # Tax Period Month
            m_tp = re.search(r'TAX\s*PERIOD\s*[:\-]?\s*([A-Z]+)', line_u)
            if m_tp and not result["month_name"]:
                raw_month = m_tp.group(1).strip()
                result["tax_period_raw"] = raw_month
                norm_m = self.MONTH_MAP.get(raw_month.upper(), "")
                if norm_m:
                    result["month_name"] = norm_m
                    result["month_index"] = self.MONTH_ORDER.index(norm_m) if norm_m in self.MONTH_ORDER else -1

            # Legal Name
            if "LEGAL NAME OF THE REGISTERED PERSON" in line_u:
                name_val = line.split("person", 1)[-1].replace(":", "").strip()
                if name_val:
                    result["legal_name"] = name_val
                elif idx + 1 < len(lines):
                    result["legal_name"] = lines[idx + 1].strip()

            # Trade Name
            if "TRADE NAME IF ANY" in line_u:
                trade_val = line.split("any", 1)[-1].replace(":", "").strip()
                if trade_val:
                    result["trade_name"] = trade_val
                elif idx + 1 < len(lines):
                    result["trade_name"] = lines[idx + 1].strip()

            # ARN
            m_arn = re.search(r'\bARN\s*[:\-]?\s*([A-Z0-9]{15,16})\b', line_u)
            if m_arn:
                result["arn"] = m_arn.group(1)

            # ARN Date
            m_arnd = re.search(r'ARN\s*DATE\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})', line_u)
            if m_arnd:
                result["arn_date"] = m_arnd.group(1)

        # Fallback for month from filename if missing from text
        if not result["month_name"]:
            fname = os.path.basename(pdf_path)
            m_fn_period = re.search(r'_(\d{2})(\d{4})\.', fname)
            if m_fn_period:
                mm = m_fn_period.group(1)
                norm_m = self.MONTH_MAP.get(mm, "")
                if norm_m:
                    result["month_name"] = norm_m
                    result["month_index"] = self.MONTH_ORDER.index(norm_m) if norm_m in self.MONTH_ORDER else -1

        # ─── 2. Number Parsing Helper ─────────────────────────────────────────
        def parse_data_line(line_text: str) -> List[float]:
            if any(k in line_text.lower() for k in [
                'ip address', 'records type', 'unregistered type', 'signature',
                'no. of document', 'financial year', 'tax period', 'gstin',
                'legal name', 'trade name', 'arn '
            ]):
                return []
            
            clean = line_text.replace(',', '').replace('₹', '').strip()
            
            # Match Document Type keyword: [optional doc count] [optional watermark] Keyword [tail numbers]
            kw_pattern = r'([A-Za-z]*\d+)?\s*[A-Za-z]*(?:Invoice|Note|Net\s*Value|Document(?:/Net\s*Value)?|All\s*Documents|NA)\s+(.*)'
            m_kw = re.search(kw_pattern, clean, re.IGNORECASE)
            if m_kw:
                doc_raw = m_kw.group(1) or '0'
                doc_nums = re.findall(r'\d+', doc_raw)
                doc_count = float(doc_nums[-1]) if doc_nums else 0.0
                tail_str = m_kw.group(2)
                tail_nums = [float(x) for x in re.findall(r'[-+]?\d+(?:\.\d+)?', tail_str)]
                return [doc_count] + tail_nums
            
            # Generic pure numbers on line - only for non-header lines
            if re.match(r'^(?:[0-9]{1,2}[A-Z]?(?:\([0-9A-Za-z]+\))?\s*[-–]|Credit / Debit notes issued|Amended Credit / Debit|Verifica\s*tion)', clean, re.IGNORECASE):
                return []
            
            nums = [float(x) for x in re.findall(r'[-+]?\d+(?:\.\d+)?', clean)]
            return nums

        def get_sec(start_pat: str, end_pat: str) -> List[str]:
            sec = []
            rec = False
            for l in lines:
                if re.search(start_pat, l, re.IGNORECASE):
                    rec = True
                    sec.append(l)
                    continue
                if rec:
                    if end_pat and re.search(end_pat, l, re.IGNORECASE):
                        break
                    sec.append(l)
            return sec

        def find_line(sec_lines: List[str], pattern: str, max_lookahead: int = 5) -> List[float]:
            for idx, l in enumerate(sec_lines):
                if re.search(pattern, l, re.IGNORECASE):
                    # Try matching line itself if it has keyword or is a sub-item like - Nil
                    if any(k in l for k in ['Invoice', 'Note', 'Net Value', 'Document', 'All Documents', 'NA']) or re.match(r'^-\s*(?:Nil|Exempted|Non-GST)', l, re.IGNORECASE):
                        same = parse_data_line(l)
                        if same:
                            return same
                    # Look ahead
                    for offset in range(1, max_lookahead + 1):
                        if idx + offset < len(sec_lines):
                            nxt = sec_lines[idx + offset]
                            nums = parse_data_line(nxt)
                            if nums:
                                return nums
            return []

        def find_after_anchor(sec_lines: List[str], anchor_pat: str, target_pat: str, max_lookahead: int = 6) -> List[float]:
            for idx, l in enumerate(sec_lines):
                if re.search(anchor_pat, l, re.IGNORECASE):
                    for offset in range(1, max_lookahead + 1):
                        if idx + offset < len(sec_lines):
                            nxt = sec_lines[idx + offset]
                            if re.search(target_pat, nxt, re.IGNORECASE):
                                nums = parse_data_line(nxt)
                                if nums:
                                    return nums
            return []

        # ─── 3. Table-by-Table Extraction (Strictly Section-Bounded) ──────────
        t = result["tables"]

        # Table 4: B2B Regular & Reverse Charge
        sec4 = get_sec(r'4A\s*-\s*Taxable outward', r'5\s*-\s*Taxable outward')
        t4a = find_line(sec4, r'4A\s*-\s*Taxable outward')
        t4b = find_line(sec4, r'4B\s*-\s*Taxable outward')
        t["4A_B2B_Regular"] = {
            "records": int(t4a[0]) if len(t4a) > 0 else 0,
            "value": t4a[1] if len(t4a) > 1 else 0.0,
            "igst": t4a[2] if len(t4a) > 2 else 0.0,
            "cgst": t4a[3] if len(t4a) > 3 else 0.0,
            "sgst": t4a[4] if len(t4a) > 4 else 0.0,
            "cess": t4a[5] if len(t4a) > 5 else 0.0
        }
        t["4B_B2B_ReverseCharge"] = {
            "records": int(t4b[0]) if len(t4b) > 0 else 0,
            "value": t4b[1] if len(t4b) > 1 else 0.0,
            "igst": t4b[2] if len(t4b) > 2 else 0.0,
            "cgst": t4b[3] if len(t4b) > 3 else 0.0,
            "sgst": t4b[4] if len(t4b) > 4 else 0.0,
            "cess": t4b[5] if len(t4b) > 5 else 0.0
        }
        t["4_Summary"] = {
            "records": t["4A_B2B_Regular"]["records"] + t["4B_B2B_ReverseCharge"]["records"],
            "value": t["4A_B2B_Regular"]["value"] + t["4B_B2B_ReverseCharge"]["value"],
            "igst": t["4A_B2B_Regular"]["igst"] + t["4B_B2B_ReverseCharge"]["igst"],
            "cgst": t["4A_B2B_Regular"]["cgst"] + t["4B_B2B_ReverseCharge"]["cgst"],
            "sgst": t["4A_B2B_Regular"]["sgst"] + t["4B_B2B_ReverseCharge"]["sgst"],
            "cess": t["4A_B2B_Regular"]["cess"] + t["4B_B2B_ReverseCharge"]["cess"]
        }

        # Table 5: B2CL (Large)
        sec5 = get_sec(r'5\s*-\s*Taxable outward', r'6A\s*[-–]\s*Exports')
        t5 = find_line(sec5, r'^Total')
        if not t5:
            t5 = find_line(sec5, r'5\s*-\s*Taxable outward')
        t["5_B2CL_Large"] = {
            "records": int(t5[0]) if len(t5) > 0 else 0,
            "value": t5[1] if len(t5) > 1 else 0.0,
            "igst": t5[2] if len(t5) > 2 else 0.0,
            "cess": t5[3] if len(t5) > 3 else 0.0
        }

        # Table 6A: Exports (EXPWP / EXPWOP)
        sec6a = get_sec(r'6A\s*[-–]\s*Exports', r'6B\s*-\s*Supplies made to SEZ')
        t6a_wp = find_line(sec6a, r'-\s*EXPWP')
        t6a_wop = find_line(sec6a, r'-\s*EXPWOP')
        t["6A_Exports"] = {
            "expwp_records": int(t6a_wp[0]) if len(t6a_wp) > 0 else 0,
            "expwp_value": t6a_wp[1] if len(t6a_wp) > 1 else 0.0,
            "expwp_igst": t6a_wp[2] if len(t6a_wp) > 2 else 0.0,
            "expwp_cess": t6a_wp[3] if len(t6a_wp) > 3 else 0.0,
            "expwop_records": int(t6a_wop[0]) if len(t6a_wop) > 0 else 0,
            "expwop_value": t6a_wop[1] if len(t6a_wop) > 1 else 0.0
        }

        # Table 6B: SEZ (SEZWP / SEZWOP)
        sec6b = get_sec(r'6B\s*-\s*Supplies made to SEZ', r'6C\s*-\s*Deemed Exports')
        t6b_wp = find_line(sec6b, r'-\s*SEZWP')
        t6b_wop = find_line(sec6b, r'-\s*SEZWOP')
        t["6B_SEZ"] = {
            "sezwp_records": int(t6b_wp[0]) if len(t6b_wp) > 0 else 0,
            "sezwp_value": t6b_wp[1] if len(t6b_wp) > 1 else 0.0,
            "sezwp_igst": t6b_wp[2] if len(t6b_wp) > 2 else 0.0,
            "sezwp_cess": t6b_wp[3] if len(t6b_wp) > 3 else 0.0,
            "sezwop_records": int(t6b_wop[0]) if len(t6b_wop) > 0 else 0,
            "sezwop_value": t6b_wop[1] if len(t6b_wop) > 1 else 0.0
        }

        # Table 6C: Deemed Exports (DE)
        sec6c = get_sec(r'6C\s*-\s*Deemed Exports', r'7-\s*Taxable supplies')
        t6c = find_line(sec6c, r'6C\s*-\s*Deemed Exports')
        t["6C_Deemed_Exports"] = {
            "records": int(t6c[0]) if len(t6c) > 0 else 0,
            "value": t6c[1] if len(t6c) > 1 else 0.0,
            "igst": t6c[2] if len(t6c) > 2 else 0.0,
            "cgst": t6c[3] if len(t6c) > 3 else 0.0,
            "sgst": t6c[4] if len(t6c) > 4 else 0.0,
            "cess": t6c[5] if len(t6c) > 5 else 0.0
        }

        # Table 7: B2CS Small
        sec7 = get_sec(r'7-\s*Taxable supplies', r'8\s*-\s*Nil rated')
        t7 = find_line(sec7, r'7-\s*Taxable supplies')
        t["7_B2CS_Small"] = {
            "records": int(t7[0]) if len(t7) > 0 else 0,
            "value": t7[1] if len(t7) > 1 else 0.0,
            "igst": t7[2] if len(t7) > 2 else 0.0,
            "cgst": t7[3] if len(t7) > 3 else 0.0,
            "sgst": t7[4] if len(t7) > 4 else 0.0,
            "cess": t7[5] if len(t7) > 5 else 0.0
        }

        # Table 8: Nil, Exempted, Non-GST
        sec8 = get_sec(r'8\s*-\s*Nil rated', r'9A\s*-\s*Amendment')
        t8_tot = find_line(sec8, r'^Total')
        t8_nil = find_line(sec8, r'^-\s*Nil')
        t8_ex = find_line(sec8, r'^-\s*Exempted')
        t8_ngst = find_line(sec8, r'^-\s*Non-GST')
        t["8_Nil_Exempted"] = {
            "total": t8_tot[0] if t8_tot else 0.0,
            "nil": t8_nil[0] if t8_nil else 0.0,
            "exempted": t8_ex[0] if t8_ex else 0.0,
            "non_gst": t8_ngst[0] if t8_ngst else 0.0
        }

        # Table 9A: Amendments
        sec9a = get_sec(r'9A\s*-\s*Amendment to taxable outward supplies', r'9B\s*-\s*Credit/Debit')
        t["9A_Amendments"] = {
            "b2b_reg_amended": find_line(sec9a, r'table 4\s*-\s*B2B Regular'),
            "b2b_reg_diff": find_after_anchor(sec9a, r'table 4\s*-\s*B2B Regular', r'Net differential'),
            "b2b_rc_amended": find_line(sec9a, r'table 4\s*-\s*B2B Reverse charge'),
            "b2b_rc_diff": find_after_anchor(sec9a, r'table 4\s*-\s*B2B Reverse charge', r'Net differential'),
            "b2cl_amended": find_line(sec9a, r'table 5\s*-\s*B2CL'),
            "b2cl_diff": find_after_anchor(sec9a, r'table 5\s*-\s*B2CL', r'Net differential'),
            "expwp_amended": find_line(sec9a, r'-\s*EXPWP'),
            "expwp_diff": find_after_anchor(sec9a, r'-\s*EXPWP', r'Net differential'),
            "expwop_amended": find_line(sec9a, r'-\s*EXPWOP'),
            "expwop_diff": find_after_anchor(sec9a, r'-\s*EXPWOP', r'Net differential'),
            "sezwp_amended": find_line(sec9a, r'-\s*SEZWP'),
            "sezwp_diff": find_after_anchor(sec9a, r'-\s*SEZWP', r'Net differential'),
            "sezwop_amended": find_line(sec9a, r'-\s*SEZWOP'),
            "sezwop_diff": find_after_anchor(sec9a, r'-\s*SEZWOP', r'Net differential'),
            "de_amended": find_line(sec9a, r'table 6C'),
            "de_diff": find_after_anchor(sec9a, r'table 6C', r'Net differential'),
        }

        # Table 9B: Credit/Debit Notes Registered & Unregistered (CDNR & CDNUR)
        sec9b = get_sec(r'9B\s*-\s*Credit/Debit Notes \(Registered\)', r'9C\s*-\s*Amended Credit/Debit')
        t["9B_CDNR"] = {
            "b2b_reg": find_after_anchor(sec9b, r'B2B Regular', r'Net Total'),
            "b2b_rc": find_after_anchor(sec9b, r'B2B Reverse charge', r'Net Total'),
            "sez": find_after_anchor(sec9b, r'SEZWP/SEZWOP', r'Net Total'),
            "de": find_after_anchor(sec9b, r'table 6C', r'Net Total'),
            "cdnur_b2cl": find_line(sec9b, r'-\s*B2CL'),
            "cdnur_expwp": find_line(sec9b, r'-\s*EXPWP'),
            "cdnur_expwop": find_line(sec9b, r'-\s*EXPWOP'),
        }

        # Table 9C: Amended Credit/Debit Notes Registered & Unregistered (CDNRA & CDNURA)
        sec9c = get_sec(r'9C\s*-\s*Amended Credit/Debit Notes \(Registered\)', r'10\s*-\s*Amendment to taxable')
        t["9C_CDNRA"] = {
            "b2b_reg": find_after_anchor(sec9c, r'B2B Regular', r'Net total'),
            "b2b_rc": find_after_anchor(sec9c, r'B2B Reverse charge', r'Net total'),
            "sez": find_after_anchor(sec9c, r'SEZWP/SEZWOP', r'Net total'),
            "de": find_after_anchor(sec9c, r'table 6C', r'Net total'),
            "cdnr_total": find_line(sec9c, r'Amended amount - Total'),
            "cdnura_b2cl": find_line(sec9c, r'-\s*B2CL'),
            "cdnura_expwp": find_line(sec9c, r'-\s*EXPWP'),
            "cdnura_expwop": find_line(sec9c, r'-\s*EXPWOP'),
            "cdnura_total": find_after_anchor(sec9c, r'9C\s*-\s*Amended Credit/Debit Notes \(Unregistered\)', r'Amended amount - Total'),
        }

        # Table 10: Amendment to B2C (Others)
        sec10 = get_sec(r'10\s*-\s*Amendment to taxable', r'11A\(1\)')
        t10_am = find_line(sec10, r'Amended amount')
        t10_diff = find_line(sec10, r'Net differential')
        t["10_Amended_B2C"] = {
            "amended_records": int(t10_am[0]) if len(t10_am) > 0 else 0,
            "amended_value": t10_am[1] if len(t10_am) > 1 else 0.0,
            "amended_igst": t10_am[2] if len(t10_am) > 2 else 0.0,
            "amended_cgst": t10_am[3] if len(t10_am) > 3 else 0.0,
            "amended_sgst": t10_am[4] if len(t10_am) > 4 else 0.0,
            "amended_cess": t10_am[5] if len(t10_am) > 5 else 0.0,
            "net_diff_value": t10_diff[0] if len(t10_diff) > 0 else 0.0,
            "net_diff_igst": t10_diff[1] if len(t10_diff) > 1 else 0.0,
            "net_diff_cgst": t10_diff[2] if len(t10_diff) > 2 else 0.0,
            "net_diff_sgst": t10_diff[3] if len(t10_diff) > 3 else 0.0,
            "net_diff_cess": t10_diff[4] if len(t10_diff) > 4 else 0.0
        }

        # Table 11: Advances Received & Adjusted
        sec11 = get_sec(r'11A\(1\),\s*11A\(2\)', r'12\s*-\s*HSN')
        t["11_Advances"] = {
            "rec_records": find_after_anchor(sec11, r'11A\(1\)', r'^Total'),
            "adj_records": find_after_anchor(sec11, r'11B\(1\)', r'^Total'),
            "11a_amend_total": find_after_anchor(sec11, r'11A\s*-\s*Amendment to advances received', r'^Total'),
            "11a_amend_amended": find_after_anchor(sec11, r'11A\s*-\s*Amendment to advances received', r'Amended amount'),
            "11b_amend_total": find_after_anchor(sec11, r'11B\s*-\s*Amendment to advances adjusted', r'^Total'),
            "11b_amend_amended": find_after_anchor(sec11, r'11B\s*-\s*Amendment to advances adjusted', r'Amended amount'),
        }

        # Table 12: HSN-wise summary
        sec12 = get_sec(r'12\s*-\s*HSN', r'13\s*-\s*Documents issued')
        t12_tot = find_line(sec12, r'^Total')
        t12_b2b = find_line(sec12, r'B2B Total')
        t12_b2c = find_line(sec12, r'B2C Total')
        t["12_HSN_Summary"] = {
            "total_records": int(t12_tot[0]) if len(t12_tot) > 0 else 0,
            "total_value": t12_tot[1] if len(t12_tot) > 1 else 0.0,
            "total_igst": t12_tot[2] if len(t12_tot) > 2 else 0.0,
            "total_cgst": t12_tot[3] if len(t12_tot) > 3 else 0.0,
            "total_sgst": t12_tot[4] if len(t12_tot) > 4 else 0.0,
            "total_cess": t12_tot[5] if len(t12_tot) > 5 else 0.0,
            "b2b_records": int(t12_b2b[0]) if len(t12_b2b) > 0 else 0,
            "b2b_value": t12_b2b[1] if len(t12_b2b) > 1 else 0.0,
            "b2b_igst": t12_b2b[2] if len(t12_b2b) > 2 else 0.0,
            "b2b_cgst": t12_b2b[3] if len(t12_b2b) > 3 else 0.0,
            "b2b_sgst": t12_b2b[4] if len(t12_b2b) > 4 else 0.0,
            "b2b_cess": t12_b2b[5] if len(t12_b2b) > 5 else 0.0,
            "b2c_records": int(t12_b2c[0]) if len(t12_b2c) > 0 else 0,
            "b2c_value": t12_b2c[1] if len(t12_b2c) > 1 else 0.0,
            "b2c_igst": t12_b2c[2] if len(t12_b2c) > 2 else 0.0,
            "b2c_cgst": t12_b2c[3] if len(t12_b2c) > 3 else 0.0,
            "b2c_sgst": t12_b2c[4] if len(t12_b2c) > 4 else 0.0,
            "b2c_cess": t12_b2c[5] if len(t12_b2c) > 5 else 0.0
        }

        # Table 13: Documents issued
        sec13 = get_sec(r'13\s*-\s*Documents issued', r'14\s*-\s*Supplies made through E-Commerce')
        t13 = find_line(sec13, r'Net issued documents')
        t["13_Documents_Issued"] = {
            "net_issued_documents": int(t13[0]) if t13 else 0
        }

        # Table 14: E-Commerce
        sec14 = get_sec(r'14\s*-\s*Supplies made through E-Commerce Operators', r'14A\s*-\s*Amended Supplies made')
        t14_tot = find_line(sec14, r'^Total')
        t14_52 = find_line(sec14, r'\(a\)\s*Liable to collect')
        t14_95 = find_line(sec14, r'\(b\)\s*Liable to pay')
        t["14_ECommerce"] = {
            "total_records": int(t14_tot[0]) if len(t14_tot) > 0 else 0,
            "total_value": t14_tot[1] if len(t14_tot) > 1 else 0.0,
            "total_igst": t14_tot[2] if len(t14_tot) > 2 else 0.0,
            "total_cgst": t14_tot[3] if len(t14_tot) > 3 else 0.0,
            "total_sgst": t14_tot[4] if len(t14_tot) > 4 else 0.0,
            "total_cess": t14_tot[5] if len(t14_tot) > 5 else 0.0,
            "sec52_records": int(t14_52[0]) if len(t14_52) > 0 else 0,
            "sec52_value": t14_52[1] if len(t14_52) > 1 else 0.0,
            "sec52_igst": t14_52[2] if len(t14_52) > 2 else 0.0,
            "sec52_cgst": t14_52[3] if len(t14_52) > 3 else 0.0,
            "sec52_sgst": t14_52[4] if len(t14_52) > 4 else 0.0,
            "sec52_cess": t14_52[5] if len(t14_52) > 5 else 0.0,
            "sec9_5_records": int(t14_95[0]) if len(t14_95) > 0 else 0,
            "sec9_5_value": t14_95[1] if len(t14_95) > 1 else 0.0,
            "sec9_5_igst": t14_95[2] if len(t14_95) > 2 else 0.0,
            "sec9_5_cgst": t14_95[3] if len(t14_95) > 3 else 0.0,
            "sec9_5_sgst": t14_95[4] if len(t14_95) > 4 else 0.0,
            "sec9_5_cess": t14_95[5] if len(t14_95) > 5 else 0.0
        }

        # Table 14A: Amended E-Commerce
        sec14a = get_sec(r'14A\s*-\s*Amended Supplies made through E-Commerce', r'15\s*-\s*Supplies U/s 9\(5\)')
        t["14A_Amended_ECommerce"] = {
            "sec52_amended": find_after_anchor(sec14a, r'\(a\)\s*Liable to collect', r'Amended amount'),
            "sec52_diff": find_after_anchor(sec14a, r'\(a\)\s*Liable to collect', r'Net differential'),
            "sec9_5_amended": find_after_anchor(sec14a, r'\(b\)\s*Liable to pay', r'Amended amount'),
            "sec9_5_diff": find_after_anchor(sec14a, r'\(b\)\s*Liable to pay', r'Net differential'),
        }

        # Table 15: Supplies U/s 9(5)
        sec15 = get_sec(r'15\s*-\s*Supplies U/s 9\(5\)', r'15A\s*\(I\)')
        t["15_Supplies_9_5"] = {
            "reg_total": find_line(sec15, r'- For Registered Recipients'),
            "reg_regular": find_line(sec15, r'-\s*Regular'),
            "reg_de": find_line(sec15, r'-\s*DE'),
            "reg_sezwp": find_line(sec15, r'-\s*SEZWP'),
            "reg_sezwop": find_line(sec15, r'-\s*SEZWOP'),
            "unregistered": find_line(sec15, r'- For Unregistered Recipient'),
        }

        # Table 15A(I): Amended Supplies U/s 9(5) - Registered
        sec15a1 = get_sec(r'15A\s*\(I\)', r'15A\s*\(II\)')
        t["15A_I_Amended_9_5"] = {
            "regular": find_line(sec15a1, r'-\s*Regular'),
            "de": find_line(sec15a1, r'-\s*DE'),
            "sezwp": find_line(sec15a1, r'-\s*SEZWP'),
            "sezwop": find_line(sec15a1, r'-\s*SEZWOP'),
        }

        # Table 15A(II): Amended Supplies U/s 9(5) - Unregistered
        sec15a2 = get_sec(r'15A\s*\(II\)', r'Total Liability')
        t["15A_II_Amended_9_5"] = {
            "amended_total": find_after_anchor(sec15a2, r'15A\s*\(II\)', r'Amended amount'),
            "diff_total": find_after_anchor(sec15a2, r'15A\s*\(II\)', r'Net differential'),
        }

        # Total Liability Outward Supplies (Summary)
        sec_tot = get_sec(r'Total Liability\s*\(Outward supplies other than Reverse charge\)', r'Verifica\s*tion')
        tot_nums = find_line(sec_tot, r'Total Liability')
        t["Total_Liability_Outward"] = {
            "value": tot_nums[0] if len(tot_nums) > 0 else 0.0,
            "igst": tot_nums[1] if len(tot_nums) > 1 else 0.0,
            "cgst": tot_nums[2] if len(tot_nums) > 2 else 0.0,
            "sgst": tot_nums[3] if len(tot_nums) > 3 else 0.0,
            "cess": tot_nums[4] if len(tot_nums) > 4 else 0.0
        }

        return result

