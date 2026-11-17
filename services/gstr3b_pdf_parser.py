import os
import re
from typing import Dict, Any, List, Optional
import pdfplumber


class GSTR3BTable(dict):
    """Table data supporting both dict and attribute access."""
    def __getattr__(self, name):
        if name == "rows":
            return list(self.values())
        return self.get(name)


class GSTR3BData(dict):
    """Parsed GSTR-3B data supporting both dict and attribute access."""
    def __getattr__(self, name):
        if name == "return_period":
            return self.get("month_name")
        return self.get(name)


class GSTR3BPDFParser:
    """Extracts structured data from GSTR-3B PDF return files."""

    MONTH_ORDER = [
        "April", "May", "June", "July", "August", "September",
        "October", "November", "December", "January", "February", "March"
    ]

    MONTH_MAP = {
        "APRIL": "April", "MAY": "May", "JUNE": "June", "JULY": "July",
        "AUGUST": "August", "SEPTEMBER": "September", "OCTOBER": "October",
        "NOVEMBER": "November", "DECEMBER": "December", "JANUARY": "January",
        "FEBRUARY": "February", "MARCH": "March",
        "APR": "April", "JUN": "June", "JUL": "July", "AUG": "August",
        "SEP": "September", "SEPT": "September", "OCTOBER": "October",
        "NOV": "November", "DEC": "December", "JAN": "January",
        "FEB": "February", "MAR": "March",
        "01": "January", "02": "February", "03": "March", "04": "April",
        "05": "May", "06": "June", "07": "July", "08": "August",
        "09": "September", "10": "October", "11": "November", "12": "December"
    }

    def __init__(self, log_callback=None):
        self.log_callback = log_callback

    def log(self, msg: str):
        if self.log_callback:
            self.log_callback(msg)

    def can_parse(self, file_path: str) -> bool:
        """Checks if the file is a valid GSTR-3B PDF."""
        if not file_path.lower().endswith(".pdf") or not os.path.exists(file_path):
            return False
        try:
            with pdfplumber.open(file_path) as pdf:
                if pdf.pages:
                    text = pdf.pages[0].extract_text() or ""
                    return "GSTR-3B" in text.upper() or "FORM GSTR-3B" in text.upper()
        except Exception:
            pass
        return False

    def parse(self, pdf_path: str) -> GSTR3BData:
        """Alias for parse_pdf."""
        return self.parse_pdf(pdf_path)

    def parse_pdf(self, pdf_path: str) -> GSTR3BData:
        """Parses a single GSTR-3B PDF and returns a structured dictionary."""
        result = GSTR3BData({
            "file_path": pdf_path,
            "filename": os.path.basename(pdf_path),
            "gstin": "",
            "legal_name": "",
            "trade_name": "",
            "financial_year": "",
            "month_name": "",
            "month_index": -1,
            "arn": "",
            "arn_date": "",
            "tables": {}
        })

        if not os.path.exists(pdf_path):
            self.log(f"[GSTR-3B Parser Error] File not found: {pdf_path}")
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
            m_gstin = re.search(r'GSTIN(?:\s*OF\s*THE\s*SUPPLIER)?\s*[:\-]?\s*([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})', line_u)
            if m_gstin and not result["gstin"]:
                result["gstin"] = m_gstin.group(1)

            # Financial Year
            m_fy = re.search(r'YEAR\s*[:\-]?\s*(\d{4}\s*-\s*\d{2,4})', line_u)
            if m_fy and not result["financial_year"]:
                result["financial_year"] = m_fy.group(1).replace(" ", "")

            # Tax Period Month
            m_tp = re.search(r'PERIOD\s*[:\-]?\s*([A-Z]+)', line_u)
            if m_tp and not result["month_name"]:
                raw_m = m_tp.group(1).strip()
                norm_m = self.MONTH_MAP.get(raw_m.upper(), "")
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
            if "TRADE NAME, IF ANY" in line_u or "TRADE NAME IF ANY" in line_u:
                trade_val = line.split("any", 1)[-1].replace(":", "").strip()
                if trade_val:
                    result["trade_name"] = trade_val
                elif idx + 1 < len(lines):
                    result["trade_name"] = lines[idx + 1].strip()

            # ARN
            m_arn = re.search(r'\bARN\s*[:\-]?\s*([A-Z0-9]{15,16})\b', line_u)
            if m_arn and not result["arn"]:
                result["arn"] = m_arn.group(1)

            # Date of ARN
            m_arnd = re.search(r'DATE\s*OF\s*ARN\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})', line_u)
            if m_arnd and not result["arn_date"]:
                result["arn_date"] = m_arnd.group(1)

        # Fallback for month from filename (e.g. GSTR3B_..._042026.pdf)
        if not result["month_name"]:
            fname = os.path.basename(pdf_path)
            m_fn = re.search(r'_(\d{2})(\d{4})\.', fname)
            if m_fn:
                norm_m = self.MONTH_MAP.get(m_fn.group(1), "")
                if norm_m:
                    result["month_name"] = norm_m
                    result["month_index"] = self.MONTH_ORDER.index(norm_m) if norm_m in self.MONTH_ORDER else -1

           # ─── 2. Number Parsing Helpers ────────────────────────────────────────
        def find_section_numbers(pattern: str, expected_count: int = 4, max_lookahead: int = 3) -> List[float]:
            """
            Finds the table row matching pattern, removes section headers & item prefixes
            like '(1)', '(2)', '(3)', '(4)', '(5)', '(i)', '(ii)', 'u/s 9(5)', '1 & 2 above',
            and returns the trailing expected_count numbers.
            """
            for i, l in enumerate(lines):
                if re.search(pattern, l, re.IGNORECASE):
                    for off in range(0, max_lookahead + 1):
                        if i + off < len(lines):
                            cand = lines[i + off]
                            if off > 0 and re.match(r'^(?:3\.[12]|4\.|5\.|6\.|Breakup|Verification)', cand.strip(), re.IGNORECASE):
                                break
                            clean_line = re.sub(r'^\s*(?:\([0-9a-zA-Z]+\)|\d+\([0-9a-zA-Z]+\)|[0-9a-zA-Z]+\.)\s*', '', cand)
                            clean_line = re.sub(r'u/s\s*9\s*\(\s*5\s*\)|section\s*9\s*\(\s*5\s*\)', '', clean_line, flags=re.IGNORECASE)
                            clean_line = re.sub(r'other\s*than\s*1\s*&\s*2|Table\s*4\s*\([A-Z]\)\s*\(\d+\)', '', clean_line, flags=re.IGNORECASE)
                            
                            words = clean_line.split()
                            nums = []
                            for w in words:
                                m = re.search(r'[-+]?\d+(?:,\d+)*(?:\.\d+)?', w)
                                if m:
                                    val_str = m.group(0).replace(',', '').strip()
                                    try:
                                        nums.append(float(val_str))
                                    except Exception:
                                        pass
                            if len(nums) >= expected_count:
                                return nums[-expected_count:]
                            elif ('0.00' in cand or '-' in cand) and len(nums) > 0:
                                padded = [0.0] * (expected_count - len(nums)) + nums
                                return padded[-expected_count:]
            return [0.0] * expected_count

        # ─── 3. Table-by-Table Extraction ─────────────────────────────────────
        t = result["tables"]

        # ── Table 3.1: Outward & Inward Reverse Charge ──
        # 3.1(a) Outward Taxable Supplies (non-zero/nil/exempted)
        t["3.1_a"] = {"val": 0.0, "igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
        n_31a = find_section_numbers(r'\(a\)\s*Outward taxable supplies\s*\(other than zero', expected_count=5)
        t["3.1_a"]["val"] = n_31a[0]
        t["3.1_a"]["igst"] = n_31a[1]
        t["3.1_a"]["cgst"] = n_31a[2]
        t["3.1_a"]["sgst"] = n_31a[3]
        t["3.1_a"]["cess"] = n_31a[4]

        # 3.1(b) Zero Rated
        t["3.1_b"] = {"val": 0.0, "igst": 0.0, "cess": 0.0}
        n_31b = find_section_numbers(r'\(b\)\s*Outward taxable supplies\s*\(zero rated\)', expected_count=3)
        t["3.1_b"]["val"] = n_31b[0]
        t["3.1_b"]["igst"] = n_31b[1]
        t["3.1_b"]["cess"] = n_31b[2]

        # 3.1(c) Other outward supplies (nil rated, exempted)
        t["3.1_c"] = {"val": 0.0}
        n_31c = find_section_numbers(r'\(c\s*\)\s*Other outward supplies\s*\(nil rated', expected_count=1)
        t["3.1_c"]["val"] = n_31c[0]

        # 3.1(d) Inward supplies (liable to reverse charge)
        t["3.1_d"] = {"val": 0.0, "igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
        n_31d = find_section_numbers(r'\(d\)\s*Inward supplies\s*\(liable to reverse charge\)', expected_count=5)
        t["3.1_d"]["val"] = n_31d[0]
        t["3.1_d"]["igst"] = n_31d[1]
        t["3.1_d"]["cgst"] = n_31d[2]
        t["3.1_d"]["sgst"] = n_31d[3]
        t["3.1_d"]["cess"] = n_31d[4]

        # 3.1(e) Non-GST outward supplies
        t["3.1_e"] = {"val": 0.0}
        n_31e = find_section_numbers(r'\(e\)\s*Non-GST outward supplies', expected_count=1)
        t["3.1_e"]["val"] = n_31e[0]

        # ── Table 3.1.1: Section 9(5) E-Commerce ──
        t["3.1.1_i"] = {"val": 0.0, "igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
        n_311i = find_section_numbers(r'\(i\)\s*Taxable supplies on which electronic commerce operator pays tax', expected_count=5)
        t["3.1.1_i"]["val"] = n_311i[0]
        t["3.1.1_i"]["igst"] = n_311i[1]
        t["3.1.1_i"]["cgst"] = n_311i[2]
        t["3.1.1_i"]["sgst"] = n_311i[3]
        t["3.1.1_i"]["cess"] = n_311i[4]

        t["3.1.1_ii"] = {"val": 0.0}
        n_311ii = find_section_numbers(r'\(ii\)\s*Taxable supplies made by registered person through electronic commerce', expected_count=1)
        t["3.1.1_ii"]["val"] = n_311ii[0]

        # ── Table 3.2: Inter-state supplies ──
        t["3.2_unregistered"] = {"val": 0.0, "igst": 0.0}
        n_32u = find_section_numbers(r'Supplies made to Unregistered Persons', expected_count=2)
        t["3.2_unregistered"]["val"] = n_32u[0]
        t["3.2_unregistered"]["igst"] = n_32u[1]

        t["3.2_composition"] = {"val": 0.0, "igst": 0.0}
        n_32c = find_section_numbers(r'Supplies made to Composition Taxable', expected_count=2)
        t["3.2_composition"]["val"] = n_32c[0]
        t["3.2_composition"]["igst"] = n_32c[1]

        t["3.2_uin"] = {"val": 0.0, "igst": 0.0}
        n_32uin = find_section_numbers(r'Supplies made to UIN holders', expected_count=2)
        t["3.2_uin"]["val"] = n_32uin[0]
        t["3.2_uin"]["igst"] = n_32uin[1]

        # ── Table 4: Eligible ITC ──
        # A. ITC Available (1 to 5)
        t["4_A_1_goods"] = {"igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
        n_4a1 = find_section_numbers(r'\(1\)\s*Import of goods', expected_count=4)
        t["4_A_1_goods"]["igst"] = n_4a1[0]
        t["4_A_1_goods"]["cgst"] = n_4a1[1]
        t["4_A_1_goods"]["sgst"] = n_4a1[2]
        t["4_A_1_goods"]["cess"] = n_4a1[3]

        t["4_A_2_services"] = {"igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
        n_4a2 = find_section_numbers(r'\(2\)\s*Import of services', expected_count=4)
        t["4_A_2_services"]["igst"] = n_4a2[0]
        t["4_A_2_services"]["cgst"] = n_4a2[1]
        t["4_A_2_services"]["sgst"] = n_4a2[2]
        t["4_A_2_services"]["cess"] = n_4a2[3]

        t["4_A_3_reverse_charge"] = {"igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
        n_4a3 = find_section_numbers(r'\(3\)\s*Inward supplies liable to reverse charge', expected_count=4)
        t["4_A_3_reverse_charge"]["igst"] = n_4a3[0]
        t["4_A_3_reverse_charge"]["cgst"] = n_4a3[1]
        t["4_A_3_reverse_charge"]["sgst"] = n_4a3[2]
        t["4_A_3_reverse_charge"]["cess"] = n_4a3[3]

        t["4_A_4_isd"] = {"igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
        n_4a4 = find_section_numbers(r'\(4\)\s*Inward supplies from ISD', expected_count=4)
        t["4_A_4_isd"]["igst"] = n_4a4[0]
        t["4_A_4_isd"]["cgst"] = n_4a4[1]
        t["4_A_4_isd"]["sgst"] = n_4a4[2]
        t["4_A_4_isd"]["cess"] = n_4a4[3]

        t["4_A_5_all_other"] = {"igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
        n_4a5 = find_section_numbers(r'\(5\)\s*All other ITC', expected_count=4)
        t["4_A_5_all_other"]["igst"] = n_4a5[0]
        t["4_A_5_all_other"]["cgst"] = n_4a5[1]
        t["4_A_5_all_other"]["sgst"] = n_4a5[2]
        t["4_A_5_all_other"]["cess"] = n_4a5[3]

        # B. ITC Reversed
        t["4_B_1_rules"] = {"igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
        n_4b1 = find_section_numbers(r'\(1\)\s*As per rules 38,42\s*&\s*43', expected_count=4)
        t["4_B_1_rules"]["igst"] = n_4b1[0]
        t["4_B_1_rules"]["cgst"] = n_4b1[1]
        t["4_B_1_rules"]["sgst"] = n_4b1[2]
        t["4_B_1_rules"]["cess"] = n_4b1[3]

        t["4_B_2_others"] = {"igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
        n_4b2 = find_section_numbers(r'\(2\)\s*Others', expected_count=4)
        t["4_B_2_others"]["igst"] = n_4b2[0]
        t["4_B_2_others"]["cgst"] = n_4b2[1]
        t["4_B_2_others"]["sgst"] = n_4b2[2]
        t["4_B_2_others"]["cess"] = n_4b2[3]

        # C. Net ITC Available (A - B)
        t["4_C_net"] = {"igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
        n_4c = find_section_numbers(r'C\.\s*Net ITC available\s*\(A-B\)', expected_count=4)
        t["4_C_net"]["igst"] = n_4c[0]
        t["4_C_net"]["cgst"] = n_4c[1]
        t["4_C_net"]["sgst"] = n_4c[2]
        t["4_C_net"]["cess"] = n_4c[3]

        # D. Other Details
        t["4_D_1_reclaimed"] = {"igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
        n_4d1 = find_section_numbers(r'\(1\)\s*ITC reclaimed which was reversed', expected_count=4)
        t["4_D_1_reclaimed"]["igst"] = n_4d1[0]
        t["4_D_1_reclaimed"]["cgst"] = n_4d1[1]
        t["4_D_1_reclaimed"]["sgst"] = n_4d1[2]
        t["4_D_1_reclaimed"]["cess"] = n_4d1[3]

        t["4_D_2_ineligible"] = {"igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
        n_4d2 = find_section_numbers(r'\(2\)\s*Ineligible ITC under section 16\(4\)', expected_count=4)
        t["4_D_2_ineligible"]["igst"] = n_4d2[0]
        t["4_D_2_ineligible"]["cgst"] = n_4d2[1]
        t["4_D_2_ineligible"]["sgst"] = n_4d2[2]
        t["4_D_2_ineligible"]["cess"] = n_4d2[3]

        # ── Table 5: Values of exempt, nil-rated & non-GST inward ──
        t["5_exempt_nil"] = {"inter": 0.0, "intra": 0.0}
        n_5en = find_section_numbers(r'From a supplier under composition scheme, Exempt, Nil rated supply', expected_count=2)
        t["5_exempt_nil"]["inter"] = n_5en[0]
        t["5_exempt_nil"]["intra"] = n_5en[1]

        t["5_non_gst"] = {"inter": 0.0, "intra": 0.0}
        n_5ng = find_section_numbers(r'Non GST supply', expected_count=2)
        t["5_non_gst"]["inter"] = n_5ng[0]
        t["5_non_gst"]["intra"] = n_5ng[1]

        # ── Table 5.1: Interest & Late Fee ──
        t["5.1_interest_paid"] = {"igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
        n_51i = find_section_numbers(r'Interest Paid', expected_count=4)
        t["5.1_interest_paid"]["igst"] = n_51i[0]
        t["5.1_interest_paid"]["cgst"] = n_51i[1]
        t["5.1_interest_paid"]["sgst"] = n_51i[2]
        t["5.1_interest_paid"]["cess"] = n_51i[3]

        t["5.1_late_fee"] = {"cgst": 0.0, "sgst": 0.0}
        n_51lf = find_section_numbers(r'Late fee', expected_count=2)
        t["5.1_late_fee"]["cgst"] = n_51lf[0]
        t["5.1_late_fee"]["sgst"] = n_51lf[1]

        # ── Table 6.1: Payment of Tax (Full Breakdown: AJ to AS) ──
        t["6.1_payment"] = self._parse_table_6_1(lines)

        # ── Breakup of Tax Liability ──
        t["breakup_tax_liability"] = {"igst": 0.0, "cgst": 0.0, "sgst": 0.0, "cess": 0.0}
        n_break = find_section_numbers(r'Breakup of tax liability declared', expected_count=4)
        t["breakup_tax_liability"]["igst"] = n_break[0]
        t["breakup_tax_liability"]["cgst"] = n_break[1]
        t["breakup_tax_liability"]["sgst"] = n_break[2]
        t["breakup_tax_liability"]["cess"] = n_break[3]

        # Convenience table grouping for tests
        t["3.1_Outward_Supplies"] = GSTR3BTable({
            "a": t.get("3.1_a", {}),
            "b": t.get("3.1_b", {}),
            "c": t.get("3.1_c", {}),
            "d": t.get("3.1_d", {}),
            "e": t.get("3.1_e", {})
        })
        t["4_Eligible_ITC"] = GSTR3BTable({
            "a_1": t.get("4_A_1_goods", {}),
            "a_5": t.get("4_A_5_all_other", {}),
            "b_1": t.get("4_B_1_rules", {}),
            "c_net": t.get("4_C_net", {})
        })

        return result

    def _parse_table_6_1(self, lines: List[str]) -> Dict[str, Any]:
        """Extracts complete Payment of Tax table (AJ to AS) with Section A and B breakdown."""
        res = {
            "A_integrated": {"tax_payable": 0.0, "adj_negative": 0.0, "net_payable": 0.0, "itc_igst": 0.0, "itc_cgst": 0.0, "itc_sgst": 0.0, "itc_cess": 0.0, "cash_paid": 0.0, "interest": 0.0, "late_fee": 0.0},
            "A_central": {"tax_payable": 0.0, "adj_negative": 0.0, "net_payable": 0.0, "itc_igst": 0.0, "itc_cgst": 0.0, "itc_sgst": 0.0, "itc_cess": 0.0, "cash_paid": 0.0, "interest": 0.0, "late_fee": 0.0},
            "A_state": {"tax_payable": 0.0, "adj_negative": 0.0, "net_payable": 0.0, "itc_igst": 0.0, "itc_cgst": 0.0, "itc_sgst": 0.0, "itc_cess": 0.0, "cash_paid": 0.0, "interest": 0.0, "late_fee": 0.0},
            "A_cess": {"tax_payable": 0.0, "adj_negative": 0.0, "net_payable": 0.0, "itc_igst": 0.0, "itc_cgst": 0.0, "itc_sgst": 0.0, "itc_cess": 0.0, "cash_paid": 0.0, "interest": 0.0, "late_fee": 0.0},
            "B_integrated": {"tax_payable": 0.0, "adj_negative": 0.0, "net_payable": 0.0, "itc_igst": 0.0, "itc_cgst": 0.0, "itc_sgst": 0.0, "itc_cess": 0.0, "cash_paid": 0.0, "interest": 0.0, "late_fee": 0.0},
            "B_central": {"tax_payable": 0.0, "adj_negative": 0.0, "net_payable": 0.0, "itc_igst": 0.0, "itc_cgst": 0.0, "itc_sgst": 0.0, "itc_cess": 0.0, "cash_paid": 0.0, "interest": 0.0, "late_fee": 0.0},
            "B_state": {"tax_payable": 0.0, "adj_negative": 0.0, "net_payable": 0.0, "itc_igst": 0.0, "itc_cgst": 0.0, "itc_sgst": 0.0, "itc_cess": 0.0, "cash_paid": 0.0, "interest": 0.0, "late_fee": 0.0},
            "B_cess": {"tax_payable": 0.0, "adj_negative": 0.0, "net_payable": 0.0, "itc_igst": 0.0, "itc_cgst": 0.0, "itc_sgst": 0.0, "itc_cess": 0.0, "cash_paid": 0.0, "interest": 0.0, "late_fee": 0.0}
        }
        in_section_a = False
        in_section_b = False

        for line in lines:
            line_u = line.upper()
            if "(A) OTHER THAN REVERSE CHARGE" in line_u:
                in_section_a = True
                in_section_b = False
                continue
            elif "(B) REVERSE CHARGE" in line_u:
                in_section_a = False
                in_section_b = True
                continue
            elif "BREAKUP OF TAX LIABILITY" in line_u or "VERIFICATION" in line_u:
                in_section_a = False
                in_section_b = False

            tokens = re.findall(r'[-+]?(?:\d+(?:,\d+)*(?:\.\d+)?|\.\d+|-)', line)
            nums = []
            for t in tokens:
                if t == "-":
                    nums.append(0.0)
                else:
                    c = t.replace(",", "").strip()
                    try:
                        nums.append(float(c))
                    except Exception:
                        pass

            if in_section_a:
                if "INTEGRATED" in line_u and len(nums) >= 3:
                    res["A_integrated"]["tax_payable"] = nums[0]
                    res["A_integrated"]["adj_negative"] = nums[1] if len(nums) > 1 else 0.0
                    res["A_integrated"]["net_payable"] = nums[2] if len(nums) > 2 else 0.0
                    res["A_integrated"]["itc_igst"] = nums[3] if len(nums) > 3 else 0.0
                    res["A_integrated"]["itc_cgst"] = nums[4] if len(nums) > 4 else 0.0
                    res["A_integrated"]["itc_sgst"] = nums[5] if len(nums) > 5 else 0.0
                    res["A_integrated"]["itc_cess"] = nums[6] if len(nums) > 6 else 0.0
                    res["A_integrated"]["cash_paid"] = nums[7] if len(nums) > 7 else 0.0
                    res["A_integrated"]["interest"] = nums[8] if len(nums) > 8 else 0.0
                    res["A_integrated"]["late_fee"] = nums[9] if len(nums) > 9 else 0.0
                elif "CENTRAL" in line_u and len(nums) >= 3:
                    res["A_central"]["tax_payable"] = nums[0]
                    res["A_central"]["adj_negative"] = nums[1] if len(nums) > 1 else 0.0
                    res["A_central"]["net_payable"] = nums[2] if len(nums) > 2 else 0.0
                    res["A_central"]["itc_igst"] = nums[3] if len(nums) > 3 else 0.0
                    res["A_central"]["itc_cgst"] = nums[4] if len(nums) > 4 else 0.0
                    res["A_central"]["itc_sgst"] = nums[5] if len(nums) > 5 else 0.0
                    res["A_central"]["itc_cess"] = nums[6] if len(nums) > 6 else 0.0
                    res["A_central"]["cash_paid"] = nums[7] if len(nums) > 7 else 0.0
                    res["A_central"]["interest"] = nums[8] if len(nums) > 8 else 0.0
                    res["A_central"]["late_fee"] = nums[9] if len(nums) > 9 else 0.0
                elif "STATE/UT" in line_u and len(nums) >= 3:
                    res["A_state"]["tax_payable"] = nums[0]
                    res["A_state"]["adj_negative"] = nums[1] if len(nums) > 1 else 0.0
                    res["A_state"]["net_payable"] = nums[2] if len(nums) > 2 else 0.0
                    res["A_state"]["itc_igst"] = nums[3] if len(nums) > 3 else 0.0
                    res["A_state"]["itc_cgst"] = nums[4] if len(nums) > 4 else 0.0
                    res["A_state"]["itc_sgst"] = nums[5] if len(nums) > 5 else 0.0
                    res["A_state"]["itc_cess"] = nums[6] if len(nums) > 6 else 0.0
                    res["A_state"]["cash_paid"] = nums[7] if len(nums) > 7 else 0.0
                    res["A_state"]["interest"] = nums[8] if len(nums) > 8 else 0.0
                    res["A_state"]["late_fee"] = nums[9] if len(nums) > 9 else 0.0
                elif "CESS" in line_u and len(nums) >= 3:
                    res["A_cess"]["tax_payable"] = nums[0]
                    res["A_cess"]["adj_negative"] = nums[1] if len(nums) > 1 else 0.0
                    res["A_cess"]["net_payable"] = nums[2] if len(nums) > 2 else 0.0
                    res["A_cess"]["cash_paid"] = nums[-3] if len(nums) >= 6 else (nums[3] if len(nums) > 3 else 0.0)

            elif in_section_b:
                if "INTEGRATED" in line_u and len(nums) >= 3:
                    res["B_integrated"]["tax_payable"] = nums[0]
                    res["B_integrated"]["cash_paid"] = nums[-1] if nums else 0.0
                elif "CENTRAL" in line_u and len(nums) >= 3:
                    res["B_central"]["tax_payable"] = nums[0]
                    res["B_central"]["cash_paid"] = nums[-1] if nums else 0.0
                elif "STATE/UT" in line_u and len(nums) >= 3:
                    res["B_state"]["tax_payable"] = nums[0]
                    res["B_state"]["cash_paid"] = nums[-1] if nums else 0.0
                elif "CESS" in line_u and len(nums) >= 3:
                    res["B_cess"]["tax_payable"] = nums[0]
                    res["B_cess"]["cash_paid"] = nums[-1] if nums else 0.0

        return res
