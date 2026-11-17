import os
import sys
import re
from typing import List, Dict, Any, Optional
import openpyxl
from openpyxl.cell.cell import MergedCell

try:
    from services.gstr1_pdf_parser import GSTR1PDFParser
    from services.template_asset import resolve_template_path, get_embedded_template_stream
except ImportError:
    from gstr1_pdf_parser import GSTR1PDFParser
    from template_asset import resolve_template_path, get_embedded_template_stream


class GSTR1ExcelConsolidator:
    """Populates the master GSTR-1 template using parsed PDF metadata and tables."""

    MONTH_LIST = [
        "April", "May", "June", "July", "August", "September",
        "October", "November", "December", "January", "February", "March"
    ]

    def __init__(self, log_callback=None):
        self.parser = GSTR1PDFParser(log_callback=log_callback)
        self.log_callback = log_callback

    def log(self, msg: str):
        if self.log_callback:
            self.log_callback(msg)

    def _set(self, ws, row: int, col: int, val: Any):
        """Safely writes to a cell, ignoring non-top-left MergedCells."""
        cell = ws.cell(row, col)
        if not isinstance(cell, MergedCell):
            cell.value = val

    def process_and_consolidate(
        self,
        pdf_paths: List[str],
        output_filepath: str,
        progress_callback=None
    ) -> Dict[str, Any]:
        """Parses GSTR-1 PDFs, populates metadata, and writes to GSTR-1 master template."""
        self.log(f"[GSTR-1 Consolidator] Parsing {len(pdf_paths)} GSTR-1 PDF return(s)...")

        parsed_months: Dict[str, Dict[str, Any]] = {}
        primary_meta = {"gstin": "", "legal_name": "", "trade_name": "", "financial_year": ""}

        for idx, p in enumerate(pdf_paths, 1):
            data = self.parser.parse_pdf(p)
            m_name = data.get("month_name", "")
            if m_name in self.MONTH_LIST:
                parsed_months[m_name] = data
                self.log(f"   [Parsed] Month: {m_name} -> {os.path.basename(p)}")
            else:
                self.log(f"   [Warning] Unknown month in {os.path.basename(p)}")

            if not primary_meta["gstin"] and data.get("gstin"):
                primary_meta["gstin"] = data["gstin"]
            if not primary_meta["legal_name"] and data.get("legal_name"):
                primary_meta["legal_name"] = data["legal_name"]
            if not primary_meta["trade_name"] and data.get("trade_name"):
                primary_meta["trade_name"] = data["trade_name"]
            if not primary_meta["financial_year"] and data.get("financial_year"):
                primary_meta["financial_year"] = data["financial_year"]

            if progress_callback:
                progress_callback(idx, len(pdf_paths), os.path.basename(p))

        # 1. Load Master GSTR 1 template
        template_file = resolve_template_path("Consolidated Report Format.xlsx")
        if template_file and os.path.exists(template_file):
            wb = openpyxl.load_workbook(template_file)
        elif get_embedded_template_stream is not None:
            wb = openpyxl.load_workbook(get_embedded_template_stream())
        else:
            raise FileNotFoundError("Master template not found on disk or embedded stream.")
        ws = wb['GSTR 1'] if 'GSTR 1' in wb.sheetnames else wb.active

        # 2. Convert all IFERROR and Google Sheets formulas to clean hardcoded text
        self._clean_all_iferror_and_ref_formulas(ws)

        # 3. Merge B:C for labels and write metadata values into D (keeping D:G merged)
        try:
            ws.merge_cells("B2:C2")
        except Exception:
            pass
        try:
            ws.merge_cells("B3:C3")
        except Exception:
            pass
        try:
            ws.merge_cells("B4:C4")
        except Exception:
            pass

        self._set(ws, 1, 3, None)  # Clean C1
        self._set(ws, 2, 3, None)  # Clean C2
        self._set(ws, 3, 3, None)  # Clean C3
        self._set(ws, 4, 3, None)  # Clean C4

        if primary_meta["gstin"]:
            self._set(ws, 2, 4, primary_meta["gstin"])        # D2 (merged D2:G2)
        if primary_meta["legal_name"]:
            self._set(ws, 3, 4, primary_meta["legal_name"])   # D3 (merged D3:G3)
        if primary_meta["trade_name"]:
            self._set(ws, 4, 4, primary_meta["trade_name"])   # D4 (merged D4:G4)

        # 4. Clear existing template dummy values in month rows
        self._clear_template_month_rows(ws)

        # 5. Populate each table section month-by-month at original coordinates
        for m_idx, month_name in enumerate(self.MONTH_LIST):
            if month_name in parsed_months:
                m_data = parsed_months[month_name]
                t = m_data.get("tables", {})
                self._populate_month_row(ws, m_idx, t)

        # 6. Set GSTR 1 as active sheet
        if 'GSTR 1' in wb.sheetnames:
            wb.active = wb['GSTR 1']

        # 7. Save workbook
        os.makedirs(os.path.dirname(os.path.abspath(output_filepath)), exist_ok=True)
        wb.save(output_filepath)
        wb.close()

        self.log(f"[GSTR-1 Consolidator] Successfully saved Master Consolidated Report:\n{output_filepath}")

        return {
            "output_file": output_filepath,
            "months_found": list(parsed_months.keys()),
            "total_months_populated": len(parsed_months),
            "metadata": primary_meta
        }

    def _clean_all_iferror_and_ref_formulas(self, ws):
        """
        Iterates over all cells in the sheet.
        Converts all IFERROR / xludf formulas into their clean hardcoded text or number,
        and clears all #REF! artifacts.
        """
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue
                val = cell.value
                if val and isinstance(val, str):
                    if "#REF!" in val:
                        cell.value = None
                    elif "IFERROR" in val or "xludf" in val or "DUMMYFUNCTION" in val:
                        cell.value = self._extract_iferror_fallback(val)

    def _extract_iferror_fallback(self, val: str) -> Any:
        """Extracts the final fallback string or number from an IFERROR formula."""
        if not isinstance(val, str):
            return val
        if "#REF!" in val:
            return None
        # Pattern 1: ... , "FALLBACK_TEXT")
        m_str = re.findall(r',\s*"([^"]*)"\s*\)', val)
        if m_str:
            res = m_str[-1].strip()
            return res if res and res != "COMPUTED_VALUE" else None
        # Pattern 2: ... , 123.45)
        m_num = re.findall(r',\s*([0-9.]+)\s*\)', val)
        if m_num:
            return float(m_num[-1]) if '.' in m_num[-1] else int(m_num[-1])
        # Pattern 3: any quoted string
        m_any = re.findall(r'"([^"]*)"', val)
        if m_any:
            res = m_any[-1].strip()
            return res if res and res != "COMPUTED_VALUE" else None
        return None

    def _clear_template_month_rows(self, ws):
        """Clears dummy values from month rows while keeping Column B Month names intact."""
        table_month_ranges = [
            (9, 20),    # Table 4A & 4B
            (25, 36),   # Table 5
            (42, 53),   # Table 6A, 6B & 6C
            (58, 69),   # Table 7
            (75, 86),   # Table 8
            (92, 103),  # Table 9A
            (110, 121), # Table 9B
            (128, 139), # Table 9C
            (145, 156), # Table 10
            (162, 173), # Table 11
            (178, 189), # Table 12
            (194, 205), # Table 13
            (211, 222), # Table 14
            (229, 240), # Table 14A
            (246, 257), # Table 15
            (264, 275), # Table 15A (I)
            (281, 292), # Table 15A (II)
        ]
        for start_r, end_r in table_month_ranges:
            for r in range(start_r, end_r + 1):
                # Clear columns C through CG (Cols 3 to 85)
                for c in range(3, 86):
                    self._set(ws, r, c, None)

    def _populate_month_row(self, ws, m_idx: int, m_data: Dict[str, Any]):
        """Populates all table values and ARN metadata at their exact original coordinates in the master template."""
        t = m_data.get("tables", m_data)

        # ─── Table 4A & 4B (Row 9 + m_idx) ──────────────────────────────────
        r4 = 9 + m_idx
        t4a = t.get("4A_B2B_Regular", {})
        t4b = t.get("4B_B2B_ReverseCharge", {})

        # 4A Regular
        self._set(ws, r4, 3, t4a.get("records", 0))                 # Col C: No. of Documents
        self._set(ws, r4, 4, "Invoice")                             # Col D: Document Type
        self._set(ws, r4, 5, t4a.get("value", 0.0))                 # Col E: Value
        self._set(ws, r4, 6, t4a.get("igst", 0.0))                  # Col F: IGST
        self._set(ws, r4, 7, t4a.get("cgst", 0.0))                  # Col G: CGST
        self._set(ws, r4, 8, t4a.get("sgst", 0.0))                  # Col H: SGST
        self._set(ws, r4, 9, t4a.get("cess", 0.0))                  # Col I: Cess

        # 4B Reverse Charge
        self._set(ws, r4, 11, t4b.get("records", 0))                # Col K: No. of Documents
        self._set(ws, r4, 12, "Invoice")                            # Col L: Document Type
        self._set(ws, r4, 13, t4b.get("value", 0.0))                # Col M: Value
        self._set(ws, r4, 14, t4b.get("igst", 0.0))                 # Col N: IGST
        self._set(ws, r4, 15, t4b.get("cgst", 0.0))                 # Col O: CGST
        self._set(ws, r4, 16, t4b.get("sgst", 0.0))                 # Col P: SGST
        self._set(ws, r4, 17, t4b.get("cess", 0.0))                 # Col Q: Cess

        # ─── Table 5: B2CL Large (Row 25 + m_idx) ───────────────────────────
        r5 = 25 + m_idx
        t5 = t.get("5_B2CL_Large", {})
        self._set(ws, r5, 3, t5.get("records", 0))                  # Col C: No. of Documents
        self._set(ws, r5, 4, "Invoice")                             # Col D: Document Type
        self._set(ws, r5, 5, t5.get("value", 0.0))                  # Col E: Value
        self._set(ws, r5, 6, t5.get("igst", 0.0))                   # Col F: IGST
        self._set(ws, r5, 7, t5.get("cess", 0.0))                   # Col G: Cess

        # ─── Table 6A, 6B & 6C (Row 42 + m_idx) ─────────────────────────────
        r6 = 42 + m_idx
        t6a = t.get("6A_Exports", {})
        t6b = t.get("6B_SEZ", {})
        t6c = t.get("6C_Deemed_Exports", {})

        # 6A EXPWP
        self._set(ws, r6, 3, t6a.get("expwp_records", 0))           # Col C: No. of Documents
        self._set(ws, r6, 4, "Invoice")
        self._set(ws, r6, 5, t6a.get("expwp_value", 0.0))
        self._set(ws, r6, 6, t6a.get("expwp_igst", 0.0))
        self._set(ws, r6, 7, t6a.get("expwp_cess", 0.0))

        # 6A EXPWOP
        self._set(ws, r6, 9, t6a.get("expwop_records", 0))          # Col I: No. of Documents
        self._set(ws, r6, 10, "Invoice")
        self._set(ws, r6, 11, t6a.get("expwop_value", 0.0))

        # 6B SEZWP
        self._set(ws, r6, 13, t6b.get("sezwp_records", 0))          # Col M: No. of Documents
        self._set(ws, r6, 14, "Invoice")
        self._set(ws, r6, 15, t6b.get("sezwp_value", 0.0))
        self._set(ws, r6, 16, t6b.get("sezwp_igst", 0.0))
        self._set(ws, r6, 17, t6b.get("sezwp_cess", 0.0))

        # 6B SEZWOP
        self._set(ws, r6, 19, t6b.get("sezwop_records", 0))         # Col S: No. of Documents
        self._set(ws, r6, 20, "Invoice")
        self._set(ws, r6, 21, t6b.get("sezwop_value", 0.0))

        # 6C Deemed Exports (DE)
        self._set(ws, r6, 23, t6c.get("records", 0))                # Col W: No. of Documents
        self._set(ws, r6, 24, "Invoice")
        self._set(ws, r6, 25, t6c.get("value", 0.0))
        self._set(ws, r6, 26, t6c.get("igst", 0.0))
        self._set(ws, r6, 27, t6c.get("cgst", 0.0))
        self._set(ws, r6, 28, t6c.get("sgst", 0.0))
        self._set(ws, r6, 29, t6c.get("cess", 0.0))

        # ─── Table 7: B2CS Small (Row 58 + m_idx) ───────────────────────────
        r7 = 58 + m_idx
        t7 = t.get("7_B2CS_Small", {})
        self._set(ws, r7, 3, t7.get("records", 0))                  # Col C: No. of Documents
        self._set(ws, r7, 4, "Invoice")                             # Col D: Document Type
        self._set(ws, r7, 5, t7.get("value", 0.0))                  # Col E: Value
        self._set(ws, r7, 6, t7.get("igst", 0.0))                   # Col F: IGST
        self._set(ws, r7, 7, t7.get("cgst", 0.0))                   # Col G: CGST
        self._set(ws, r7, 8, t7.get("sgst", 0.0))                   # Col H: SGST
        self._set(ws, r7, 9, t7.get("cess", 0.0))                   # Col I: Cess

        # ─── Table 8: Nil, Exempted & Non-GST (Row 75 + m_idx) ──────────────
        r8 = 75 + m_idx
        t8 = t.get("8_Nil_Exempted", {})
        self._set(ws, r8, 3, t8.get("nil", 0.0))                    # Col C: Nil Rated
        self._set(ws, r8, 4, t8.get("exempted", 0.0))               # Col D: Exempted
        self._set(ws, r8, 5, t8.get("non_gst", 0.0))                # Col E: Non-GST

        # ─── Table 9A: Amendments (Row 92 + m_idx) ───────────────────────────
        r9a = 92 + m_idx
        t9a = t.get("9A_Amendments", {})
        
        # 9A B2B Regular
        b2b_reg_am = t9a.get("b2b_reg_amended") or []
        self._set(ws, r9a, 3, int(b2b_reg_am[0]) if len(b2b_reg_am) > 0 else 0)
        self._set(ws, r9a, 4, "Invoice")
        self._set(ws, r9a, 5, b2b_reg_am[1] if len(b2b_reg_am) > 1 else 0.0)
        self._set(ws, r9a, 6, b2b_reg_am[2] if len(b2b_reg_am) > 2 else 0.0)
        self._set(ws, r9a, 7, b2b_reg_am[3] if len(b2b_reg_am) > 3 else 0.0)
        self._set(ws, r9a, 8, b2b_reg_am[4] if len(b2b_reg_am) > 4 else 0.0)
        self._set(ws, r9a, 9, b2b_reg_am[5] if len(b2b_reg_am) > 5 else 0.0)
        
        b2b_reg_df = t9a.get("b2b_reg_diff") or []
        self._set(ws, r9a, 11, b2b_reg_df[0] if len(b2b_reg_df) > 0 else 0.0)
        self._set(ws, r9a, 12, b2b_reg_df[1] if len(b2b_reg_df) > 1 else 0.0)
        self._set(ws, r9a, 13, b2b_reg_df[2] if len(b2b_reg_df) > 2 else 0.0)
        self._set(ws, r9a, 14, b2b_reg_df[3] if len(b2b_reg_df) > 3 else 0.0)
        self._set(ws, r9a, 15, b2b_reg_df[4] if len(b2b_reg_df) > 4 else 0.0)

        # 9A B2B Reverse Charge
        b2b_rc_am = t9a.get("b2b_rc_amended") or []
        self._set(ws, r9a, 17, int(b2b_rc_am[0]) if len(b2b_rc_am) > 0 else 0)
        self._set(ws, r9a, 18, "Invoice")
        self._set(ws, r9a, 19, b2b_rc_am[1] if len(b2b_rc_am) > 1 else 0.0)
        self._set(ws, r9a, 20, b2b_rc_am[2] if len(b2b_rc_am) > 2 else 0.0)
        self._set(ws, r9a, 21, b2b_rc_am[3] if len(b2b_rc_am) > 3 else 0.0)
        self._set(ws, r9a, 22, b2b_rc_am[4] if len(b2b_rc_am) > 4 else 0.0)
        self._set(ws, r9a, 23, b2b_rc_am[5] if len(b2b_rc_am) > 5 else 0.0)

        b2b_rc_df = t9a.get("b2b_rc_diff") or []
        self._set(ws, r9a, 25, b2b_rc_df[0] if len(b2b_rc_df) > 0 else 0.0)
        self._set(ws, r9a, 26, b2b_rc_df[1] if len(b2b_rc_df) > 1 else 0.0)
        self._set(ws, r9a, 27, b2b_rc_df[2] if len(b2b_rc_df) > 2 else 0.0)
        self._set(ws, r9a, 28, b2b_rc_df[3] if len(b2b_rc_df) > 3 else 0.0)
        self._set(ws, r9a, 29, b2b_rc_df[4] if len(b2b_rc_df) > 4 else 0.0)

        # 9A B2CL Large
        b2cl_am = t9a.get("b2cl_amended") or []
        self._set(ws, r9a, 31, int(b2cl_am[0]) if len(b2cl_am) > 0 else 0)
        self._set(ws, r9a, 32, "Invoice")
        self._set(ws, r9a, 33, b2cl_am[1] if len(b2cl_am) > 1 else 0.0)
        self._set(ws, r9a, 34, b2cl_am[2] if len(b2cl_am) > 2 else 0.0)
        self._set(ws, r9a, 35, b2cl_am[3] if len(b2cl_am) > 3 else 0.0)

        b2cl_df = t9a.get("b2cl_diff") or []
        self._set(ws, r9a, 37, b2cl_df[0] if len(b2cl_df) > 0 else 0.0)
        self._set(ws, r9a, 38, b2cl_df[1] if len(b2cl_df) > 1 else 0.0)
        self._set(ws, r9a, 39, b2cl_df[2] if len(b2cl_df) > 2 else 0.0)

        # 9A EXPWP
        expwp_am = t9a.get("expwp_amended") or []
        self._set(ws, r9a, 41, int(expwp_am[0]) if len(expwp_am) > 0 else 0)
        self._set(ws, r9a, 42, "Invoice")
        self._set(ws, r9a, 43, expwp_am[1] if len(expwp_am) > 1 else 0.0)
        self._set(ws, r9a, 44, expwp_am[2] if len(expwp_am) > 2 else 0.0)
        self._set(ws, r9a, 45, expwp_am[3] if len(expwp_am) > 3 else 0.0)

        expwp_df = t9a.get("expwp_diff") or []
        self._set(ws, r9a, 47, expwp_df[0] if len(expwp_df) > 0 else 0.0)
        self._set(ws, r9a, 48, expwp_df[1] if len(expwp_df) > 1 else 0.0)
        self._set(ws, r9a, 49, expwp_df[2] if len(expwp_df) > 2 else 0.0)

        # 9A EXPWOP
        expwop_am = t9a.get("expwop_amended") or []
        self._set(ws, r9a, 51, int(expwop_am[0]) if len(expwop_am) > 0 else 0)
        self._set(ws, r9a, 52, "Invoice")
        self._set(ws, r9a, 53, expwop_am[1] if len(expwop_am) > 1 else 0.0)

        expwop_df = t9a.get("expwop_diff") or []
        self._set(ws, r9a, 55, int(expwop_df[0]) if len(expwop_df) > 0 else 0)
        self._set(ws, r9a, 56, expwop_df[1] if len(expwop_df) > 1 else 0.0)

        # 9A SEZWP
        sezwp_am = t9a.get("sezwp_amended") or []
        self._set(ws, r9a, 58, int(sezwp_am[0]) if len(sezwp_am) > 0 else 0)
        self._set(ws, r9a, 59, "Invoice")
        self._set(ws, r9a, 60, sezwp_am[1] if len(sezwp_am) > 1 else 0.0)
        self._set(ws, r9a, 61, sezwp_am[2] if len(sezwp_am) > 2 else 0.0)
        self._set(ws, r9a, 62, sezwp_am[3] if len(sezwp_am) > 3 else 0.0)

        sezwp_df = t9a.get("sezwp_diff") or []
        self._set(ws, r9a, 64, sezwp_df[0] if len(sezwp_df) > 0 else 0.0)
        self._set(ws, r9a, 65, sezwp_df[1] if len(sezwp_df) > 1 else 0.0)
        self._set(ws, r9a, 66, sezwp_df[2] if len(sezwp_df) > 2 else 0.0)

        # 9A SEZWOP
        sezwop_am = t9a.get("sezwop_amended") or []
        self._set(ws, r9a, 68, int(sezwop_am[0]) if len(sezwop_am) > 0 else 0)
        self._set(ws, r9a, 69, "Invoice")
        self._set(ws, r9a, 70, sezwop_am[1] if len(sezwop_am) > 1 else 0.0)

        sezwop_df = t9a.get("sezwop_diff") or []
        self._set(ws, r9a, 72, int(sezwop_df[0]) if len(sezwop_df) > 0 else 0)
        self._set(ws, r9a, 73, sezwop_df[1] if len(sezwop_df) > 1 else 0.0)

        # 9A DE
        de_am = t9a.get("de_amended") or []
        self._set(ws, r9a, 75, int(de_am[0]) if len(de_am) > 0 else 0)
        self._set(ws, r9a, 76, "Invoice")
        self._set(ws, r9a, 77, de_am[1] if len(de_am) > 1 else 0.0)
        self._set(ws, r9a, 78, de_am[2] if len(de_am) > 2 else 0.0)
        self._set(ws, r9a, 79, de_am[3] if len(de_am) > 3 else 0.0)
        self._set(ws, r9a, 80, de_am[4] if len(de_am) > 4 else 0.0)
        self._set(ws, r9a, 81, de_am[5] if len(de_am) > 5 else 0.0)

        de_df = t9a.get("de_diff") or []
        self._set(ws, r9a, 83, de_df[0] if len(de_df) > 0 else 0.0)
        self._set(ws, r9a, 84, de_df[1] if len(de_df) > 1 else 0.0)
        self._set(ws, r9a, 85, de_df[2] if len(de_df) > 2 else 0.0)
        self._set(ws, r9a, 86, de_df[3] if len(de_df) > 3 else 0.0)
        self._set(ws, r9a, 87, de_df[4] if len(de_df) > 4 else 0.0)

        # ─── Table 9B: Credit/Debit Notes (Row 110 + m_idx) ──────────────────
        r9b = 110 + m_idx
        t9b = t.get("9B_CDNR", {})
        
        # 9B B2B Regular
        b2b_reg_9b = t9b.get("b2b_reg") or []
        self._set(ws, r9b, 3, int(b2b_reg_9b[0]) if len(b2b_reg_9b) > 0 else 0)
        self._set(ws, r9b, 4, "Note")
        self._set(ws, r9b, 5, b2b_reg_9b[1] if len(b2b_reg_9b) > 1 else 0.0)
        self._set(ws, r9b, 6, b2b_reg_9b[2] if len(b2b_reg_9b) > 2 else 0.0)
        self._set(ws, r9b, 7, b2b_reg_9b[3] if len(b2b_reg_9b) > 3 else 0.0)
        self._set(ws, r9b, 8, b2b_reg_9b[4] if len(b2b_reg_9b) > 4 else 0.0)
        self._set(ws, r9b, 9, b2b_reg_9b[5] if len(b2b_reg_9b) > 5 else 0.0)

        # 9B B2B Reverse Charge
        b2b_rc_9b = t9b.get("b2b_rc") or []
        self._set(ws, r9b, 11, int(b2b_rc_9b[0]) if len(b2b_rc_9b) > 0 else 0)
        self._set(ws, r9b, 12, "Note")
        self._set(ws, r9b, 13, b2b_rc_9b[1] if len(b2b_rc_9b) > 1 else 0.0)
        self._set(ws, r9b, 14, b2b_rc_9b[2] if len(b2b_rc_9b) > 2 else 0.0)
        self._set(ws, r9b, 15, b2b_rc_9b[3] if len(b2b_rc_9b) > 3 else 0.0)
        self._set(ws, r9b, 16, b2b_rc_9b[4] if len(b2b_rc_9b) > 4 else 0.0)
        self._set(ws, r9b, 17, b2b_rc_9b[5] if len(b2b_rc_9b) > 5 else 0.0)

        # 9B SEZ
        sez_9b = t9b.get("sez") or []
        self._set(ws, r9b, 19, int(sez_9b[0]) if len(sez_9b) > 0 else 0)
        self._set(ws, r9b, 20, "Note")
        self._set(ws, r9b, 21, sez_9b[1] if len(sez_9b) > 1 else 0.0)
        self._set(ws, r9b, 22, sez_9b[2] if len(sez_9b) > 2 else 0.0)
        self._set(ws, r9b, 23, sez_9b[3] if len(sez_9b) > 3 else 0.0)

        # 9B DE
        de_9b = t9b.get("de") or []
        self._set(ws, r9b, 25, int(de_9b[0]) if len(de_9b) > 0 else 0)
        self._set(ws, r9b, 26, "Note")
        self._set(ws, r9b, 27, de_9b[1] if len(de_9b) > 1 else 0.0)
        self._set(ws, r9b, 28, de_9b[2] if len(de_9b) > 2 else 0.0)
        self._set(ws, r9b, 29, de_9b[3] if len(de_9b) > 3 else 0.0)
        self._set(ws, r9b, 30, de_9b[4] if len(de_9b) > 4 else 0.0)
        self._set(ws, r9b, 31, de_9b[5] if len(de_9b) > 5 else 0.0)

        # 9B CDNUR B2CL
        cdnur_b2cl = t9b.get("cdnur_b2cl") or []
        self._set(ws, r9b, 33, int(cdnur_b2cl[0]) if len(cdnur_b2cl) > 0 else 0)
        self._set(ws, r9b, 34, "Note")
        self._set(ws, r9b, 35, cdnur_b2cl[1] if len(cdnur_b2cl) > 1 else 0.0)
        self._set(ws, r9b, 36, cdnur_b2cl[2] if len(cdnur_b2cl) > 2 else 0.0)
        self._set(ws, r9b, 37, cdnur_b2cl[3] if len(cdnur_b2cl) > 3 else 0.0)

        # 9B CDNUR EXPWP
        cdnur_expwp = t9b.get("cdnur_expwp") or []
        self._set(ws, r9b, 39, int(cdnur_expwp[0]) if len(cdnur_expwp) > 0 else 0)
        self._set(ws, r9b, 40, "Note")
        self._set(ws, r9b, 41, cdnur_expwp[1] if len(cdnur_expwp) > 1 else 0.0)
        self._set(ws, r9b, 42, cdnur_expwp[2] if len(cdnur_expwp) > 2 else 0.0)
        self._set(ws, r9b, 43, cdnur_expwp[3] if len(cdnur_expwp) > 3 else 0.0)

        # 9B CDNUR EXPWOP
        cdnur_expwop = t9b.get("cdnur_expwop") or []
        self._set(ws, r9b, 45, int(cdnur_expwop[0]) if len(cdnur_expwop) > 0 else 0)
        self._set(ws, r9b, 46, "Note")
        self._set(ws, r9b, 47, cdnur_expwop[1] if len(cdnur_expwop) > 1 else 0.0)

        # ─── Table 9C: Amended Credit/Debit Notes (Row 128 + m_idx) ──────────
        r9c = 128 + m_idx
        t9c = t.get("9C_CDNRA", {})

        # 9C B2B Regular
        b2b_reg_9c = t9c.get("b2b_reg") or []
        self._set(ws, r9c, 3, int(b2b_reg_9c[0]) if len(b2b_reg_9c) > 0 else 0)
        self._set(ws, r9c, 4, "Note")
        self._set(ws, r9c, 5, b2b_reg_9c[1] if len(b2b_reg_9c) > 1 else 0.0)
        self._set(ws, r9c, 6, b2b_reg_9c[2] if len(b2b_reg_9c) > 2 else 0.0)
        self._set(ws, r9c, 7, b2b_reg_9c[3] if len(b2b_reg_9c) > 3 else 0.0)
        self._set(ws, r9c, 8, b2b_reg_9c[4] if len(b2b_reg_9c) > 4 else 0.0)
        self._set(ws, r9c, 9, b2b_reg_9c[5] if len(b2b_reg_9c) > 5 else 0.0)

        # 9C B2B Reverse Charge
        b2b_rc_9c = t9c.get("b2b_rc") or []
        self._set(ws, r9c, 11, int(b2b_rc_9c[0]) if len(b2b_rc_9c) > 0 else 0)
        self._set(ws, r9c, 12, "Note")
        self._set(ws, r9c, 13, b2b_rc_9c[1] if len(b2b_rc_9c) > 1 else 0.0)
        self._set(ws, r9c, 14, b2b_rc_9c[2] if len(b2b_rc_9c) > 2 else 0.0)
        self._set(ws, r9c, 15, b2b_rc_9c[3] if len(b2b_rc_9c) > 3 else 0.0)
        self._set(ws, r9c, 16, b2b_rc_9c[4] if len(b2b_rc_9c) > 4 else 0.0)
        self._set(ws, r9c, 17, b2b_rc_9c[5] if len(b2b_rc_9c) > 5 else 0.0)

        # 9C SEZ
        sez_9c = t9c.get("sez") or []
        self._set(ws, r9c, 19, int(sez_9c[0]) if len(sez_9c) > 0 else 0)
        self._set(ws, r9c, 20, "Note")
        self._set(ws, r9c, 21, sez_9c[1] if len(sez_9c) > 1 else 0.0)
        self._set(ws, r9c, 22, sez_9c[2] if len(sez_9c) > 2 else 0.0)
        self._set(ws, r9c, 23, sez_9c[3] if len(sez_9c) > 3 else 0.0)

        # 9C DE
        de_9c = t9c.get("de") or []
        self._set(ws, r9c, 25, int(de_9c[0]) if len(de_9c) > 0 else 0)
        self._set(ws, r9c, 26, "Note")
        self._set(ws, r9c, 27, de_9c[1] if len(de_9c) > 1 else 0.0)
        self._set(ws, r9c, 28, de_9c[2] if len(de_9c) > 2 else 0.0)
        self._set(ws, r9c, 29, de_9c[3] if len(de_9c) > 3 else 0.0)
        self._set(ws, r9c, 30, de_9c[4] if len(de_9c) > 4 else 0.0)
        self._set(ws, r9c, 31, de_9c[5] if len(de_9c) > 5 else 0.0)

        # 9C CDNR Amended Total
        cdnr_tot_9c = t9c.get("cdnr_total") or []
        self._set(ws, r9c, 33, int(cdnr_tot_9c[0]) if len(cdnr_tot_9c) > 0 else 0)
        self._set(ws, r9c, 34, "Note")
        self._set(ws, r9c, 35, cdnr_tot_9c[1] if len(cdnr_tot_9c) > 1 else 0.0)
        self._set(ws, r9c, 36, cdnr_tot_9c[2] if len(cdnr_tot_9c) > 2 else 0.0)
        self._set(ws, r9c, 37, cdnr_tot_9c[3] if len(cdnr_tot_9c) > 3 else 0.0)
        self._set(ws, r9c, 38, cdnr_tot_9c[4] if len(cdnr_tot_9c) > 4 else 0.0)
        self._set(ws, r9c, 39, cdnr_tot_9c[5] if len(cdnr_tot_9c) > 5 else 0.0)

        # 9C CDNURA B2CL
        cdnura_b2cl = t9c.get("cdnura_b2cl") or []
        self._set(ws, r9c, 41, int(cdnura_b2cl[0]) if len(cdnura_b2cl) > 0 else 0)
        self._set(ws, r9c, 42, "Note")
        self._set(ws, r9c, 43, cdnura_b2cl[1] if len(cdnura_b2cl) > 1 else 0.0)
        self._set(ws, r9c, 44, cdnura_b2cl[2] if len(cdnura_b2cl) > 2 else 0.0)
        self._set(ws, r9c, 45, cdnura_b2cl[3] if len(cdnura_b2cl) > 3 else 0.0)

        # 9C CDNURA EXPWP
        cdnura_expwp = t9c.get("cdnura_expwp") or []
        self._set(ws, r9c, 47, int(cdnura_expwp[0]) if len(cdnura_expwp) > 0 else 0)
        self._set(ws, r9c, 48, "Note")
        self._set(ws, r9c, 49, cdnura_expwp[1] if len(cdnura_expwp) > 1 else 0.0)
        self._set(ws, r9c, 50, cdnura_expwp[2] if len(cdnura_expwp) > 2 else 0.0)
        self._set(ws, r9c, 51, cdnura_expwp[3] if len(cdnura_expwp) > 3 else 0.0)

        # 9C CDNURA EXPWOP
        cdnura_expwop = t9c.get("cdnura_expwop") or []
        self._set(ws, r9c, 53, int(cdnura_expwop[0]) if len(cdnura_expwop) > 0 else 0)
        self._set(ws, r9c, 54, "Note")
        self._set(ws, r9c, 55, cdnura_expwop[1] if len(cdnura_expwop) > 1 else 0.0)

        # 9C CDNURA Amended Total
        cdnura_tot_9c = t9c.get("cdnura_total") or []
        self._set(ws, r9c, 57, int(cdnura_tot_9c[0]) if len(cdnura_tot_9c) > 0 else 0)
        self._set(ws, r9c, 58, "Note")
        self._set(ws, r9c, 59, cdnura_tot_9c[1] if len(cdnura_tot_9c) > 1 else 0.0)
        self._set(ws, r9c, 60, cdnura_tot_9c[2] if len(cdnura_tot_9c) > 2 else 0.0)
        self._set(ws, r9c, 61, cdnura_tot_9c[3] if len(cdnura_tot_9c) > 3 else 0.0)

        # ─── Table 10: Amendment to B2C (Row 145 + m_idx) ────────────────────
        r10 = 145 + m_idx
        t10 = t.get("10_Amended_B2C", {})
        self._set(ws, r10, 3, t10.get("net_diff_value", 0.0))       # Col C: Net Diff Value
        self._set(ws, r10, 4, t10.get("net_diff_igst", 0.0))
        self._set(ws, r10, 5, t10.get("net_diff_cgst", 0.0))
        self._set(ws, r10, 6, t10.get("net_diff_sgst", 0.0))
        self._set(ws, r10, 7, t10.get("net_diff_cess", 0.0))
        self._set(ws, r10, 9, t10.get("amended_records", 0))        # Col I: Amended Docs
        self._set(ws, r10, 10, "Net Value")
        self._set(ws, r10, 11, t10.get("amended_value", 0.0))
        self._set(ws, r10, 12, t10.get("amended_igst", 0.0))
        self._set(ws, r10, 13, t10.get("amended_cgst", 0.0))
        self._set(ws, r10, 14, t10.get("amended_sgst", 0.0))
        self._set(ws, r10, 15, t10.get("amended_cess", 0.0))

        # ─── Table 11: Advances Received & Adjusted (Row 162 + m_idx) ───────
        r11 = 162 + m_idx
        t11 = t.get("11_Advances", {})
        
        # 11A Received
        rec_11a = t11.get("rec_records") or []
        self._set(ws, r11, 3, int(rec_11a[0]) if len(rec_11a) > 0 else 0)
        self._set(ws, r11, 4, "Net Value")
        self._set(ws, r11, 5, rec_11a[1] if len(rec_11a) > 1 else 0.0)
        self._set(ws, r11, 6, rec_11a[2] if len(rec_11a) > 2 else 0.0)
        self._set(ws, r11, 7, rec_11a[3] if len(rec_11a) > 3 else 0.0)
        self._set(ws, r11, 8, rec_11a[4] if len(rec_11a) > 4 else 0.0)
        self._set(ws, r11, 9, rec_11a[5] if len(rec_11a) > 5 else 0.0)

        # 11B Adjusted
        adj_11b = t11.get("adj_records") or []
        self._set(ws, r11, 11, int(adj_11b[0]) if len(adj_11b) > 0 else 0)
        self._set(ws, r11, 12, "Net Value")
        self._set(ws, r11, 13, adj_11b[1] if len(adj_11b) > 1 else 0.0)
        self._set(ws, r11, 14, adj_11b[2] if len(adj_11b) > 2 else 0.0)
        self._set(ws, r11, 15, adj_11b[3] if len(adj_11b) > 3 else 0.0)
        self._set(ws, r11, 16, adj_11b[4] if len(adj_11b) > 4 else 0.0)
        self._set(ws, r11, 17, adj_11b[5] if len(adj_11b) > 5 else 0.0)

        # 11A Amendment Total
        amend_tot_11a = t11.get("11a_amend_total") or []
        self._set(ws, r11, 19, amend_tot_11a[0] if len(amend_tot_11a) > 0 else 0.0)
        self._set(ws, r11, 20, amend_tot_11a[1] if len(amend_tot_11a) > 1 else 0.0)
        self._set(ws, r11, 21, amend_tot_11a[2] if len(amend_tot_11a) > 2 else 0.0)
        self._set(ws, r11, 22, amend_tot_11a[3] if len(amend_tot_11a) > 3 else 0.0)
        self._set(ws, r11, 23, amend_tot_11a[4] if len(amend_tot_11a) > 4 else 0.0)

        # 11A Amendment Amended Total
        amend_am_11a = t11.get("11a_amend_amended") or []
        self._set(ws, r11, 25, int(amend_am_11a[0]) if len(amend_am_11a) > 0 else 0)
        self._set(ws, r11, 26, "Net Value")
        self._set(ws, r11, 27, amend_am_11a[1] if len(amend_am_11a) > 1 else 0.0)
        self._set(ws, r11, 28, amend_am_11a[2] if len(amend_am_11a) > 2 else 0.0)
        self._set(ws, r11, 29, amend_am_11a[3] if len(amend_am_11a) > 3 else 0.0)
        self._set(ws, r11, 30, amend_am_11a[4] if len(amend_am_11a) > 4 else 0.0)
        self._set(ws, r11, 31, amend_am_11a[5] if len(amend_am_11a) > 5 else 0.0)

        # 11B Amendment Total
        amend_tot_11b = t11.get("11b_amend_total") or []
        self._set(ws, r11, 33, amend_tot_11b[0] if len(amend_tot_11b) > 0 else 0.0)
        self._set(ws, r11, 34, amend_tot_11b[1] if len(amend_tot_11b) > 1 else 0.0)
        self._set(ws, r11, 35, amend_tot_11b[2] if len(amend_tot_11b) > 2 else 0.0)
        self._set(ws, r11, 36, amend_tot_11b[3] if len(amend_tot_11b) > 3 else 0.0)
        self._set(ws, r11, 37, amend_tot_11b[4] if len(amend_tot_11b) > 4 else 0.0)

        # 11B Amendment Amended Total
        amend_am_11b = t11.get("11b_amend_amended") or []
        self._set(ws, r11, 39, int(amend_am_11b[0]) if len(amend_am_11b) > 0 else 0)
        self._set(ws, r11, 40, "Net Value")
        self._set(ws, r11, 41, amend_am_11b[1] if len(amend_am_11b) > 1 else 0.0)
        self._set(ws, r11, 42, amend_am_11b[2] if len(amend_am_11b) > 2 else 0.0)
        self._set(ws, r11, 43, amend_am_11b[3] if len(amend_am_11b) > 3 else 0.0)
        self._set(ws, r11, 44, amend_am_11b[4] if len(amend_am_11b) > 4 else 0.0)
        self._set(ws, r11, 45, amend_am_11b[5] if len(amend_am_11b) > 5 else 0.0)

        # ─── Table 12: HSN Summary (Row 178 + m_idx) ────────────────────────
        r12 = 178 + m_idx
        t12 = t.get("12_HSN_Summary", {})
        self._set(ws, r12, 3, t12.get("total_records", 0))          # Col C: No. of Documents
        self._set(ws, r12, 4, "NA")
        self._set(ws, r12, 5, t12.get("total_value", 0.0))          # Col E: Value
        self._set(ws, r12, 6, t12.get("total_igst", 0.0))           # Col F: IGST
        self._set(ws, r12, 7, t12.get("total_cgst", 0.0))           # Col G: CGST
        self._set(ws, r12, 8, t12.get("total_sgst", 0.0))           # Col H: SGST
        self._set(ws, r12, 9, t12.get("total_cess", 0.0))           # Col I: Cess

        # ─── Table 13: Documents Issued & ARN (Row 194 + m_idx) ─────────────
        r13 = 194 + m_idx
        t13 = t.get("13_Documents_Issued", {})
        self._set(ws, r13, 2, self.MONTH_LIST[m_idx])               # Col B: Month
        self._set(ws, r13, 3, t13.get("net_issued_documents", 0))   # Col C: No. of Documents
        self._set(ws, r13, 4, "All Documents")                      # Col D: Document Type
        self._set(ws, r13, 6, m_data.get("financial_year") or "")   # Col F: Financial Year
        self._set(ws, r13, 8, m_data.get("arn") or "")              # Col H: ARN
        self._set(ws, r13, 10, m_data.get("arn_date") or "")        # Col J: ARN Date

        # ─── Table 14: E-Commerce (Row 211 + m_idx) ─────────────────────────
        r14 = 211 + m_idx
        t14 = t.get("14_ECommerce", {})
        self._set(ws, r14, 3, t14.get("sec52_records", 0))          # Col C: Sec 52 Docs
        self._set(ws, r14, 4, "Net Value")
        self._set(ws, r14, 5, t14.get("sec52_value", 0.0))          # Col E: Sec 52 Value
        self._set(ws, r14, 6, t14.get("sec52_igst", 0.0))           # Col F: IGST
        self._set(ws, r14, 7, t14.get("sec52_cgst", 0.0))           # Col G: CGST
        self._set(ws, r14, 8, t14.get("sec52_sgst", 0.0))           # Col H: SGST
        self._set(ws, r14, 9, t14.get("sec52_cess", 0.0))           # Col I: Cess
        self._set(ws, r14, 11, t14.get("sec9_5_records", 0))        # Col K: Sec 9(5) Docs
        self._set(ws, r14, 12, "Net Value")
        self._set(ws, r14, 13, t14.get("sec9_5_value", 0.0))
        self._set(ws, r14, 14, t14.get("sec9_5_igst", 0.0))
        self._set(ws, r14, 15, t14.get("sec9_5_cgst", 0.0))
        self._set(ws, r14, 16, t14.get("sec9_5_sgst", 0.0))
        self._set(ws, r14, 17, t14.get("sec9_5_cess", 0.0))

        # ─── Table 14A: Amended E-Commerce (Row 229 + m_idx) ────────────────
        r14a = 229 + m_idx
        t14a = t.get("14A_Amended_ECommerce", {})
        
        # Sec 52 Amended
        sec52_am = t14a.get("sec52_amended") or []
        self._set(ws, r14a, 3, int(sec52_am[0]) if len(sec52_am) > 0 else 0)
        self._set(ws, r14a, 4, "Net Value")
        self._set(ws, r14a, 5, sec52_am[1] if len(sec52_am) > 1 else 0.0)
        self._set(ws, r14a, 6, sec52_am[2] if len(sec52_am) > 2 else 0.0)
        self._set(ws, r14a, 7, sec52_am[3] if len(sec52_am) > 3 else 0.0)
        self._set(ws, r14a, 8, sec52_am[4] if len(sec52_am) > 4 else 0.0)
        self._set(ws, r14a, 9, sec52_am[5] if len(sec52_am) > 5 else 0.0)

        # Sec 52 Net Diff
        sec52_df = t14a.get("sec52_diff") or []
        self._set(ws, r14a, 11, int(sec52_df[0]) if len(sec52_df) > 0 else 0)
        self._set(ws, r14a, 12, "Net Value")
        self._set(ws, r14a, 13, sec52_df[1] if len(sec52_df) > 1 else 0.0)
        self._set(ws, r14a, 14, sec52_df[2] if len(sec52_df) > 2 else 0.0)
        self._set(ws, r14a, 15, sec52_df[3] if len(sec52_df) > 3 else 0.0)
        self._set(ws, r14a, 16, sec52_df[4] if len(sec52_df) > 4 else 0.0)
        self._set(ws, r14a, 17, sec52_df[5] if len(sec52_df) > 5 else 0.0)

        # Sec 9(5) Amended
        sec95_am = t14a.get("sec9_5_amended") or []
        self._set(ws, r14a, 19, int(sec95_am[0]) if len(sec95_am) > 0 else 0)
        self._set(ws, r14a, 20, "Net Value")
        self._set(ws, r14a, 21, sec95_am[1] if len(sec95_am) > 1 else 0.0)
        self._set(ws, r14a, 22, sec95_am[2] if len(sec95_am) > 2 else 0.0)
        self._set(ws, r14a, 23, sec95_am[3] if len(sec95_am) > 3 else 0.0)
        self._set(ws, r14a, 24, sec95_am[4] if len(sec95_am) > 4 else 0.0)
        self._set(ws, r14a, 25, sec95_am[5] if len(sec95_am) > 5 else 0.0)

        # Sec 9(5) Net Diff
        sec95_df = t14a.get("sec9_5_diff") or []
        self._set(ws, r14a, 27, int(sec95_df[0]) if len(sec95_df) > 0 else 0)
        self._set(ws, r14a, 28, "Net Value")
        self._set(ws, r14a, 29, sec95_df[1] if len(sec95_df) > 1 else 0.0)
        self._set(ws, r14a, 30, sec95_df[2] if len(sec95_df) > 2 else 0.0)
        self._set(ws, r14a, 31, sec95_df[3] if len(sec95_df) > 3 else 0.0)
        self._set(ws, r14a, 32, sec95_df[4] if len(sec95_df) > 4 else 0.0)
        self._set(ws, r14a, 33, sec95_df[5] if len(sec95_df) > 5 else 0.0)

        # ─── Table 15: Supplies U/s 9(5) (Row 246 + m_idx) ──────────────────
        r15 = 246 + m_idx
        t15 = t.get("15_Supplies_9_5", {})
        
        # For Registered Recipients Total
        reg_tot_15 = t15.get("reg_total") or []
        self._set(ws, r15, 3, int(reg_tot_15[0]) if len(reg_tot_15) > 0 else 0)
        self._set(ws, r15, 4, "Document")
        self._set(ws, r15, 5, reg_tot_15[1] if len(reg_tot_15) > 1 else 0.0)
        self._set(ws, r15, 6, reg_tot_15[2] if len(reg_tot_15) > 2 else 0.0)
        self._set(ws, r15, 7, reg_tot_15[3] if len(reg_tot_15) > 3 else 0.0)
        self._set(ws, r15, 8, reg_tot_15[4] if len(reg_tot_15) > 4 else 0.0)
        self._set(ws, r15, 9, reg_tot_15[5] if len(reg_tot_15) > 5 else 0.0)

        # Regular
        reg_reg_15 = t15.get("reg_regular") or []
        self._set(ws, r15, 11, int(reg_reg_15[0]) if len(reg_reg_15) > 0 else 0)
        self._set(ws, r15, 12, "Document")
        self._set(ws, r15, 13, reg_reg_15[1] if len(reg_reg_15) > 1 else 0.0)
        self._set(ws, r15, 14, reg_reg_15[2] if len(reg_reg_15) > 2 else 0.0)
        self._set(ws, r15, 15, reg_reg_15[3] if len(reg_reg_15) > 3 else 0.0)
        self._set(ws, r15, 16, reg_reg_15[4] if len(reg_reg_15) > 4 else 0.0)
        self._set(ws, r15, 17, reg_reg_15[5] if len(reg_reg_15) > 5 else 0.0)

        # DE
        reg_de_15 = t15.get("reg_de") or []
        self._set(ws, r15, 19, int(reg_de_15[0]) if len(reg_de_15) > 0 else 0)
        self._set(ws, r15, 20, "Document")
        self._set(ws, r15, 21, reg_de_15[1] if len(reg_de_15) > 1 else 0.0)
        self._set(ws, r15, 22, reg_de_15[2] if len(reg_de_15) > 2 else 0.0)
        self._set(ws, r15, 23, reg_de_15[3] if len(reg_de_15) > 3 else 0.0)
        self._set(ws, r15, 24, reg_de_15[4] if len(reg_de_15) > 4 else 0.0)
        self._set(ws, r15, 25, reg_de_15[5] if len(reg_de_15) > 5 else 0.0)

        # SEZWP
        reg_sezwp_15 = t15.get("reg_sezwp") or []
        self._set(ws, r15, 27, int(reg_sezwp_15[0]) if len(reg_sezwp_15) > 0 else 0)
        self._set(ws, r15, 28, "Document")
        self._set(ws, r15, 29, reg_sezwp_15[1] if len(reg_sezwp_15) > 1 else 0.0)
        self._set(ws, r15, 30, reg_sezwp_15[2] if len(reg_sezwp_15) > 2 else 0.0)
        self._set(ws, r15, 31, reg_sezwp_15[3] if len(reg_sezwp_15) > 3 else 0.0)

        # SEZWOP
        reg_sezwop_15 = t15.get("reg_sezwop") or []
        self._set(ws, r15, 33, int(reg_sezwop_15[0]) if len(reg_sezwop_15) > 0 else 0)
        self._set(ws, r15, 34, "Document")
        self._set(ws, r15, 35, reg_sezwop_15[1] if len(reg_sezwop_15) > 1 else 0.0)

        # For Unregistered Recipient
        unreg_15 = t15.get("unregistered") or []
        self._set(ws, r15, 37, int(unreg_15[0]) if len(unreg_15) > 0 else 0)
        self._set(ws, r15, 38, "Net Value")
        self._set(ws, r15, 39, unreg_15[1] if len(unreg_15) > 1 else 0.0)
        self._set(ws, r15, 40, unreg_15[2] if len(unreg_15) > 2 else 0.0)
        self._set(ws, r15, 41, unreg_15[3] if len(unreg_15) > 3 else 0.0)
        self._set(ws, r15, 42, unreg_15[4] if len(unreg_15) > 4 else 0.0)
        self._set(ws, r15, 43, unreg_15[5] if len(unreg_15) > 5 else 0.0)

        # ─── Table 15A(I): Amended 9(5) Registered (Row 264 + m_idx) ────────
        r15a1 = 264 + m_idx
        t15a1 = t.get("15A_I_Amended_9_5", {})

        # Regular
        reg_15a1 = t15a1.get("regular") or []
        self._set(ws, r15a1, 3, int(reg_15a1[0]) if len(reg_15a1) > 0 else 0)
        self._set(ws, r15a1, 4, "Document")
        self._set(ws, r15a1, 5, reg_15a1[1] if len(reg_15a1) > 1 else 0.0)
        self._set(ws, r15a1, 6, reg_15a1[2] if len(reg_15a1) > 2 else 0.0)
        self._set(ws, r15a1, 7, reg_15a1[3] if len(reg_15a1) > 3 else 0.0)
        self._set(ws, r15a1, 8, reg_15a1[4] if len(reg_15a1) > 4 else 0.0)
        self._set(ws, r15a1, 9, reg_15a1[5] if len(reg_15a1) > 5 else 0.0)

        # DE
        de_15a1 = t15a1.get("de") or []
        self._set(ws, r15a1, 11, int(de_15a1[0]) if len(de_15a1) > 0 else 0)
        self._set(ws, r15a1, 12, "Document")
        self._set(ws, r15a1, 13, de_15a1[1] if len(de_15a1) > 1 else 0.0)
        self._set(ws, r15a1, 14, de_15a1[2] if len(de_15a1) > 2 else 0.0)
        self._set(ws, r15a1, 15, de_15a1[3] if len(de_15a1) > 3 else 0.0)
        self._set(ws, r15a1, 16, de_15a1[4] if len(de_15a1) > 4 else 0.0)
        self._set(ws, r15a1, 17, de_15a1[5] if len(de_15a1) > 5 else 0.0)

        # SEZWP
        sezwp_15a1 = t15a1.get("sezwp") or []
        self._set(ws, r15a1, 19, int(sezwp_15a1[0]) if len(sezwp_15a1) > 0 else 0)
        self._set(ws, r15a1, 20, "Document")
        self._set(ws, r15a1, 21, sezwp_15a1[1] if len(sezwp_15a1) > 1 else 0.0)
        self._set(ws, r15a1, 22, sezwp_15a1[2] if len(sezwp_15a1) > 2 else 0.0)
        self._set(ws, r15a1, 23, sezwp_15a1[3] if len(sezwp_15a1) > 3 else 0.0)

        # SEZWOP
        sezwop_15a1 = t15a1.get("sezwop") or []
        self._set(ws, r15a1, 25, int(sezwop_15a1[0]) if len(sezwop_15a1) > 0 else 0)
        self._set(ws, r15a1, 26, "Document")
        self._set(ws, r15a1, 27, sezwop_15a1[1] if len(sezwop_15a1) > 1 else 0.0)

        # ─── Table 15A(II): Amended 9(5) Unregistered (Row 281 + m_idx) ──────
        r15a2 = 281 + m_idx
        t15a2 = t.get("15A_II_Amended_9_5", {})

        # Amended Total
        am_15a2 = t15a2.get("amended_total") or []
        self._set(ws, r15a2, 3, int(am_15a2[0]) if len(am_15a2) > 0 else 0)
        self._set(ws, r15a2, 4, "Net Value")
        self._set(ws, r15a2, 5, am_15a2[1] if len(am_15a2) > 1 else 0.0)
        self._set(ws, r15a2, 6, am_15a2[2] if len(am_15a2) > 2 else 0.0)
        self._set(ws, r15a2, 7, am_15a2[3] if len(am_15a2) > 3 else 0.0)
        self._set(ws, r15a2, 8, am_15a2[4] if len(am_15a2) > 4 else 0.0)
        self._set(ws, r15a2, 9, am_15a2[5] if len(am_15a2) > 5 else 0.0)

        # Net Diff
        df_15a2 = t15a2.get("diff_total") or []
        self._set(ws, r15a2, 11, int(df_15a2[0]) if len(df_15a2) > 0 else 0)
        self._set(ws, r15a2, 12, "Net Value")
        self._set(ws, r15a2, 13, df_15a2[1] if len(df_15a2) > 1 else 0.0)
        self._set(ws, r15a2, 14, df_15a2[2] if len(df_15a2) > 2 else 0.0)
        self._set(ws, r15a2, 15, df_15a2[3] if len(df_15a2) > 3 else 0.0)
        self._set(ws, r15a2, 16, df_15a2[4] if len(df_15a2) > 4 else 0.0)
        self._set(ws, r15a2, 17, df_15a2[5] if len(df_15a2) > 5 else 0.0)

