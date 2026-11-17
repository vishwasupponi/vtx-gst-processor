import os
import sys

_CUR_DIR = os.path.dirname(os.path.abspath(__file__))
_VENDOR_DIR = os.path.join(_CUR_DIR, "_vendor")
if os.path.isdir(_VENDOR_DIR) and _VENDOR_DIR not in sys.path:
    sys.path.insert(0, _VENDOR_DIR)
if _CUR_DIR not in sys.path:
    sys.path.insert(0, _CUR_DIR)


class GSTProcessorModule:
    """Vertex Addon Module entrypoint for GST Processor & Consolidator."""

    def get_manifest(self) -> dict:
        return {
            "id": "gst_processor",
            "name": "GST Processor & Consolidator",
            "version": "1.0.0",
            "author": "Vishwas Upponi",
            "category": "Tax & Compliance"
        }
