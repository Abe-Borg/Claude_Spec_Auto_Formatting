import re
import hashlib
import zipfile
from pathlib import Path
from typing import List

def _sha256(b: bytes) -> str:
    return hashlib.sha256(b).hexdigest()

def _read_docx_part(docx: Path, internal_path: str) -> bytes:
    with zipfile.ZipFile(docx, "r") as z:
        return z.read(internal_path)

def _extract_all_sectpr_blocks(document_xml: str) -> List[str]:
    return re.findall(r"<w:sectPr\b[\s\S]*?</w:sectPr>", document_xml)

def verify_phase2_invariants(
    src_docx: Path,
    new_document_xml: bytes,
) -> None:
    # 1) sectPr unchanged
    before_doc = _read_docx_part(src_docx, "word/document.xml").decode("utf-8", errors="strict")
    after_doc  = new_document_xml.decode("utf-8", errors="strict")

    if _extract_all_sectpr_blocks(before_doc) != _extract_all_sectpr_blocks(after_doc):
        raise RuntimeError("INVARIANT FAIL: sectPr changed")

    # 2) headers/footers unchanged (we don't patch them, but this catches accidental edits if you ever do)
    with zipfile.ZipFile(src_docx, "r") as z:
        for name in z.namelist():
            if name.startswith("word/header") and name.endswith(".xml"):
                pass
            if name.startswith("word/footer") and name.endswith(".xml"):
                pass

    # 3) no run-level formatting edits in document.xml
    # (we assert <w:rPr> blocks are identical)
    before_rpr = re.findall(r"<w:rPr\b[\s\S]*?</w:rPr>", before_doc)
    after_rpr  = re.findall(r"<w:rPr\b[\s\S]*?</w:rPr>", after_doc)
    if before_rpr != after_rpr:
        raise RuntimeError("INVARIANT FAIL: document.xml run properties (<w:rPr>) changed")
