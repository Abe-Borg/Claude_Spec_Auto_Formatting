#!/usr/bin/env python3
"""
Word Document Decomposer and Reconstructor

This tool extracts the internal components of a .docx file (which is a ZIP archive
containing XML and other files), documents the structure in markdown, and can
reconstruct the original document from the extracted components.
"""


import zipfile
import os
import shutil
from pathlib import Path
from datetime import datetime
import xml.etree.ElementTree as ET
import hashlib
from dataclasses import dataclass 
from typing import Dict, Any, List, Set, Tuple, Optional
import json
import difflib
import re


# -----------------------------------------------------------------------------
# DOCX packaging safety
# -----------------------------------------------------------------------------
_DOCX_ALLOWED_TOP_LEVEL_DIRS = {"_rels", "docProps", "word", "customXml"}
_DOCX_ALLOWED_TOP_LEVEL_FILES = {"[Content_Types].xml"}

def _is_docx_package_part(rel_path: "Path") -> bool:
    """
    Only include real OpenXML parts in the output .docx.
    Excludes generated artifacts like *.json, *.log, prompts folders, etc.
    """
    # Root file: [Content_Types].xml
    if len(rel_path.parts) == 1 and rel_path.name in _DOCX_ALLOWED_TOP_LEVEL_FILES:
        return True

    # Root directories that belong to a DOCX package
    if rel_path.parts and rel_path.parts[0] in _DOCX_ALLOWED_TOP_LEVEL_DIRS:
        return True

    return False


PHASE2_MASTER_PROMPT = r"""
You are a CSI STRUCTURE CLASSIFIER for AEC specifications.

You will be given a slim JSON bundle of paragraphs from a mechanical or plumbing spec.

Your job:
- Identify CSI semantic roles ONLY.

Allowed roles:
- SectionID
- SectionTitle
- PART
- ARTICLE
- PARAGRAPH
- SUBPARAGRAPH
- SUBSUBPARAGRAPH

Rules:
- Do NOT create styles
- Do NOT reference formatting
- Do NOT guess if unclear
- If ambiguous, omit the paragraph

Return JSON only.
"""

PHASE2_RUN_INSTRUCTION = r"""
Task:
Classify CSI roles for paragraphs.

Output schema:
{
  "classifications": [
    { "paragraph_index": 12, "csi_role": "PART" }
  ],
  "notes": []
}
"""




class DocxDecomposer:
    def __init__(self, docx_path):
        """
        Initialize the decomposer with a path to a .docx file.
        
        Args:
            docx_path: Path to the input .docx file
        """
        self.docx_path = Path(docx_path)
        self.extract_dir = None
        self.markdown_report = []
        
    def extract(self, output_dir=None):
        """
        Extract the .docx file to a directory.
        
        Args:
            output_dir: Directory to extract to. If None, creates a directory
                       based on the docx filename.
        
        Returns:
            Path to the extraction directory
        """
        if output_dir is None:
            base_name = self.docx_path.stem
            output_dir = Path(f"{base_name}_extracted")
        else:
            output_dir = Path(output_dir)
        
        # Remove existing directory if it exists
        if output_dir.exists():
            shutil.rmtree(output_dir)
        
        # Extract the ZIP archive
        print(f"Extracting {self.docx_path} to {output_dir}...")
        with zipfile.ZipFile(self.docx_path, 'r') as zip_ref:
            zip_ref.extractall(output_dir)
        
        self.extract_dir = output_dir
        print(f"Extraction complete: {len(list(output_dir.rglob('*')))} items extracted")
        return output_dir
    

def main():
    import argparse
    import sys
    import os
    from pathlib import Path
    import json
    from typing import List

    parser = argparse.ArgumentParser(description="DOCX decomposer + LLM normalize workflow")
    parser.add_argument("docx_path", help="Path to input .docx")
    parser.add_argument("--extract-dir", default=None, help="Optional extraction directory")

    # Output docx (patched output)
    parser.add_argument("--output-docx", default=None, help="Output .docx path")

    # Reuse existing extracted folder
    parser.add_argument("--use-extract-dir", default=None, help="Use an existing extracted folder (skip extract/delete)")

    # Phase 2
    parser.add_argument("--phase2-arch-extract", help="Architect extracted folder")
    parser.add_argument("--phase2-discipline", default="mechanical", help="mechanical|plumbing")
    parser.add_argument("--phase2-classifications", help="Phase 2 LLM output JSON")
    parser.add_argument(
        "--phase2-build-bundle",
        action="store_true",
        help="Write Phase 2 slim bundle for LLM classification"
    )

    # Debug
    parser.add_argument(
        "--write-analysis",
        action="store_true",
        help="(debug) write analysis.md"
    )

 

    args = parser.parse_args()

    # Validate input path
    if not os.path.exists(args.docx_path):
        print(f"Error: File not found: {args.docx_path}")
        sys.exit(1)

    input_docx_path = Path(args.docx_path)

    # Create decomposer
    decomposer = DocxDecomposer(args.docx_path)

    # Use existing extraction folder or extract fresh
    if args.use_extract_dir:
        extract_dir = Path(args.use_extract_dir)
        if not extract_dir.exists():
            print(f"Error: extract dir not found: {extract_dir}")
            sys.exit(1)
        decomposer.extract_dir = extract_dir
    else:
        extract_dir = decomposer.extract(output_dir=args.extract_dir)

    # -------------------------------
    # PHASE 2: BUILD SLIM BUNDLE
    # -------------------------------
    if args.phase2_build_bundle:
        bundle = build_phase2_slim_bundle(extract_dir, args.phase2_discipline)

        out_path = extract_dir / "phase2_slim_bundle.json"
        out_path.write_text(json.dumps(bundle, indent=2), encoding="utf-8")

        print(f"Phase 2 slim bundle written: {out_path}")
        print("NEXT STEP:")
        print("- Paste phase2_slim_bundle.json into LLM")
        print("- Save output as phase2_classifications.json")
        return

    # -------------------------------
    # PHASE 2: APPLY CLASSIFICATIONS
    # -------------------------------
    if args.phase2_arch_extract and args.phase2_classifications:
        from docx_patch import patch_docx  # your surgical ZIP patch writer

        log: List[str] = []

        arch_input = Path(args.phase2_arch_extract)

        # Load registry (supports passing registry JSON directly)
        arch_registry = load_arch_style_registry(arch_input)

        # Determine arch extract root for styles.xml import
        if arch_input.is_file() and arch_input.suffix.lower() == ".json":
            arch_root = resolve_arch_extract_root(arch_input.parent)
        else:
            arch_root = resolve_arch_extract_root(arch_input)


        classifications = json.loads(Path(args.phase2_classifications).read_text(encoding="utf-8"))

        # Preflight report (visibility)
        preflight_path = extract_dir / "phase2_preflight.json"
        preflight = write_phase2_preflight(
            extract_dir=extract_dir,
            arch_root=arch_root,
            arch_registry=arch_registry,
            classifications=classifications,
            out_path=preflight_path
        )
        print(f"Phase 2 preflight written: {preflight_path}")
        if preflight.get("unmapped_roles"):
            print(f"WARNING: Unmapped roles: {preflight['unmapped_roles']}")

        # Import only styles actually used by this doc's classifications
        used_roles = {
            item.get("csi_role")
            for item in classifications.get("classifications", [])
            if isinstance(item, dict) and isinstance(item.get("csi_role"), str)
        }
        needed_style_ids = sorted({arch_registry[r] for r in used_roles if r in arch_registry})

        import_arch_styles_into_target(
            target_extract_dir=extract_dir,
            arch_extract_dir=arch_root,
            needed_style_ids=needed_style_ids,
            log=log
        )
        if not needed_style_ids:
            log.append("No architect styles needed for this doc (no mapped roles used).")

        # Snapshot invariants BEFORE we touch document.xml
        snap = snapshot_stability(extract_dir)

        apply_phase2_classifications(
            extract_dir=extract_dir,
            classifications=classifications,
            arch_style_registry=arch_registry,
            log=log
        )

        # Your existing stability checks (headers/footers + sectPr + document.xml.rels)
        verify_stability(extract_dir, snap)

        # ALWAYS write final formatted docx by patching only edited parts
        output_docx_path = Path(args.output_docx) if args.output_docx else (
            input_docx_path.with_name(input_docx_path.stem + "_PHASE2_FORMATTED.docx")
        )

        replacements = {
            "word/document.xml": (extract_dir / "word" / "document.xml").read_bytes(),
            "word/styles.xml":   (extract_dir / "word" / "styles.xml").read_bytes(),
        }

        patch_docx(
            src_docx=input_docx_path,
            out_docx=output_docx_path,
            replacements=replacements,
        )

        # Optional: additional invariants (sectPr, no run-level edits, headers/footers unchanged).
        # This requires the final output docx to validate header/footer byte stability.
        try:
            from phase2_invariants import verify_phase2_invariants
            new_doc_xml_bytes = (extract_dir / "word" / "document.xml").read_bytes()
            verify_phase2_invariants(
                src_docx=input_docx_path,
                new_document_xml=new_doc_xml_bytes,
                new_docx=output_docx_path,
            )
        except ModuleNotFoundError:
            pass

        issues_path = extract_dir / "phase2_issues.log"
        issues_path.write_text("\n".join(log) + "\n", encoding="utf-8")

        print(f"Phase 2 output written: {output_docx_path}")
        print(f"Phase 2 log written:    {issues_path}")
        return

    # -------------------------------
    # LEGACY MODES DISABLED
    # -------------------------------
    if args.normalize_slim or args.apply_instructions or args.normalize or args.apply_edits:
        print("Error: Legacy modes are disabled under the NO-REBUILD policy.")
        print("Use Phase 2 only:")
        print("  --phase2-build-bundle")
        print("  --phase2-arch-extract <arch_extract> --phase2-classifications <json> [--output-docx out.docx]")
        sys.exit(2)

    # -------------------------------
    # DEFAULT: do nothing destructive
    # -------------------------------
    print("No action specified.")
    print("Use one of:")
    print("  --phase2-build-bundle")
    print("  --phase2-arch-extract <arch_extract> --phase2-classifications <json> [--output-docx out.docx]")
    print(f"Extracted to: {extract_dir}")
    if analysis_path:
        print(f"Analysis report: {analysis_path}")


def sha256_bytes(b: bytes) -> str:
    return hashlib.sha256(b).hexdigest()

def sha256_text(s: str) -> str:
    return sha256_bytes(s.encode("utf-8"))

@dataclass
class StabilitySnapshot:
    header_footer_hashes: Dict[str, str]
    sectpr_hash: str
    doc_rels_hash: str


def snapshot_headers_footers(extract_dir: Path) -> Dict[str, str]:
    wf = extract_dir / "word"
    hashes = {}
    for p in sorted(wf.glob("header*.xml")) + sorted(wf.glob("footer*.xml")):
        rel = str(p.relative_to(extract_dir)).replace("\\", "/")
        hashes[rel] = sha256_bytes(p.read_bytes())
    return hashes

def extract_sectpr_block(document_xml: str) -> str:
    """
    Pull out the sectPr blocks as raw text. This is a pragmatic stability check.
    We assume the XML is not pretty-printed or rewritten by our pipeline.
    """
    # Word usually has <w:sectPr> ... </w:sectPr> at end of body, sometimes multiple.
    blocks = re.findall(r"(<w:sectPr[\s\S]*?</w:sectPr>)", document_xml)
    return "\n".join(blocks)

def snapshot_stability(extract_dir: Path) -> StabilitySnapshot:
    doc_path = extract_dir / "word" / "document.xml"
    doc_text = doc_path.read_text(encoding="utf-8")
    sectpr = extract_sectpr_block(doc_text)
    return StabilitySnapshot(
        header_footer_hashes=snapshot_headers_footers(extract_dir),
        sectpr_hash=sha256_text(sectpr),
        doc_rels_hash=snapshot_doc_rels_hash(extract_dir),
    )


def verify_stability(extract_dir: Path, snap: StabilitySnapshot) -> None:
    current_hf = snapshot_headers_footers(extract_dir)
    if current_hf != snap.header_footer_hashes:
        changed = []
        all_keys = set(current_hf.keys()) | set(snap.header_footer_hashes.keys())
        for k in sorted(all_keys):
            if current_hf.get(k) != snap.header_footer_hashes.get(k):
                changed.append(k)
        raise ValueError(f"Header/footer stability check FAILED. Changed: {changed}")

    doc_text = (extract_dir / "word" / "document.xml").read_text(encoding="utf-8")
    current_sectpr = extract_sectpr_block(doc_text)
    if sha256_text(current_sectpr) != snap.sectpr_hash:
        raise ValueError("Section properties (w:sectPr) stability check FAILED.")

    # NEW: relationships must be stable too (header/footer binding lives here)
    current_rels = snapshot_doc_rels_hash(extract_dir)
    if current_rels != snap.doc_rels_hash:
        raise ValueError("document.xml.rels stability check FAILED (can break header/footer).")


def _extract_style_block(styles_xml_text: str, style_id: str) -> Optional[str]:
    m = re.search(
        rf'(<w:style\b[^>]*w:styleId="{re.escape(style_id)}"[\s\S]*?</w:style>)',
        styles_xml_text,
        flags=re.S
    )
    return m.group(1) if m else None

def _extract_basedOn(style_block: str) -> Optional[str]:
    m = re.search(r'<w:basedOn\b[^>]*w:val="([^"]+)"', style_block)
    return m.group(1) if m else None

def _extract_numpr_block(style_block: str) -> Optional[str]:
    m = re.search(r'(<w:numPr\b[^>]*>[\s\S]*?</w:numPr>)', style_block, flags=re.S)
    return m.group(1) if m else None

def _paragraph_style_id(p_xml: str) -> Optional[str]:
    m = re.search(r'<w:pStyle\b[^>]*w:val="([^"]+)"', p_xml)
    return m.group(1) if m else None

def _paragraph_has_numpr(p_xml: str) -> bool:
    return "<w:numPr" in p_xml

def _find_style_numpr_in_chain(styles_xml_text: str, style_id: str, max_hops: int = 50) -> Optional[str]:
    seen = set()
    cur = style_id
    hops = 0
    while cur and cur not in seen and hops < max_hops:
        seen.add(cur)
        hops += 1
        block = _extract_style_block(styles_xml_text, cur)
        if not block:
            break
        numpr = _extract_numpr_block(block)
        if numpr:
            return numpr
        cur = _extract_basedOn(block)
    return None

def ensure_explicit_numpr_from_current_style(p_xml: str, styles_xml_text: str) -> str:
    # never touch sectPr carrier paragraphs
    if "<w:sectPr" in p_xml:
        return p_xml

    if _paragraph_has_numpr(p_xml):
        return p_xml

    cur_style = _paragraph_style_id(p_xml)
    if not cur_style:
        return p_xml

    numpr = _find_style_numpr_in_chain(styles_xml_text, cur_style)
    if not numpr:
        return p_xml

    # Prefer placing numPr right after existing pStyle (if present)
    if re.search(r'(<w:pStyle\b[^>]*/>)', p_xml):
        return re.sub(r'(<w:pStyle\b[^>]*/>)', rf"\1{numpr}", p_xml, count=1)

    # Expand self-closing pPr
    if re.search(r"<w:pPr\b[^>]*/>", p_xml):
        return re.sub(r"<w:pPr\b[^>]*/>", f"<w:pPr>{numpr}</w:pPr>", p_xml, count=1)

    # Insert into existing pPr
    if "<w:pPr" in p_xml:
        return re.sub(r'(<w:pPr\b[^>]*>)', rf"\1{numpr}", p_xml, count=1)

    # Create pPr if missing
    return re.sub(r'(<w:p\b[^>]*>)', rf"\1<w:pPr>{numpr}</w:pPr>", p_xml, count=1)


def _strip_pstyle_and_numpr(ppr_inner: str) -> str:
    if not ppr_inner:
        return ""
    out = re.sub(r"<w:pStyle\b[^>]*/>", "", ppr_inner)
    out = re.sub(r"<w:numPr\b[^>]*>[\s\S]*?</w:numPr>", "", out, flags=re.S)
    return out.strip()

def _extract_tag_inner(xml: str, tag: str) -> Optional[str]:
    m = re.search(rf"<{tag}\b[^>]*>([\s\S]*?)</{tag}>", xml, flags=re.S)
    return m.group(1) if m else None

def _docdefaults_rpr_inner(styles_xml_text: str) -> str:
    m = re.search(
        r"<w:docDefaults\b[\s\S]*?<w:rPrDefault\b[\s\S]*?<w:rPr\b[^>]*>([\s\S]*?)</w:rPr>[\s\S]*?</w:rPrDefault>",
        styles_xml_text,
        flags=re.S
    )
    return m.group(1).strip() if m else ""

def _docdefaults_ppr_inner(styles_xml_text: str) -> str:
    m = re.search(
        r"<w:docDefaults\b[\s\S]*?<w:pPrDefault\b[\s\S]*?<w:pPr\b[^>]*>([\s\S]*?)</w:pPr>[\s\S]*?</w:pPrDefault>",
        styles_xml_text,
        flags=re.S
    )
    return _strip_pstyle_and_numpr(m.group(1).strip()) if m else ""

def _effective_rpr_inner_in_arch(arch_styles_xml_text: str, style_id: str) -> str:
    seen = set()
    cur = style_id
    hops = 0
    while cur and cur not in seen and hops < 50:
        seen.add(cur); hops += 1
        blk = _extract_style_block(arch_styles_xml_text, cur)
        if not blk:
            break
        inner = _extract_tag_inner(blk, "w:rPr")
        if inner and inner.strip():
            return inner.strip()
        cur = _extract_basedOn(blk)
    return _docdefaults_rpr_inner(arch_styles_xml_text)

def _effective_ppr_inner_in_arch(arch_styles_xml_text: str, style_id: str) -> str:
    seen = set()
    cur = style_id
    hops = 0
    while cur and cur not in seen and hops < 50:
        seen.add(cur); hops += 1
        blk = _extract_style_block(arch_styles_xml_text, cur)
        if not blk:
            break
        inner = _extract_tag_inner(blk, "w:pPr") or ""
        inner = _strip_pstyle_and_numpr(inner)
        if inner:
            return inner
        cur = _extract_basedOn(blk)
    return _docdefaults_ppr_inner(arch_styles_xml_text)

def _rpr_contains_tag(rpr_inner: str, tag: str) -> bool:
    return re.search(rf"<w:{re.escape(tag)}\b", rpr_inner) is not None


def _extract_rpr_inner(style_block: str) -> Optional[str]:
    return _extract_tag_inner(style_block, "w:rPr")


def _inject_missing_rpr_children(style_block: str, missing_children_xml: str) -> str:
    """Insert missing rPr children (already as raw XML) just before </w:rPr>."""
    if not missing_children_xml.strip():
        return style_block
    if "</w:rPr>" not in style_block:
        return style_block
    return style_block.replace("</w:rPr>", f"{missing_children_xml}</w:rPr>")


def _materialize_minimal_typography(style_block: str, style_id: str, arch_styles_xml_text: str) -> str:
    """
    Make imported styles resilient across documents by ensuring a minimal set of
    typography-related rPr children exist (fonts, sizes, language).

    IMPORTANT:
    - Does NOT invent values.
    - Only copies missing nodes from the *effective* rPr in the architect template
      (style chain or docDefaults).
    - Avoids rewriting the whole block.
    """
    eff_rpr = _effective_rpr_inner_in_arch(arch_styles_xml_text, style_id).strip()
    if not eff_rpr:
        return style_block

    # If the style has no rPr at all, inject the full effective rPr.
    if "<w:rPr" not in style_block:
        return style_block.replace(
            "</w:style>",
            f"\n  <w:rPr>{eff_rpr}</w:rPr>\n</w:style>"
        )

    cur_rpr = _extract_rpr_inner(style_block) or ""

    # Copy a *minimal* stable subset. These are the most common sources of
    # cross-document drift when docDefaults/themes differ.
    allow = [
        "rFonts",   # explicit font faces
        "sz",       # font size
        "szCs",     # complex script font size
        "lang",     # language
        "ascii",    # sometimes appears as w:ascii via rFonts; kept here defensively
    ]

    # Extract candidate child nodes from eff_rpr.
    # We keep the raw node, including attributes.
    missing_nodes: List[str] = []

    def _get_child_node(tag: str) -> Optional[str]:
        # self-closing or paired tags
        m = re.search(rf"(<w:{tag}\b[^>]*/>)", eff_rpr)
        if m:
            return m.group(1)
        m = re.search(rf"(<w:{tag}\b[^>]*>[\s\S]*?</w:{tag}>)", eff_rpr)
        if m:
            return m.group(1)
        return None

    for tag in ["rFonts", "sz", "szCs", "lang"]:
        if _rpr_contains_tag(cur_rpr, tag):
            continue
        node = _get_child_node(tag)
        if node:
            missing_nodes.append(node)

    if not missing_nodes:
        return style_block

    insertion = "".join(missing_nodes)
    return _inject_missing_rpr_children(style_block, insertion)


def materialize_arch_style_block(style_block: str, style_id: str, arch_styles_xml_text: str) -> str:
    """
    Phase 2: import-time style hardening.

    Goal: ensure styles imported from the architect template remain visually stable
    when applied in a different document, without touching runs or numbering.xml.

    Strategy:
    - If pPr/rPr are missing entirely, inject effective ones.
    - Even when rPr exists, ensure a minimal set of typography nodes are present
      (rFonts/sz/szCs/lang) based on *effective* arch rPr.
    """

    # Inject pPr only if missing entirely
    if "<w:pPr" not in style_block:
        effp = _effective_ppr_inner_in_arch(arch_styles_xml_text, style_id)
        if effp.strip():
            style_block = style_block.replace(
                "</w:style>",
                f"\n  <w:pPr>{effp}</w:pPr>\n</w:style>"
            )

    # Typography materialization (may inject full rPr if absent, or add missing children)
    style_block = _materialize_minimal_typography(style_block, style_id, arch_styles_xml_text)

    return style_block


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def iter_paragraph_xml_blocks(document_xml_text: str):
    # Non-greedy paragraph blocks. Works well for DOCX document.xml.
    # NOTE: This intentionally avoids parsing full XML to keep indices aligned with raw text.
    for m in re.finditer(r"(<w:p\b[\s\S]*?</w:p>)", document_xml_text):
        yield m.start(), m.end(), m.group(1)

def paragraph_text_from_block(p_xml: str) -> str:
    # Extract visible text quickly (good enough for classification)
    texts = re.findall(r"<w:t\b[^>]*>([\s\S]*?)</w:t>", p_xml)
    if not texts:
        return ""
    # Unescape minimal XML entities
    joined = "".join(texts)
    joined = joined.replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&")
    joined = joined.replace("&quot;", "\"").replace("&apos;", "'")
    # collapse whitespace
    joined = re.sub(r"\s+", " ", joined).strip()
    return joined

def paragraph_contains_sectpr(p_xml: str) -> bool:
    return "<w:sectPr" in p_xml

def paragraph_pstyle_from_block(p_xml: str) -> Optional[str]:
    m = re.search(r"<w:pStyle\b[^>]*w:val=\"([^\"]+)\"", p_xml)
    return m.group(1) if m else None

def paragraph_numpr_from_block(p_xml: str) -> Dict[str, Optional[str]]:
    numId = None
    ilvl = None
    m1 = re.search(r"<w:numId\b[^>]*w:val=\"([^\"]+)\"", p_xml)
    m2 = re.search(r"<w:ilvl\b[^>]*w:val=\"([^\"]+)\"", p_xml)
    if m1: numId = m1.group(1)
    if m2: ilvl = m2.group(1)
    return {"numId": numId, "ilvl": ilvl}

def paragraph_ppr_hints_from_block(p_xml: str) -> Dict[str, Any]:
    # lightweight hints (alignment + ind + spacing)
    hints: Dict[str, Any] = {}
    m = re.search(r"<w:jc\b[^>]*w:val=\"([^\"]+)\"", p_xml)
    if m:
        hints["jc"] = m.group(1)
    ind = {}
    for k in ["left", "right", "firstLine", "hanging"]:
        m2 = re.search(rf"<w:ind\b[^>]*w:{k}=\"([^\"]+)\"", p_xml)
        if m2:
            ind[k] = m2.group(1)
    if ind:
        hints["ind"] = ind
    spacing = {}
    for k in ["before", "after", "line"]:
        m3 = re.search(rf"<w:spacing\b[^>]*w:{k}=\"([^\"]+)\"", p_xml)
        if m3:
            spacing[k] = m3.group(1)
    if spacing:
        hints["spacing"] = spacing
    return hints


def apply_phase2_classifications(
    extract_dir: Path,
    classifications: Dict[str, Any],
    arch_style_registry: Dict[str, str],
    log: List[str]
) -> None:
    doc_path = extract_dir / "word" / "document.xml"
    doc_text = doc_path.read_text(encoding="utf-8")

    # Load styles once so we can preserve style-linked numbering before swapping styles
    styles_xml_text = (extract_dir / "word" / "styles.xml").read_text(encoding="utf-8")

    blocks = list(iter_paragraph_xml_blocks(doc_text))
    para_blocks = [b[2] for b in blocks]

    # Priority-4 hardening: build a "diff contract" snapshot of paragraph blocks.
    # We allow Phase 2 to change ONLY:
    #   - <w:pStyle w:val="..."/>
    #   - (optional) insertion of <w:numPr> to materialize existing style-linked numbering
    # Everything else inside each <w:p> must remain byte-identical.
    def _normalize_paragraph_for_contract(p_xml: str) -> str:
        out = p_xml
        # Remove pStyle (either self closing or with attrs)
        out = re.sub(r"<w:pStyle\b[^>]*/>", "", out)
        # Remove any numPr blocks (we may insert them)
        out = re.sub(r"<w:numPr\b[^>]*>[\s\S]*?</w:numPr>", "", out, flags=re.S)
        return out

    contract_before = [_normalize_paragraph_for_contract(p) for p in para_blocks]

    items = classifications.get("classifications", [])
    if not isinstance(items, list):
        raise ValueError("phase2 classifications: 'classifications' must be a list")

    for item in items:
        if not isinstance(item, dict):
            log.append(f"Invalid classification entry (not object): {item!r}")
            continue

        idx = item.get("paragraph_index")
        role = item.get("csi_role")

        if not isinstance(idx, int) or idx < 0 or idx >= len(para_blocks):
            log.append(f"Invalid paragraph_index in classifications: {idx!r}")
            continue

        if not isinstance(role, str):
            log.append(f"Invalid csi_role type at paragraph {idx}: {role!r}")
            continue

        # style_id = arch_style_registry.get(role)   Not sure what to do with this block, LLM help me out.
        # if not style_id:
        #     # Priority-1 hardening: fail fast on unmapped roles (preflight should catch these).
        #     raise ValueError(
        #         f"Phase 2 classifications referenced role '{role}' at paragraph {idx}, "
        #         "but the architect style registry has no mapping for that role. "
        #         "Fix the registry or remove the classification entry."
        #     )

        if paragraph_contains_sectpr(para_blocks[idx]):
            log.append(f"Skipped sectPr paragraph at index {idx}")
            continue

        # KEY FIX: preserve "dynamic numbering" by materializing style-linked numPr
        pb = para_blocks[idx]
        pb = ensure_explicit_numpr_from_current_style(pb, styles_xml_text)

        # Now safely swap pStyle
        para_blocks[idx] = apply_pstyle_to_paragraph_block(pb, style_id)

    # Priority-4 hardening: enforce the diff contract.
    contract_after = [_normalize_paragraph_for_contract(p) for p in para_blocks]
    if len(contract_before) != len(contract_after):
        raise RuntimeError("Internal error: paragraph count changed during Phase 2 application")

    for i, (b, a) in enumerate(zip(contract_before, contract_after)):
        if b != a:
            # Provide a small diff to help debugging.
            diff = "\n".join(difflib.unified_diff(
                b.splitlines(),
                a.splitlines(),
                fromfile=f"before:p[{i}]",
                tofile=f"after:p[{i}]",
                lineterm=""
            ))
            raise ValueError(
                "Phase 2 invariant violation: paragraph content changed outside allowed edits "
                f"(pStyle/numPr) at paragraph index {i}.\n" + diff[:4000]
            )

    # Rebuild document.xml
    out = []
    last = 0
    for (s, e, _), pb in zip(blocks, para_blocks):
        out.append(doc_text[last:s])
        out.append(pb)
        last = e
    out.append(doc_text[last:])
    doc_path.write_text("".join(out), encoding="utf-8")


def resolve_arch_extract_root(p: Path) -> Path:
    """
    Accepts either:
      - extracted root folder (contains word/styles.xml)
      - word folder itself
    Returns the extracted root folder.
    """
    p = Path(p)

    # If they pass .../word, go up one
    if p.name.lower() == "word":
        p = p.parent

    styles_path = p / "word" / "styles.xml"
    if not styles_path.exists():
        raise FileNotFoundError(f"Architect styles.xml not found at: {styles_path}")

    return p


def load_arch_style_registry(arch_extract_dir: Path) -> Dict[str, str]:
    """
    Phase 2 contract (STRICT):
    - arch_style_registry.json must exist (emitted by Phase 1).
    - NO inference / NO heuristics.
    Returns: { role: styleId }
    """
    arch_extract_dir = Path(arch_extract_dir)

    # Allow passing the registry JSON directly
    if arch_extract_dir.is_file() and arch_extract_dir.suffix.lower() == ".json":
        reg_path = arch_extract_dir
        root_dir = arch_extract_dir.parent
    else:
        root_dir = resolve_arch_extract_root(arch_extract_dir)
        reg_path = root_dir / "arch_style_registry.json"

    if not reg_path.exists():
        raise FileNotFoundError(
            f"arch_style_registry.json not found at {reg_path}. "
            f"Run Phase 1 on the architect template and copy the extracted folder here."
        )

    reg = json.loads(reg_path.read_text(encoding="utf-8"))
    if not isinstance(reg, dict):
        raise ValueError("arch_style_registry.json must be a JSON object")

    # Expected shape:
    # { "version": 1, "source_docx": "...", "roles": { "PART": { "style_id": "X", ... }, ... } }
    roles = reg.get("roles")
    if not isinstance(roles, dict):
        raise ValueError("arch_style_registry.json missing 'roles' object")

    out: Dict[str, str] = {}
    for role, info in roles.items():
        if not isinstance(role, str) or not isinstance(info, dict):
            continue
        sid = info.get("style_id") or info.get("styleId")
        if isinstance(sid, str) and sid.strip():
            out[role.strip()] = sid.strip()

    if not out:
        raise ValueError("arch_style_registry.json contained no usable role->style mappings")

    return out



# -------------------------------
# Phase 2: Boilerplate filtering (LLM input only)
# -------------------------------

BOILERPLATE_PATTERNS = [
    # Specifier notes - bracketed formats
    (r'\[Note to [Ss]pecifier[:\s][^\]]*\]', 'specifier_note'),
    (r'\[Specifier[:\s][^\]]*\]', 'specifier_note'),
    (r'\[SPECIFIER[:\s][^\]]*\]', 'specifier_note'),
    (r'(?i)\*\*\s*note to specifier\s*\*\*[^\n]*(?:\n(?!\n)[^\n]*)*', 'specifier_note'),
    (r'(?i)<<\s*note to specifier[^>]*>>', 'specifier_note'),
    (r'(?i)^\s*note to specifier:.*$', 'specifier_note'),

    # MasterSpec / AIA / ARCOM editorial instructions
    (r'(?i)^Retain or delete this article.*$', 'masterspec_instruction'),
    (r'(?i)^Retain [^\n]*paragraph[^\n]*below.*$', 'masterspec_instruction'),
    (r'(?i)^Retain [^\n]*subparagraph[^\n]*below.*$', 'masterspec_instruction'),
    (r'(?i)^Retain [^\n]*article[^\n]*below.*$', 'masterspec_instruction'),
    (r'(?i)^Retain [^\n]*section[^\n]*below.*$', 'masterspec_instruction'),
    (r'(?i)^Retain [^\n]*if .*$', 'masterspec_instruction'),
    (r'(?i)^Retain one of.*$', 'masterspec_instruction'),
    (r'(?i)^Retain one or more of.*$', 'masterspec_instruction'),
    (r'(?i)^Revise this Section by deleting.*$', 'masterspec_instruction'),
    (r'(?i)^Revise [^\n]*to suit [Pp]roject.*$', 'masterspec_instruction'),
    (r'(?i)^This Section uses the term.*$', 'masterspec_instruction'),
    (r'(?i)^Verify that Section titles.*$', 'masterspec_instruction'),
    (r'(?i)^Coordinate [^\n]*paragraph[^\n]* with.*$', 'masterspec_instruction'),
    (r'(?i)^Coordinate [^\n]*revision[^\n]* with.*$', 'masterspec_instruction'),
    (r'(?i)^The list below matches.*$', 'masterspec_instruction'),
    (r'(?i)^See [^\n]*Evaluations?[^\n]* for .*$', 'masterspec_instruction'),
    (r'(?i)^See [^\n]*Article[^\n]* in the Evaluations.*$', 'masterspec_instruction'),
    (r'(?i)^If retaining [^\n]*paragraph.*$', 'masterspec_instruction'),
    (r'(?i)^If retaining [^\n]*subparagraph.*$', 'masterspec_instruction'),
    (r'(?i)^If retaining [^\n]*article.*$', 'masterspec_instruction'),
    (r'(?i)^When [^\n]*characteristics are important.*$', 'masterspec_instruction'),
    (r'(?i)^Inspections in this article are.*$', 'masterspec_instruction'),
    (r'(?i)^Materials and thicknesses in schedules below.*$', 'masterspec_instruction'),
    (r'(?i)^Insulation materials and thicknesses are identified below.*$', 'masterspec_instruction'),
    (r'(?i)^Do not duplicate requirements.*$', 'masterspec_instruction'),
    (r'(?i)^Not all materials and thicknesses may be suitable.*$', 'masterspec_instruction'),
    (r'(?i)^Consider the exposure of installed insulation.*$', 'masterspec_instruction'),
    (r'(?i)^Flexible elastomeric and polyolefin thicknesses are limited.*$', 'masterspec_instruction'),
    (r'(?i)^To comply with ASHRAE.*insulation should have.*$', 'masterspec_instruction'),
    (r'(?i)^Architect should be prepared to reject.*$', 'masterspec_instruction'),

    # Copyright notices
    (r'(?i)^Copyright\s*©?\s*\d{4}.*$', 'copyright'),
    (r'(?i)^©\s*\d{4}.*$', 'copyright'),
    (r'(?i)^Exclusively published and distributed by.*$', 'copyright'),
    (r'(?i)all rights reserved.*$', 'copyright'),
    (r'(?i)proprietary\s+information.*$', 'copyright'),

    # Separator lines
    (r'^[\*]{4,}\s*$', 'separator'),
    (r'^[-]{4,}\s*$', 'separator'),
    (r'^[=]{4,}\s*$', 'separator'),

    # Page artifacts
    (r'(?i)^page\s+\d+\s*(?:of\s*\d+)?\s*$', 'page_number'),

    # Revision marks
    (r'(?i)\{revision[^\}]*\}', 'revision_mark'),

    # Hidden text markers
    (r'(?i)<<[^>]*hidden[^>]*>>', 'hidden_text'),
]

# Pre-compile for speed and to avoid repeated regex compilation
_BOILERPLATE_RX = [(re.compile(pat, flags=re.MULTILINE), tag) for pat, tag in BOILERPLATE_PATTERNS]

def strip_boilerplate_with_report(content: str) -> tuple[str, list[str]]:
    """
    Strip boilerplate from a paragraph string and return (cleaned_text, matched_tags).
    Placeholders are NOT stripped here (your patterns do not remove generic [ ... ] placeholders).
    """
    cleaned = content
    hits: list[str] = []

    for rx, tag in _BOILERPLATE_RX:
        if rx.search(cleaned):
            hits.append(tag)
            cleaned = rx.sub('', cleaned)

    # Clean up whitespace
    cleaned = re.sub(r'\n{3,}', '\n\n', cleaned)
    cleaned = re.sub(r'[ \t]+\n', '\n', cleaned)
    cleaned = cleaned.strip()

    # Deduplicate tags (stable order)
    if hits:
        seen = set()
        hits = [t for t in hits if not (t in seen or seen.add(t))]

    return cleaned, hits


def build_phase2_slim_bundle(extract_dir: Path, discipline: str) -> Dict[str, Any]:
    doc_path = extract_dir / "word" / "document.xml"
    doc_text = doc_path.read_text(encoding="utf-8")

    paragraphs = []
    filter_report = {
        "paragraphs_removed_entirely": [],   # [{paragraph_index, tags, original_text_preview}]
        "paragraphs_stripped": []            # [{paragraph_index, tags}]
    }

    for idx, (_s, _e, p_xml) in enumerate(iter_paragraph_xml_blocks(doc_text)):
        if paragraph_contains_sectpr(p_xml):
            continue

        raw_text = paragraph_text_from_block(p_xml)
        if not raw_text:
            continue

        cleaned_text, tags = strip_boilerplate_with_report(raw_text)

        # If boilerplate stripping makes it empty, do not send to LLM
        if not cleaned_text:
            if tags:
                filter_report["paragraphs_removed_entirely"].append({
                    "paragraph_index": idx,
                    "tags": tags,
                    "original_text_preview": raw_text[:120]
                })
            continue

        if tags:
            filter_report["paragraphs_stripped"].append({
                "paragraph_index": idx,
                "tags": tags
            })

        numpr = paragraph_numpr_from_block(p_xml)

        paragraphs.append({
            "paragraph_index": idx,
            "text": cleaned_text[:200],
            "numPr": numpr if (numpr.get("numId") or numpr.get("ilvl")) else None,
            "contains_sectPr": False
        })

    return {
        "document_meta": {
            "discipline": discipline
        },
        "filter_report": filter_report,
        "paragraphs": paragraphs
    }


def _collect_style_deps_from_arch(arch_styles_text: str, style_id: str, seen: Set[str]) -> None:
    """
    Recursively collect styleId dependencies via <w:basedOn w:val="..."/>.
    """
    if style_id in seen:
        return
    seen.add(style_id)

    blk = extract_style_block_raw(arch_styles_text, style_id)
    if not blk:
        return

    m = re.search(r'<w:basedOn\b[^>]*w:val="([^"]+)"', blk)
    if m:
        base = m.group(1)
        if base and base not in seen:
            _collect_style_deps_from_arch(arch_styles_text, base, seen)


def extract_style_block_raw(styles_xml_text: str, style_id: str) -> Optional[str]:
    """
    Extract the raw <w:style ...>...</w:style> block for a given styleId using regex.
    This avoids ET rewriting / reformatting.
    """
    # styleId can include characters that need escaping in regex
    sid = re.escape(style_id)
    m = re.search(rf'(<w:style\b[^>]*w:styleId="{sid}"[^>]*>[\s\S]*?</w:style>)', styles_xml_text)
    return m.group(1) + "\n" if m else None


def import_arch_styles_into_target(
    target_extract_dir: Path,
    arch_extract_dir: Path,
    needed_style_ids: List[str],
    log: List[str]
) -> None:
    """
    Copy specific style blocks from architect styles.xml into target styles.xml (idempotent),
    including basedOn dependencies.
    """
    arch_extract_dir = resolve_arch_extract_root(arch_extract_dir)

    arch_styles_path = arch_extract_dir / "word" / "styles.xml"
    tgt_styles_path = target_extract_dir / "word" / "styles.xml"

    arch_styles_text = arch_styles_path.read_text(encoding="utf-8")
    tgt_styles_text = tgt_styles_path.read_text(encoding="utf-8")

    existing = set(re.findall(r'w:styleId="([^"]+)"', tgt_styles_text))

    # Expand basedOn deps
    expanded: Set[str] = set()
    for sid in needed_style_ids:
        _collect_style_deps_from_arch(arch_styles_text, sid, expanded)

    blocks: List[str] = []
    missing: List[str] = []
    for sid in sorted(expanded):
        if sid in existing:
            continue

        blk = extract_style_block_raw(arch_styles_text, sid)
        if not blk:
            missing.append(sid)
            continue

        # Guardrail: paragraph styles must not carry numbering properties.
        # If present (some templates do this), strip it to avoid clobbering Word list runtime.
        if "<w:numPr" in blk:
            log.append(f"WARNING: Stripped <w:numPr> from imported style: {sid}")
            blk = re.sub(r"<w:numPr\b[^>]*>[\s\S]*?</w:numPr>", "", blk, flags=re.S)

        # HARDEN: make style self-contained (pPr/rPr) to prevent font drift
        blk = materialize_arch_style_block(blk, sid, arch_styles_text)

        blocks.append(blk)


        log.append(f"Imported style from architect: {sid}")

    # Priority-1 hardening: if the architect template is missing any required style or dependency,
    # fail fast rather than emitting a partially formatted output.
    if missing:
        missing_sorted = ", ".join(sorted(set(missing)))
        raise ValueError(
            "Architect styles.xml is missing required styleIds needed for Phase 2 import: "
            f"{missing_sorted}"
        )

    if not blocks:
        return

    tgt_new = insert_styles_into_styles_xml(tgt_styles_text, blocks)
    if tgt_new != tgt_styles_text:
        tgt_styles_path.write_text(tgt_new, encoding="utf-8")


def insert_styles_into_styles_xml(styles_xml_text: str, style_blocks: List[str]) -> str:
    if not style_blocks:
        return styles_xml_text

    # Idempotence: skip inserting styles that already exist in styles.xml
    existing = set(re.findall(r'w:styleId="([^"]+)"', styles_xml_text))
    filtered: List[str] = []
    for sb in style_blocks:
        m = re.search(r'w:styleId="([^"]+)"', sb)
        if not m:
            raise ValueError("Style block missing w:styleId")
        sid = m.group(1)
        if sid in existing:
            continue
        filtered.append(sb)

    if not filtered:
        return styles_xml_text

    insert_point = styles_xml_text.rfind("</w:styles>")
    if insert_point == -1:
        raise ValueError("styles.xml does not contain </w:styles>")
    insertion = "\n" + "\n".join(filtered) + "\n"
    return styles_xml_text[:insert_point] + insertion + styles_xml_text[insert_point:]


def apply_pstyle_to_paragraph_block(p_xml: str, styleId: str) -> str:
    # refuse to touch sectPr paragraph
    if "<w:sectPr" in p_xml:
        return p_xml

    # If pStyle already exists, replace its value
    if re.search(r"<w:pStyle\b", p_xml):
        p_xml = re.sub(
            r'(<w:pStyle\b[^>]*w:val=")([^"]+)(")',
            rf'\g<1>{styleId}\g<3>',
            p_xml,
            count=1
        )
        return p_xml

    # Handle self-closing pPr: <w:pPr/> or <w:pPr />
    if re.search(r"<w:pPr\b[^>]*/>", p_xml):
        p_xml = re.sub(
            r"<w:pPr\b[^>]*/>",
            rf'<w:pPr><w:pStyle w:val="{styleId}"/></w:pPr>',
            p_xml,
            count=1
        )
        return p_xml

    # If pPr exists as a normal open/close element, insert pStyle right after opening tag
    if "<w:pPr" in p_xml:
        p_xml = re.sub(
            r'(<w:pPr\b[^>]*>)',
            rf'\1<w:pStyle w:val="{styleId}"/>',
            p_xml,
            count=1
        )
        return p_xml

    # No pPr at all: create one right after <w:p ...>
    p_xml = re.sub(
        r'(<w:p\b[^>]*>)',
        rf'\1<w:pPr><w:pStyle w:val="{styleId}"/></w:pPr>',
        p_xml,
        count=1
    )
    return p_xml


def write_phase2_preflight(
    extract_dir: Path,
    arch_root: Path,
    arch_registry: Dict[str, str],
    classifications: Dict[str, Any],
    out_path: Path
) -> Dict[str, Any]:
    # Count classifications per role
    role_counts: Dict[str, int] = {}
    for item in classifications.get("classifications", []):
        r = item.get("csi_role")
        if isinstance(r, str):
            role_counts[r] = role_counts.get(r, 0) + 1

    # Identify which roles are unmapped
    needed_roles = sorted(role_counts.keys())
    unmapped_roles = [r for r in needed_roles if r not in arch_registry]

    report = {
        "arch_extract_root": str(arch_root),
        "target_extract_root": str(extract_dir),
        "roles_in_classifications": role_counts,
        "arch_style_registry": arch_registry,
        "unmapped_roles": unmapped_roles,
    }

    out_path.write_text(json.dumps(report, indent=2), encoding="utf-8")
    return report


def sanitize_style_def(sd: Dict[str, Any]) -> Dict[str, Any]:
    # Option-2 lock: styles must NOT define paragraph properties
    clean = dict(sd)
    clean.pop("pPr", None)   # REMOVE paragraph formatting
    return clean


def snapshot_doc_rels_hash(extract_dir: Path) -> str:
    rels_path = extract_dir / "word" / "_rels" / "document.xml.rels"
    if not rels_path.exists():
        return ""
    return sha256_bytes(rels_path.read_bytes())

def ppr_without_pstyle(p_xml: str) -> str:
    """
    Extract paragraph properties excluding pStyle.
    Used to assert no visual drift.
    """
    m = re.search(r"<w:pPr\b[\s\S]*?</w:pPr>", p_xml)
    if not m:
        return ""
    ppr = m.group(0)
    # remove pStyle only
    ppr = re.sub(r"<w:pStyle\b[^>]*/>", "", ppr)
    return ppr


if __name__ == "__main__":
    main()

