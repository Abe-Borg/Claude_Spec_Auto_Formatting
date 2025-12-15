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
    
    def _add_directory_tree(self):
        """Add directory tree structure."""
        self.markdown_report.append("## Directory Structure\n")
        self.markdown_report.append("```")
        self._print_tree(self.extract_dir, prefix="")
        self.markdown_report.append("```\n")
    
    def _print_tree(self, directory, prefix="", is_last=True):
        """Recursively print directory tree."""
        items = sorted(directory.iterdir(), key=lambda x: (not x.is_dir(), x.name))
        
        for i, item in enumerate(items):
            is_last_item = (i == len(items) - 1)
            current_prefix = "└── " if is_last_item else "├── "
            self.markdown_report.append(f"{prefix}{current_prefix}{item.name}")
            
            if item.is_dir():
                extension = "    " if is_last_item else "│   "
                self._print_tree(item, prefix + extension, is_last_item)
    
    
    def _parse_xml_with_namespaces(self, file_path):
        """Parse XML and return tree with namespace mapping."""
        tree = ET.parse(file_path)
        root = tree.getroot()
        
        # Extract all namespaces
        namespaces = {}
        for event, elem in ET.iterparse(file_path, events=['start-ns']):
            prefix, uri = elem
            if prefix:
                namespaces[prefix] = uri
            else:
                namespaces['default'] = uri
        
        return tree, root, namespaces
    
    def _element_to_dict(self, element, namespaces):
        """Convert XML element to detailed dict representation."""
        result = {
            'tag': element.tag,
            'attributes': dict(element.attrib),
            'text': element.text.strip() if element.text and element.text.strip() else None,
            'tail': element.tail.strip() if element.tail and element.tail.strip() else None,
            'children': []
        }
        
        for child in element:
            result['children'].append(self._element_to_dict(child, namespaces))
        
        return result
    
   
    
    def _add_all_relationships(self):
        """COMPLETE analysis of ALL relationship files."""
        self.markdown_report.append("## Relationships - COMPLETE ANALYSIS\n")
        
        # Find all .rels files
        rels_files = list(self.extract_dir.rglob('*.rels'))
        
        for rels_file in sorted(rels_files):
            rel_path = rels_file.relative_to(self.extract_dir)
            self.markdown_report.append(f"### `{rel_path}`\n")
            
            try:
                tree, root, namespaces = self._parse_xml_with_namespaces(rels_file)
                
                self.markdown_report.append(f"**File Size:** {rels_file.stat().st_size:,} bytes")
                self.markdown_report.append(f"**Namespaces:** {namespaces}")
                self.markdown_report.append("")
                
                # Remove namespace for easier parsing
                for elem in root.iter():
                    if '}' in elem.tag:
                        elem.tag = elem.tag.split('}', 1)[1]
                
                relationships = root.findall('.//Relationship')
                
                self.markdown_report.append(f"**Total Relationships:** {len(relationships)}\n")
                
                for i, rel in enumerate(relationships, 1):
                    rel_id = rel.get('Id')
                    rel_type = rel.get('Type')
                    target = rel.get('Target')
                    target_mode = rel.get('TargetMode', 'Internal')
                    
                    self.markdown_report.append(f"{i}. **Relationship ID:** `{rel_id}`")
                    self.markdown_report.append(f"   - **Type:** `{rel_type}`")
                    self.markdown_report.append(f"   - **Target:** `{target}`")
                    self.markdown_report.append(f"   - **Target Mode:** `{target_mode}`")
                    self.markdown_report.append("")
            
            except Exception as e:
                self.markdown_report.append(f"Error parsing: {e}\n")
    

    
    def _document_element_recursive(self, element, depth, max_depth=5):
        """Recursively document an XML element and its children."""
        if depth > max_depth:
            return
        
        indent = "  " * depth
        tag_name = element.tag.split('}')[-1] if '}' in element.tag else element.tag
        attrs = ', '.join([f"{k.split('}')[-1]}={v}" for k, v in element.attrib.items()])
        
        self.markdown_report.append(f"{indent}- **{tag_name}** {f'({attrs})' if attrs else ''}")
        
        if element.text and element.text.strip():
            self.markdown_report.append(f"{indent}  - Text: `{element.text.strip()[:100]}`")
        
        for child in element:
            self._document_element_recursive(child, depth + 1, max_depth)
    
    
    
    
    
    
    def _add_binary_files(self):
        """Analyze binary files (images, etc.)."""
        self.markdown_report.append("## Binary Files Analysis\n")
        
        binary_extensions = {'.jpeg', '.jpg', '.png', '.gif', '.bmp', '.tiff', '.emf', '.wmf'}
        all_files = list(self.extract_dir.rglob('*'))
        binary_files = [f for f in all_files if f.is_file() and f.suffix.lower() in binary_extensions]
        
        if not binary_files:
            self.markdown_report.append("No binary files found.\n")
            return
        
        for bin_file in sorted(binary_files):
            rel_path = bin_file.relative_to(self.extract_dir)
            size = bin_file.stat().st_size
            
            self.markdown_report.append(f"### `{rel_path}`")
            self.markdown_report.append(f"- **Type:** {bin_file.suffix.upper()}")
            self.markdown_report.append(f"- **Size:** {size:,} bytes ({size/1024:.2f} KB)")
            
            # Read file signature (magic bytes)
            with open(bin_file, 'rb') as f:
                magic = f.read(16)
                hex_magic = ' '.join([f'{b:02x}' for b in magic])
                self.markdown_report.append(f"- **Magic Bytes:** `{hex_magic}`")
            
            self.markdown_report.append("")
    
    def _add_raw_xml_dumps(self):
        """Add complete raw XML dumps for all XML files."""
        self.markdown_report.append("## RAW XML DUMPS\n")
        self.markdown_report.append("Complete, unprocessed XML content for every XML file.\n")
        
        all_xml = sorted(self.extract_dir.rglob('*.xml'))
        
        for xml_file in all_xml:
            rel_path = xml_file.relative_to(self.extract_dir)
            self.markdown_report.append(f"### `{rel_path}` - RAW XML\n")
            
            try:
                with open(xml_file, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                self.markdown_report.append("```xml")
                self.markdown_report.append(content)
                self.markdown_report.append("```\n")
            
            except Exception as e:
                self.markdown_report.append(f"Error reading file: {e}\n")
    
    def save_analysis(self, output_path=None):
        """
        Save the markdown analysis to a file.
        
        Args:
            output_path: Path to save the markdown file. If None, uses default name.
        
        Returns:
            Path to the saved markdown file
        """
        if not self.markdown_report:
            self.analyze_structure()
        
        if output_path is None:
            output_path = self.extract_dir.parent / f"{self.extract_dir.name}_analysis.md"
        else:
            output_path = Path(output_path)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write("\n".join(self.markdown_report))
        
        print(f"Analysis saved to: {output_path}")
        return output_path
    
    def reconstruct(self, output_path=None):
        """
        Reconstruct the .docx file from the extracted components.
        
        Args:
            output_path: Path for the reconstructed .docx file. If None, uses default name.
        
        Returns:
            Path to the reconstructed .docx file
        """
        if self.extract_dir is None:
            raise ValueError("Must call extract() before reconstruct()")
        
        if output_path is None:
            output_path = self.extract_dir.parent / f"{self.extract_dir.name}_reconstructed.docx"
        else:
            output_path = Path(output_path)
        
        print(f"Reconstructing document from {self.extract_dir}...")
        
        # Create a new ZIP file
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as docx:
            for file_path in self.extract_dir.rglob("*"):
                if not file_path.is_file():
                    continue

                rel_path = file_path.relative_to(self.extract_dir)

                # Only include legitimate DOCX parts
                if not _is_docx_package_part(rel_path):
                    continue

                # Force forward slashes inside the ZIP (Word expects this)
                docx.write(file_path, arcname=rel_path.as_posix())


        print(f"Reconstruction complete: {output_path}")
        return output_path


    def write_normalize_bundle(self, bundle_path=None, prompts_dir=None):
        """
        Create a focused bundle for LLM editing and write prompts to disk.
        Produces:
          - bundle.json (editable + read_only xml + sha256)
          - prompts/master_prompt.txt
          - prompts/run_instruction.txt
        """
        if self.extract_dir is None:
            raise ValueError("Must call extract() before write_normalize_bundle()")

        if bundle_path is None:
            bundle_path = self.extract_dir / "bundle.json"
        else:
            bundle_path = Path(bundle_path)

        if prompts_dir is None:
            prompts_dir = self.extract_dir / "prompts"
        else:
            prompts_dir = Path(prompts_dir)

        prompts_dir.mkdir(parents=True, exist_ok=True)

        bundle = build_llm_bundle(self.extract_dir)

        bundle_path.write_text(json.dumps(bundle, indent=2), encoding="utf-8")
        (prompts_dir / "master_prompt.txt").write_text(MASTER_PROMPT, encoding="utf-8")
        (prompts_dir / "run_instruction.txt").write_text(RUN_INSTRUCTION_DEFAULT, encoding="utf-8")

        print(f"Normalize bundle written: {bundle_path}")
        print(f"Prompts written in: {prompts_dir}")
        return bundle_path, prompts_dir

    def apply_edits_and_rebuild(self, edits_json_path, output_docx_path=None):
        """
        Apply LLM edits, verify stability (headers/footers + sectPr), then rebuild docx.
        Emits diffs under extract_dir/patches by default.
        """
        if self.extract_dir is None:
            raise ValueError("Must call extract() before apply_edits_and_rebuild()")

        edits_json_path = Path(edits_json_path)
        if not edits_json_path.exists():
            raise FileNotFoundError(f"Edits JSON not found: {edits_json_path}")

        # Take stability snapshot BEFORE applying edits
        snap = snapshot_stability(self.extract_dir)

        # Load edits JSON
        edits = json.loads(edits_json_path.read_text(encoding="utf-8"))

        # Apply edits + write diffs
        patches_dir = self.extract_dir / "patches"
        apply_llm_edits(self.extract_dir, edits, patches_dir)

        # Verify stability AFTER applying edits
        verify_stability(self.extract_dir, snap)

        # Rebuild docx
        return self.reconstruct(output_path=output_docx_path)

    def write_slim_normalize_bundle(self, output_path=None):
        if self.extract_dir is None:
            raise ValueError("Must call extract() before write_slim_normalize_bundle()")

        if output_path is None:
            output_path = self.extract_dir / "slim_bundle.json"
        else:
            output_path = Path(output_path)

        prompts_dir = self.extract_dir / "prompts_slim"
        prompts_dir.mkdir(parents=True, exist_ok=True)

        bundle = build_slim_bundle(self.extract_dir)
        output_path.write_text(json.dumps(bundle, indent=2), encoding="utf-8")

        (prompts_dir / "master_prompt.txt").write_text(SLIM_MASTER_PROMPT, encoding="utf-8")
        (prompts_dir / "run_instruction.txt").write_text(SLIM_RUN_INSTRUCTION_DEFAULT, encoding="utf-8")

        print(f"Slim bundle written: {output_path}")
        print(f"Slim prompts written: {prompts_dir}")
        return output_path, prompts_dir

    def apply_instructions_and_rebuild(self, instructions_json_path, output_docx_path=None):
        if self.extract_dir is None:
            raise ValueError("Must call extract() before apply_instructions_and_rebuild()")

        instructions_json_path = Path(instructions_json_path)
        instructions = json.loads(instructions_json_path.read_text(encoding="utf-8"))

        apply_instructions(self.extract_dir, instructions)
        return self.reconstruct(output_path=output_docx_path)



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

    # Legacy args still accepted but *disabled*
    parser.add_argument("--normalize", action="store_true", help="(LEGACY) disabled")
    parser.add_argument("--apply-edits", default=None, help="(LEGACY) disabled")
    parser.add_argument("--normalize-slim", action="store_true", help="(LEGACY) disabled")
    parser.add_argument("--apply-instructions", default=None, help="(LEGACY) disabled")

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

    analysis_path = None
    if args.write_analysis and not (args.phase2_arch_extract or args.phase2_build_bundle):
        analysis_path = decomposer.save_analysis()

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

        # Optional: your separate invariants module (if you created it)
        # (This is the "no run-level <w:rPr> edits" guard, etc.)
        try:
            from phase2_invariants import verify_phase2_invariants
            new_doc_xml_bytes = (extract_dir / "word" / "document.xml").read_bytes()
            verify_phase2_invariants(src_docx=input_docx_path, new_document_xml=new_doc_xml_bytes)
        except ModuleNotFoundError:
            pass

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


ALLOWED_EDIT_PATHS = {
    "word/document.xml",
    "word/styles.xml",
    "word/numbering.xml",
    # NOTE: headers/footers are intentionally NOT allowed to be edited by LLM.
}

def apply_llm_edits(extract_dir: Path, edits_json: dict, diff_dir: Path) -> None:
    diff_dir.mkdir(parents=True, exist_ok=True)

    edits = edits_json.get("edits", [])
    if not isinstance(edits, list) or not edits:
        raise ValueError("No edits found in LLM response JSON.")

    for edit in edits:
        rel_path = edit["path"].replace("\\", "/")
        if rel_path not in ALLOWED_EDIT_PATHS:
            raise ValueError(f"Edit path not allowed: {rel_path}")

        target = extract_dir / rel_path
        if not target.exists():
            raise FileNotFoundError(f"Target file not found in package: {rel_path}")

        before = target.read_text(encoding="utf-8")
        after = edit["content"]

        sha_before = edit.get("sha256_before")
        if sha_before and sha256_text(before) != sha_before:
            raise ValueError(f"sha256_before mismatch for {rel_path}")

        # Write replacement exactly as provided
        target.write_text(after, encoding="utf-8")

        # Emit diff
        diff = difflib.unified_diff(
            before.splitlines(keepends=True),
            after.splitlines(keepends=True),
            fromfile=f"a/{rel_path}",
            tofile=f"b/{rel_path}",
        )
        diff_path = diff_dir / (rel_path.replace("/", "__") + ".diff")
        diff_path.write_text("".join(diff), encoding="utf-8")

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

def materialize_arch_style_block(style_block: str, style_id: str, arch_styles_xml_text: str) -> str:
    # Inject rPr only if missing entirely
    if "<w:rPr" not in style_block:
        eff = _effective_rpr_inner_in_arch(arch_styles_xml_text, style_id)
        if eff.strip():
            style_block = style_block.replace(
                "</w:style>",
                f"\n  <w:rPr>{eff}</w:rPr>\n</w:style>"
            )

    # Inject pPr only if missing entirely
    if "<w:pPr" not in style_block:
        effp = _effective_ppr_inner_in_arch(arch_styles_xml_text, style_id)
        if effp.strip():
            style_block = style_block.replace(
                "</w:style>",
                f"\n  <w:pPr>{effp}</w:pPr>\n</w:style>"
            )

    return style_block




def build_llm_bundle(extract_dir: Path) -> dict:
    paths = [
        "word/document.xml",
        "word/styles.xml",
        "word/numbering.xml",
    ]

    # Include headers/footers for analysis ONLY (not editable)
    hf_paths = []
    wf = extract_dir / "word"
    for p in sorted(wf.glob("header*.xml")) + sorted(wf.glob("footer*.xml")):
        hf_paths.append(str(p.relative_to(extract_dir)).replace("\\", "/"))

    payload = {"editable": {}, "read_only": {}}

    for rel in paths:
        p = extract_dir / rel
        if p.exists():
            txt = p.read_text(encoding="utf-8")
            payload["editable"][rel] = {
                "sha256": sha256_text(txt),
                "content": txt
            }

    for rel in hf_paths:
        p = extract_dir / rel
        txt = p.read_text(encoding="utf-8")
        payload["read_only"][rel] = {
            "sha256": sha256_text(txt),
            "content": txt
        }

    # settings.xml is optional; include read-only unless you are debugging layout
    settings_rel = "word/settings.xml"
    sp = extract_dir / settings_rel
    if sp.exists():
        txt = sp.read_text(encoding="utf-8")
        payload["read_only"][settings_rel] = {
            "sha256": sha256_text(txt),
            "content": txt
        }

    return payload



W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def _get_attr(elem, local_name: str) -> Optional[str]:
    # WordprocessingML attributes use w:val etc (namespaced). ET shows {ns}val.
    return elem.get(f"{{{W_NS}}}{local_name}")

def _q(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"

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

def build_style_catalog(styles_xml_path: Path, used_style_ids: Set[str]) -> Dict[str, Any]:
    # Parse styles.xml and extract compact info only for used styles + inheritance chain
    tree = ET.parse(styles_xml_path)
    root = tree.getroot()

    # index by styleId
    styles_by_id: Dict[str, ET.Element] = {}
    for st in root.findall(f".//{_q('style')}"):
        sid = _get_attr(st, "styleId")
        if sid:
            styles_by_id[sid] = st

    # expand basedOn chain
    to_include = set(used_style_ids)
    changed = True
    while changed:
        changed = False
        for sid in list(to_include):
            st = styles_by_id.get(sid)
            if st is None:
                continue
            based = st.find(_q("basedOn"))
            if based is not None:
                base_id = _get_attr(based, "val")
                if base_id and base_id not in to_include:
                    to_include.add(base_id)
                    changed = True

    def extract_pr(st: ET.Element) -> Dict[str, Any]:
        out: Dict[str, Any] = {"pPr": {}, "rPr": {}}
        pPr = st.find(_q("pPr"))
        rPr = st.find(_q("rPr"))

        if pPr is not None:
            jc = pPr.find(_q("jc"))
            if jc is not None:
                out["pPr"]["jc"] = _get_attr(jc, "val")
            spacing = pPr.find(_q("spacing"))
            if spacing is not None:
                out["pPr"]["spacing"] = {k: spacing.get(f"{{{W_NS}}}{k}") for k in ["before","after","line"] if spacing.get(f"{{{W_NS}}}{k}") is not None}
            ind = pPr.find(_q("ind"))
            if ind is not None:
                out["pPr"]["ind"] = {k: ind.get(f"{{{W_NS}}}{k}") for k in ["left","right","firstLine","hanging"] if ind.get(f"{{{W_NS}}}{k}") is not None}

        if rPr is not None:
            rFonts = rPr.find(_q("rFonts"))
            if rFonts is not None:
                out["rPr"]["rFonts"] = {k: rFonts.get(f"{{{W_NS}}}{k}") for k in ["ascii","hAnsi","cs"] if rFonts.get(f"{{{W_NS}}}{k}") is not None}
            sz = rPr.find(_q("sz"))
            if sz is not None:
                out["rPr"]["sz"] = _get_attr(sz, "val")
            b = rPr.find(_q("b"))
            if b is not None:
                out["rPr"]["b"] = True
            i = rPr.find(_q("i"))
            if i is not None:
                out["rPr"]["i"] = True
            u = rPr.find(_q("u"))
            if u is not None:
                out["rPr"]["u"] = _get_attr(u, "val") or True
            color = rPr.find(_q("color"))
            if color is not None:
                out["rPr"]["color"] = _get_attr(color, "val")

        return out

    catalog: Dict[str, Any] = {}
    for sid in sorted(to_include):
        st = styles_by_id.get(sid)
        if st is None:
            continue
        name_el = st.find(_q("name"))
        based_el = st.find(_q("basedOn"))
        st_type = _get_attr(st, "type")
        catalog[sid] = {
            "styleId": sid,
            "type": st_type,
            "name": _get_attr(name_el, "val") if name_el is not None else None,
            "basedOn": _get_attr(based_el, "val") if based_el is not None else None,
            **extract_pr(st),
        }
    return catalog

def build_numbering_catalog(numbering_xml_path: Path, used_num_ids: Set[str]) -> Dict[str, Any]:
    if not numbering_xml_path.exists():
        return {}

    tree = ET.parse(numbering_xml_path)
    root = tree.getroot()

    # map numId -> abstractNumId
    num_map: Dict[str, str] = {}
    for num in root.findall(f".//{_q('num')}"):
        numId = _get_attr(num, "numId")
        abs_el = num.find(_q("abstractNumId"))
        if numId and abs_el is not None:
            absId = _get_attr(abs_el, "val")
            if absId:
                num_map[numId] = absId

    abs_needed = {num_map[n] for n in used_num_ids if n in num_map}

    # extract abstractNum level patterns
    abstracts: Dict[str, Any] = {}
    for absn in root.findall(f".//{_q('abstractNum')}"):
        absId = _get_attr(absn, "abstractNumId")
        if not absId or absId not in abs_needed:
            continue
        lvls = []
        for lvl in absn.findall(_q("lvl")):
            ilvl = _get_attr(lvl, "ilvl")
            numFmt = lvl.find(_q("numFmt"))
            lvlText = lvl.find(_q("lvlText"))
            pPr = lvl.find(_q("pPr"))
            lvl_entry = {
                "ilvl": ilvl,
                "numFmt": _get_attr(numFmt, "val") if numFmt is not None else None,
                "lvlText": _get_attr(lvlText, "val") if lvlText is not None else None,
                "pPr": {}
            }
            if pPr is not None:
                ind = pPr.find(_q("ind"))
                if ind is not None:
                    lvl_entry["pPr"]["ind"] = {k: ind.get(f"{{{W_NS}}}{k}") for k in ["left","hanging","firstLine"] if ind.get(f"{{{W_NS}}}{k}") is not None}
                jc = pPr.find(_q("jc"))
                if jc is not None:
                    lvl_entry["pPr"]["jc"] = _get_attr(jc, "val")
            lvls.append(lvl_entry)
        abstracts[absId] = {"abstractNumId": absId, "levels": lvls}

    nums: Dict[str, Any] = {}
    for numId in sorted(used_num_ids):
        absId = num_map.get(numId)
        nums[numId] = {"numId": numId, "abstractNumId": absId}

    return {"nums": nums, "abstracts": abstracts}

def build_slim_bundle(extract_dir: Path) -> Dict[str, Any]:
    # Stability hashes
    snap = snapshot_stability(extract_dir)

    doc_path = extract_dir / "word" / "document.xml"
    doc_text = doc_path.read_text(encoding="utf-8")

    paragraphs = []
    used_style_ids: Set[str] = set()
    used_num_ids: Set[str] = set()

    for idx, (_s, _e, p_xml) in enumerate(iter_paragraph_xml_blocks(doc_text)):
        txt = paragraph_text_from_block(p_xml)
        pStyle = paragraph_pstyle_from_block(p_xml)
        numpr = paragraph_numpr_from_block(p_xml)
        hints = paragraph_ppr_hints_from_block(p_xml)
        contains_sect = paragraph_contains_sectpr(p_xml)

        if pStyle:
            used_style_ids.add(pStyle)
        if numpr.get("numId"):
            used_num_ids.add(numpr["numId"])

        # Keep summary compact: cap text length
        if len(txt) > 200:
            txt = txt[:200] + "…"

        paragraphs.append({
            "paragraph_index": idx,
            "text": txt,
            "pStyle": pStyle,
            "numPr": numpr if (numpr.get("numId") or numpr.get("ilvl")) else None,
            "pPr_hints": hints if hints else None,
            "contains_sectPr": contains_sect
        })

    styles_path = extract_dir / "word" / "styles.xml"
    style_catalog = build_style_catalog(styles_path, used_style_ids) if styles_path.exists() else {}

    numbering_path = extract_dir / "word" / "numbering.xml"
    numbering_catalog = build_numbering_catalog(numbering_path, used_num_ids)

    return {
        "stability": {
            "header_footer_hashes": snap.header_footer_hashes,
            "sectPr_hash": snap.sectpr_hash
        },
        "paragraphs": paragraphs,
        "style_catalog": style_catalog,
        "numbering_catalog": numbering_catalog
    }


def apply_phase2_classifications(
    extract_dir: Path,
    classifications: Dict[str, Any],
    arch_style_registry: Dict[str, str],
    log: List[str]
) -> None:
    doc_path = extract_dir / "word" / "document.xml"
    styles_xml_text = (extract_dir / "word" / "styles.xml").read_text(encoding="utf-8")
    doc_text = doc_path.read_text(encoding="utf-8")

    # Load styles once so we can preserve style-linked numbering before swapping styles
    styles_xml_text = (extract_dir / "word" / "styles.xml").read_text(encoding="utf-8")

    blocks = list(iter_paragraph_xml_blocks(doc_text))
    para_blocks = [b[2] for b in blocks]

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

        style_id = arch_style_registry.get(role)
        if not style_id:
            # Your preflight already expects SKIP / END_OF_SECTION to be unmapped :contentReference[oaicite:0]{index=0}
            log.append(f"Missing architect style for role: {role} (paragraph {idx})")
            continue

        if paragraph_contains_sectpr(para_blocks[idx]):
            log.append(f"Skipped sectPr paragraph at index {idx}")
            continue

        # KEY FIX: preserve "dynamic numbering" by materializing style-linked numPr
        pb = para_blocks[idx]
        pb = ensure_explicit_numpr_from_current_style(pb, styles_xml_text)

        # Now safely swap pStyle
        para_blocks[idx] = apply_pstyle_to_paragraph_block(pb, style_id)

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



def build_style_xml_block(style_def: Dict[str, Any]) -> str:
    """Build a <w:style> block. Formatting is supplied ONLY by local derivation."""
    sid = style_def.get("styleId")
    name = style_def.get("name") or sid
    based_on = style_def.get("basedOn")
    stype = style_def.get("type") or "paragraph"
    ppr_inner = style_def.get("pPr_inner") or ""
    rpr_inner = style_def.get("rPr_inner") or ""

    if not sid or not isinstance(sid, str):
        raise ValueError("styleId is required")
    if stype != "paragraph":
        raise ValueError("Only paragraph styles are supported")

    parts: List[str] = []
    parts.append(f'<w:style w:type="{stype}" w:styleId="{sid}">')
    parts.append(f'  <w:name w:val="{xml_escape(name)}"/>')
    if based_on:
        parts.append(f'  <w:basedOn w:val="{xml_escape(based_on)}"/>')
    parts.append('  <w:qFormat/>')

    # Paragraph properties (captured from exemplar)
    if ppr_inner.strip():
        parts.append('  <w:pPr>')
        parts.append(ppr_inner.strip())
        parts.append('  </w:pPr>')

    # Run properties (captured from exemplar)
    if rpr_inner.strip():
        parts.append('  <w:rPr>')
        parts.append(rpr_inner.strip())
        parts.append('  </w:rPr>')

    parts.append('</w:style>')
    return "\n".join(parts) + "\n"


def xml_escape(s: str) -> str:
    return (s.replace("&", "&amp;")
             .replace("<", "&lt;")
             .replace(">", "&gt;")
             .replace('"', "&quot;")
             .replace("'", "&apos;"))


def strip_pstyle_from_paragraph(p_xml: str) -> str:
    # Remove any <w:pStyle .../> tags for drift comparison
    return re.sub(r"<w:pStyle\b[^>]*/>", "", p_xml)


def extract_paragraph_ppr_inner(p_xml: str) -> str:
    """Return inner XML of <w:pPr>..</w:pPr> in a paragraph, or '' if none."""
    # Self-closing
    if re.search(r"<w:pPr\b[^>]*/>", p_xml):
        return ""
    m = re.search(r"<w:pPr\b[^>]*>(.*?)</w:pPr>", p_xml, flags=re.S)
    if not m:
        return ""
    inner = m.group(1)
    # Remove style assignment and numbering from captured style attributes
    inner = re.sub(r"<w:pStyle\b[^>]*/>", "", inner)
    inner = re.sub(r"<w:numPr\b[^>]*>.*?</w:numPr>", "", inner, flags=re.S)
    return inner.strip()


def extract_paragraph_rpr_inner(p_xml: str) -> str:
    """Return inner XML of the first meaningful <w:rPr> inside a paragraph, or '' if none."""
    # Find first run that contains a text node
    for rm in re.finditer(r"<w:r\b[^>]*>(.*?)</w:r>", p_xml, flags=re.S):
        run_inner = rm.group(1)
        if "<w:t" not in run_inner:
            continue
        m = re.search(r"<w:rPr\b[^>]*>(.*?)</w:rPr>", run_inner, flags=re.S)
        if m:
            return m.group(1).strip()
        # If the run has no rPr, keep searching
    return ""


def derive_style_def_from_paragraph(styleId: str, name: str, p_xml: str, based_on: Optional[str] = None) -> Dict[str, Any]:
    """Derive a paragraph style definition from an exemplar paragraph block."""
    ppr_inner = extract_paragraph_ppr_inner(p_xml)
    rpr_inner = extract_paragraph_rpr_inner(p_xml)
    return {
        "styleId": styleId,
        "name": name,
        "type": "paragraph",
        "basedOn": based_on,
        "pPr_inner": ppr_inner,
        "rPr_inner": rpr_inner,
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
    for sid in sorted(expanded):
        if sid in existing:
            continue

        blk = extract_style_block_raw(arch_styles_text, sid)
        if not blk:
            log.append(f"Architect styles.xml missing styleId: {sid}")
            continue

        # HARDEN: make style self-contained (pPr/rPr) to prevent font drift
        blk = materialize_arch_style_block(blk, sid, arch_styles_text)

        blocks.append(blk)


        log.append(f"Imported style from architect: {sid}")

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


def validate_instructions(instructions: Dict[str, Any]) -> None:
    allowed_keys = {"create_styles", "apply_pStyle", "notes"}
    extra = set(instructions.keys()) - allowed_keys
    if extra:
        raise ValueError(f"Invalid instruction keys: {extra}")

    # Validate create_styles (LLM must NOT provide formatting; only exemplar mapping)
    seen_style_ids = set()
    for sd in instructions.get("create_styles", []):
        if not isinstance(sd, dict):
            raise ValueError("create_styles entries must be objects")

        sid = sd.get("styleId")
        if not sid or not isinstance(sid, str):
            raise ValueError("create_styles entries must have styleId (string)")
        if sid in seen_style_ids:
            raise ValueError(f"Duplicate styleId: {sid}")
        seen_style_ids.add(sid)

        # LLM is forbidden from specifying formatting directly
        if "pPr" in sd or "rPr" in sd or "pPr_inner" in sd or "rPr_inner" in sd:
            raise ValueError(f"Style {sid}: LLM formatting fields are forbidden (pPr/rPr). Use derive_from_paragraph_index only.")

        allowed_style_fields = {"styleId", "name", "type", "derive_from_paragraph_index", "basedOn", "role"}
        extra_style_fields = set(sd.keys()) - allowed_style_fields
        if extra_style_fields:
            raise ValueError(f"Style {sid}: invalid fields: {extra_style_fields}")

        stype = sd.get("type", "paragraph")
        if stype != "paragraph":
            raise ValueError(f"Style {sid}: only paragraph styles are supported (type='paragraph').")

        src = sd.get("derive_from_paragraph_index")
        if src is None or not isinstance(src, int) or src < 0:
            raise ValueError(f"Style {sid}: derive_from_paragraph_index must be a non-negative integer.")

        if "basedOn" in sd and sd["basedOn"] is not None and not isinstance(sd["basedOn"], str):
            raise ValueError(f"Style {sid}: basedOn must be a string if provided.")

        if "name" in sd and sd["name"] is not None and not isinstance(sd["name"], str):
            raise ValueError(f"Style {sid}: name must be a string if provided.")

        if "role" in sd and sd["role"] is not None and not isinstance(sd["role"], str):
            raise ValueError(f"Style {sid}: role must be a string if provided.")

    # Validate apply_pStyle
    seen_para = set()
    for ap in instructions.get("apply_pStyle", []):
        if not isinstance(ap, dict):
            raise ValueError("apply_pStyle entries must be objects")
        idx = ap.get("paragraph_index")
        sid = ap.get("styleId")
        if not isinstance(idx, int) or idx < 0:
            raise ValueError(f"Invalid paragraph_index: {idx}")
        if not isinstance(sid, str):
            raise ValueError(f"Invalid styleId for paragraph {idx}")
        if idx in seen_para:
            raise ValueError(f"Duplicate paragraph_index: {idx}")
        seen_para.add(idx)


def apply_instructions(extract_dir: Path, instructions: Dict[str, Any]) -> None:
    validate_instructions(instructions)

    # Stability snapshot before
    snap = snapshot_stability(extract_dir)

    # Load styles.xml and document.xml
    styles_path = extract_dir / "word" / "styles.xml"
    styles_text = styles_path.read_text(encoding="utf-8")

    doc_path = extract_dir / "word" / "document.xml"
    doc_text = doc_path.read_text(encoding="utf-8")

    # Split document.xml into paragraph blocks
    blocks = list(iter_paragraph_xml_blocks(doc_text))
    para_blocks = [b[2] for b in blocks]
    original_para_blocks = list(para_blocks)

    # 1) Create derived styles (formatting captured locally from exemplar paragraphs)
    style_defs = instructions.get("create_styles") or []
    derived_blocks: List[str] = []
    for sd in style_defs:
        style_id = sd["styleId"]
        style_name = sd.get("name") or style_id
        src_idx = sd["derive_from_paragraph_index"]
        based_on = sd.get("basedOn")

        if src_idx >= len(para_blocks):
            raise ValueError(f"Style {style_id}: derive_from_paragraph_index out of range: {src_idx}")

        exemplar_p = para_blocks[src_idx]
        derived_def = derive_style_def_from_paragraph(style_id, style_name, exemplar_p, based_on=based_on)
        derived_blocks.append(build_style_xml_block(derived_def))

    styles_new = insert_styles_into_styles_xml(styles_text, derived_blocks)
    if styles_new != styles_text:
        styles_path.write_text(styles_new, encoding="utf-8")

    # 2) Apply paragraph styles by index (pStyle insertion ONLY)
    apply_list = instructions.get("apply_pStyle") or []
    idx_map: Dict[int, str] = {}
    for item in apply_list:
        idx = int(item["paragraph_index"])
        sid = item["styleId"]
        idx_map[idx] = sid

    # Capture original pPr (minus pStyle) for drift detection
    original_ppr = {i: ppr_without_pstyle(pb) for i, pb in enumerate(para_blocks)}

    # Apply
    for idx, sid in idx_map.items():
        if idx < 0 or idx >= len(para_blocks):
            raise ValueError(f"paragraph_index out of range: {idx}")
        if paragraph_contains_sectpr(para_blocks[idx]):
            raise ValueError(f"Refusing to apply style to paragraph {idx} because it contains sectPr.")
        
        pb = ensure_explicit_numpr_from_current_style(para_blocks[idx], styles_xml_text)
        para_blocks[idx] = apply_pstyle_to_paragraph_block(pb, style_id)




    # STEP 4: Full paragraph drift check (only pStyle may differ)
    for idx in idx_map.keys():
        before = strip_pstyle_from_paragraph(original_para_blocks[idx])
        after = strip_pstyle_from_paragraph(para_blocks[idx])
        if before != after:
            raise ValueError(f"Paragraph drift detected at index {idx}: changes beyond <w:pStyle>.")

    # Paragraph-level pPr drift check (only pStyle change allowed inside pPr)
    for i, pb in enumerate(para_blocks):
        if original_ppr[i] != ppr_without_pstyle(pb):
            raise ValueError(f"Paragraph properties drift detected at index {i} (beyond w:pStyle).")

    # Reassemble document.xml with updated paragraph blocks
    out_parts: List[str] = []
    last_end = 0
    for i, (s, e, _p) in enumerate(blocks):
        out_parts.append(doc_text[last_end:s])
        out_parts.append(para_blocks[i])
        last_end = e
    out_parts.append(doc_text[last_end:])
    doc_new = "".join(out_parts)
    doc_path.write_text(doc_new, encoding="utf-8")

    # Verify stability (headers/footers + sectPr + rels)
    verify_stability(extract_dir, snap)


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

