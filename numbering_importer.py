#!/usr/bin/env python3
"""
numbering_importer.py

Imports architect's numbering definitions (abstractNum + num) into target document's
numbering.xml, handling ID collisions by remapping.

This allows imported styles to reference the architect's exact numbering definitions,
preserving list number formatting (fonts, indents, prefixes).
"""

import re
import json
import random
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any
from copy import deepcopy


def _generate_unique_nsid() -> str:
    """Generate a unique nsid (8 hex chars) for abstractNum."""
    return f"{random.randint(0, 0xFFFFFFFF):08X}"


def _generate_unique_durable_id() -> str:
    """Generate a unique durableId for num."""
    return str(random.randint(1, 2147483647))


def find_max_ids_in_numbering(numbering_xml: str) -> Tuple[int, int]:
    """
    Find the maximum abstractNumId and numId in existing numbering.xml.
    Returns (max_abstract_num_id, max_num_id).
    """
    abstract_ids = [int(m) for m in re.findall(r'w:abstractNumId="(\d+)"', numbering_xml)]
    num_ids = [int(m) for m in re.findall(r'<w:num\s+w:numId="(\d+)"', numbering_xml)]
    
    max_abstract = max(abstract_ids) if abstract_ids else -1
    max_num = max(num_ids) if num_ids else 0
    
    return max_abstract, max_num


def extract_used_num_ids_from_styles(styles_xml: str) -> Dict[str, int]:
    """
    Extract which numIds are referenced by which styles.
    Returns dict of styleId -> numId.
    """
    result = {}
    # Find all style definitions with numPr
    style_pattern = r'<w:style[^>]*w:styleId="([^"]+)"[^>]*>[\s\S]*?</w:style>'
    for match in re.finditer(style_pattern, styles_xml):
        style_xml = match.group(0)
        style_id = match.group(1)
        
        # Look for numId in this style
        num_match = re.search(r'<w:numId\s+w:val="(\d+)"', style_xml)
        if num_match:
            result[style_id] = int(num_match.group(1))
    
    return result


def build_numbering_import_plan(
    arch_template_registry: Dict[str, Any],
    arch_styles_xml: str,
    target_numbering_xml: str,
    style_ids_to_import: List[str]
) -> Dict[str, Any]:
    """
    Build a plan for importing numbering definitions.
    
    Returns:
    {
        "abstract_nums_to_import": [
            {"old_id": 6, "new_id": 15, "xml": "..."},
            ...
        ],
        "nums_to_import": [
            {"old_id": 2, "new_id": 12, "old_abstract_id": 6, "new_abstract_id": 15, "xml": "..."},
            ...
        ],
        "style_numid_remap": {
            "CSILevel1": {"old_numId": 2, "new_numId": 12},
            ...
        }
    }
    """
    # Find which numIds the styles we're importing reference
    style_to_numid = extract_used_num_ids_from_styles(arch_styles_xml)
    
    # Filter to only styles we're importing
    relevant_numids = set()
    style_numid_usage = {}
    for style_id in style_ids_to_import:
        if style_id in style_to_numid:
            num_id = style_to_numid[style_id]
            relevant_numids.add(num_id)
            style_numid_usage[style_id] = num_id
    
    if not relevant_numids:
        return {
            "abstract_nums_to_import": [],
            "nums_to_import": [],
            "style_numid_remap": {}
        }
    
    # Get the numbering data from arch_template_registry
    numbering = arch_template_registry.get("numbering", {})
    abstract_nums = {an["abstractNumId"]: an for an in numbering.get("abstract_nums", [])}
    nums = {n["numId"]: n for n in numbering.get("nums", [])}
    
    # Find max IDs in target to avoid collisions
    max_abstract_id, max_num_id = find_max_ids_in_numbering(target_numbering_xml)
    
    # Build import lists
    abstract_num_id_remap = {}  # old_id -> new_id
    num_id_remap = {}  # old_id -> new_id
    
    abstract_nums_to_import = []
    nums_to_import = []
    
    # First, determine which abstractNums we need (referenced by the nums we need)
    needed_abstract_ids = set()
    for num_id in relevant_numids:
        if num_id in nums:
            needed_abstract_ids.add(nums[num_id]["abstractNumId"])
    
    # Assign new IDs to abstractNums
    next_abstract_id = max_abstract_id + 1
    for old_abstract_id in sorted(needed_abstract_ids):
        if old_abstract_id in abstract_nums:
            new_abstract_id = next_abstract_id
            abstract_num_id_remap[old_abstract_id] = new_abstract_id
            
            # Get XML and remap the abstractNumId
            xml = abstract_nums[old_abstract_id]["xml"]
            xml = re.sub(
                r'w:abstractNumId="' + str(old_abstract_id) + '"',
                f'w:abstractNumId="{new_abstract_id}"',
                xml
            )
            # Generate new nsid to avoid conflicts
            xml = re.sub(
                r'<w:nsid\s+w:val="[^"]+"/>',
                f'<w:nsid w:val="{_generate_unique_nsid()}"/>',
                xml
            )
            
            abstract_nums_to_import.append({
                "old_id": old_abstract_id,
                "new_id": new_abstract_id,
                "xml": xml
            })
            next_abstract_id += 1
    
    # Assign new IDs to nums
    next_num_id = max_num_id + 1
    for old_num_id in sorted(relevant_numids):
        if old_num_id in nums:
            new_num_id = next_num_id
            num_id_remap[old_num_id] = new_num_id
            
            num_data = nums[old_num_id]
            old_abstract_id = num_data["abstractNumId"]
            new_abstract_id = abstract_num_id_remap.get(old_abstract_id, old_abstract_id)
            
            # Get XML and remap IDs
            xml = num_data["xml"]
            xml = re.sub(
                r'w:numId="' + str(old_num_id) + '"',
                f'w:numId="{new_num_id}"',
                xml
            )
            xml = re.sub(
                r'<w:abstractNumId\s+w:val="' + str(old_abstract_id) + '"',
                f'<w:abstractNumId w:val="{new_abstract_id}"',
                xml
            )
            # Generate new durableId
            xml = re.sub(
                r'w16cid:durableId="[^"]*"',
                f'w16cid:durableId="{_generate_unique_durable_id()}"',
                xml
            )
            
            nums_to_import.append({
                "old_id": old_num_id,
                "new_id": new_num_id,
                "old_abstract_id": old_abstract_id,
                "new_abstract_id": new_abstract_id,
                "xml": xml
            })
            next_num_id += 1
    
    # Build style remap
    style_numid_remap = {}
    for style_id, old_num_id in style_numid_usage.items():
        if old_num_id in num_id_remap:
            style_numid_remap[style_id] = {
                "old_numId": old_num_id,
                "new_numId": num_id_remap[old_num_id]
            }
    
    return {
        "abstract_nums_to_import": abstract_nums_to_import,
        "nums_to_import": nums_to_import,
        "style_numid_remap": style_numid_remap
    }


def inject_font_into_numbering_rpr(xml: str, font_name: str = "Arial", font_size: str = "20") -> str:
    """
    Inject font specifications into all <w:rPr> blocks within numbering levels.
    
    This ensures list numbers render in the correct font, regardless of
    what Word's cascade would otherwise inherit.
    
    Args:
        xml: The numbering XML (abstractNum or num block)
        font_name: Font family name (default "Arial")
        font_size: Font size in half-points (default "20" = 10pt)
    
    Returns:
        XML with font specs injected into rPr blocks
    """
    font_xml = f'<w:rFonts w:ascii="{font_name}" w:hAnsi="{font_name}" w:eastAsia="{font_name}" w:cs="{font_name}"/><w:sz w:val="{font_size}"/><w:szCs w:val="{font_size}"/>'
    
    def inject_into_rpr(match):
        rpr_block = match.group(0)
        # Check if rFonts already exists
        if '<w:rFonts' in rpr_block:
            return rpr_block  # Already has font, don't override
        # Inject font specs right after <w:rPr>
        return rpr_block.replace('<w:rPr>', f'<w:rPr>{font_xml}', 1)
    
    # Find and process all rPr blocks in lvl elements
    result = re.sub(r'<w:rPr>[\s\S]*?</w:rPr>', inject_into_rpr, xml)
    
    return result


def inject_numbering_into_xml(
    target_numbering_xml: str,
    abstract_nums_to_import: List[Dict],
    nums_to_import: List[Dict]
) -> str:
    """
    Inject imported abstractNums and nums into target numbering.xml.
    
    abstractNums go before the first <w:num> element.
    nums go at the end, before </w:numbering>.
    
    Also injects font specifications into rPr blocks to ensure list numbers
    render in the correct font.
    """
    result = target_numbering_xml
    
    # Inject fonts into the numbering definitions before adding them
    abstract_nums_with_fonts = []
    for an in abstract_nums_to_import:
        xml_with_fonts = inject_font_into_numbering_rpr(an["xml"])
        abstract_nums_with_fonts.append(xml_with_fonts)
    
    nums_with_fonts = []
    for n in nums_to_import:
        xml_with_fonts = inject_font_into_numbering_rpr(n["xml"])
        nums_with_fonts.append(xml_with_fonts)
    
    # Find insertion point for abstractNums (before first <w:num>)
    first_num_match = re.search(r'<w:num\s', result)
    if first_num_match:
        insert_pos = first_num_match.start()
        abstract_xml = "\n".join(abstract_nums_with_fonts)
        if abstract_xml:
            result = result[:insert_pos] + abstract_xml + "\n" + result[insert_pos:]
    
    # Find insertion point for nums (before </w:numbering>)
    end_match = re.search(r'</w:numbering>', result)
    if end_match:
        insert_pos = end_match.start()
        num_xml = "\n".join(nums_with_fonts)
        if num_xml:
            result = result[:insert_pos] + num_xml + "\n" + result[insert_pos:]
    
    return result


def remap_numid_in_style_xml(style_xml: str, old_num_id: int, new_num_id: int) -> str:
    """
    Update a style's numPr to reference the new numId.
    """
    return re.sub(
        r'(<w:numId\s+w:val=")' + str(old_num_id) + r'"',
        f'\\g<1>{new_num_id}"',
        style_xml
    )


def import_numbering(
    arch_extract_dir: Path,
    target_extract_dir: Path,
    arch_template_registry: Dict[str, Any],
    style_ids_to_import: List[str],
    log: List[str]
) -> Dict[str, Dict[str, int]]:
    """
    Main entry point: import architect's numbering into target.
    
    Returns style_numid_remap for use when importing styles.
    """
    # Check if registry has numbering data
    if "numbering" not in arch_template_registry:
        log.append("No numbering data in arch_template_registry, skipping numbering import")
        return {}
    
    numbering_data = arch_template_registry.get("numbering", {})
    if not numbering_data.get("abstract_nums") and not numbering_data.get("nums"):
        log.append("No numbering definitions in arch_template_registry")
        return {}
    
    # Read architect's styles.xml to find numId references
    arch_styles_path = arch_extract_dir / "word" / "styles.xml"
    if not arch_styles_path.exists():
        log.append(f"WARNING: Architect styles.xml not found at {arch_styles_path}")
        return {}
    arch_styles_xml = arch_styles_path.read_text(encoding="utf-8")
    
    # Read target's numbering.xml
    target_numbering_path = target_extract_dir / "word" / "numbering.xml"
    if not target_numbering_path.exists():
        log.append("WARNING: Target has no numbering.xml - skipping numbering import")
        return {}
    target_numbering_xml = target_numbering_path.read_text(encoding="utf-8")
    
    # Build import plan
    plan = build_numbering_import_plan(
        arch_template_registry,
        arch_styles_xml,
        target_numbering_xml,
        style_ids_to_import
    )
    
    if not plan["abstract_nums_to_import"] and not plan["nums_to_import"]:
        log.append("No numbering definitions need to be imported")
        return plan.get("style_numid_remap", {})
    
    # Log what we're importing
    log.append(f"Importing {len(plan['abstract_nums_to_import'])} abstractNum definitions")
    log.append(f"Importing {len(plan['nums_to_import'])} num definitions")
    for num in plan["nums_to_import"]:
        log.append(f"  numId {num['old_id']} -> {num['new_id']} (abstractNum {num['old_abstract_id']} -> {num['new_abstract_id']})")
    
    # Inject into target numbering.xml
    new_numbering_xml = inject_numbering_into_xml(
        target_numbering_xml,
        plan["abstract_nums_to_import"],
        plan["nums_to_import"]
    )
    
    # Write updated numbering.xml
    target_numbering_path.write_text(new_numbering_xml, encoding="utf-8")
    log.append(f"Updated {target_numbering_path}")
    
    return plan["style_numid_remap"]


# =============================================================================
# UPDATED STYLE IMPORT FUNCTION
# =============================================================================
# This replaces the style import logic to KEEP numPr and remap numIds

def import_arch_style_with_numbering(
    style_xml: str,
    style_numid_remap: Dict[str, Dict[str, int]],
    style_id: str,
    log: List[str]
) -> str:
    """
    Import a style, keeping numPr but remapping numId to the imported numbering.
    
    OLD behavior: Strip numPr entirely
    NEW behavior: Keep numPr, update numId reference
    """
    if style_id in style_numid_remap:
        remap = style_numid_remap[style_id]
        old_num_id = remap["old_numId"]
        new_num_id = remap["new_numId"]
        
        # Remap the numId
        style_xml = remap_numid_in_style_xml(style_xml, old_num_id, new_num_id)
        log.append(f"  Remapped numId {old_num_id} -> {new_num_id} in style {style_id}")
    
    return style_xml