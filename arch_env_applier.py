#!/usr/bin/env python3
"""
arch_env_applier.py — Phase 2 Environment Application

Applies the formatting environment captured in arch_template_registry.json
to a target document. This ensures that imported styles render correctly
by providing the same context (theme fonts, docDefaults, etc.) they expect.

Application order (deterministic):
1. Theme (fonts/colors foundation)
2. Settings + compat flags (rendering behavior)  
3. Font table (declared fonts)
4. docDefaults (baseline rPr/pPr)
5. Styles (with materialized typography)

NOTE: This module does NOT touch:
- numbering.xml (handled separately with explicit numPr materialization)
- headers/footers (preserved from source)
- sectPr (preserved from source)

Usage:
    from arch_env_applier import apply_environment_to_target
    
    apply_environment_to_target(
        target_extract_dir=Path("mech_spec_extracted"),
        registry=loaded_registry_dict,
        log=[]
    )
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, List, Optional


# ─────────────────────────────────────────────────────────────────────────────
# docDefaults application
# ─────────────────────────────────────────────────────────────────────────────

def _extract_doc_defaults_block(styles_xml: str) -> Optional[str]:
    """Extract existing <w:docDefaults>...</w:docDefaults> block."""
    m = re.search(r'(<w:docDefaults\b[\s\S]*?</w:docDefaults>)', styles_xml)
    return m.group(1) if m else None


def _build_doc_defaults_block(
    default_rpr: Optional[str],
    default_ppr: Optional[str]
) -> str:
    """
    Build a complete <w:docDefaults> block from rPr and pPr.
    """
    parts = ["<w:docDefaults>"]
    
    if default_rpr:
        parts.append(f"  <w:rPrDefault>{default_rpr}</w:rPrDefault>")
    else:
        parts.append("  <w:rPrDefault><w:rPr/></w:rPrDefault>")
    
    if default_ppr:
        parts.append(f"  <w:pPrDefault>{default_ppr}</w:pPrDefault>")
    else:
        parts.append("  <w:pPrDefault><w:pPr/></w:pPrDefault>")
    
    parts.append("</w:docDefaults>")
    return "\n".join(parts)


def apply_doc_defaults(
    styles_xml: str,
    registry: Dict[str, Any],
    log: List[str]
) -> str:
    """
    Replace or insert docDefaults in styles.xml with values from registry.
    
    This is critical because styles inherit from docDefaults, and if the
    target document has different defaults, fonts/spacing will be wrong.
    """
    doc_defaults = registry.get("doc_defaults", {})
    
    arch_rpr = doc_defaults.get("default_run_props", {}).get("rPr")
    arch_ppr = doc_defaults.get("default_paragraph_props", {}).get("pPr")
    
    if not arch_rpr and not arch_ppr:
        log.append("No docDefaults in registry; skipping docDefaults application")
        return styles_xml
    
    new_defaults = _build_doc_defaults_block(arch_rpr, arch_ppr)
    
    existing = _extract_doc_defaults_block(styles_xml)
    if existing:
        # Replace existing docDefaults
        styles_xml = styles_xml.replace(existing, new_defaults, 1)
        log.append("Replaced existing docDefaults with architect values")
    else:
        # Insert after <w:styles ...> opening tag
        m = re.search(r'(<w:styles\b[^>]*>)', styles_xml)
        if m:
            insert_point = m.end()
            styles_xml = (
                styles_xml[:insert_point] + 
                "\n" + new_defaults + "\n" + 
                styles_xml[insert_point:]
            )
            log.append("Inserted docDefaults from architect (none existed)")
        else:
            log.append("WARNING: Could not find <w:styles> tag to insert docDefaults")
    
    return styles_xml


# ─────────────────────────────────────────────────────────────────────────────
# Theme application
# ─────────────────────────────────────────────────────────────────────────────

def apply_theme(
    target_extract_dir: Path,
    registry: Dict[str, Any],
    log: List[str]
) -> None:
    """
    Copy theme1.xml from registry to target.
    
    Theme defines majorFont/minorFont which styles reference via
    w:asciiTheme="majorHAnsi" etc. Without the correct theme,
    font resolution fails.
    """
    theme_data = registry.get("theme", {})
    theme_xml = theme_data.get("theme1_xml")
    
    if not theme_xml:
        log.append("No theme in registry; skipping theme application")
        return
    
    theme_dir = target_extract_dir / "word" / "theme"
    theme_dir.mkdir(parents=True, exist_ok=True)
    
    theme_path = theme_dir / "theme1.xml"
    
    # Check if target already has a theme
    if theme_path.exists():
        log.append("Replacing target theme1.xml with architect theme")
    else:
        log.append("Adding theme1.xml from architect (none existed)")
        # May need to update [Content_Types].xml and relationships
        _ensure_theme_in_content_types(target_extract_dir, log)
        _ensure_theme_in_rels(target_extract_dir, log)
    
    theme_path.write_text(theme_xml, encoding="utf-8")


def _ensure_theme_in_content_types(extract_dir: Path, log: List[str]) -> None:
    """Ensure [Content_Types].xml has an entry for theme1.xml."""
    ct_path = extract_dir / "[Content_Types].xml"
    if not ct_path.exists():
        return
    
    ct_xml = ct_path.read_text(encoding="utf-8")
    
    # Check if theme override already exists
    if 'PartName="/word/theme/theme1.xml"' in ct_xml:
        return
    
    # Add override for theme
    theme_override = (
        '<Override PartName="/word/theme/theme1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
    )
    
    # Insert before </Types>
    if "</Types>" in ct_xml:
        ct_xml = ct_xml.replace("</Types>", f"  {theme_override}\n</Types>")
        ct_path.write_text(ct_xml, encoding="utf-8")
        log.append("Added theme1.xml to [Content_Types].xml")


def _ensure_theme_in_rels(extract_dir: Path, log: List[str]) -> None:
    """Ensure word/_rels/document.xml.rels has a relationship for theme."""
    rels_path = extract_dir / "word" / "_rels" / "document.xml.rels"
    if not rels_path.exists():
        return
    
    rels_xml = rels_path.read_text(encoding="utf-8")
    
    # Check if theme relationship exists
    if 'Target="theme/theme1.xml"' in rels_xml:
        return
    
    # Find highest rId
    rids = re.findall(r'Id="rId(\d+)"', rels_xml)
    max_rid = max(int(r) for r in rids) if rids else 0
    new_rid = f"rId{max_rid + 1}"
    
    theme_rel = (
        f'<Relationship Id="{new_rid}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" '
        f'Target="theme/theme1.xml"/>'
    )
    
    if "</Relationships>" in rels_xml:
        rels_xml = rels_xml.replace("</Relationships>", f"  {theme_rel}\n</Relationships>")
        rels_path.write_text(rels_xml, encoding="utf-8")
        log.append(f"Added theme relationship ({new_rid}) to document.xml.rels")


# ─────────────────────────────────────────────────────────────────────────────
# Settings/compat application
# ─────────────────────────────────────────────────────────────────────────────

def apply_settings(
    target_extract_dir: Path,
    registry: Dict[str, Any],
    log: List[str]
) -> None:
    """
    Apply settings.xml from registry, focusing on compat flags.
    
    Compat flags affect rendering behavior (list spacing, line breaking, etc.)
    and can cause subtle visual differences if not matched.
    """
    settings_data = registry.get("settings", {})
    
    # For now, we focus on compat flags rather than replacing entire settings.xml
    # (full replacement could break other document-specific settings)
    
    compat_xml = settings_data.get("compat", {}).get("compat_xml")
    if not compat_xml:
        log.append("No compat flags in registry; skipping settings application")
        return
    
    settings_path = target_extract_dir / "word" / "settings.xml"
    if not settings_path.exists():
        log.append("Target has no settings.xml; skipping compat application")
        return
    
    settings_xml = settings_path.read_text(encoding="utf-8")
    
    # Find and replace existing <w:compat> block
    existing_compat = re.search(r'<w:compat\b[\s\S]*?</w:compat>', settings_xml)
    
    if existing_compat:
        settings_xml = settings_xml.replace(existing_compat.group(0), compat_xml, 1)
        log.append("Replaced compat flags with architect values")
    else:
        # Insert before </w:settings>
        if "</w:settings>" in settings_xml:
            settings_xml = settings_xml.replace(
                "</w:settings>",
                f"  {compat_xml}\n</w:settings>"
            )
            log.append("Inserted compat flags from architect")
    
    settings_path.write_text(settings_xml, encoding="utf-8")


# ─────────────────────────────────────────────────────────────────────────────
# Font table application
# ─────────────────────────────────────────────────────────────────────────────

def apply_font_table(
    target_extract_dir: Path,
    registry: Dict[str, Any],
    log: List[str]
) -> None:
    """
    Merge font declarations from registry into target fontTable.xml.
    
    This ensures fonts referenced by architect styles are declared,
    which helps Word resolve them correctly.
    """
    fonts_data = registry.get("fonts", {})
    arch_font_xml = fonts_data.get("font_table_xml")
    
    if not arch_font_xml:
        log.append("No fontTable in registry; skipping font table application")
        return
    
    font_path = target_extract_dir / "word" / "fontTable.xml"
    
    if not font_path.exists():
        # Just copy the architect's font table
        font_path.write_text(arch_font_xml, encoding="utf-8")
        log.append("Added fontTable.xml from architect")
        return
    
    # Merge: add fonts from architect that don't exist in target
    target_font_xml = font_path.read_text(encoding="utf-8")
    
    # Extract font names from both
    target_fonts = set(re.findall(r'<w:font\s+w:name="([^"]+)"', target_font_xml))
    arch_fonts = re.findall(r'(<w:font\s+w:name="([^"]+)"[\s\S]*?</w:font>)', arch_font_xml)
    
    fonts_to_add = []
    for font_block, font_name in arch_fonts:
        if font_name not in target_fonts:
            fonts_to_add.append(font_block)
    
    if not fonts_to_add:
        log.append("All architect fonts already present in target fontTable")
        return
    
    # Insert before </w:fonts>
    if "</w:fonts>" in target_font_xml:
        insertion = "\n".join(fonts_to_add)
        target_font_xml = target_font_xml.replace(
            "</w:fonts>",
            f"{insertion}\n</w:fonts>"
        )
        font_path.write_text(target_font_xml, encoding="utf-8")
        log.append(f"Added {len(fonts_to_add)} font declarations from architect")


# ─────────────────────────────────────────────────────────────────────────────
# Style materialization helpers (for styles not already in target)
# ─────────────────────────────────────────────────────────────────────────────

def resolve_effective_rpr(
    style_id: str,
    style_defs: List[Dict[str, Any]],
    doc_defaults: Dict[str, Any],
    force_tags: tuple = ("rFonts", "sz", "szCs", "lang")
) -> str:
    """
    Resolve effective rPr for a style by walking basedOn chain + docDefaults.
    Returns raw XML string with just the force tags.
    """
    # Build lookup
    by_id = {s["style_id"]: s for s in style_defs}
    
    def _extract_child(rpr_xml: Optional[str], tag: str) -> Optional[str]:
        if not rpr_xml:
            return None
        # Self-closing
        m = re.search(rf'(<w:{tag}\b[^>]*/>)', rpr_xml)
        if m:
            return m.group(1)
        # Paired
        m = re.search(rf'(<w:{tag}\b[^>]*>[\s\S]*?</w:{tag}>)', rpr_xml, re.S)
        if m:
            return m.group(1)
        return None
    
    resolved = {}
    
    for tag in force_tags:
        # Walk basedOn chain
        seen = set()
        cur = style_id
        found = None
        
        while cur and cur not in seen:
            seen.add(cur)
            style_def = by_id.get(cur)
            if not style_def:
                break
            
            rpr = style_def.get("rPr")
            node = _extract_child(rpr, tag)
            if node:
                found = node
                break
            
            cur = style_def.get("based_on")
        
        # Fall back to docDefaults
        if not found:
            default_rpr = doc_defaults.get("default_run_props", {}).get("rPr")
            found = _extract_child(default_rpr, tag)
        
        if found:
            resolved[tag] = found
    
    return "".join(resolved.values())


def materialize_style_for_import(
    style_def: Dict[str, Any],
    all_style_defs: List[Dict[str, Any]],
    doc_defaults: Dict[str, Any]
) -> str:
    """
    Build a complete <w:style> block with materialized typography.
    
    This ensures the style is self-contained enough to render correctly
    without depending on the basedOn chain or docDefaults being present.
    """
    style_id = style_def["style_id"]
    style_type = style_def.get("type", "paragraph")
    name = style_def.get("name", style_id)
    based_on = style_def.get("based_on")
    next_style = style_def.get("next")
    link = style_def.get("link")
    
    # Start building style block
    attrs = f'w:type="{style_type}" w:styleId="{style_id}"'
    if style_def.get("qformat"):
        attrs += ' w:customStyle="1"'
    
    parts = [f'<w:style {attrs}>']
    parts.append(f'  <w:name w:val="{name}"/>')
    
    if based_on:
        parts.append(f'  <w:basedOn w:val="{based_on}"/>')
    if next_style:
        parts.append(f'  <w:next w:val="{next_style}"/>')
    if link:
        parts.append(f'  <w:link w:val="{link}"/>')
    if style_def.get("ui_priority") is not None:
        parts.append(f'  <w:uiPriority w:val="{style_def["ui_priority"]}"/>')
    if style_def.get("qformat"):
        parts.append('  <w:qFormat/>')
    
    # Add pPr (strip numPr to avoid list conflicts)
    ppr = style_def.get("pPr")
    if ppr:
        # Strip numPr from pPr
        ppr = re.sub(r'<w:numPr\b[^>]*>[\s\S]*?</w:numPr>', '', ppr, flags=re.S)
        if ppr.strip():
            parts.append(f'  {ppr}')
    
    # Add rPr with materialized typography
    rpr = style_def.get("rPr") or ""
    effective_rpr = resolve_effective_rpr(style_id, all_style_defs, doc_defaults)
    
    if effective_rpr:
        if rpr:
            # Merge: inject missing tags from effective_rpr
            for tag in ("rFonts", "sz", "szCs", "lang"):
                if f"<w:{tag}" not in rpr:
                    node_m = re.search(rf'(<w:{tag}\b[^/>]*(?:/>|>[\s\S]*?</w:{tag}>))', effective_rpr)
                    if node_m:
                        # Insert before </w:rPr>
                        rpr = rpr.replace("</w:rPr>", f"{node_m.group(1)}</w:rPr>")
            parts.append(f'  {rpr}')
        else:
            parts.append(f'  <w:rPr>{effective_rpr}</w:rPr>')
    elif rpr:
        parts.append(f'  {rpr}')
    
    # Table properties if present
    if style_def.get("tblPr"):
        parts.append(f'  {style_def["tblPr"]}')
    if style_def.get("trPr"):
        parts.append(f'  {style_def["trPr"]}')
    if style_def.get("tcPr"):
        parts.append(f'  {style_def["tcPr"]}')
    
    parts.append('</w:style>')
    
    return "\n".join(parts)


# ─────────────────────────────────────────────────────────────────────────────
# Main environment application
# ─────────────────────────────────────────────────────────────────────────────

def apply_environment_to_target(
    target_extract_dir: Path,
    registry: Dict[str, Any],
    log: List[str],
    apply_theme_flag: bool = True,
    apply_settings_flag: bool = True,
    apply_doc_defaults_flag: bool = True,
    apply_fonts_flag: bool = True,
) -> None:
    """
    Apply the formatting environment from arch_template_registry to target.
    
    Application order (deterministic):
    1. Theme (font/color definitions)
    2. Settings + compat (rendering behavior)
    3. Font table (font declarations)
    4. docDefaults in styles.xml (baseline formatting)
    
    Args:
        target_extract_dir: Extracted target document folder
        registry: Loaded arch_template_registry.json
        log: List to append log messages
        apply_*: Flags to selectively disable parts of application
    """
    target_extract_dir = Path(target_extract_dir)
    
    log.append("=" * 60)
    log.append("BEGIN ENVIRONMENT APPLICATION")
    log.append("=" * 60)
    
    # 1. Theme
    if apply_theme_flag:
        log.append("\n[1/4] Applying theme...")
        apply_theme(target_extract_dir, registry, log)
    else:
        log.append("\n[1/4] Theme application skipped")
    
    # 2. Settings/compat
    if apply_settings_flag:
        log.append("\n[2/4] Applying settings/compat...")
        apply_settings(target_extract_dir, registry, log)
    else:
        log.append("\n[2/4] Settings application skipped")
    
    # 3. Font table
    if apply_fonts_flag:
        log.append("\n[3/4] Applying font table...")
        apply_font_table(target_extract_dir, registry, log)
    else:
        log.append("\n[3/4] Font table application skipped")
    
    # 4. docDefaults in styles.xml
    if apply_doc_defaults_flag:
        log.append("\n[4/4] Applying docDefaults...")
        styles_path = target_extract_dir / "word" / "styles.xml"
        if styles_path.exists():
            styles_xml = styles_path.read_text(encoding="utf-8")
            styles_xml = apply_doc_defaults(styles_xml, registry, log)
            styles_path.write_text(styles_xml, encoding="utf-8")
        else:
            log.append("WARNING: No styles.xml in target; cannot apply docDefaults")
    else:
        log.append("\n[4/4] docDefaults application skipped")
    
    log.append("\n" + "=" * 60)
    log.append("END ENVIRONMENT APPLICATION")
    log.append("=" * 60)


def get_style_def_by_id(registry: Dict[str, Any], style_id: str) -> Optional[Dict[str, Any]]:
    """Look up a style definition from the registry by styleId."""
    style_defs = registry.get("styles", {}).get("style_defs", [])
    for sd in style_defs:
        if sd.get("style_id") == style_id:
            return sd
    return None


def get_styles_with_dependencies(
    registry: Dict[str, Any],
    needed_style_ids: List[str]
) -> List[Dict[str, Any]]:
    """
    Get style definitions including basedOn dependencies.
    Returns styles in dependency order (base styles first).
    """
    style_defs = registry.get("styles", {}).get("style_defs", [])
    by_id = {s["style_id"]: s for s in style_defs}
    
    # Expand dependencies
    expanded = set()
    to_process = list(needed_style_ids)
    
    while to_process:
        sid = to_process.pop()
        if sid in expanded:
            continue
        expanded.add(sid)
        
        sd = by_id.get(sid)
        if sd and sd.get("based_on"):
            to_process.append(sd["based_on"])
    
    # Sort by dependency order (styles with no basedOn first)
    result = []
    remaining = list(expanded)
    added = set()
    
    # Simple topological sort
    max_iterations = len(remaining) * 2
    iterations = 0
    while remaining and iterations < max_iterations:
        iterations += 1
        for sid in list(remaining):
            sd = by_id.get(sid)
            if not sd:
                remaining.remove(sid)
                continue
            
            base = sd.get("based_on")
            if not base or base in added or base not in expanded:
                result.append(sd)
                added.add(sid)
                remaining.remove(sid)
    
    # Add any remaining (shouldn't happen with valid data)
    for sid in remaining:
        sd = by_id.get(sid)
        if sd:
            result.append(sd)
    
    return result


# ─────────────────────────────────────────────────────────────────────────────
# CLI (for testing)
# ─────────────────────────────────────────────────────────────────────────────

def main():
    import argparse
    import json
    
    parser = argparse.ArgumentParser(
        description="Apply arch_template_registry.json environment to target document"
    )
    parser.add_argument(
        "target_dir",
        help="Path to extracted target document folder"
    )
    parser.add_argument(
        "registry",
        help="Path to arch_template_registry.json"
    )
    parser.add_argument(
        "--no-theme", action="store_true",
        help="Skip theme application"
    )
    parser.add_argument(
        "--no-settings", action="store_true",
        help="Skip settings/compat application"
    )
    parser.add_argument(
        "--no-fonts", action="store_true",
        help="Skip font table application"
    )
    parser.add_argument(
        "--no-doc-defaults", action="store_true",
        help="Skip docDefaults application"
    )
    
    args = parser.parse_args()
    
    target_dir = Path(args.target_dir)
    if not target_dir.exists():
        raise FileNotFoundError(f"Target directory not found: {target_dir}")
    
    registry_path = Path(args.registry)
    if not registry_path.exists():
        raise FileNotFoundError(f"Registry not found: {registry_path}")
    
    registry = json.loads(registry_path.read_text(encoding="utf-8"))
    
    log: List[str] = []
    
    apply_environment_to_target(
        target_extract_dir=target_dir,
        registry=registry,
        log=log,
        apply_theme_flag=not args.no_theme,
        apply_settings_flag=not args.no_settings,
        apply_fonts_flag=not args.no_fonts,
        apply_doc_defaults_flag=not args.no_doc_defaults,
    )
    
    print("\n".join(log))


if __name__ == "__main__":
    main()