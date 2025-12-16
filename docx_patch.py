# docx_patch.py
from __future__ import annotations

from pathlib import Path
import zipfile
from typing import Dict, Union

BytesOrStr = Union[bytes, str]

def patch_docx(
    src_docx: Path,
    out_docx: Path,
    replacements: Dict[str, BytesOrStr],
) -> None:
    """
    Create out_docx by copying every ZIP entry from src_docx unchanged,
    except for entries whose internal paths match keys in `replacements`.

    This is NOT a "rebuild from extracted folder".
    It's a surgical patch: swap specific parts, preserve everything else.
    """
    src_docx = Path(src_docx)
    out_docx = Path(out_docx)

    rep_bytes: Dict[str, bytes] = {}
    for k, v in replacements.items():
        if isinstance(v, str):
            rep_bytes[k] = v.encode("utf-8")
        else:
            rep_bytes[k] = v

    # Phase 2 hard invariants — enforce at patch boundary
    FORBIDDEN_PREFIXES = (
        "word/header",
        "word/footer",
    )

    FORBIDDEN_EXACT = {
        "word/numbering.xml",
    }

    # Expanded to include environment parts
    ALLOWED_PATCHES = {
        "word/document.xml",
        "word/styles.xml",
        "word/theme/theme1.xml",
        "word/settings.xml",
        "word/fontTable.xml",
        "[Content_Types].xml",
        "word/_rels/document.xml.rels",
    }

    for name in rep_bytes:
        if name in FORBIDDEN_EXACT:
            raise RuntimeError(f"Forbidden patch target: {name}")

        if name.startswith(FORBIDDEN_PREFIXES):
            raise RuntimeError(f"Forbidden patch target: {name}")

        if name not in ALLOWED_PATCHES:
            raise RuntimeError(
                f"Illegal patch target: {name}\n"
                f"Allowed: {sorted(ALLOWED_PATCHES)}"
            )


    out_docx.parent.mkdir(parents=True, exist_ok=True)
    if out_docx.exists():
        out_docx.unlink()

    with zipfile.ZipFile(src_docx, "r") as zin:
        with zipfile.ZipFile(out_docx, "w") as zout:
            # preserve archive comment if any
            zout.comment = zin.comment

            src_names = set(zin.namelist())
            
            # For new parts (like theme1.xml if it didn't exist), we'll add them
            new_parts = [name for name in rep_bytes.keys() if name not in src_names]
            existing_replacements = [name for name in rep_bytes.keys() if name in src_names]
            
            # Ensure we are not accidentally dropping entries
            assert len(src_names) == len(zin.infolist())

            for info in zin.infolist():
                name = info.filename
                data = rep_bytes.get(name, zin.read(name))

                # Preserve per-entry compression type where possible
                zout.writestr(info, data, compress_type=info.compress_type)
            
            # Add any new parts that didn't exist in source
            for new_name in new_parts:
                zout.writestr(new_name, rep_bytes[new_name])