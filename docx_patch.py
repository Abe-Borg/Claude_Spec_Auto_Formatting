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

    out_docx.parent.mkdir(parents=True, exist_ok=True)
    if out_docx.exists():
        out_docx.unlink()

    with zipfile.ZipFile(src_docx, "r") as zin:
        with zipfile.ZipFile(out_docx, "w") as zout:
            # preserve archive comment if any
            zout.comment = zin.comment

            src_names = set(zin.namelist())
            missing = [name for name in rep_bytes.keys() if name not in src_names]
            if missing:
                raise FileNotFoundError(f"Replacement targets not found in docx: {missing}")

            for info in zin.infolist():
                name = info.filename
                data = rep_bytes.get(name, zin.read(name))

                # Preserve per-entry compression type where possible
                # (Word doesn't require it, but this reduces drift)
                zout.writestr(info, data, compress_type=info.compress_type)
