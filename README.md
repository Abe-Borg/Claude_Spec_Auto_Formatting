📘 README — Phase 2: MEP Specification Styling Engine
Overview

Phase 2 applies architect-defined CSI paragraph styles to mechanical and plumbing specifications while preserving exact Word behavior and appearance.

This phase does not design styles.
It consumes styles produced by Phase 1 and applies them deterministically.

The core principle is:

Change as little Word XML as possible while achieving exact visual and behavioral alignment with the architect’s template.

Inputs

Phase 2 requires three inputs:

Target MEP DOCX
A mechanical or plumbing specification document (often inconsistent or poorly styled).

Architect Style Registry
Produced by Phase 1:

arch_style_registry.json


Phase 2 Classifications JSON
Output from the LLM:

{
  "classifications": [
    { "paragraph_index": 42, "csi_role": "PARAGRAPH" }
  ]
}

Responsibilities (What Phase 2 Does)
1. Extract DOCX Safely

Unzips the DOCX into an editable workspace

Preserves:

headers

footers

w:sectPr

numbering definitions

2. Build Slim Bundle

Sends only paragraph text + indices to the LLM

No XML

No formatting

No styles

3. Classify CSI Roles (LLM)

The LLM:

Assigns CSI semantic roles only

Never references formatting

Never creates styles

Never guesses

Allowed roles:

SectionID
SectionTitle
PART
ARTICLE
PARAGRAPH
SUBPARAGRAPH
SUBSUBPARAGRAPH

4. Load Architect Style Registry (STRICT)

Reads arch_style_registry.json

No heuristics

No guessing

Registry is the only source of truth

If a role is missing → skip safely and log.

5. Import Architect Styles

Imports only styles actually used in the document

Includes all basedOn dependencies

Materializes inherited properties:

<w:rPr> (font, size, etc.)

<w:pPr> (spacing, alignment)

Does not:

modify numbering definitions

modify base styles

6. Preserve Dynamic Word Numbering

Before swapping w:pStyle, Phase 2:

Detects if numbering comes from the current style

Copies <w:numPr> from the style chain onto the paragraph

Ensures:

Pressing Enter continues a / b / c, 1 / 2 / 3, etc.

This is critical for Word usability.

7. Apply Styles

Applies architect styles using only w:pStyle

Never edits runs

Never invents formatting

8. Verify Stability

Confirms headers, footers, and sectPr are unchanged

Fails loudly if invariants are violated

9. Optional Outputs

--rebuild-docx → rebuild final DOCX

--write-analysis → debug markdown (off by default)

Explicit Non-Goals

Phase 2 does not:

Create styles

Fix bad architect templates

Modify numbering.xml

Apply run-level formatting

Reconstruct DOCX XML

Infer style intent

Merge Phase 1 logic

Relationship to Phase 1
Phase	Responsibility
Phase 1	Defines and labels architect intent
Phase 2	Applies that intent to MEP specs

Phase 1 emits:

arch_style_registry.json


Phase 2 consumes it verbatim.

Philosophy

Word is stateful, undocumented, and fragile.
We respect it by touching as little as possible.

⚠️ Known Invariants & Failure Modes

This document defines what must never break, and what to watch for when it does.

🔒 Hard Invariants (Must Always Hold)

If any of these break, the run is invalid.

1. Headers / Footers Unchanged

No XML drift

No relationship changes

2. w:sectPr Untouched

No page setup changes

No section breaks altered

3. Numbering Definitions Untouched

numbering.xml is never edited

All numbering preservation happens at paragraph-level only

4. No Run-Level Formatting

No <w:rPr> edits inside document.xml

All formatting via paragraph styles only

5. Registry-Only Styling

No guessing style IDs

No scanning styles.xml heuristically

Missing role → skip + log

⚠️ Known Failure Modes (And Why They Happen)
1. Numbering Stops When Pressing Enter

Cause

Numbering was style-linked

w:pStyle was swapped without materializing <w:numPr>

Fix

Ensure ensure_explicit_numpr_from_current_style() runs before restyling

2. Fonts Change After Styling

Cause

Architect style relied on inheritance / docDefaults

Imported style lacked explicit <w:rPr>

Fix

Materialize effective <w:rPr> and <w:pPr> when importing styles

3. Some Paragraphs Don’t Get Styled

Cause

Role missing from registry

Role intentionally skipped (e.g., SKIP, END_OF_SECTION)

Fix

Expected behavior

Logged in preflight

4. Architect Template Uses Only “Normal”

Cause

Architect never defined styles

Phase 1 Responsibility

Derive styles from exemplars anyway

Registry still emitted

Phase 2 Behavior

Works normally once registry exists

5. Word Opens With “Repair” Warning

Cause

Invalid XML insertion

Broken style dependency chain

Fix

Check:

imported style blocks are intact

all basedOn styles imported

no workspace artifacts zipped into DOCX

🧪 Acceptance Checklist (Before Cleanup)

Before trimming code further, confirm:

 Lists continue correctly on Enter

 Fonts match architect intent

 Styles pane shows CSI_*__ARCH styles

 No extra files produced by default

 Preflight reports only expected unmapped roles

 Word opens cleanly without warnings

📉 Cleanup Guidance (Future)

When trimming Phase 2:

Remove:

heuristic code paths

legacy analysis reports

unused CLI modes

Keep:

extract

classify

import styles

apply styles

verify invariants