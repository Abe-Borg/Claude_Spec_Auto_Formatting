DOCX CSI Structural Normalizer — Architect-Template Learning Engine

Overview

This project implements a safe, deterministic system for extracting structural meaning and architect-defined formatting from Microsoft Word (.docx) specification documents used in AEC workflows.

It was built to operate in environments where:

Architect Word templates are sacred

Visual drift is unacceptable

Formatting is often inconsistent, implicit, or hand-applied

CSI MasterFormat / SectionFormat / PageFormat structure is desired

Word’s internal formatting behavior is fragile and non-obvious

The system does not normalize appearance.
It learns structure and formatting from the architect’s document without changing how it looks.

Core Problem

AEC specification documents must satisfy two competing requirements:

Structural correctness

Sections, PARTs, Articles, Paragraphs, Subparagraphs, etc. must be explicitly identifiable

Required for automation, analysis, downstream processing, and consistency

Visual immutability

Architect-controlled templates must not change

Hanging indents, spacing, numbering, headers/footers, and alignment must remain pixel-identical

Naïve approaches (DOCX reconstruction, formatting enforcement, XML rewriting) consistently fail because Word relies on implicit inheritance and undocumented behavior.

Final Architecture (Successful & Locked In)
Principle

The architect’s DOCX controls appearance.
The system may only annotate structure and learn formatting — never invent it.

What This Tool Does
High-Level Capabilities

Analyzes a DOCX spec without altering its appearance

Identifies CSI structural roles:

Section Title

PART

Article

Paragraph

Subparagraph

Sub-subparagraph

Learns actual formatting used by the architect for each role

Creates real Word styles derived from exemplar paragraphs

Applies styles in place with zero visual drift

Produces a reusable formatting profile for downstream use

What This Tool Explicitly Does NOT Do

These are intentional and enforced:

❌ No DOCX reconstruction

❌ No formatting normalization

❌ No paragraph spacing/alignment/indent edits

❌ No numbering changes

❌ No header/footer changes

❌ No sectPr changes

❌ No LLM-authored XML

❌ No CSI visual enforcement

If any of the above occur, the run fails.

Pipeline Summary
1. Extract (Once)

DOCX is unzipped into an extracted directory

Stability hashes are recorded for:

headers

footers

section properties (sectPr)

relationships

content types

2. Slim Structural Bundle

The system generates a minimal JSON representation:

Paragraph index

Raw text

Numbering hints (read-only context)

Existing styles (catalog only)

Flags for sectPr containment

No formatting data is sent to the LLM.

3. LLM Classification (Structure Only)

The LLM:

Sees only the slim bundle

Classifies paragraphs into CSI roles

Chooses exemplar paragraph indices per role

Returns JSON instructions only

The LLM is forbidden from:

specifying formatting

emitting XML

emitting pPr / rPr

proposing visual changes

4. Local Style Derivation (Critical Step)

The local script:

Locates exemplar paragraphs chosen by the LLM

Extracts their effective formatting:

paragraph properties (w:pPr)

excluding w:pStyle

excluding w:numPr

run properties (w:rPr) from representative runs

Synthesizes real Word styles in styles.xml

Preserves all existing formatting behavior

These styles reflect exactly what the architect authored, whether via:

real styles

hand formatting

inconsistent usage

5. In-Place Mutation

Inserts w:pStyle into paragraphs by index

Does not modify any other paragraph properties

Does not touch numbering, headers, footers, or section properties

6. Stability Verification (Hard Fail on Drift)

Every run verifies:

Paragraph XML unchanged except for w:pStyle

Headers unchanged

Footers unchanged

sectPr unchanged

Relationships unchanged

[Content_Types].xml unchanged

If anything drifts → execution stops.

Resulting Output

After a successful run, the document:

Looks pixel-identical

Contains:

Explicit CSI structural tagging

Architect-derived paragraph styles:

CSI_SectionTitle__ARCH

CSI_Part__ARCH

CSI_Article__ARCH

CSI_Paragraph__ARCH

CSI_Subparagraph__ARCH

CSI_Subsubparagraph__ARCH

These styles now represent the architect’s formatting intent in a reusable, machine-readable way.

Why This Matters

This tool solves a long-standing AEC problem:

“How do we match an architect’s spec formatting without guessing, rewriting Word, or breaking their template?”

Answer:

Learn their formatting

Encode it as styles

Apply it mechanically later

Known Limitations (Intentional)

No numbering normalization

No list-definition rewriting

No spec generation

No visual enforcement

These belong to later, opt-in phases.

Design Philosophy

Determinism over cleverness

Safety over convenience

Structure before appearance

Learn from the document — never override it

Fail fast if invariants are violated

Status

✅ Production-ready for structural learning
✅ Successfully tested against multiple architect templates
✅ Robust even on poorly authored specs