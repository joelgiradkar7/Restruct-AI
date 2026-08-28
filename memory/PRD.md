# ResructAI Product Requirements Document

## Original problem statement
Build me a AI for suggesting the best technique can use for rehabilitation of the blast damane structure.

## Architecture decisions
- React frontend with FastAPI backend and MongoDB persistence.
- AI recommendations are generated server-side through `emergentintegrations` using OpenAI GPT-5.4 Mini and the universal LLM key.
- Optional photos are sent as base64 only to the backend/AI request and are not persisted in MongoDB.
- API response models use Pydantic and Mongo queries exclude MongoDB `_id`.
- Indian Standards considerations are requested per technique without invented clause numbers.

## User personas
- Structural engineer or inspector performing an early blast-damage triage.
- Recovery coordinator who needs a concise comparison of possible rehabilitation pathways.

## Core requirements (static)
- Guided structural damage intake with component, material, blast context, cracks, spalling, reinforcement, integrity, and notes.
- Optional JPG/PNG/WEBP photo preview and strict server-side validation.
- Ranked techniques with reasoning, urgency, materials, procedure, safety notes, timeframe, and IS-code considerations.
- Safety gate, professional disclaimer, saved history, and prior-result reopening.

## What's been implemented
- 2026-08-24: Built the ResructAI intake, result, history, and responsive navigation experience.
- 2026-08-24: Added live GPT-5.4 Mini streaming-backed generation and MongoDB persistence.
- 2026-08-24: Added Indian Standards-aware recommendations and verified a representative RC blast case.
- 2026-08-24: Added strict actual-image verification and rejection of malformed/near-uniform uploads.
- 2026-08-24: Added concise result presentation, retained uploaded site photos in the result view, and added PDF, DOCX, and XLSX report downloads.
- 2026-02: Evidence-based reporting upgrade — Quick Assessment, Evidence & Confidence (Observed/Estimated/Unknown), possible damage mechanism, image-based limitations, required verification, insufficient-evidence handling; all propagated to PDF/DOCX/XLSX.
- 2026-02: Backend `normalize_advice` hardening — enforces enum values (status, classification, confidence, urgency), clamps confidence percentage 0-100, coerces list types, pads to exactly 3 recommendations with a verification-focused pathway when AI returns fewer.
- 2026-02: Mobile 390px fix — added `overflow-wrap: anywhere` to result/section headings and history item text to prevent long underscored structure names from overflowing.
- 2026-02: Phase 1 general damage assessment — generalized from blast-only to 20 damage types (cracking, spalling, delamination, corrosion, fire, seismic, blast, impact, water, chemical, settlement, foundation, overloading, defects, deterioration, deformation, masonry, connection, unknown, other), 16 component types, 8 materials. Added new AI output fields: `observed_conditions`, `user_measurements`, `preliminary_severity` (Low/Moderate/High/Critical/Indeterminate), `evidence_quality` (Good/Moderate/Poor/Insufficient), `missing_information`. AI prompt enforces observed vs mechanism separation with hedged wording and forbids invented measurements/capacities. All new sections included in PDF/DOCX/XLSX exports and rendered on the results screen.

## Prioritized backlog
- P0: None remaining.
- P1: Printable engineering review report; structured engineer sign-off/review status; multi-component site assessment grouping. Phase 2 rehabilitation decision-support (repair/strengthen/replace/reconstruct routing) — deferred by user.
- P2: Evidence gallery for multiple labelled before/after photos; side-by-side technique comparison; one-page field executive summary.

## Next tasks list
1. Add a printable engineering review report.
2. Add structured engineer sign-off and review status.
3. Add multi-component site assessment grouping.