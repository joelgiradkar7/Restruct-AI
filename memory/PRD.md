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

## Prioritized backlog
- P0: None remaining for the current scope.
- P1: Add engineer-authored review notes and report branding/custom cover details.
- P2: Add multi-component assessments and side-by-side technique comparison.

## Next tasks list
1. Add a printable engineering review report.
2. Add structured engineer sign-off and review status.
3. Add multi-component site assessment grouping.