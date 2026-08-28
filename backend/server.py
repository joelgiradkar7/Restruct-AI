from fastapi import FastAPI, APIRouter, HTTPException
from fastapi.responses import StreamingResponse
from dotenv import load_dotenv
from starlette.middleware.cors import CORSMiddleware
from motor.motor_asyncio import AsyncIOMotorClient
import os
import logging
from pathlib import Path
from pydantic import BaseModel, Field, ConfigDict
from typing import List, Optional, Any
import uuid
import json
import base64
import binascii
from io import BytesIO
from datetime import datetime, timezone
from PIL import Image, UnidentifiedImageError
from emergentintegrations.llm.chat import LlmChat, UserMessage, TextDelta
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, ListFlowable, ListItem
from docx import Document
from openpyxl import Workbook


ROOT_DIR = Path(__file__).parent
load_dotenv(ROOT_DIR / '.env')

# MongoDB connection
mongo_url = os.environ['MONGO_URL']
client = AsyncIOMotorClient(mongo_url)
db = client[os.environ['DB_NAME']]

# Create the main app without a prefix
app = FastAPI()

# Create a router with the /api prefix
api_router = APIRouter(prefix="/api")


# Define Models
class StatusCheck(BaseModel):
    model_config = ConfigDict(extra="ignore")  # Ignore MongoDB's _id field
    
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    client_name: str
    timestamp: datetime = Field(default_factory=lambda: datetime.now(timezone.utc))

class StatusCheckCreate(BaseModel):
    client_name: str

class AssessmentCreate(BaseModel):
    structure_name: str
    component_type: str
    construction_material: str
    blast_context: str
    crack_width: str
    spalling_depth: str
    reinforcement_exposure: str
    integrity_status: str
    site_notes: str
    image_base64: Optional[str] = None
    image_mime: Optional[str] = None

class Assessment(BaseModel):
    id: str
    structure_name: str
    component_type: str
    integrity_status: str
    created_at: str
    recommendations: List[dict]
    immediate_actions: List[str]
    disclaimer: str
    assessment_status: Optional[str] = None
    damage_classification: Optional[str] = None
    quick_assessment: Optional[dict] = None
    evidence_confidence: Optional[dict] = None
    possible_damage_mechanism: Optional[dict] = None
    limitations: Optional[List[str]] = None
    required_verification: Optional[List[str]] = None

# Add your routes to the router instead of directly to app
@api_router.get("/")
async def root():
    return {"message": "ResructAI API ready"}

def clean_json_text(raw: str) -> dict:
    text = raw.strip().replace("```json", "").replace("```", "").strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        start, end = text.find("{"), text.rfind("}")
        if start >= 0 and end > start:
            return json.loads(text[start:end + 1])
        raise

STATUS_VALUES = {"Sufficient evidence", "Insufficient evidence"}
CLASSIFICATION_VALUES = {"No apparent structural damage", "Minor / localized", "Moderate", "Severe", "Potentially unstable", "Insufficient evidence"}
CONFIDENCE_VALUES = {"Low", "Moderate", "High"}
URGENCY_VALUES = {"Critical", "High", "Moderate", "Monitor"}

def _pick(value: Any, allowed: set, fallback: str) -> str:
    return value if isinstance(value, str) and value in allowed else fallback

def _string_list(value: Any) -> List[str]:
    if isinstance(value, list):
        return [str(x).strip() for x in value if isinstance(x, (str, int, float)) and str(x).strip()]
    if isinstance(value, str) and value.strip():
        return [value.strip()]
    return []

def _int_percent(value: Any) -> int:
    try:
        return max(0, min(100, int(value)))
    except (TypeError, ValueError):
        return 0

def normalize_advice(advice: dict, component_type: str) -> dict:
    if not isinstance(advice, dict):
        advice = {}
    status = _pick(advice.get("assessment_status"), STATUS_VALUES, "Insufficient evidence")
    classification = _pick(advice.get("damage_classification"), CLASSIFICATION_VALUES, "Insufficient evidence")
    if status == "Insufficient evidence":
        classification = "Insufficient evidence"

    quick = advice.get("quick_assessment") if isinstance(advice.get("quick_assessment"), dict) else {}
    quick_norm = {
        "element": str(quick.get("element") or component_type),
        "apparent_condition": str(quick.get("apparent_condition") or "Cannot be determined from available evidence"),
        "key_visible_evidence": _string_list(quick.get("key_visible_evidence")),
        "severity": str(quick.get("severity") or classification),
        "confidence": _pick(quick.get("confidence"), CONFIDENCE_VALUES, "Low"),
        "immediate_concern": str(quick.get("immediate_concern") or "Requires engineer review before further action"),
        "engineer_verification_required": str(quick.get("engineer_verification_required") or "Yes — licensed structural engineer must inspect and verify"),
    }

    conf = advice.get("evidence_confidence") if isinstance(advice.get("evidence_confidence"), dict) else {}
    conf_norm = {
        "overall": _pick(conf.get("overall"), CONFIDENCE_VALUES, "Low"),
        "percentage": _int_percent(conf.get("percentage")),
        "evidence_detected": _string_list(conf.get("evidence_detected")),
        "estimated_inferred": _string_list(conf.get("estimated_inferred")),
        "cannot_be_determined": _string_list(conf.get("cannot_be_determined")) or ["Residual capacity and hidden damage cannot be determined from available evidence."],
    }

    mech = advice.get("possible_damage_mechanism") if isinstance(advice.get("possible_damage_mechanism"), dict) else {}
    mech_norm = {
        "most_likely": str(mech.get("most_likely") or "Cannot be determined from available evidence."),
        "alternatives": _string_list(mech.get("alternatives")),
        "confidence": _pick(mech.get("confidence"), CONFIDENCE_VALUES, "Low"),
        "supporting_evidence": _string_list(mech.get("supporting_evidence")),
    }

    limitations = _string_list(advice.get("limitations")) or ["Image-based assessment cannot establish residual load-bearing capacity, reinforcement loss, concrete strength, hidden delamination, or complete structural stability."]
    verification = _string_list(advice.get("required_verification")) or ["Detailed visual inspection and structural calculations by a licensed engineer."]
    immediate = _string_list(advice.get("immediate_actions")) or ["Restrict access to the affected zone until a licensed structural engineer inspects the component."]

    raw_recs = advice.get("recommendations") if isinstance(advice.get("recommendations"), list) else []
    recs: List[dict] = []
    for idx, item in enumerate(raw_recs[:3]):
        if not isinstance(item, dict):
            continue
        recs.append({
            "rank": idx + 1,
            "technique": str(item.get("technique") or "Engineer-led verification and interim protection"),
            "fit": str(item.get("fit") or "Requires engineer confirmation"),
            "reasoning": str(item.get("reasoning") or "Evidence is insufficient to commit to a specific rehabilitation technique."),
            "evidence_link": str(item.get("evidence_link") or "UNKNOWN — the engineer must confirm the condition that makes this option appropriate."),
            "appropriate_when": str(item.get("appropriate_when") or "Only after the listed verification confirms the observed condition."),
            "do_not_use_when": str(item.get("do_not_use_when") or "Do not use until instability, hidden damage, and load-path effects are checked."),
            "urgency": _pick(item.get("urgency"), URGENCY_VALUES, "Moderate"),
            "materials": _string_list(item.get("materials")),
            "procedure": _string_list(item.get("procedure")),
            "safety_notes": _string_list(item.get("safety_notes")),
            "timeframe": str(item.get("timeframe") or "To be confirmed after engineer review"),
            "is_code_considerations": _string_list(item.get("is_code_considerations")) or ["Verify the current applicable Indian Standards and project provisions with the engineer."],
        })
    while len(recs) < 3:
        rank = len(recs) + 1
        recs.append({
            "rank": rank,
            "technique": "Engineer-led verification pathway",
            "fit": "Requires engineer confirmation",
            "reasoning": "Insufficient evidence to recommend a specific rehabilitation technique for this pathway.",
            "evidence_link": "UNKNOWN — evidence does not yet support a distinct technique for this pathway.",
            "appropriate_when": "Only after site inspection, material testing, and structural analysis by a licensed engineer.",
            "do_not_use_when": "Do not proceed until instability, hidden damage, and load paths are verified.",
            "urgency": "Monitor",
            "materials": [],
            "procedure": ["Await engineer inspection and testing before selecting a repair technique."],
            "safety_notes": ["Maintain exclusion zone and shoring until verification is complete."],
            "timeframe": "To be confirmed after engineer review",
            "is_code_considerations": ["Verify the current applicable Indian Standards and project provisions with the engineer."],
        })

    return {
        "assessment_status": status,
        "damage_classification": classification,
        "quick_assessment": quick_norm,
        "evidence_confidence": conf_norm,
        "possible_damage_mechanism": mech_norm,
        "limitations": limitations,
        "required_verification": verification,
        "immediate_actions": immediate,
        "recommendations": recs,
    }

async def generate_advice(payload: AssessmentCreate) -> dict:
    key = os.environ.get("EMERGENT_LLM_KEY")
    if not key:
        raise HTTPException(status_code=503, detail="AI service is not configured")
    if payload.image_base64:
        if payload.image_mime not in {"image/jpeg", "image/png", "image/webp"}:
            raise HTTPException(status_code=422, detail="Only JPG, PNG, or WEBP photos are supported")
        try:
            raw_image = base64.b64decode(payload.image_base64.split(",")[-1], validate=True)
            with Image.open(BytesIO(raw_image)) as image:
                image.verify()
            with Image.open(BytesIO(raw_image)) as image:
                sample = image.convert("L").resize((64, 64)).getextrema()
                if sample[1] - sample[0] <= 2:
                    raise HTTPException(status_code=422, detail="Please upload a photo with visible structural detail")
        except HTTPException:
            raise
        except (ValueError, binascii.Error, UnidentifiedImageError, OSError) as exc:
            raise HTTPException(status_code=422, detail="The uploaded photo is not a valid image") from exc
    prompt = f"""You are a cautious structural rehabilitation advisor supporting work in India. Analyze the supplied user fields and optional image for a qualified engineer. Never claim a final design, occupancy approval, structural safety guarantee, remaining capacity, or confirmed cause. Return ONLY valid JSON with this exact shape:
{{"assessment_status":"Sufficient evidence|Insufficient evidence","damage_classification":"No apparent structural damage|Minor / localized|Moderate|Severe|Potentially unstable|Insufficient evidence","quick_assessment":{{"element":"...","apparent_condition":"...","key_visible_evidence":["..."],"severity":"...","confidence":"Low|Moderate|High","immediate_concern":"...","engineer_verification_required":"..."}},"evidence_confidence":{{"overall":"Low|Moderate|High","percentage":0,"evidence_detected":["..."],"estimated_inferred":["..."],"cannot_be_determined":["..."]}},"possible_damage_mechanism":{{"most_likely":"...","alternatives":["..."],"confidence":"Low|Moderate|High","supporting_evidence":["..."]}},"limitations":["..."],"required_verification":["..."],"immediate_actions":["..."],"recommendations":[{{"rank":1,"technique":"...","fit":"...","reasoning":"...","evidence_link":"...","appropriate_when":"...","do_not_use_when":"...","urgency":"Critical|High|Moderate|Monitor","materials":["..."],"procedure":["..."],"safety_notes":["..."],"timeframe":"...","is_code_considerations":["..." ]}}]}}
Internally classify every important statement as OBSERVED (visible or explicitly user-provided), ESTIMATED (approximate), INFERRED (possible explanation), or UNKNOWN (not determinable), then communicate those labels naturally. Evidence_detected may include only visible image features or explicit user fields. Never invent measurements, damage depth, material properties, loading causes, or structural capacities. If the image is poor, irrelevant, ambiguous, or no structural evidence is available, set assessment_status and damage_classification to Insufficient evidence, lower confidence, explain what is visible/provided and what is unknown, request better photographs (overall context, close-up, scale, multiple angles), and do not issue a confident diagnosis. A reported blast event is user-provided context, not visual proof of blast damage.
Provide 3 concise ranked pathways only when evidence supports decision support; otherwise keep recommendations conservative and verification-led. Connect each pathway to evidence, state when it is appropriate, and when it must not be used. For each option identify applicable Indian Standard families only when relevant (for example IS 456 for reinforced concrete, IS 15988 for evaluation/strengthening, IS 13920 for ductile detailing, IS 800 for structural steel, or IS 516 for material testing). Do not invent clause numbers; say current editions and project-specific provisions must be verified.
Assessment: {payload.model_dump_json(exclude={"image_base64"})}"""
    chat = LlmChat(api_key=key, session_id=str(uuid.uuid4()), system_message="You provide cautious, evidence-aware structural rehabilitation decision support.").with_model("openai", "gpt-5.4-mini")
    contents = []
    if payload.image_base64 and payload.image_mime in {"image/jpeg", "image/png", "image/webp"}:
        from emergentintegrations.llm.chat import ImageContent
        contents.append(ImageContent(image_base64=payload.image_base64.split(",")[-1]))
        prompt += "\nA site photo is attached. Use it only as supplementary visual evidence; do not infer hidden damage."
    output = []
    async for event in chat.stream_message(UserMessage(text=prompt, file_contents=contents)):
        if isinstance(event, TextDelta):
            output.append(event.content)
    try:
        return clean_json_text("".join(output))
    except Exception as exc:
        logging.exception("AI returned invalid JSON")
        raise HTTPException(status_code=502, detail="The AI returned an unreadable assessment. Please try again.") from exc

@api_router.post("/assessments", response_model=Assessment)
async def create_assessment(payload: AssessmentCreate):
    advice = await generate_advice(payload)
    normalized = normalize_advice(advice, payload.component_type)
    record = {
        "id": str(uuid.uuid4()), "structure_name": payload.structure_name,
        "component_type": payload.component_type, "integrity_status": payload.integrity_status,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "recommendations": normalized["recommendations"],
        "immediate_actions": normalized["immediate_actions"],
        "disclaimer": "AI provides preliminary decision support, not structural certification. A licensed structural engineer must inspect, verify calculations, assess load paths, and approve all work; this report does not guarantee safety, occupancy, capacity, or repair performance.",
        "assessment_status": normalized["assessment_status"],
        "damage_classification": normalized["damage_classification"],
        "quick_assessment": normalized["quick_assessment"],
        "evidence_confidence": normalized["evidence_confidence"],
        "possible_damage_mechanism": normalized["possible_damage_mechanism"],
        "limitations": normalized["limitations"],
        "required_verification": normalized["required_verification"],
    }
    await db.assessments.insert_one(record.copy())
    return record

@api_router.get("/assessments", response_model=List[Assessment])
async def list_assessments():
    records = await db.assessments.find({}, {"_id": 0}).sort("created_at", -1).to_list(50)
    return records

async def get_assessment_or_404(assessment_id: str) -> dict:
    record = await db.assessments.find_one({"id": assessment_id}, {"_id": 0})
    if not record:
        raise HTTPException(status_code=404, detail="Assessment not found")
    return record

def report_filename(record: dict, extension: str) -> str:
    safe_name = "".join(c if c.isalnum() or c in "-_" else "_" for c in record["structure_name"])[:48]
    return f"{safe_name}_resructai_report.{extension}"

@api_router.get("/assessments/{assessment_id}/report/{file_format}")
async def download_report(assessment_id: str, file_format: str):
    record = await get_assessment_or_404(assessment_id)
    if file_format not in {"pdf", "docx", "xlsx"}:
        raise HTTPException(status_code=400, detail="Choose pdf, docx, or xlsx")
    buffer = BytesIO()
    title = f"ResructAI | {record['structure_name']}"
    if file_format == "pdf":
        styles = getSampleStyleSheet()
        doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=42, leftMargin=42, topMargin=42, bottomMargin=42)
        quick = record.get("quick_assessment", {}); confidence = record.get("evidence_confidence", {}); mechanism = record.get("possible_damage_mechanism", {})
        story = [Paragraph("RESRUCTAI / STRUCTURAL REVIEW", styles["Title"]), Paragraph(title, styles["Heading2"]), Paragraph(f"{record['component_type']} · {record['integrity_status']} · {record['created_at'][:10]}", styles["Normal"]), Spacer(1, 16), Paragraph("QUICK ASSESSMENT", styles["Heading3"]), Paragraph(f"Element: {quick.get('element', record['component_type'])} | Condition: {quick.get('apparent_condition', 'Insufficient evidence')} | Severity: {quick.get('severity', record.get('damage_classification', 'Insufficient evidence'))} | Confidence: {confidence.get('overall', 'Low')} ({confidence.get('percentage', 0)}%)", styles["BodyText"]), Paragraph("EVIDENCE & CONFIDENCE", styles["Heading3"]), Paragraph("Observed / provided: " + "; ".join(confidence.get("evidence_detected", [])), styles["BodyText"]), Paragraph("Estimated / inferred: " + "; ".join(confidence.get("estimated_inferred", [])), styles["BodyText"]), Paragraph("Cannot be determined: " + "; ".join(confidence.get("cannot_be_determined", [])), styles["BodyText"]), Paragraph("POSSIBLE DAMAGE MECHANISM", styles["Heading3"]), Paragraph(f"Most likely: {mechanism.get('most_likely', 'Cannot be determined')} | Alternatives: {'; '.join(mechanism.get('alternatives', []))} | Confidence: {mechanism.get('confidence', 'Low')}", styles["BodyText"]), Paragraph("IMMEDIATE ACTIONS", styles["Heading3"]), ListFlowable([ListItem(Paragraph(x, styles["BodyText"])) for x in record.get("immediate_actions", [])], bulletType="bullet")]
        for item in record.get("recommendations", []):
            story += [Spacer(1, 12), Paragraph(f"{item.get('rank')}. {item.get('technique', '')} — {item.get('urgency', '')}", styles["Heading3"]), Paragraph(item.get("reasoning", ""), styles["BodyText"]), Paragraph("Evidence link: " + item.get("evidence_link", ""), styles["BodyText"]), Paragraph("Appropriate when: " + item.get("appropriate_when", ""), styles["BodyText"]), Paragraph("Do not use when: " + item.get("do_not_use_when", ""), styles["BodyText"]), Paragraph("Materials: " + "; ".join(item.get("materials", [])), styles["BodyText"]), Paragraph("Procedure: " + "; ".join(item.get("procedure", [])), styles["BodyText"]), Paragraph("Safety: " + "; ".join(item.get("safety_notes", [])), styles["BodyText"]), Paragraph("IS code considerations: " + "; ".join(item.get("is_code_considerations", [])), styles["BodyText"])]
        story += [Spacer(1, 16), Paragraph("LIMITATIONS OF IMAGE-BASED ASSESSMENT", styles["Heading3"]), Paragraph("; ".join(record.get("limitations", [])), styles["BodyText"]), Paragraph("REQUIRED VERIFICATION", styles["Heading3"]), Paragraph("; ".join(record.get("required_verification", [])), styles["BodyText"]), Paragraph("PROFESSIONAL REVIEW REQUIRED", styles["Heading3"]), Paragraph(record["disclaimer"], styles["BodyText"])]
        doc.build(story)
        media_type = "application/pdf"
    elif file_format == "docx":
        document = Document()
        document.add_heading("RESRUCTAI / STRUCTURAL REVIEW", 0); document.add_heading(title, 1)
        document.add_paragraph(f"{record['component_type']} · {record['integrity_status']} · {record['created_at'][:10]}")
        quick = record.get("quick_assessment", {}); confidence = record.get("evidence_confidence", {}); mechanism = record.get("possible_damage_mechanism", {})
        document.add_heading("Quick assessment", 2); document.add_paragraph(f"Element: {quick.get('element', record['component_type'])} | Condition: {quick.get('apparent_condition', 'Insufficient evidence')} | Severity: {record.get('damage_classification', 'Insufficient evidence')} | Confidence: {confidence.get('overall', 'Low')} ({confidence.get('percentage', 0)}%)")
        document.add_heading("Evidence & confidence", 2); document.add_paragraph("Observed / provided: " + "; ".join(confidence.get("evidence_detected", []))); document.add_paragraph("Estimated / inferred: " + "; ".join(confidence.get("estimated_inferred", []))); document.add_paragraph("Cannot be determined: " + "; ".join(confidence.get("cannot_be_determined", [])))
        document.add_heading("Possible damage mechanism", 2); document.add_paragraph(f"Most likely: {mechanism.get('most_likely', 'Cannot be determined')} | Alternatives: {'; '.join(mechanism.get('alternatives', []))} | Confidence: {mechanism.get('confidence', 'Low')}")
        document.add_heading("Immediate actions", 2)
        for action in record.get("immediate_actions", []): document.add_paragraph(action, style="List Bullet")
        for item in record.get("recommendations", []):
            document.add_heading(f"{item.get('rank')}. {item.get('technique', '')} — {item.get('urgency', '')}", 2)
            document.add_paragraph(item.get("reasoning", "")); document.add_paragraph("Evidence link: " + item.get("evidence_link", "")); document.add_paragraph("Appropriate when: " + item.get("appropriate_when", "")); document.add_paragraph("Do not use when: " + item.get("do_not_use_when", "")); document.add_paragraph("Materials: " + "; ".join(item.get("materials", []))); document.add_paragraph("Procedure: " + "; ".join(item.get("procedure", []))); document.add_paragraph("Safety: " + "; ".join(item.get("safety_notes", []))); document.add_paragraph("IS code considerations: " + "; ".join(item.get("is_code_considerations", [])))
        document.add_heading("Limitations of image-based assessment", 2); document.add_paragraph("; ".join(record.get("limitations", []))); document.add_heading("Required verification", 2); document.add_paragraph("; ".join(record.get("required_verification", []))); document.add_heading("Professional review required", 2); document.add_paragraph(record["disclaimer"]); document.save(buffer)
        media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    else:
        workbook = Workbook(); sheet = workbook.active; sheet.title = "Assessment"
        sheet.append(["ResructAI structural review", record["structure_name"]]); sheet.append(["Component", record["component_type"]]); sheet.append(["Integrity", record["integrity_status"]]); sheet.append([])
        sheet.append(["Assessment status", record.get("assessment_status", "Insufficient evidence")]); sheet.append(["Damage classification", record.get("damage_classification", "Insufficient evidence")]); confidence = record.get("evidence_confidence", {}); sheet.append(["Confidence", f"{confidence.get('overall', 'Low')} ({confidence.get('percentage', 0)}%)"]); sheet.append(["Evidence detected", "\n".join(confidence.get("evidence_detected", []))]); sheet.append(["Estimated / inferred", "\n".join(confidence.get("estimated_inferred", []))]); sheet.append(["Cannot be determined", "\n".join(confidence.get("cannot_be_determined", []))]); sheet.append([])
        sheet.append(["Rank", "Technique", "Urgency", "Reasoning", "Evidence link", "Appropriate when", "Do not use when", "Materials", "Procedure", "Safety notes", "IS code considerations"])
        for item in record.get("recommendations", []): sheet.append([item.get("rank"), item.get("technique"), item.get("urgency"), item.get("reasoning"), item.get("evidence_link"), item.get("appropriate_when"), item.get("do_not_use_when"), "\n".join(item.get("materials", [])), "\n".join(item.get("procedure", [])), "\n".join(item.get("safety_notes", [])), "\n".join(item.get("is_code_considerations", []))])
        sheet.append([]); sheet.append(["Required verification", "\n".join(record.get("required_verification", []))]); sheet.append(["Limitations", "\n".join(record.get("limitations", []))]); sheet.append(["Disclaimer", record["disclaimer"]]); sheet.column_dimensions["A"].width = 24; sheet.column_dimensions["B"].width = 30
        workbook.save(buffer); media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    buffer.seek(0)
    return StreamingResponse(buffer, media_type=media_type, headers={"Content-Disposition": f'attachment; filename="{report_filename(record, file_format)}"'})

@api_router.post("/status", response_model=StatusCheck)
async def create_status_check(input: StatusCheckCreate):
    status_dict = input.model_dump()
    status_obj = StatusCheck(**status_dict)
    
    # Convert to dict and serialize datetime to ISO string for MongoDB
    doc = status_obj.model_dump()
    doc['timestamp'] = doc['timestamp'].isoformat()
    
    _ = await db.status_checks.insert_one(doc)
    return status_obj

@api_router.get("/status", response_model=List[StatusCheck])
async def get_status_checks():
    # Exclude MongoDB's _id field from the query results
    status_checks = await db.status_checks.find({}, {"_id": 0}).to_list(1000)
    
    # Convert ISO string timestamps back to datetime objects
    for check in status_checks:
        if isinstance(check['timestamp'], str):
            check['timestamp'] = datetime.fromisoformat(check['timestamp'])
    
    return status_checks

# Include the router in the main app
app.include_router(api_router)

app.add_middleware(
    CORSMiddleware,
    allow_credentials=True,
    allow_origins=os.environ.get('CORS_ORIGINS', '*').split(','),
    allow_methods=["*"],
    allow_headers=["*"],
)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

@app.on_event("shutdown")
async def shutdown_db_client():
    client.close()