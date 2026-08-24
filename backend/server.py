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
    prompt = f"""You are a senior structural rehabilitation advisor supporting work in India. Analyze this blast-damage assessment for a qualified engineer. Never claim a final design or approve occupancy. Return ONLY valid JSON with this exact shape:
{{"immediate_actions":["..."],"recommendations":[{{"rank":1,"technique":"...","fit":"...","reasoning":"...","urgency":"Critical|High|Moderate|Monitor","materials":["..."],"procedure":["..."],"safety_notes":["..."],"timeframe":"...","is_code_considerations":["..." ]}}]}}
Provide 3 ranked, technically plausible approaches and make uncertainty explicit. If integrity is unsafe, put isolation and an engineer inspection first. For each option, identify applicable Indian Standard families only when relevant (for example IS 456 for reinforced concrete, IS 15988 for seismic evaluation/strengthening, IS 13920 for ductile detailing, IS 800 for structural steel, or IS 516 for material testing). Do not invent clause numbers; say that the current edition and project-specific provisions must be verified by the engineer.
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
    record = {
        "id": str(uuid.uuid4()), "structure_name": payload.structure_name,
        "component_type": payload.component_type, "integrity_status": payload.integrity_status,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "recommendations": advice.get("recommendations", []),
        "immediate_actions": advice.get("immediate_actions", []),
        "disclaimer": "Decision support only. A licensed structural engineer must inspect, verify calculations, and approve all work before action.",
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
        story = [Paragraph("RESRUCTAI / STRUCTURAL REVIEW", styles["Title"]), Paragraph(title, styles["Heading2"]), Paragraph(f"{record['component_type']} · {record['integrity_status']} · {record['created_at'][:10]}", styles["Normal"]), Spacer(1, 16), Paragraph("IMMEDIATE ACTIONS", styles["Heading3"]), ListFlowable([ListItem(Paragraph(x, styles["BodyText"])) for x in record.get("immediate_actions", [])], bulletType="bullet")]
        for item in record.get("recommendations", []):
            story += [Spacer(1, 12), Paragraph(f"{item.get('rank')}. {item.get('technique', '')} — {item.get('urgency', '')}", styles["Heading3"]), Paragraph(item.get("reasoning", ""), styles["BodyText"]), Paragraph("Materials: " + "; ".join(item.get("materials", [])), styles["BodyText"]), Paragraph("Procedure: " + "; ".join(item.get("procedure", [])), styles["BodyText"]), Paragraph("Safety: " + "; ".join(item.get("safety_notes", [])), styles["BodyText"]), Paragraph("IS code considerations: " + "; ".join(item.get("is_code_considerations", [])), styles["BodyText"])]
        story += [Spacer(1, 16), Paragraph("PROFESSIONAL REVIEW REQUIRED", styles["Heading3"]), Paragraph(record["disclaimer"], styles["BodyText"])]
        doc.build(story)
        media_type = "application/pdf"
    elif file_format == "docx":
        document = Document()
        document.add_heading("RESRUCTAI / STRUCTURAL REVIEW", 0); document.add_heading(title, 1)
        document.add_paragraph(f"{record['component_type']} · {record['integrity_status']} · {record['created_at'][:10]}")
        document.add_heading("Immediate actions", 2)
        for action in record.get("immediate_actions", []): document.add_paragraph(action, style="List Bullet")
        for item in record.get("recommendations", []):
            document.add_heading(f"{item.get('rank')}. {item.get('technique', '')} — {item.get('urgency', '')}", 2)
            document.add_paragraph(item.get("reasoning", "")); document.add_paragraph("Materials: " + "; ".join(item.get("materials", []))); document.add_paragraph("Procedure: " + "; ".join(item.get("procedure", []))); document.add_paragraph("Safety: " + "; ".join(item.get("safety_notes", []))); document.add_paragraph("IS code considerations: " + "; ".join(item.get("is_code_considerations", [])))
        document.add_heading("Professional review required", 2); document.add_paragraph(record["disclaimer"]); document.save(buffer)
        media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    else:
        workbook = Workbook(); sheet = workbook.active; sheet.title = "Assessment"
        sheet.append(["ResructAI structural review", record["structure_name"]]); sheet.append(["Component", record["component_type"]]); sheet.append(["Integrity", record["integrity_status"]]); sheet.append([])
        sheet.append(["Rank", "Technique", "Urgency", "Reasoning", "Materials", "Procedure", "Safety notes", "IS code considerations"])
        for item in record.get("recommendations", []): sheet.append([item.get("rank"), item.get("technique"), item.get("urgency"), item.get("reasoning"), "\n".join(item.get("materials", [])), "\n".join(item.get("procedure", [])), "\n".join(item.get("safety_notes", [])), "\n".join(item.get("is_code_considerations", []))])
        sheet.append([]); sheet.append(["Disclaimer", record["disclaimer"]]); sheet.column_dimensions["A"].width = 12; sheet.column_dimensions["B"].width = 30
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