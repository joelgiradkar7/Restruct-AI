import os
import requests
import base64
from io import BytesIO
from PIL import Image


BASE_URL = os.environ["TEST_BASE_URL"].rstrip("/")

PAYLOAD = {
    "structure_name": "TEST_image_review",
    "component_type": "Reinforced concrete beam",
    "construction_material": "Reinforced concrete",
    "blast_context": "External blast approximately 12 m east; inspection after event.",
    "crack_width": "0.3–1 mm", "spalling_depth": "Under 25 mm",
    "reinforcement_exposure": "Not visible", "integrity_status": "Needs assessment",
    "site_notes": "Restrict access pending engineer inspection.",
}

def image_data(fmt, uniform=False):
    image = Image.new("RGB", (32, 32), "white")
    if not uniform:
        for x in range(8, 24):
            for y in range(7, 25):
                image.putpixel((x, y), (220, 30, 30))
    output = BytesIO(); image.save(output, format=fmt)
    return base64.b64encode(output.getvalue()).decode()


def test_live_assessment_create_and_persist():
    payload = {
        "structure_name": "TEST_live_beam_review",
        "component_type": "Reinforced concrete beam",
        "construction_material": "Reinforced concrete",
        "blast_context": "External blast approximately 12 m east; inspection after event.",
        "crack_width": "0.3–1 mm",
        "spalling_depth": "Under 25 mm",
        "reinforcement_exposure": "Not visible",
        "integrity_status": "Needs assessment",
        "site_notes": "Restrict access pending engineer inspection.",
    }
    response = requests.post(f"{BASE_URL}/api/assessments", json=payload, timeout=180)
    assert response.status_code == 200, response.text
    record = response.json()
    assert record["structure_name"] == payload["structure_name"]
    assert len(record["recommendations"]) == 3
    assert record["immediate_actions"]
    for item in record["recommendations"]:
        for key in ("rank", "technique", "reasoning", "urgency", "materials", "procedure", "safety_notes", "timeframe"):
            assert item.get(key) not in (None, "", [])

    history = requests.get(f"{BASE_URL}/api/assessments", timeout=30)
    assert history.status_code == 200
    assert any(item["id"] == record["id"] for item in history.json())


def test_allowed_nonuniform_image_formats_are_accepted():
    for fmt, mime in (("JPEG", "image/jpeg"), ("PNG", "image/png"), ("WEBP", "image/webp")):
        response = requests.post(f"{BASE_URL}/api/assessments", json={**PAYLOAD, "structure_name": f"TEST_{fmt}", "image_base64": image_data(fmt), "image_mime": mime}, timeout=180)
        assert response.status_code == 200, f"{fmt}: {response.text}"
        assert len(response.json()["recommendations"]) == 3


def test_malformed_and_uniform_images_are_rejected():
    for encoded in ("not-base64", image_data("PNG", uniform=True)):
        response = requests.post(f"{BASE_URL}/api/assessments", json={**PAYLOAD, "image_base64": encoded, "image_mime": "image/png"}, timeout=30)
        assert response.status_code == 422, response.text