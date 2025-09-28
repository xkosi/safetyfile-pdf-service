# server.py
from fastapi import FastAPI, Request
from fastapi.responses import FileResponse, JSONResponse
import tempfile
import subprocess
import os
import json

app = FastAPI()

@app.post("/generate")
async def generate(request: Request):
    try:
        body = await request.json()
    except Exception:
        return JSONResponse({"error": "Invalid JSON"}, status_code=400)

    # tijdelijk JSON-bestand
    with tempfile.NamedTemporaryFile(delete=False, suffix=".json", mode="w", encoding="utf-8") as f:
        json.dump(body, f, indent=2, ensure_ascii=False)
        preview_path = f.name

    out_path = preview_path.replace(".json", ".pdf")

    try:
        # run jouw generate_pdf.py script
        subprocess.run(
            ["python3", "generate_pdf.py", "--preview", preview_path, "--out", out_path],
            check=True
        )
    except subprocess.CalledProcessError as e:
        return JSONResponse({"error": "PDF generation failed", "detail": str(e)}, status_code=500)

    return FileResponse(out_path, media_type="application/pdf", filename="dossier.pdf")
