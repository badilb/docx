import os
import uuid
from typing import List

import uvicorn
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse

import config
from docx_service import DocxService

app = FastAPI(title="DocX QR Footer Service")


@app.post("/generate-pdf/", summary="Принять один или несколько .docx → вернуть PDF с QR в футере")
async def generate_pdf(
    templates: List[UploadFile] = File(..., description="Word-шаблоны (.docx)"),
):
    if not os.path.exists(config.QR_CODE_PATH):
        raise HTTPException(status_code=500, detail=f"QR-код не найден: {config.QR_CODE_PATH}")

    for t in templates:
        if not t.filename.endswith(".docx"):
            raise HTTPException(status_code=400, detail=f"Файл '{t.filename}' не является .docx")

    session_id = uuid.uuid4().hex
    session_dir = os.path.join(config.OUTPUT_DIR, session_id)
    os.makedirs(session_dir, exist_ok=True)

    logo = config.LOGO_PATH if os.path.exists(config.LOGO_PATH) else None
    service = DocxService(output_dir=session_dir)

    # Один файл — отдаём PDF напрямую
    if len(templates) == 1:
        try:
            tmpl_path = await _save_upload(templates[0], session_dir)
            pdf_path = service.process_and_convert(tmpl_path, config.QR_CODE_PATH, logo)
            return FileResponse(path=pdf_path, media_type="application/pdf", filename=os.path.basename(pdf_path))
        except (RuntimeError, FileNotFoundError) as e:
            raise HTTPException(status_code=500, detail=str(e))

    # Несколько файлов — возвращаем ссылки
    results, errors = [], []
    for tmpl in templates:
        try:
            tmpl_path = await _save_upload(tmpl, session_dir)
            pdf_path = service.process_and_convert(tmpl_path, config.QR_CODE_PATH, logo)
            results.append({
                "template": tmpl.filename,
                "pdf_url": f"/download/{session_id}/{os.path.basename(pdf_path)}",
            })
        except Exception as e:
            errors.append({"template": tmpl.filename, "error": str(e)})

    return JSONResponse(content={"session_id": session_id, "generated": results, "errors": errors})


@app.get("/download/{session_id}/{filename}", response_class=FileResponse)
async def download_pdf(session_id: str, filename: str):
    if ".." in session_id or ".." in filename:
        raise HTTPException(status_code=400, detail="Недопустимый путь")
    file_path = os.path.join(config.OUTPUT_DIR, session_id, filename)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="Файл не найден")
    return FileResponse(path=file_path, media_type="application/pdf", filename=filename)


async def _save_upload(upload: UploadFile, dest_dir: str) -> str:
    dest_path = os.path.join(dest_dir, upload.filename)
    with open(dest_path, "wb") as f:
        f.write(await upload.read())
    return dest_path


if __name__ == "__main__":
    uvicorn.run("main:app", host="0.0.0.0", port=8500, reload=True)