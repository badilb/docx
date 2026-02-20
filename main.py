import json
import os
import config
import uvicorn
from fastapi import FastAPI, HTTPException
from fastapi.responses import JSONResponse, FileResponse
from pydantic import BaseModel
from docx_service import DocxService

app = FastAPI()

class GenerateRequest(BaseModel):
    json_data: dict

# Генерация документа
@app.post("/generate-documents/")
async def generate_documents(request: GenerateRequest):
    json_path = os.path.join(config.UPLOAD_DIR, "data.json")

    # Сохраняем JSON-файл
    with open(json_path, "w", encoding="utf-8") as json_file:
        json.dump(request.json_data, json_file, ensure_ascii=False, indent=4)

    # Список шаблонов в папке uploads
    templates = [
        os.path.join(config.UPLOAD_DIR, file)
        for file in os.listdir(config.UPLOAD_DIR)
        if file.endswith(".docx")
    ]

    if not templates:
        raise HTTPException(status_code=400, detail="Не найдено ни одного docx-шаблона!")

    # Генерируем документы
    doc_service = DocxService(templates, json_path, config.OUTPUT_DIR)
    doc_service.generate_documents(qr_code_path="uploads/2025-03-13 14.34.05.jpg")

    return JSONResponse(content={"message": "Документы сгенерированы!", "output_dir": config.OUTPUT_DIR})

# Получение сгенерированного документа
@app.get("/get-document/{filename}")
async def get_document(filename: str):
    filename = filename.strip()
    file_path = os.path.join(config.OUTPUT_DIR, filename)

    print(f"Пытаемся найти файл: {file_path}")  # Логируем путь
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="Файл не найден!")
    return FileResponse(file_path, filename=filename)


if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8500, reload=True)
