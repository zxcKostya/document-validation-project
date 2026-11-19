import os
import datetime
from fastapi import FastAPI, UploadFile, File, Depends
from fastapi.responses import HTMLResponse, JSONResponse
from docx import Document
from docx.shared import Mm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from sqlalchemy import create_engine, Column, Integer, String, DateTime, JSON
from sqlalchemy.orm import declarative_base, sessionmaker, Session

# ==========================================
# Настройка Базы Данных (SQLAlchemy)
# ==========================================
DATABASE_URL = "sqlite:///./docx_checker.db"
engine = create_engine(DATABASE_URL, connect_args={"check_same_thread": False})
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

class ValidationHistory(Base):
    __tablename__ = "validation_history"
    
    id = Column(Integer, primary_key=True, index=True)
    filename = Column(String, nullable=False)
    timestamp = Column(DateTime, default=datetime.datetime.utcnow)
    status = Column(String, nullable=False)
    errors = Column(JSON, nullable=False)

Base.metadata.create_all(bind=engine)

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

# ==========================================
# Модуль автоматических проверок DOCX
# ==========================================
def check_docx_formatting(file_path: str) -> list:
    violations = []
    try:
        doc = Document(file_path)
    except Exception as e:
        return [f"Не удалось прочитать файл как DOCX: {str(e)}"]

    # Получаем базовый стиль документа (Normal) для проверки наследования
    normal_style = doc.styles['Normal'] if 'Normal' in doc.styles else None
    default_font_name = normal_style.font.name if normal_style else None
    default_font_size = normal_style.font.size if normal_style else None

    # 1. Проверка формата страницы (А4: 210x297 мм) и точных полей
    for i, section in enumerate(doc.sections):
        # Допуск 1 мм на погрешности конвертации Word
        if abs(section.page_width - Mm(210)) > Mm(1) or abs(section.page_height - Mm(297)) > Mm(1):
            violations.append(f"Раздел {i+1}: Формат страницы не соответствует А4 (210x297 мм).")

        # Точные пороги с микродопуском 0.2 мм на округление
        if section.top_margin < Mm(19.8):
            violations.append(f"Раздел {i+1}: Верхнее поле меньше нормы (должно быть >= 20 мм).")
        if section.bottom_margin < Mm(19.8):
            violations.append(f"Раздел {i+1}: Нижнее поле меньше нормы (должно быть >= 20 мм).")
        if section.left_margin < Mm(29.8):
            violations.append(f"Раздел {i+1}: Левое поле меньше нормы (должно быть >= 30 мм).")
        if section.right_margin < Mm(9.8):
            violations.append(f"Раздел {i+1}: Правое поле меньше нормы (должно быть >= 10 мм).")

    # 2. Проверка шрифтов (включая наследование) и заголовков
    for p_idx, paragraph in enumerate(doc.paragraphs):
        text = paragraph.text.strip()
        if not text:
            continue

        style_name = paragraph.style.name if paragraph.style else ""
        
        # Проверка заголовков (Heading)
        if style_name.startswith("Heading") or style_name.startswith("Заголовок"):
            if text.endswith("."):
                violations.append(f"Абзац {p_idx+1} (Заголовок): Точка в конце заголовка не допускается по ГОСТ.")
            if paragraph.alignment not in [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER, None]:
                violations.append(f"Абзац {p_idx+1} (Заголовок): Некорректное выравнивание заголовка.")

        # Проверка шрифтов с учетом каскада: Run -> Paragraph Style -> Normal Style
        for run in paragraph.runs:
            if not run.text.strip():
                continue
            
            font_name = run.font.name or (paragraph.style.font.name if paragraph.style else None) or default_font_name
            if font_name != "Times New Roman":
                actual_font = font_name if font_name else "по умолчанию (не задан явно)"
                violations.append(f"Абзац {p_idx+1}: Некорректный шрифт '{actual_font}' вместо Times New Roman.")
                break
                
            font_size = run.font.size or (paragraph.style.font.size if paragraph.style else None) or default_font_size
            if font_size != Pt(14):
                actual_size = f"{font_size.pt}pt" if font_size else "по умолчанию (не задан явно)"
                violations.append(f"Абзац {p_idx+1}: Неверный размер шрифта {actual_size} (ожидается 14pt).")
                break

    # 3. Расширенная проверка структуры титульного листа (первые 25 абзацев)
    if len(doc.paragraphs) > 0:
        title_text = " ".join([p.text for p in doc.paragraphs[:25]]).lower()
        required_elements = {
            "реквизиты вуза (ТУСУР / Министерство)": ["тусур", "министерство"],
            "указание кафедры": ["кафедра"],
            "тип работы (Отчет / Задание)": ["отчет", "отчёт", "задание"],
            "город и год (Томск)": ["томск"]
        }
        for element_name, keywords in required_elements.items():
            if not any(kw in title_text for kw in keywords):
                violations.append(f"Титульный лист: Отсутствует обязательный элемент — {element_name}.")

    return violations

# ==========================================
# Инициализация API и GUI (FastAPI)
# ==========================================
app = FastAPI(title="Автоматизированная проверка отчетов DOCX")

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <title>Проверка DOCX отчетов</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 40px; background-color: #f4f6f9; color: #333; }
        .container { max-width: 800px; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        h1 { color: #1a5c96; }
        .upload-box { border: 2px dashed #1a5c96; padding: 30px; text-align: center; border-radius: 6px; cursor: pointer; background: #f9fbfd; }
        button { background: #1a5c96; color: white; border: none; padding: 10px 20px; font-size: 16px; border-radius: 4px; cursor: pointer; margin-top: 15px; }
        button:hover { background: #134470; }
        .result { margin-top: 20px; padding: 15px; border-radius: 4px; display: none; }
        .success { background-color: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
        .danger { background-color: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
        ul { padding-left: 20px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Автоматическая проверка оформления отчетов (DOCX)</h1>
        <p>Выберите файл отчета для автоматической проверки формата страницы, полей, шрифтов, заголовков и структуры титульного листа.</p>
        
        <form id="uploadForm" enctype="multipart/form-data">
            <div class="upload-box">
                <input type="file" id="fileInput" name="file" accept=".docx" required><br>
            </div>
            <button type="submit">Запустить проверку</button>
        </form>

        <div id="resultBox" class="result"></div>
    </div>

    <script>
        document.getElementById('uploadForm').onsubmit = async (e) => {
            e.preventDefault();
            const fileInput = document.getElementById('fileInput');
            const formData = new FormData();
            formData.append('file', fileInput.files[0]);

            const resultBox = document.getElementById('resultBox');
            resultBox.style.display = 'block';
            resultBox.className = 'result';
            resultBox.innerHTML = 'Обработка и анализ документа...';

            try {
                const response = await fetch('/api/check', {
                    method: 'POST',
                    body: formData
                });
                const data = await response.json();

                if (data.status === "Passed") {
                    resultBox.className = 'result success';
                    resultBox.innerHTML = '<b>Успешно!</b> Ошибок оформления не обнаружено.';
                } else {
                    resultBox.className = 'result danger';
                    let html = '<b>Обнаружены нарушения оформления:</b><ul>';
                    data.violations.forEach(v => { html += `<li>${v}</li>`; });
                    html += '</ul>';
                    resultBox.innerHTML = html;
                }
            } catch (err) {
                resultBox.className = 'result danger';
                resultBox.innerText = 'Ошибка при отправке файла на сервер.';
            }
        };
    </script>
</body>
</html>
"""

@app.get("/", response_class=HTMLResponse)
def get_gui():
    return HTML_TEMPLATE

@app.post("/api/check")
async def api_check_document(file: UploadFile = File(...), db: Session = Depends(get_db)):
    temp_path = f"temp_{file.filename}"
    with open(temp_path, "wb") as buffer:
        buffer.write(await file.read())
    
    violations = check_docx_formatting(temp_path)
    
    if os.path.exists(temp_path):
        os.remove(temp_path)
        
    status = "Passed" if len(violations) == 0 else "Failed"
    
    history_entry = ValidationHistory(
        filename=file.filename,
        status=status,
        errors=violations
    )
    db.add(history_entry)
    db.commit()
    
    return JSONResponse(content={
        "filename": file.filename,
        "status": status,
        "violations": violations,
        "timestamp": datetime.datetime.now().isoformat()
    })

@app.get("/api/history")
def get_history(db: Session = Depends(get_db)):
    history = db.query(ValidationHistory).order_by(ValidationHistory.id.desc()).all()
    return [{
        "id": h.id,
        "filename": h.filename,
        "timestamp": h.timestamp.isoformat(),
        "status": h.status,
        "errors": h.errors
    } for h in history]

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port = 8000)