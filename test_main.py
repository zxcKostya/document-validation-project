import os
from fastapi.testclient import TestClient
from docx import Document
from docx.shared import Mm, Pt
from main import app

client = TestClient(app)

def create_test_docx(filename: str, valid: bool = True):
    doc = Document()
    section = doc.sections[0]
    if valid:
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.left_margin = Mm(30)
        section.right_margin = Mm(10)
        section.top_margin = Mm(20)
        section.bottom_margin = Mm(20)
        
        p = doc.add_paragraph("Министерство науки и высшего образования РФ\nТУСУР\nКафедра ТУ\nОтчет по практике\nТомск 2026")
        for run in p.runs:
            run.font.name = "Times New Roman"
            run.font.size = Pt(14)
    else:
        # Формат Letter и неверные поля/шрифт
        section.page_width = Mm(216)
        section.page_height = Mm(279)
        section.left_margin = Mm(15)
        p = doc.add_paragraph("Текст без титульника и неверным шрифтом")
        for run in p.runs:
            run.font.name = "Arial"
            run.font.size = Pt(12)
            
    doc.save(filename)

def test_valid_document():
    test_file = "test_valid.docx"
    create_test_docx(test_file, valid=True)
    with open(test_file, "rb") as f:
        response = client.post("/api/check", files={"file": (test_file, f, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")})
    os.remove(test_file)
    assert response.status_code == 200
    assert response.json()["status"] == "Passed"

def test_invalid_document():
    test_file = "test_invalid.docx"
    create_test_docx(test_file, valid=False)
    with open(test_file, "rb") as f:
        response = client.post("/api/check", files={"file": (test_file, f, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")})
    os.remove(test_file)
    data = response.json()
    assert response.status_code == 200
    assert data["status"] == "Failed"
    assert len(data["violations"]) >= 4
