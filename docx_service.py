import os
import json
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT


class DocxService:
    def __init__(self, templates, json_file, output_dir):

        self.templates = templates
        self.json_file = json_file
        self.output_dir = output_dir

        # Загружаем данные из JSON-файла
        with open(self.json_file, "r", encoding="utf-8") as file:
            self.replacements = json.load(file)

    def replace_text_in_paragraph(self, paragraph, data):
        full_text = "".join(run.text for run in paragraph.runs)
        modified_text = full_text

        for key, value in data.items():
            modified_text = modified_text.replace(f"[{key}]", str(value))

        if full_text != modified_text:
            for run in paragraph.runs:
                run.clear()

            lines = modified_text.split("\n")
            for i, line in enumerate(lines):
                new_run = paragraph.add_run(line.strip())
                new_run.font.name = "Calibri"
                if i < len(lines) - 1:
                    new_run.add_break()

    def process_document(self, doc):
        """Проходит по абзацам и таблицам документа и выполняет замену текста."""
        for paragraph in doc.paragraphs:
            self.replace_text_in_paragraph(paragraph, self.replacements)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        self.replace_text_in_paragraph(paragraph, self.replacements)

    def create_document(self, template_path, output_filename, qr_code_path):
        """Создает документ с QR-кодом, логотипом и рамкой в подвале."""
        doc = Document(template_path)

        section = doc.sections[0]
        section.footer_distance = Cm(0.5) # Немного отступим от края
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(2)

        footer = section.footer
        
        # Очищаем старые параграфы в футере, чтобы не было лишних отступов
        for p in footer.paragraphs:
            p_element = p._element
            p_element.getparent().remove(p_element)

        # Создаем таблицу для красивой рамки (1 строка, 2 колонки)
        table = footer.add_table(rows=1, cols=2, width=Inches(7))
        table.autofit = False
        
        # --- МАГИЯ РАМКИ (Пунктирная или сплошная) ---
        # Чтобы сделать рамку как на картинке, нужно задать границы ячейкам
        for cell in table.cells:
            for side in ['top', 'left', 'bottom', 'right']:
                from docx.oxml.shared import qn
                from docx.oxml import OxmlElement
                tcPr = cell._element.get_or_add_tcPr()
                borders = tcPr.find(qn('w:tcBorders'))
                if borders is None:
                    borders = OxmlElement('w:tcBorders')
                    tcPr.append(borders)
                border = OxmlElement(f'w:{side}')
                border.set(qn('w:val'), 'dotted') # 'dotted' для пунктира или 'single' для сплошной
                border.set(qn('w:sz'), '4')       # толщина
                border.set(qn('w:color'), 'A6A6A6') # серый цвет
                borders.append(border)

        qr_cell = table.cell(0, 0)
        text_cell = table.cell(0, 1)

        qr_cell.width = Inches(0.9)
        text_cell.width = Inches(5.5)

        # 1. Вставляем QR-код
        qr_paragraph = qr_cell.paragraphs[0]
        qr_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run_qr = qr_paragraph.add_run()
        run_qr.add_picture(qr_code_path, width=Inches(0.7))

        # 2. Вставляем текст и логотип
        text_paragraph = text_cell.paragraphs[0]
        text_paragraph.paragraph_format.space_before = Pt(2)
        
        # Первая строка
        run1 = text_paragraph.add_run("Подписи ЭЦП проверены НУЦ РК")
        run1.font.name = "Calibri"
        run1.font.size = Pt(8)
        
        # Перенос на вторую строку
        text_paragraph.add_run("\n") 
        
        # Вторая строка: Текст + ЛОГОТИП
        run2 = text_paragraph.add_run("Документ подписан в сервис ")
        run2.font.name = "Calibri"
        run2.font.size = Pt(8)

        logo_path = "uploads/dcx logo png.png"
        if os.path.exists(logo_path):
            run_logo = text_paragraph.add_run()
            # Вставляем картинку прямо в строку
            run_logo.add_picture(logo_path, width=Inches(0.4)) 
        else:
            run_err = text_paragraph.add_run("dcx.kz")
            run_err.font.name = "Calibri"
            run_err.font.size = Pt(8)

        # Обработка основного текста документа (замена данных из ERPNext)
        self.process_document(doc)

        output_path = os.path.join(self.output_dir, output_filename)
        doc.save(output_path)
        return output_path
    
    
    def generate_documents(self, qr_code_path):
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)

        created_files = []
        for template in self.templates:
            if os.path.exists(template):
                output_filename = f"processed_{os.path.basename(template)}"
                path = self.create_document(template, output_filename, qr_code_path)
                created_files.append(path)

        return created_files


