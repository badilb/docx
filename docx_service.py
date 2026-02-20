import os
import docx
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
        """Создает новый документ на основе шаблона и заменяет текст."""
        doc = Document(template_path)

        section = doc.sections[0]
        section.footer_distance = Cm(0)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(2)

        footer = section.footer
        footer.add_paragraph()
        table = footer.add_table(rows=1, cols=2, width=Inches(9))
        table.autofit = False

        qr_cell = table.cell(0, 0)
        text_cell = table.cell(0, 1)

        qr_cell.width = Inches(1)
        text_cell.width = Inches(6)

        if footer.paragraphs:
            for paragraph in footer.paragraphs:
                if not paragraph.text.strip():
                    p = paragraph._element
                    p.getparent().remove(p)


        qr_paragraph = qr_cell.paragraphs[0]
        run = qr_paragraph.add_run()
        run.add_picture(qr_code_path, width=Inches(0.7))

        text_paragraph = text_cell.paragraphs[0]
        text_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        text_paragraph.paragraph_format.left_indent = Inches(0)
        text_paragraph.paragraph_format.space_before = Pt(0)
        text_paragraph.paragraph_format.space_after = Pt(0)

        run = text_paragraph.add_run(
            "Данный документ подписан электронной цифровой подписью Удостоверяющего центра НУЦ Республики Казахстан "
            "в системе электронного документооборота ТОО 'Рога и Копыта' «DocX». Проверить подлинность электронного "
            "документа Вы можете по ссылке https://tesla.kz/verify, указав идентификатор – iF9jFZPp."
        )
        run.font.name = "Calibri"
        run.font.size = Pt(8)

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


