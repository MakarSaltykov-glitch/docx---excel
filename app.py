import io
import os
import subprocess
import tempfile
from flask import Flask, render_template, request, send_file, flash, redirect
from docx import Document
from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image as PILImage

app = Flask(__name__)
app.secret_key = "super_secret_key_for_flash_messages"  # Нужно для работы уведомлений


def convert_to_docx(input_path):
    if input_path.endswith(".docx"):
        return input_path

    output_dir = os.path.dirname(input_path)

    # Внимание: на сервере должен быть установлен soffice (LibreOffice)
    subprocess.run([
        "soffice",
        "--headless",
        "--convert-to", "docx",
        "--outdir", output_dir,
        input_path
    ])

    new_path = os.path.splitext(input_path)[0] + ".docx"
    return new_path


def extract_content_to_stream(word_path):
    """Модифицированная функция: возвращает файл Excel в виде потока байт в памяти"""
    word_path = convert_to_docx(word_path)

    doc = Document(word_path)
    wb = Workbook()
    ws = wb.active

    row_offset = 1

    # --- Таблицы ---
    for table in doc.tables:
        for i, row in enumerate(table.rows):
            for j, cell in enumerate(row.cells):
                ws.cell(row=row_offset + i, column=j + 1, value=cell.text)
        row_offset += len(table.rows) + 2

    # --- Картинки ---
    img_index = 1

    with tempfile.TemporaryDirectory() as temp_dir:
        for rel in doc.part.rels.values():
            if "image" in rel.target_ref:
                try:
                    img_data = rel.target_part.blob
                    image = PILImage.open(io.BytesIO(img_data))

                    temp_path = os.path.join(temp_dir, f"temp_img_{img_index}.png")
                    image.save(temp_path)

                    xl_img = OpenpyxlImage(temp_path)
                    ws.row_dimensions[row_offset].height = 80
                    ws.add_image(xl_img, f"A{row_offset}")

                    row_offset += 6
                    img_index += 1
                except Exception as e:
                    print(f"Пропущено изображение: {e}")
                    continue

        # Сохраняем результат в виртуальный файл в памяти
        excel_stream = io.BytesIO()
        wb.save(excel_stream)
        excel_stream.seek(0)  # Сбрасываем указатель в начало файла
        return excel_stream


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # Проверяем, прикрепил ли пользователь файл
        if "word_file" not in request.files:
            flash("Файл не найден в запросе")
            return redirect(request.url)

        file = request.files["word_file"]

        if file.filename == "":
            flash("Файл не выбран")
            return redirect(request.url)

        if file:
            # Создаем временный файл для входящего документа Word
            with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(file.filename)[1]) as temp_word:
                file.save(temp_word.name)
                temp_word_path = temp_word.name

            try:
                # Обрабатываем файл
                excel_stream = extract_content_to_stream(temp_word_path)

                # Формируем имя для скачивания
                base_name = os.path.splitext(file.filename)[0]
                output_filename = f"{base_name}.xlsx"

                # Удаляем временный файл Word после обработки
                os.unlink(temp_word_path)

                # Отправляем Excel-файл пользователю в браузер
                return send_file(
                    excel_stream,
                    attachment_filename=output_filename, # Для старых версий Flask
                    download_name=output_filename,       # Для новых версий Flask
                    as_attachment=True,
                    mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                if os.path.exists(temp_word_path):
                    os.unlink(temp_word_path)
                flash(f"Произошла ошибка при обработке: {e}")
                return redirect(request.url)

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)
