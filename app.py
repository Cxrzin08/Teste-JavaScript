import os
from flask import Flask, request, render_template, send_file, url_for
from docx import Document
from docx2pdf import convert as word_to_pdf
import pdfplumber
from PyPDF2 import PdfReader
from pdf2docx import Converter as Pdf2DocxConverter
import fitz  # PyMuPDF

# Configuração do Flask
app = Flask(__name__)
UPLOAD_FOLDER = os.path.join(os.getcwd(), "uploads")  # Caminho absoluto
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html", download_link=None)

@app.route("/convert", methods=["POST"])
def convert():
    file = request.files.get("file")
    conversion_type = request.form.get("conversionType")

    if not file or not conversion_type:
        return "Arquivo ou tipo de conversão não selecionado.", 400

    # Salvar o arquivo enviado
    input_filename = file.filename
    input_path = os.path.join(UPLOAD_FOLDER, input_filename)
    file.save(input_path)

    # Definir o caminho de saída
    output_filename = f"converted_{input_filename.rsplit('.', 1)[0]}"
    output_path = None
    try:
        if conversion_type == "pdf-to-word":
            if not input_path.lower().endswith(".pdf"):
                return "Erro: Para PDF para Word, o arquivo enviado deve ser um PDF.", 400
            output_path = os.path.join(UPLOAD_FOLDER, output_filename + ".docx")
            convert_pdf_to_word(input_path, output_path)
        elif conversion_type == "word-to-pdf":
            if not input_path.lower().endswith(".docx"):
                return "Erro: Para Word para PDF, o arquivo enviado deve ser um DOCX.", 400
            output_path = os.path.join(UPLOAD_FOLDER, output_filename + ".pdf")
            if not convert_word_to_pdf(input_path, output_path):
                return "Erro ao converter o arquivo Word para PDF.", 500
        else:
            return "Tipo de conversão inválido.", 400
    except Exception as e:
        return f"Erro inesperado: {str(e)}", 500

    # Gerar link para download
    download_link = url_for("download_file", filename=os.path.basename(output_path))
    return render_template("index.html", download_link=download_link)

@app.route("/download/<filename>")
def download_file(filename):
    file_path = os.path.join(UPLOAD_FOLDER, filename)
    if not os.path.exists(file_path):
        return "Arquivo não encontrado.", 404
    return send_file(file_path, as_attachment=True)

def convert_pdf_to_word(input_path, output_path):
    errors = []
    try:
        # 1. Tentar pdfplumber
        with pdfplumber.open(input_path) as pdf:
            document = Document()
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    document.add_paragraph(text)
                else:
                    document.add_paragraph("Página vazia ou sem texto legível.")
            document.save(output_path)
            print("Conversão bem-sucedida com pdfplumber.")
            return
    except Exception as e:
        errors.append(f"Erro com pdfplumber: {str(e)}")

    try:
        # 2. Tentar pdf2docx
        converter = Pdf2DocxConverter(input_path)
        converter.convert(output_path)
        converter.close()
        print("Conversão bem-sucedida com pdf2docx.")
        return
    except Exception as e:
        errors.append(f"Erro com pdf2docx: {str(e)}")

    try:
        # 3. Tentar PyMuPDF
        pdf_document = fitz.open(input_path)
        document = Document()
        for page in pdf_document:
            text = page.get_text("text")
            document.add_paragraph(text if text.strip() else "[Página sem texto legível]")
        document.save(output_path)
        print("Conversão bem-sucedida com PyMuPDF.")
        return
    except Exception as e:
        errors.append(f"Erro com PyMuPDF: {str(e)}")

    raise Exception(f"Todas as tentativas falharam: {errors}")

def convert_word_to_pdf(input_path, output_path):
    try:
        word_to_pdf(input_path)
        os.rename(input_path.replace(".docx", ".pdf"), output_path)
        return True
    except Exception as e:
        print(f"Erro ao converter Word para PDF: {str(e)}")
        return False

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=True)