import fitz  # PyMuPDF para leer PDF
from docx import Document
from deep_translator import GoogleTranslator
import re

def clean_text(text):
    # Elimina caracteres no válidos para XML/Word
    if text is None:
        return ""
    # Sustituye caracteres de control y nulos
    return re.sub(r'[\x00-\x08\x0B-\x0C\x0E-\x1F]', '', text)

def pdf_to_word(pdf_path, word_path, target_lang="es"):
    # Abrir PDF
    doc_pdf = fitz.open(pdf_path)
    doc_word = Document()

    translator = GoogleTranslator(source="auto", target=target_lang)

    for page_num in range(len(doc_pdf)):
        page = doc_pdf[page_num]
        text = page.get_text("text")

        # Limpiar texto antes de traducir
        text = clean_text(text)

        # Traducir texto al español
        translated_text = translator.translate(text)

        # Limpiar texto traducido también
        translated_text = clean_text(translated_text)

        # Agregar al documento Word
        doc_word.add_paragraph(translated_text)

    # Guardar archivo Word
    doc_word.save(word_path)
    print(f"Archivo Word generado en: {word_path}")

# Ejemplo de uso
pdf_to_word("DP-700 Asociado Ingeniero de datos de Fabric.pdf", "DP-700_espanol.docx")
