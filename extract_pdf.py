"""
Скрипт для извлечения таблиц нормативных затрат из PDF
Приказ Минстроя №281/пр от 12.05.2025

Запуск:
pip install pymupdf
python extract_pdf.py
"""

import fitz  # PyMuPDF
import re
import json
from pathlib import Path


def extract_text_from_pdf(pdf_path: str) -> str:
    """Извлечь весь текст из PDF"""
    doc = fitz.open(pdf_path)
    full_text = ""
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        text = page.get_text()
        full_text += f"\n--- Страница {page_num + 1} ---\n{text}"
    
    doc.close()
    return full_text


def save_text_to_file(text: str, output_path: str):
    """Сохранить текст в файл для анализа"""
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(text)
    print(f"Текст сохранён в: {output_path}")


def main():
    # Путь к PDF
    pdf_path = Path(__file__).parent / "Приказ  Минстроя России от 12-05-2025 N 281пр (1).pdf"
    
    if not pdf_path.exists():
        print(f"Файл не найден: {pdf_path}")
        return
    
    print(f"Извлечение текста из: {pdf_path}")
    
    # Извлекаем текст
    text = extract_text_from_pdf(str(pdf_path))
    
    # Сохраняем для анализа
    output_path = Path(__file__).parent / "pdf_extracted_text.txt"
    save_text_to_file(text, str(output_path))
    
    print(f"\nИзвлечено {len(text)} символов")
    print("Теперь можно проанализировать pdf_extracted_text.txt")


if __name__ == "__main__":
    main()
