"""
Экспорт сметы в PDF
"""

from pathlib import Path
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os


def register_fonts():
    """Регистрация кириллических шрифтов (Windows + Linux)"""
    font_paths = [
        # Windows
        os.path.join(os.environ.get('WINDIR', 'C:\\Windows'), 'Fonts', 'arial.ttf'),
        os.path.join(os.environ.get('WINDIR', 'C:\\Windows'), 'Fonts', 'times.ttf'),
        # Linux / Streamlit Cloud (Debian-based)
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/usr/share/fonts/truetype/freefont/FreeSans.ttf"
    ]
    
    for path in font_paths:
        if os.path.exists(path):
            try:
                # Регистрируем основной шрифт
                pdfmetrics.registerFont(TTFont('CustomCyrillic', path))
                
                # Пытаемся найти жирную версию (эвристика)
                bold_path = None
                if "arial.ttf" in path.lower(): 
                    bold_path = path.replace("arial.ttf", "arialbd.ttf")
                elif "times.ttf" in path.lower(): 
                    bold_path = path.replace("times.ttf", "timesbd.ttf")
                elif "DejaVuSans.ttf" in path: 
                    bold_path = path.replace("DejaVuSans.ttf", "DejaVuSans-Bold.ttf")
                elif "LiberationSans-Regular.ttf" in path: 
                    bold_path = path.replace("LiberationSans-Regular.ttf", "LiberationSans-Bold.ttf")
                elif "FreeSans.ttf" in path:
                     bold_path = path.replace("FreeSans.ttf", "FreeSansBold.ttf")

                if bold_path and os.path.exists(bold_path):
                    pdfmetrics.registerFont(TTFont('CustomCyrillic-Bold', bold_path))
                else:
                    # Фоллбэк: используем обычный шрифт как жирный
                    pdfmetrics.registerFont(TTFont('CustomCyrillic-Bold', path))
                    
                return 'CustomCyrillic'
            except Exception as e:
                print(f"Font loading error ({path}): {e}")
                continue
    
    # Если шрифты не найдены, используем встроенный (без кириллицы)
    return 'Helvetica'


def export_to_pdf(estimate, filename: str = None) -> Path:
    """Экспортировать смету в PDF"""
    
    if filename is None:
        filename = f"Смета_{estimate.project_name}_{estimate.date_created}.pdf"
    
    output_path = Path(filename)
    
    # Регистрируем шрифты
    font_name = register_fonts()
    
    # Создаём документ
    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        rightMargin=15*mm,
        leftMargin=15*mm,
        topMargin=15*mm,
        bottomMargin=15*mm
    )
    
    # Стили
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'Title',
        parent=styles['Heading1'],
        fontName=font_name,
        fontSize=14,
        alignment=1,  # center
        spaceAfter=5*mm
    )
    
    subtitle_style = ParagraphStyle(
        'Subtitle',
        parent=styles['Normal'],
        fontName=font_name,
        fontSize=11,
        alignment=1,
        spaceAfter=10*mm
    )
    
    normal_style = ParagraphStyle(
        'CustomNormal',
        parent=styles['Normal'],
        fontName=font_name,
        fontSize=9
    )
    
    bold_style = ParagraphStyle(
        'CustomBold',
        parent=styles['Normal'],
        fontName=font_name if font_name == 'Helvetica' else f'{font_name}-Bold',
        fontSize=9
    )
    
    elements = []
    
    # Заголовок
    template_label = f" ({estimate.template_name})" if getattr(estimate, "template_name", "") else ""
    elements.append(Paragraph(f"ЛОКАЛЬНАЯ СМЕТА{template_label}", title_style))
    elements.append(Paragraph("на инженерно-геологические изыскания", subtitle_style))
    
    # Информация о проекте
    project_data = [
        ["Проект:", estimate.project_name],
        ["Шифр:", estimate.project_code or "-"],
        ["Объект:", estimate.object_name or "-"],
        ["Заказчик:", estimate.customer or "-"],
        ["Дата:", estimate.date_created],
        ["Индекс пересчёта:", f"{float(estimate.price_index):.2f}"],
    ]
    
    project_table = Table(project_data, colWidths=[40*mm, 130*mm])
    project_table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (0, -1), font_name if font_name == 'Helvetica' else f'{font_name}-Bold'),
        ('FONTNAME', (1, 0), (1, -1), font_name),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))
    
    elements.append(project_table)
    elements.append(Spacer(1, 10*mm))
    
    # Таблица сметы
    header = ["№", "Код", "Наименование", "Ед.", "Кол-во", "Цена", "Сумма"]
    col_widths = [8*mm, 12*mm, 70*mm, 12*mm, 15*mm, 22*mm, 25*mm]
    
    table_data = [header]
    
    # Группировка
    categories = {
        "field": {"name": "ПОЛЕВЫЕ РАБОТЫ", "items": []},
        "laboratory": {"name": "ЛАБОРАТОРНЫЕ РАБОТЫ", "items": []},
        "office": {"name": "КАМЕРАЛЬНЫЕ РАБОТЫ", "items": []}
    }
    
    for item in estimate.items:
        if item.code.startswith(("01", "02", "03", "04")):
            categories["field"]["items"].append(item)
        elif item.code.startswith(("05", "06", "07")):
            categories["laboratory"]["items"].append(item)
        else:
            categories["office"]["items"].append(item)
    
    row_styles = []
    current_row = 1
    item_num = 1
    
    for cat_key, cat_data in categories.items():
        if not cat_data["items"]:
            continue
        
        # Заголовок раздела
        table_data.append([cat_data["name"], "", "", "", "", "", ""])
        row_styles.append(('SPAN', (0, current_row), (6, current_row)))
        row_styles.append(('BACKGROUND', (0, current_row), (6, current_row), colors.lightgrey))
        row_styles.append(('FONTNAME', (0, current_row), (6, current_row), 
                          font_name if font_name == 'Helvetica' else f'{font_name}-Bold'))
        current_row += 1
        
        # Позиции
        for item in cat_data["items"]:
            # Обрезаем длинные названия
            name = item.name[:50] + "..." if len(item.name) > 50 else item.name
            
            table_data.append([
                str(item_num),
                item.code,
                name,
                item.unit,
                f"{float(item.quantity):.1f}",
                f"{float(item.unit_cost):,.0f}",
                f"{float(item.total_cost):,.0f}"
            ])
            item_num += 1
            current_row += 1
        
        # Подитог
        subtotal = sum(float(item.total_cost) for item in cat_data["items"])
        table_data.append(["", "", f"Итого {cat_data['name'].lower()}:", "", "", "", f"{subtotal:,.0f}"])
        row_styles.append(('SPAN', (0, current_row), (1, current_row)))
        row_styles.append(('BACKGROUND', (0, current_row), (6, current_row), colors.Color(1, 0.95, 0.9)))
        row_styles.append(('FONTNAME', (2, current_row), (6, current_row),
                          font_name if font_name == 'Helvetica' else f'{font_name}-Bold'))
        current_row += 1
    
    # Итого
    base_total = sum(float(item.total_cost) for item in estimate.items)
    table_data.append(["", "", "ИТОГО (в ценах 01.01.2024):", "", "", "", f"{base_total:,.0f}"])
    row_styles.append(('SPAN', (0, current_row), (1, current_row)))
    row_styles.append(('BACKGROUND', (0, current_row), (6, current_row), colors.Color(0.9, 0.95, 0.9)))
    row_styles.append(('FONTNAME', (2, current_row), (6, current_row),
                      font_name if font_name == 'Helvetica' else f'{font_name}-Bold'))
    current_row += 1
    
    # С индексом
    table_data.append(["", "", f"ВСЕГО с индексом ({float(estimate.price_index):.2f}):", "", "", "", f"{float(estimate.total):,.0f}"])
    row_styles.append(('SPAN', (0, current_row), (1, current_row)))
    row_styles.append(('BACKGROUND', (0, current_row), (6, current_row), colors.Color(0.8, 0.9, 0.8)))
    row_styles.append(('FONTNAME', (0, current_row), (6, current_row),
                      font_name if font_name == 'Helvetica' else f'{font_name}-Bold'))
    
    # Создаём таблицу
    main_table = Table(table_data, colWidths=col_widths, repeatRows=1)
    
    base_style = [
        ('FONTNAME', (0, 0), (-1, -1), font_name),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('FONTNAME', (0, 0), (-1, 0), font_name if font_name == 'Helvetica' else f'{font_name}-Bold'),
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('ALIGN', (0, 1), (0, -1), 'CENTER'),
        ('ALIGN', (3, 1), (3, -1), 'CENTER'),
        ('ALIGN', (4, 1), (6, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]
    
    main_table.setStyle(TableStyle(base_style + row_styles))
    
    elements.append(main_table)
    
    # Подписи
    elements.append(Spacer(1, 15*mm))
    elements.append(Paragraph("Составил: __________________ / __________________ /", normal_style))
    elements.append(Spacer(1, 8*mm))
    elements.append(Paragraph("Проверил: __________________ / __________________ /", normal_style))
    
    # Генерируем PDF
    doc.build(elements)
    
    return output_path
