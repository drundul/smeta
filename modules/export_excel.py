"""
Экспорт сметы в Excel
"""

from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter


def export_to_excel(estimate, filename: str = None) -> Path:
    """Экспортировать смету в Excel"""
    
    if filename is None:
        filename = f"Смета_{estimate.project_name}_{estimate.date_created}.xlsx"
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Смета ИГИ"
    
    # Стили
    header_font = Font(bold=True, size=12)
    title_font = Font(bold=True, size=14)
    money_format = '#,##0.00 ₽'
    
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    header_fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
    subtotal_fill = PatternFill(start_color="FFF3E0", end_color="FFF3E0", fill_type="solid")
    total_fill = PatternFill(start_color="C8E6C9", end_color="C8E6C9", fill_type="solid")
    
    # Заголовок сметы
    ws.merge_cells('A1:G1')
    ws['A1'] = "ЛОКАЛЬНАЯ СМЕТА"
    ws['A1'].font = title_font
    ws['A1'].alignment = Alignment(horizontal='center')
    
    ws.merge_cells('A2:G2')
    ws['A2'] = f"на инженерно-геологические изыскания"
    ws['A2'].alignment = Alignment(horizontal='center')
    
    # Информация о проекте
    row = 4
    project_info = [
        ("Проект:", estimate.project_name),
        ("Шифр:", estimate.project_code or "-"),
        ("Объект:", estimate.object_name or "-"),
        ("Заказчик:", estimate.customer or "-"),
        ("Подрядчик:", estimate.contractor or "-"),
        ("Дата:", estimate.date_created),
        ("Индекс пересчёта:", f"{float(estimate.price_index):.2f}"),
    ]
    
    for label, value in project_info:
        ws[f'A{row}'] = label
        ws[f'A{row}'].font = Font(bold=True)
        ws[f'B{row}'] = value
        row += 1
    
    row += 1
    
    # Заголовок таблицы
    headers = ["№ п/п", "Наименование работ", "Ед. изм.", "Кол-во", "Обоснование", "Расчёт", "Стоимость, руб."]
    col_widths = [8, 50, 10, 10, 20, 25, 18]
    
    for col, (header, width) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=row, column=col, value=header)
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
        cell.fill = header_fill
        ws.column_dimensions[get_column_letter(col)].width = width
    
    row += 1
    start_data_row = row
    
    # Группировка по категориям
    categories = {
        "field": {"name": "ПОЛЕВЫЕ РАБОТЫ", "items": []},
        "laboratory": {"name": "ЛАБОРАТОРНЫЕ РАБОТЫ", "items": []},
        "office": {"name": "КАМЕРАЛЬНЫЕ РАБОТЫ", "items": []}
    }
    
    for item in estimate.items:
        # Определяем категорию по коду
        if item.code.startswith(("01", "02", "03", "04")):
            categories["field"]["items"].append(item)
        elif item.code.startswith(("05", "06", "07")):
            categories["laboratory"]["items"].append(item)
        else:
            categories["office"]["items"].append(item)
    
    item_num = 1
    
    for cat_key, cat_data in categories.items():
        if not cat_data["items"]:
            continue
        
        # Заголовок раздела
        ws.merge_cells(f'A{row}:G{row}')
        ws[f'A{row}'] = cat_data["name"]
        ws[f'A{row}'].font = Font(bold=True)
        ws[f'A{row}'].fill = header_fill
        for col in range(1, 8):
            ws.cell(row=row, column=col).border = thin_border
        row += 1
        
        # Позиции
        for item in cat_data["items"]:
            ws.cell(row=row, column=1, value=item_num).border = thin_border
            # Наименование
            ws.cell(row=row, column=2, value=item.name).border = thin_border
            ws.cell(row=row, column=2).alignment = Alignment(wrap_text=True)
            
            # Ед. изм.
            ws.cell(row=row, column=3, value=item.unit).border = thin_border
            ws.cell(row=row, column=3).alignment = Alignment(horizontal='center')
            
            # Кол-во
            qty_cell = ws.cell(row=row, column=4, value=float(item.quantity))
            qty_cell.border = thin_border
            qty_cell.alignment = Alignment(horizontal='right')
            
            # Обоснование (НЗ, таблица)
            ref_text = f"НЗ №281/пр, {item.table_ref}" if item.table_ref else "НЗ №281/пр"
            ref_cell = ws.cell(row=row, column=5, value=ref_text)
            ref_cell.border = thin_border
            ref_cell.alignment = Alignment(wrap_text=True, vertical='center')

            
            # Расчёт (Формула)
            formula_text = item.formula if item.formula else f"{float(item.base_cost):,.0f} x {float(item.quantity):,.1f}"
            calc_cell = ws.cell(row=row, column=6, value=formula_text)
            calc_cell.border = thin_border
            calc_cell.alignment = Alignment(horizontal='right')
            
            # Стоимость
            total_cell = ws.cell(row=row, column=7, value=float(item.total_cost))
            total_cell.border = thin_border
            total_cell.number_format = money_format
            
            item_num += 1
            row += 1
        
        # Подитог раздела
        subtotal = sum(float(item.total_cost) for item in cat_data["items"])
        ws.merge_cells(f'A{row}:F{row}')
        ws[f'A{row}'] = f"Итого по разделу «{cat_data['name'].lower()}»:"
        ws[f'A{row}'].font = Font(bold=True)
        ws[f'A{row}'].alignment = Alignment(horizontal='right')
        for col in range(1, 7):
            ws.cell(row=row, column=col).border = thin_border
            ws.cell(row=row, column=col).fill = subtotal_fill
        
        subtotal_cell = ws.cell(row=row, column=7, value=subtotal)
        subtotal_cell.font = Font(bold=True)
        subtotal_cell.number_format = money_format
        subtotal_cell.border = thin_border
        subtotal_cell.fill = subtotal_fill
        row += 1
    
    # -------------------------------------------------------------
    # ИТОГИ
    # -------------------------------------------------------------
    
    # 1. Базовые затраты (Сумма всех работ)
    base_total = sum(float(item.total_cost) for item in estimate.items)
    
    row += 1
    ws.merge_cells(f'A{row}:F{row}')
    ws[f'A{row}'] = "ИТОГО базовые затраты (СП + СЛ + СК):"
    ws[f'A{row}'].font = Font(bold=True)
    ws[f'A{row}'].alignment = Alignment(horizontal='right')
    for col in range(1, 7):
        ws.cell(row=row, column=col).border = thin_border
        ws.cell(row=row, column=col).fill = subtotal_fill
        
    total_cell = ws.cell(row=row, column=7, value=base_total)
    total_cell.font = Font(bold=True)
    total_cell.number_format = money_format
    total_cell.border = thin_border
    total_cell.fill = subtotal_fill
    
    # 2. Дополнительные затраты (построчно)
    dz_sum = 0
    if estimate.additional_costs:
        row += 1
        ws.merge_cells(f'A{row}:G{row}')
        ws[f'A{row}'] = "Дополнительные затраты:"
        ws[f'A{row}'].font = Font(bold=True, italic=True)
        ws[f'A{row}'].border = thin_border
        row += 1
        
        dz_item_num = 1
        for cost in estimate.additional_costs:
            # 1. № п/п (можно продолжить или свой)
            ws.cell(row=row, column=1, value=f"ДЗ-{dz_item_num}").border = thin_border
            
            # 2. Наименование
            ws.cell(row=row, column=2, value=cost.get('name', 'ДЗ')).border = thin_border
            ws.cell(row=row, column=2).alignment = Alignment(wrap_text=True)
            
            # 3. Ед. изм.
            ws.cell(row=row, column=3, value="-").border = thin_border
            ws.cell(row=row, column=3).alignment = Alignment(horizontal='center')
            
            # 4. Кол-во
            ws.cell(row=row, column=4, value="-").border = thin_border
            ws.cell(row=row, column=4).alignment = Alignment(horizontal='center')
            
            # 5. Обоснование
            ws.cell(row=row, column=5, value=cost.get('basis', '-')).border = thin_border
            ws.cell(row=row, column=5).alignment = Alignment(wrap_text=True)
            
            # 6. Расчёт
            ws.cell(row=row, column=6, value=cost.get('formula', '-')).border = thin_border
            ws.cell(row=row, column=6).alignment = Alignment(horizontal='right')
            
            # 7. Стоимость
            val = float(cost.get('value', 0))
            dz_sum += val
            val_cell = ws.cell(row=row, column=7, value=val)
            val_cell.number_format = money_format
            val_cell.border = thin_border
            
            row += 1
            dz_item_num += 1
            
    # 3. Итого с учетом ДЗ
    total_with_dz = base_total + dz_sum
    row += 1
    ws.merge_cells(f'A{row}:F{row}')
    ws[f'A{row}'] = "ИТОГО с учетом дополнительных затрат:"
    ws[f'A{row}'].font = Font(bold=True)
    ws[f'A{row}'].alignment = Alignment(horizontal='right')
    for col in range(1, 7):
        ws.cell(row=row, column=col).border = thin_border
        ws.cell(row=row, column=col).fill = subtotal_fill

    total_dz_cell = ws.cell(row=row, column=7, value=total_with_dz)
    total_dz_cell.font = Font(bold=True)
    total_dz_cell.number_format = money_format
    total_dz_cell.border = thin_border
    total_dz_cell.fill = subtotal_fill
    
    # 4. С индексом пересчета
    idx = float(estimate.price_index)
    total_indexed = total_with_dz * idx
    
    row += 1
    ws.merge_cells(f'A{row}:F{row}')
    ws[f'A{row}'] = f"ИТОГО с индексом пересчёта ({idx:.2f}):"
    ws[f'A{row}'].font = Font(bold=True)
    ws[f'A{row}'].alignment = Alignment(horizontal='right')
    for col in range(1, 7):
        ws.cell(row=row, column=col).border = thin_border
        ws.cell(row=row, column=col).fill = total_fill
    
    total_indexed_cell = ws.cell(row=row, column=7, value=total_indexed)
    total_indexed_cell.font = Font(bold=True)
    total_indexed_cell.number_format = money_format
    total_indexed_cell.border = thin_border
    total_indexed_cell.fill = total_fill

    # 5. Коэффициент договорной цены
    k_contract = float(estimate.contract_coefficient)
    if k_contract != 1.0:
        row += 1
        ws.merge_cells(f'A{row}:F{row}')
        ws[f'A{row}'] = f"Коэффициент договорной цены ({k_contract:.3f}):"
        ws[f'A{row}'].font = Font(bold=True)
        ws[f'A{row}'].alignment = Alignment(horizontal='right')
        for col in range(1, 7):
            ws.cell(row=row, column=col).border = thin_border
        
        # Здесь выводим просто сумму
        # Но если по образцу, то это применяется к итогу
        final_total = total_indexed * k_contract
        
        k_cell = ws.cell(row=row, column=7, value=final_total)
        k_cell.font = Font(bold=True)
        k_cell.number_format = money_format
        k_cell.border = thin_border
    else:
        final_total = total_indexed

    # ВСЕГО ПО СМЕТЕ
    row += 1
    ws.merge_cells(f'A{row}:F{row}')
    ws[f'A{row}'] = "ВСЕГО по смете:"
    ws[f'A{row}'].font = Font(bold=True, size=12)
    ws[f'A{row}'].alignment = Alignment(horizontal='right')
    for col in range(1, 7):
        ws.cell(row=row, column=col).border = thin_border
        ws.cell(row=row, column=col).fill = total_fill
    
    final_cell = ws.cell(row=row, column=7, value=final_total)
    final_cell.font = Font(bold=True, size=12)
    final_cell.number_format = money_format
    final_cell.border = thin_border
    final_cell.fill = total_fill
    
    # Подпись
    row += 3
    ws[f'A{row}'] = "Составил: __________________ / __________________ /"
    row += 2
    ws[f'A{row}'] = "Проверил: __________________ / __________________ /"
    
    # Сохранение
    output_path = Path(filename)
    wb.save(output_path)
    
    return output_path
