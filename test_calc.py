import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from modules.calculator import Calculator

def run_tests():
    calc = Calculator()
    
    # Ищем подходящие ID в базе
    recon_id = None
    drill_id = None
    for wt in calc.work_types.get("work_types", []):
        if wt.get("group") == "reconnaissance" and recon_id is None:
            recon_id = wt["id"]
        # Колонковое бурение (менее 160 мм)
        if wt.get("group") == "drilling" and drill_id is None:
            if "колонков" in wt.get("name", "").lower() and "168" not in wt.get("name", ""):
                 drill_id = wt["id"]
    
    print(f"Используем ID рекогносцировки: {recon_id}")
    print(f"Используем ID бурения: {drill_id}\n")

    print("=== ТЕСТ 1: Рекогносцировка (двухкомпонентная формула) ===")
    if recon_id:
        qty = 2.5 # га
        item = calc.create_work_item(recon_id, qty)
        print(f"Ожидаем ненулевую сумму по формуле ПЗ1п + ПЗ2п * S")
        print(f"Формула работы: {item.formula}")
        print(f"Получено: {float(item.total_cost)}\n")
    
    print("=== ТЕСТ 2: Технический отчёт (Таблица 65) ===")
    cost, desc = calc.calculate_report_cost(45000, "II")
    print(f"Камералка 45 тыс., Категория II -> {cost} руб. ({desc})")
    cost, desc = calc.calculate_report_cost(800000, "I")
    print(f"Камералка 800 тыс., Категория I -> {cost} руб. ({desc})")
    cost, desc = calc.calculate_report_cost(4000000, "III")
    print(f"Камералка 4 млн., Категория III -> {cost} руб. ({desc})\n")

    print("=== ТЕСТ 3: Корректирующие коэффициенты (К1 и К2) ===")
    if drill_id:
        item_normal = calc.create_work_item(drill_id, 10, climate_zone="IV", is_local_work=False)
        print(f"Обычные условия (10 м): {float(item_normal.total_cost)}")
        
        item_north = calc.create_work_item(drill_id, 10, climate_zone="II", is_local_work=False)
        print(f"Крайний Север (К2=1.1): {float(item_north.total_cost)} (Коэфф: {item_north.coefficients.get('K2_climate')})")
        
        item_local = calc.create_work_item(drill_id, 10, climate_zone="IV", is_local_work=True)
        print(f"Местная работа (К1=0.88): {float(item_local.total_cost)} (Коэфф: {item_local.coefficients.get('K1_local')})")

if __name__ == '__main__':
    run_tests()
