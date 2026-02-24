"""
Модуль расчёта сметной стоимости ИГИ
по Приказу Минстроя РФ №281/пр от 12.05.2025
"""

import json
from pathlib import Path
from decimal import Decimal, ROUND_HALF_UP
from dataclasses import dataclass, field
from typing import Optional, Dict, Any
import datetime


def load_json(filename: str) -> dict:
    """Загрузка JSON-файла из папки data"""
    base_path = Path(__file__).parent.parent / "data"
    with open(base_path / filename, "r", encoding="utf-8") as f:
        return json.load(f)


def get_nested_value(data: dict, key_path: str, default=None):
    """Получить значение по вложенному пути (a.b.c)"""
    keys = key_path.split(".")
    value = data
    for key in keys:
        if isinstance(value, dict) and key in value:
            value = value[key]
        else:
            return default
    return value


@dataclass
class WorkItem:
    """Позиция сметы"""
    work_id: str
    code: str
    name: str
    unit: str
    quantity: Decimal
    base_cost: Decimal
    coefficients: dict = field(default_factory=dict)
    total_coefficient: Decimal = Decimal("1.0")
    unit_cost: Decimal = Decimal("0")
    total_cost: Decimal = Decimal("0")
    notes: str = ""
    table_ref: str = ""
    formula: str = ""
    
    def calculate(self):
        """Рассчитать стоимость позиции"""
        self.total_coefficient = Decimal("1.0")
        for name, value in self.coefficients.items():
            self.total_coefficient *= Decimal(str(value))
        
        self.unit_cost = (self.base_cost * self.total_coefficient).quantize(
            Decimal("0.01"), rounding=ROUND_HALF_UP
        )
        self.total_cost = (self.unit_cost * self.quantity).quantize(
            Decimal("0.01"), rounding=ROUND_HALF_UP
        )
        return self


@dataclass
class Estimate:
    """Смета на ИГИ"""
    project_name: str
    project_code: str = ""
    object_name: str = ""
    customer: str = ""
    contractor: str = ""
    date_created: str = field(default_factory=lambda: datetime.date.today().isoformat())
    items: list = field(default_factory=list)
    global_coefficients: dict = field(default_factory=dict)
    price_index: Decimal = Decimal("1.0")
    additional_costs: list = field(default_factory=list)
    contract_coefficient: Decimal = Decimal("1.0")
    
    @property
    def base_total(self) -> Decimal:
        """Базовая стоимость работ без ДЗ"""
        return sum(item.total_cost for item in self.items)

    @property
    def total_with_dz(self) -> Decimal:
        """Стоимость с учетом дополнительных затрат (в базовых ценах)"""
        base = self.base_total
        dz_sum = sum(Decimal(str(item.get("value", 0))) for item in self.additional_costs)
        return base + dz_sum

    @property
    def total(self) -> Decimal:
        """Итоговая стоимость с учетом индекса и коэффициента договорной цены"""
        total_indexed = self.total_with_dz * self.price_index
        return (total_indexed * self.contract_coefficient).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    
    @property
    def subtotal_field(self) -> Decimal:
        """Итого по полевым работам"""
        return sum(
            item.total_cost for item in self.items 
            if self._get_work_category(item.work_id) == "field"
        )
    
    @property
    def subtotal_laboratory(self) -> Decimal:
        """Итого по лабораторным работам"""
        return sum(
            item.total_cost for item in self.items 
            if self._get_work_category(item.work_id) == "laboratory"
        )
    
    @property
    def subtotal_office(self) -> Decimal:
        """Итого по камеральным работам"""
        return sum(
            item.total_cost for item in self.items 
            if self._get_work_category(item.work_id) == "office"
        )
    

    
    def _get_work_category(self, work_id: str) -> str:
        """Получить категорию работы по ID"""
        work_types = load_json("work_types.json")
        for work in work_types.get("work_types", []):
            if work["id"] == work_id:
                return work.get("category", "")
        return ""
    
    def add_item(self, item: WorkItem):
        """Добавить позицию в смету"""
        item.calculate()
        self.items.append(item)
    
    def to_dict(self) -> dict:
        """Конвертировать в словарь для экспорта"""
        return {
            "project_name": self.project_name,
            "project_code": self.project_code,
            "object_name": self.object_name,
            "customer": self.customer,
            "contractor": self.contractor,
            "date_created": self.date_created,
            "price_index": float(self.price_index),
            "items": [
                {
                    "code": item.code,
                    "name": item.name,
                    "unit": item.unit,
                    "quantity": float(item.quantity),
                    "unit_cost": float(item.unit_cost),
                    "total_cost": float(item.total_cost),
                    "coefficients": item.coefficients,
                    "notes": item.notes
                }
                for item in self.items
            ],
            "subtotals": {
                "field": float(self.subtotal_field),
                "laboratory": float(self.subtotal_laboratory),
                "office": float(self.subtotal_office)
            },
            "total": float(self.total)
        }


class Calculator:
    """Калькулятор сметной стоимости ИГИ"""
    
    def __init__(self):
        self.work_types = load_json("work_types.json")
        self.normative_costs = load_json("normative_costs.json")
        self.coefficients = load_json("coefficients.json")
    
    def get_work_types_by_category(self, category: str = None) -> list:
        """Получить виды работ (опционально по категории)"""
        works = self.work_types.get("work_types", [])
        if category:
            works = [w for w in works if w.get("category") == category]
        return works
    
    def get_work_type(self, work_id: str) -> dict:
        """Получить вид работы по ID"""
        for work in self.work_types.get("work_types", []):
            if work["id"] == work_id:
                return work
        return {}
    
    def get_base_cost(self, work_id: str) -> Decimal:
        """Получить базовую стоимость по ID работы.
        
        Для рекогносцировки возвращает ПЗ2п (удельная стоимость за ед.).
        Для получения обеих компонент используйте get_reconnaissance_components().
        """
        work_type = self.get_work_type(work_id)
        
        # Прямая стоимость
        if "base_cost" in work_type:
            return Decimal(str(work_type["base_cost"]))
        
        # Стоимость по ключу из normative_costs
        if "cost_key" in work_type:
            cost_data = get_nested_value(self.normative_costs, work_type["cost_key"])
            if isinstance(cost_data, dict):
                # Двухкомпонентная формула (рекогносцировка):
                # СПреког = ПЗ1п + ПЗ2п × S (п.49, ф.16)
                # Возвращаем ПЗ2п как "удельную стоимость" — ПЗ1п добавляется отдельно
                pz2 = cost_data.get("PZ2p", 0)
                return Decimal(str(pz2))
            elif cost_data is not None:
                return Decimal(str(cost_data))
        
        return Decimal("0")
    
    def get_reconnaissance_components(self, work_id: str) -> tuple:
        """Получить компоненты ПЗ1п и ПЗ2п для рекогносцировки.
        
        Returns:
            (PZ1p, PZ2p) — постоянная и удельная часть стоимости.
            Формула: СПреког = ПЗ1п + ПЗ2п × Sреког (п.49 НЗ)
        """
        work_type = self.get_work_type(work_id)
        if "cost_key" in work_type:
            cost_data = get_nested_value(self.normative_costs, work_type["cost_key"])
            if isinstance(cost_data, dict):
                pz1 = Decimal(str(cost_data.get("PZ1p", 0)))
                pz2 = Decimal(str(cost_data.get("PZ2p", 0)))
                return (pz1, pz2)
        return (Decimal("0"), Decimal("0"))
    
    def is_reconnaissance(self, work_id: str) -> bool:
        """Проверить, является ли работа рекогносцировкой (двухкомпонентная формула)"""
        work_type = self.get_work_type(work_id)
        return work_type.get("group") == "reconnaissance" and "cost_key" in work_type
    
    def get_soil_coefficient(self, work_id: str, category: str) -> Decimal:
        """Получить коэффициент по категории грунта"""
        work_type = self.get_work_type(work_id)
        
        # Для работ с встроенной категорией грунта в названии
        if "_cat" in work_id:
            return Decimal("1.0")  # Уже учтено в базовой стоимости
        
        soil_coefs = work_type.get("soil_category_coefficients", {})
        return Decimal(str(soil_coefs.get(category, 1.0)))
    
    def get_climate_coefficient(self, zone: str) -> Decimal:
        """Получить коэффициент К2 климатической зоны (п.13 НЗ, Таблица 2)"""
        k2_data = self.coefficients.get("K2_climate", {}).get("values", {})
        zone_data = k2_data.get(zone, {})
        if isinstance(zone_data, dict):
            return Decimal(str(zone_data.get("value", 1.0)))
        return Decimal("1.0")
    
    def get_K1_coefficient(self, work_id: str) -> Decimal:
        """Получить коэффициент К1 для работы по месту постоянной работы (п.12 НЗ, Таблица 1).
        
        Определяет тип К1 на основе вида работы:
        - Бурение/зондирование ⌀≤160 мм: 0.88
        - Бурение ⌀>160 мм: 0.82
        - Рекогносцировка и прочие: 0.85
        """
        work_type = self.get_work_type(work_id)
        group = work_type.get("group", "")
        name = work_type.get("name", "").lower()
        k1_data = self.coefficients.get("K1_local_work", {}).get("values", {})
        
        # Бурение или зондирование
        if group in ("drilling", "field_tests"):
            # Проверяем диаметр > 160 мм
            if "168" in name or "273" in name or ">160" in name or "свыше 160" in name:
                return Decimal(str(k1_data.get("drilling_gt160mm", {}).get("value", 0.82)))
            else:
                return Decimal(str(k1_data.get("drilling_sounding_le160mm", {}).get("value", 0.88)))
        else:
            # Рекогносцировка и прочие полевые работы
            return Decimal(str(k1_data.get("reconnaissance_and_other", {}).get("value", 0.85)))
    
    def get_unfavorable_period_duration(self, region: str) -> float:
        """Получить продолжительность неблагоприятного периода (месяцы)"""
        regions = self.coefficients.get("unfavorable_periods_by_region", {}).get("regions", {})
        return regions.get(region, 6.0)
    
    def get_price_index(self, quarter: str = None) -> Decimal:
        """Получить индекс пересчёта цен"""
        if quarter is None:
            # Определяем текущий квартал
            today = datetime.date.today()
            quarter = f"{today.year}-Q{(today.month - 1) // 3 + 1}"
        
        # Индекс пока равен 1.0 (базовые цены на 01.01.2024)
        return Decimal("1.0")
    
    def calculate_additional_costs(
        self,
        field_cost: Decimal,
        region: str = None,
        distance_km: float = 0,
        winter_days: int = 0,
        night_days: int = 0
    ) -> Dict[str, Decimal]:
        """Рассчитать дополнительные затраты"""
        additional = {}
        
        # Дополнительные затраты на неблагоприятный период
        if region:
            duration = self.get_unfavorable_period_duration(region)
            unfav_coefs = self.coefficients.get("unfavorable_period", {}).get("coefficients_by_duration_months", {})
            
            # Находим подходящий диапазон
            for range_key, percents in unfav_coefs.items():
                if self._check_duration_range(duration, range_key):
                    # Определяем стоимостной диапазон
                    cost_key = self._get_cost_range_key(float(field_cost))
                    percent = percents.get(cost_key, 0)
                    additional["unfavorable_period"] = (field_cost * Decimal(str(percent)) / 100).quantize(
                        Decimal("0.01"), rounding=ROUND_HALF_UP
                    )
                    break
        
        # Дополнительные затраты на проезд
        if distance_km > 0:
            travel_coefs = self.coefficients.get("travel_costs_NZ", {}).get("coefficients_by_distance_km", {})
            for dist_key, percents in travel_coefs.items():
                if self._check_distance_range(distance_km, dist_key):
                    cost_key = self._get_cost_range_key(float(field_cost))
                    percent = percents.get(cost_key, 0)
                    if percent:
                        additional["travel"] = (field_cost * Decimal(str(percent)) / 100).quantize(
                            Decimal("0.01"), rounding=ROUND_HALF_UP
                        )
                    break
        
        return additional
    
    def _check_duration_range(self, duration: float, range_key: str) -> bool:
        """Проверить попадание в диапазон продолжительности"""
        ranges = {
            "up_to_3": (0, 3),
            "3_to_4": (3, 4),
            "4_to_5": (4, 5),
            "5_to_6": (5, 6),
            "6_to_7": (6, 7),
            "7_to_8": (7, 8),
            "8_to_9": (8, 9),
            "9_to_10": (9, 10),
            "over_10": (10, 100)
        }
        if range_key in ranges:
            min_val, max_val = ranges[range_key]
            return min_val <= duration < max_val
        return False
    
    def _check_distance_range(self, distance: float, range_key: str) -> bool:
        """Проверить попадание в диапазон расстояния"""
        ranges = {
            # Для проезда (Таблицы 4-7, 16-17)
            "up_to_200": (0, 200),
            "200_to_500": (200, 500),
            "500_to_1000": (500, 1000),
            "1000_to_2000": (1000, 2000),
            "2000_to_4000": (2000, 4000),
            "over_4000": (4000, 100000),
            
            # Для организации (Таблица 8/20)
            "up_to_100": (0, 100),
            "100_to_500": (100, 500)
        }
        if range_key in ranges:
            min_val, max_val = ranges[range_key]
            return min_val <= distance < max_val
        return False
    
    def _get_distance_range_boundaries(self, range_key: str) -> tuple:
        """Получить границы диапазона расстояния (для интерполяции)"""
        ranges = {
            "up_to_200": (0, 200),
            "200_to_500": (200, 500),
            "500_to_1000": (500, 1000),
            "1000_to_2000": (1000, 2000),
            "2000_to_4000": (2000, 4000),
            "over_4000": (4000, 100000),
            "up_to_100": (0, 100),
            "100_to_500": (100, 500)
        }
        return ranges.get(range_key, (0, 0))
    
    def _get_cost_range_key(self, cost: float) -> str:
        """Определить ключ стоимостного диапазона (общий)"""
        cost_k = cost / 1000  # В тысячах рублей
        if cost_k <= 300:
            return "up_to_300k"
        elif cost_k <= 500:
            return "up_to_500k"
        elif cost_k <= 1000:
            return "1000k"
        elif cost_k <= 2000:
            return "2000k"
        elif cost_k <= 5000:
            return "5000k"
        elif cost_k <= 10000:
            return "10000k"
        else:
            return "over_20000k"

    def _get_travel_cost_range_key(self, cost: float) -> str:
        """Определить ключ стоимостного диапазона для проезда (НЗ)"""
        cost_k = cost / 1000
        if cost_k <= 500:
            return "up_to_500k"
        elif cost_k <= 2000:
            return "2000k"
        elif cost_k <= 5000:
            return "5000k"
        elif cost_k <= 10000:
            return "10000k"
        else:
            return "over_20000k"
    
    def interpolate_coefficient(self, distance: float, coefs_by_distance: dict, cost_key: str) -> float:
        """Линейная интерполяция коэффициента проезда между значениями таблицы.
        
        Согласно п.160, примечание 3 Приказа 281/пр:
        промежуточные значения определяются методом линейной интерполяции.
        
        Args:
            distance: расстояние в км
            coefs_by_distance: словарь {range_key: {cost_key: percent}}
            cost_key: ключ стоимостного диапазона
        
        Returns:
            интерполированный процент
        """
        # Собираем отсортированные диапазоны с их опорными точками
        ordered_ranges = [
            ("up_to_200", 100),      # опорная точка — середина диапазона
            ("200_to_500", 350),
            ("500_to_1000", 750),
            ("1000_to_2000", 1500),
            ("2000_to_4000", 3000),
            ("over_4000", 5000),
        ]
        
        # Собираем точки (distance_midpoint, percent) для интерполяции
        points = []
        for rng_key, midpoint in ordered_ranges:
            if rng_key in coefs_by_distance:
                val = coefs_by_distance[rng_key].get(cost_key)
                if val is not None:
                    points.append((midpoint, val))
        
        if not points:
            return 0
        
        # Если расстояние <= первой точки или >= последней — экстраполяция не делается
        if distance <= points[0][0]:
            return points[0][1]
        if distance >= points[-1][0]:
            return points[-1][1]
        
        # Интерполяция между двумя ближайшими точками
        for i in range(len(points) - 1):
            x0, y0 = points[i]
            x1, y1 = points[i + 1]
            if x0 <= distance <= x1:
                # Линейная интерполяция
                ratio = (distance - x0) / (x1 - x0)
                return round(y0 + ratio * (y1 - y0), 2)
        
        return points[-1][1]
    
    def get_regional_coefficient(self, region: str) -> float:
        """Получить районный коэффициент (ПДЗр) по региону.
        
        Для регионов с подвариантами (например, 'Тюменская область (<60°)')
        ищет совпадение по базовому имени региона.
        
        Returns:
            Районный коэффициент (1.0 если не найден)
        """
        reg_data = self.coefficients.get("regional_allowances", {})
        reg_coefs = reg_data.get("regional_coefficients", {})
        
        # Точное совпадение
        if region in reg_coefs:
            val = reg_coefs[region]
            return val if isinstance(val, (int, float)) else 1.0
        
        # Поиск по базовому имени (без уточнений в скобках)
        for key, val in reg_coefs.items():
            if key.startswith("_"):
                continue
            if not isinstance(val, (int, float)):
                continue
            # Проверяем, содержит ли регион название из справочника
            if key in region or region in key:
                return val
        
        # Для городов — попробуем найти область
        # напр. "г. Москва" → "Московская область" => 1.0
        return 1.0
    
    def create_work_item(
        self,
        work_id: str,
        quantity: float,
        soil_category: str = "II",
        climate_zone: str = "III",
        additional_coefficients: dict = None,
        override_base_cost: Optional[Decimal] = None,
        formula: str = "",
        is_local_work: bool = False
    ) -> WorkItem:
        """Создать позицию сметы
        
        Args:
            is_local_work: Работы по месту постоянной работы (п.12, применяется К1)
        """
        work_type = self.get_work_type(work_id)
        
        if override_base_cost is not None:
             base_cost = Decimal(str(override_base_cost))
        else:
             base_cost = self.get_base_cost(work_id)
        
        coefficients = {}
        
        # Коэффициенты для полевых работ
        if work_type.get("category") == "field":
            soil_coef = self.get_soil_coefficient(work_id, soil_category)
            if soil_coef != Decimal("1.0"):
                coefficients["soil_category"] = float(soil_coef)
            
            # К2 — климатический коэффициент (п.13 НЗ, Таблица 2)
            climate_coef = self.get_climate_coefficient(climate_zone)
            if climate_coef != Decimal("1.0"):
                coefficients["K2_climate"] = float(climate_coef)
            
            # К1 — работы по месту постоянной работы (п.12 НЗ, Таблица 1)
            if is_local_work:
                k1_coef = self.get_K1_coefficient(work_id)
                if k1_coef != Decimal("1.0"):
                    coefficients["K1_local"] = float(k1_coef)
        
        # Дополнительные коэффициенты
        if additional_coefficients:
            coefficients.update(additional_coefficients)
        
        # Для рекогносцировки — двухкомпонентная формула (п.49, ф.16):
        # СПреког = ПЗ1п + ПЗ2п × Sреког
        # base_cost = PZ2p (удельная), PZ1p добавляется как фиксированная часть
        pz1p_fixed = Decimal("0")
        if self.is_reconnaissance(work_id) and override_base_cost is None:
            pz1p, pz2p = self.get_reconnaissance_components(work_id)
            pz1p_fixed = pz1p
            base_cost = pz2p
            if not formula:
                formula = f"ПЗ1п({pz1p:,.0f}) + ПЗ2п({pz2p:,.0f}) × {quantity:.1f}"
        
        item = WorkItem(
            work_id=work_id,
            code=work_type.get("code", ""),
            name=work_type.get("name", ""),
            unit=work_type.get("unit", ""),
            quantity=Decimal(str(quantity)),
            base_cost=base_cost,
            coefficients=coefficients,
            table_ref=work_type.get("table_ref", ""),
            formula=formula
        )
        
        # Рекогносцировка: после calculate() добавляем фиксированную часть ПЗ1п, умноженную на коэффициенты
        if pz1p_fixed > 0:
            item.calculate()
            item.total_cost = item.total_cost + (pz1p_fixed * item.total_coefficient).quantize(
                Decimal("0.01"), rounding=ROUND_HALF_UP)
            return item
        
        return item
    
    def create_estimate(
        self,
        project_name: str,
        items_data: list,
        soil_category: str = "II",
        climate_zone: str = "III",
        apply_price_index: bool = True,
        is_local_work: bool = False
    ) -> Estimate:
        """Создать смету
        
        items_data: список словарей вида {"work_id": "...", "quantity": 10, "override_base_cost": 123.45, "formula": "..."}
        is_local_work: Работы по месту постоянной работы (п.12 НЗ, применяется К1)
        """
        estimate = Estimate(project_name=project_name)
        
        if apply_price_index:
            estimate.price_index = self.get_price_index()
        
        for item_data in items_data:
            work_id = item_data.get("work_id")
            quantity = item_data.get("quantity", 0)
            additional_coefs = item_data.get("additional_coefficients", {})
            override_cost = item_data.get("override_base_cost")
            formula = item_data.get("formula", "")
            
            if work_id and quantity > 0:
                work_item = self.create_work_item(
                    work_id=work_id,
                    quantity=quantity,
                    soil_category=soil_category,
                    climate_zone=climate_zone,
                    additional_coefficients=additional_coefs,
                    override_base_cost=override_cost,
                    formula=formula,
                    is_local_work=is_local_work
                )
                estimate.add_item(work_item)
        
        return estimate
        
    def calculate_report_cost(self, cameral_sum: float, complexity: str) -> tuple[float, str]:
        """Рассчитать стоимость технического отчёта по Таблице 65.
        
        Args:
            cameral_sum: Суммарная стоимость камеральных работ
            complexity: Категория сложности (I, II, III)
            
        Returns:
            (стоимость, описание_диапазона)
        """
        table_65 = self.normative_costs.get("technical_report", {})
        comp_key = f"complexity_{complexity}"
        data = table_65.get(comp_key, {})
        
        if not data:
            return 0.0, "не найдено"
            
        # Правила разбора ключей (напр. "up_to_20k", "50k", "over_3500k")
        # Для простоты преобразуем ключи в (limit_value, cost_value, display_text)
        # Отсортируем по лимиту
        ranges = []
        for key, value in data.items():
            if key.startswith("up_to_"):
                k_val = int(key.replace("up_to_", "").replace("k", "")) * 1000
                display = f"до {k_val//1000} тыс. руб."
            elif key.startswith("over_"):
                k_val = float('inf')
                val = int(key.replace("over_", "").replace("k", "")) * 1000
                display = f"свыше {val//1000} тыс. руб."
            else:
                k_val = int(key.replace("k", "")) * 1000
                display = f"до {k_val//1000} тыс. руб."
            ranges.append((k_val, value, display))
            
        ranges.sort(key=lambda x: x[0])
        
        for limit, cost, display in ranges:
            if cameral_sum <= limit:
                return float(cost), display
                
        # Если почему-то не подошло, возвращаем последнее
        if ranges:
            return float(ranges[-1][1]), ranges[-1][2]
        return 0.0, ""


# Пример использования
if __name__ == "__main__":
    calc = Calculator()
    
    # Создаём тестовую смету
    items = [
        {"work_id": "drill_core_15m_cat2", "quantity": 50},  # 50 п.м. бурения
        {"work_id": "lab_moisture", "quantity": 20},         # 20 определений влажности
        {"work_id": "lab_density_ring", "quantity": 20},     # 20 определений плотности
        {"work_id": "report_cat2_100k", "quantity": 1},      # 1 отчёт
    ]
    
    estimate = calc.create_estimate(
        project_name="Тестовый проект",
        items_data=items,
        soil_category="II",
        climate_zone="III"
    )
    
    print(f"Смета: {estimate.project_name}")
    print(f"Индекс цен: {estimate.price_index}")
    print("-" * 70)
    for item in estimate.items:
        print(f"{item.code} {item.name[:45]}: {item.quantity} {item.unit} x {item.unit_cost} = {item.total_cost} руб.")
    print("-" * 70)
    print(f"Полевые: {estimate.subtotal_field} руб.")
    print(f"Лабораторные: {estimate.subtotal_laboratory} руб.")
    print(f"Камеральные: {estimate.subtotal_office} руб.")
    print(f"ИТОГО: {estimate.total} руб.")
 
