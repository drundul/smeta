"""
Расчёт сметной стоимости ИГИ по НЗ
Приказ Минстроя РФ №281/пр от 12.05.2025

Главное приложение Streamlit
"""

import streamlit as st
import json
from pathlib import Path
from decimal import Decimal
import datetime
import tempfile
import os

# Добавляем путь к модулям
import sys
sys.path.insert(0, str(Path(__file__).parent))

from modules.calculator import Calculator, Estimate, WorkItem
from modules.export_excel import export_to_excel
from modules.export_pdf import export_to_pdf
from modules.export_word import export_to_word
from config import (
    APP_TITLE, APP_ICON, APP_LAYOUT, 
    SOIL_CATEGORIES, COMPLEXITY_CATEGORIES, FIELD_WORK_CATEGORIES,
    REGIONS, DIFFICULT_ACCESS_TYPES
)


# Конфигурация страницы
st.set_page_config(
    page_title=APP_TITLE,
    page_icon=APP_ICON,
    layout=APP_LAYOUT,
    initial_sidebar_state="expanded"
)

# Инициализация калькулятора
@st.cache_resource
def get_calculator_v7():
    return Calculator()

calc = get_calculator_v7()


# Инициализация состояния
if "estimate_items" not in st.session_state:
    st.session_state.estimate_items = []

if "project_info" not in st.session_state:
    st.session_state.project_info = {
        "name": "",
        "code": "",
        "object": "",
        "customer": "",
        "contractor": "",
        "soil_category": "II",
        "complexity": "II",
        "region": "moscow",
        "distance_km": 50
    }


def load_coefficients():
    """Загрузка коэффициентов"""
    data_path = Path(__file__).parent / "data" / "coefficients.json"
    with open(data_path, "r", encoding="utf-8") as f:
        return json.load(f)


def get_region_list():
    """Получить список регионов из коэффициентов"""
    coefficients = load_coefficients()
    regions = coefficients.get("unfavorable_periods_by_region", {}).get("regions", {})
    return regions


# Заголовок
st.title(f"{APP_ICON} {APP_TITLE}")
st.markdown("**Приказ Минстроя РФ №281/пр от 12.05.2025**")
st.markdown("*Базовый уровень цен: 01.01.2024*")
st.divider()

# Боковая панель - информация о проекте
with st.sidebar:
    st.header("📋 Данные проекта")
    
    st.session_state.project_info["name"] = st.text_input(
        "Наименование проекта",
        value=st.session_state.project_info["name"],
        placeholder="Введите название проекта"
    )
    
    st.session_state.project_info["code"] = st.text_input(
        "Шифр проекта",
        value=st.session_state.project_info["code"],
        placeholder="Например: 2024-ИГИ-001"
    )
    
    st.session_state.project_info["object"] = st.text_input(
        "Объект",
        value=st.session_state.project_info["object"],
        placeholder="Адрес или название объекта"
    )
    
    st.session_state.project_info["customer"] = st.text_input(
        "Заказчик",
        value=st.session_state.project_info["customer"]
    )
    
    st.session_state.project_info["contractor"] = st.text_input(
        "Подрядчик",
        value=st.session_state.project_info["contractor"]
    )
    
    st.divider()
    st.subheader("⚙️ Условия работ")
    
    # Сортируем регионы: Сначала приоритетные, потом остальные по алфавиту
    regions = get_region_list()
    priority_regions = ["г. Москва", "Московская область", "г. Санкт-Петербург", "Ленинградская область"]
    all_regions = list(regions.keys())
    other_regions = sorted([r for r in all_regions if r not in priority_regions])
    region_options = priority_regions + other_regions
    
    # Поиск региона
    search_region = st.text_input("🔍 Поиск региона", placeholder="Начните вводить название...")
    
    if search_region:
        filtered_regions = [r for r in region_options if search_region.lower() in r.lower()]
    else:
        filtered_regions = region_options
    
    # Определяем индекс по умолчанию
    default_region = "г. Санкт-Петербург" # Меняем на Питер по умолчанию по просьбе (контекст)
    if default_region in filtered_regions:
        default_idx = filtered_regions.index(default_region)
    else:
        default_idx = 0
    
    selected_region = st.selectbox(
        "Регион производства работ",
        options=filtered_regions if filtered_regions else region_options,
        index=default_idx if filtered_regions else 0
    )
    st.session_state.project_info["region"] = selected_region
    
    # Показываем неблагоприятный период
    unfav_duration = regions.get(selected_region, 6.0)
    
    st.session_state.project_info["is_unfavorable_period_active"] = st.checkbox(
        f"Учесть неблагоприятный период ({unfav_duration} мес.)",
        value=st.session_state.project_info.get("is_unfavorable_period_active", False)
    )
    # st.info(f"⏱️ Неблагоприятный период: **{unfav_duration} мес.**")
    
    # Категория сложности ИГУ
    complexity_options = list(COMPLEXITY_CATEGORIES.keys())
    selected_complexity_code = st.selectbox(
        "Категория сложности ИГУ",
        options=complexity_options,
        format_func=lambda x: COMPLEXITY_CATEGORIES[x],
        index=1 if "II" in complexity_options else 0
    )
    
    st.session_state.project_info["complexity"] = selected_complexity_code
    
    # Расстояние до объекта
    st.session_state.project_info["distance_km"] = st.number_input(
        "Расстояние до объекта (км)",
        value=st.session_state.project_info.get("distance_km", 50),
        min_value=0,
        step=5
    )
    
    st.divider()
    
    # Индекс цен
    coefficients = load_coefficients()
    current_index = st.number_input(
        "Индекс пересчёта (к ценам 01.01.2024)",
        value=1.0,
        min_value=0.01,
        step=0.01,
        format="%.2f",
        help="Индекс изменения сметной стоимости к уровню цен 01.01.2024"
    )
    
    k_contract = st.number_input(
        "Коэффициент договорной цены",
        value=1.0,
        min_value=0.001,
        step=0.001,
        format="%.3f",
        help="Понижающий/повышающий коэффициент (тендерное снижение и т.п.)"
    )


def load_templates():
    """Загрузка шаблонов смет"""
    data_path = Path(__file__).parent / "data" / "templates.json"
    with open(data_path, "r", encoding="utf-8") as f:
        return json.load(f)



def calculate_additional_costs(field_cost: float, project_info: dict) -> list:
    """Расчет дополнительных затрат"""
    coefficients = load_coefficients()
    
    # 1. ДЗ на неблагоприятный период
    if project_info.get("is_unfavorable_period_active", False): # По умолчанию выключено
        region = project_info.get("region", "г. Москва")
        regions = get_region_list()
        unfav_duration = regions.get(region, 6.0)
        
        unfav_coefs = coefficients.get("unfavorable_period", {}).get("coefficients_by_duration_months", {})
        unfav_percent = 0
        
        for range_key, percents in unfav_coefs.items():
            if calc._check_duration_range(unfav_duration, range_key):
                cost_key = calc._get_cost_range_key(field_cost)
                unfav_percent = percents.get(cost_key, 0)
                break
        
        dz_unfav = field_cost * unfav_percent / 100
    else:
        dz_unfav = 0
        unfav_percent = 0
    
    # 2. ДЗ на проезд (Таблица 4)
    distance = project_info.get("distance_km", 50)
    
    # Используем обновленные данные Таблицы 4 (в базе поле travel_costs_IZ)
    travel_coefs = coefficients.get("travel_costs_IZ", {}).get("coefficients_by_distance_km", {})
    travel_percent = 0
    
    # Используем стандартные диапазоны стоимости (до 300к, до 1000к...)
    cost_key = calc._get_cost_range_key(field_cost)
    
    for dist_key, percents in travel_coefs.items():
        if calc._check_distance_range(distance, dist_key):
            travel_percent = percents.get(cost_key, 0) or 0
            break
    
    dz_travel = field_cost * travel_percent / 100
    
    # 3. ДЗ на организацию полевых работ (Таблица 8)
    org_coefs = coefficients.get("organization_costs", {}).get("coefficients_by_distance_km", {})
    org_percent = 0
    org_cost_key = calc._get_cost_range_key(field_cost)
    
    for dist_key, percents in org_coefs.items():
        if calc._check_distance_range(distance, dist_key):
            org_percent = percents.get(org_cost_key, 0) or 0
            break
            
    dz_org = field_cost * org_percent / 100
    
    # Формируем список словарей
    additional_costs = []
    if dz_unfav > 0:
        additional_costs.append({
            "name": f"ДЗ на неблагоприятный период ({unfav_percent}%)",
            "value": dz_unfav,
            "percent": unfav_percent,
            "basis": f"НЗ №281/пр, Приложение (неблагоприятный период)",
            "formula": f"{field_cost:,.0f} × {unfav_percent/100:.4f}"
        })
    if dz_travel > 0:
        additional_costs.append({
            "name": f"ДЗ на проезд ({travel_percent}%)",
            "value": dz_travel,
            "percent": travel_percent,
            "basis": f"НЗ №281/пр, Таблица 4 (проезд)",
            "formula": f"{field_cost:,.0f} × {travel_percent/100:.4f}"
        })
    if dz_org > 0:
        additional_costs.append({
            "name": f"ДЗ на организацию полевых работ ({org_percent}%)",
            "value": dz_org,
            "percent": org_percent,
            "basis": f"НЗ №281/пр, Таблица 8 (организация)",
            "formula": f"{field_cost:,.0f} × {org_percent/100:.4f}"
        })
        
    return additional_costs


# Основная область - добавление работ
tab0, tab1, tab2, tab3, tab4 = st.tabs([
    "📋 Шаблоны", 
    "📝 Добавление работ", 
    "📊 Текущая смета", 
    "💰 Дополнительные затраты", 
    "📥 Экспорт"
])

with tab0:
    st.subheader("📋 Готовые шаблоны смет")
    st.markdown("Выберите типовой шаблон для быстрого создания сметы")
    
    templates_data = load_templates()
    templates = templates_data.get("templates", [])
    
    # Группировка по категориям
    categories = templates_data.get("template_categories", {})
    
    for cat_id, cat_name in categories.items():
        cat_templates = [t for t in templates if t.get("category") == cat_id]
        if cat_templates:
            st.markdown(f"### {cat_name}")
            
            for template in cat_templates:
                with st.expander(f"**{template['name']}** — {template['description']}"):
                    # Нормативные документы
                    st.markdown("**📚 Нормативные документы:**")
                    for doc in template.get("normative_docs", []):
                        st.markdown(f"- {doc}")
                    
                    # Методика расчёта
                    if template.get("methodology"):
                        st.divider()
                        st.markdown("**📋 Методика (требования):**")
                        for method in template["methodology"]:
                            st.markdown(f"- **{method['item']}**: {method['requirement']}")
                            st.caption(f"   _Источник: {method['source']}_")
                    
                    st.divider()
                    
                    # Состав работ с ссылками на НЗ
                    st.markdown("**📝 Состав работ:**")
                    for item in template.get("items", []):
                        work_info = calc.get_work_type(item["work_id"])
                        base_cost = calc.get_base_cost(item["work_id"])
                        item_cost = float(base_cost) * item["quantity"]
                        
                        # Название работы
                        work_name = work_info.get('name', item['work_id'])
                        nz_ref = item.get('nz_ref', '')
                        
                        col_a, col_b = st.columns([3, 1])
                        with col_a:
                            st.markdown(f"**{work_name}**")
                            if item.get("description"):
                                st.caption(f"_{item['description']}_")
                            if nz_ref:
                                st.caption(f"📖 _{nz_ref}_")
                        with col_b:
                            st.write(f"{item['quantity']} {work_info.get('unit', 'ед.')}")
                            st.write(f"**{item_cost:,.0f} ₽**")
                    
                    # Дополнительные затраты
                    if template.get("additional_costs"):
                        st.divider()
                        st.markdown("**➕ Дополнительные затраты:**")
                        for add_cost in template["additional_costs"]:
                            if add_cost.get("percent"):
                                st.markdown(f"- **{add_cost['description']}**: {add_cost['percent']}%")
                            else:
                                st.markdown(f"- **{add_cost['description']}**")
                            if add_cost.get("source"):
                                st.caption(f"   _Источник: {add_cost['source']}_")
                            if add_cost.get("note"):
                                st.caption(f"   _{add_cost['note']}_")
                    
                    st.divider()
                    
                    # Примечания
                    if template.get("notes"):
                        st.markdown("**📌 Примечания:**")
                        for note in template["notes"]:
                            st.markdown(f"- {note}")
                    
                    st.divider()
                    
                    # Предварительный расчёт
                    total_cost = 0
                    for item in template.get("items", []):
                        base_cost = calc.get_base_cost(item["work_id"])
                        total_cost += float(base_cost) * item["quantity"]
                    
                    # Учитываем ДЗрежим если есть
                    regime_surcharge = 0
                    for add_cost in template.get("additional_costs", []):
                        if add_cost.get("type") == "regime_surcharge":
                            regime_surcharge = total_cost * add_cost.get("percent", 0) / 100
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("💰 Базовая стоимость", f"{total_cost:,.0f} ₽")
                    with col2:
                        if regime_surcharge > 0:
                            st.metric("⚡ С учётом ДЗрежим", f"{total_cost + regime_surcharge:,.0f} ₽")
                    
                    st.caption("_Без учёта ДЗ на неблагоприятный период, проезд, привязку_")
                    
                    # Кнопка применения шаблона
                    if st.button(f"✅ Применить шаблон", key=f"apply_{template['id']}", type="primary"):
                        # Очищаем текущую смету
                        st.session_state.estimate_items = []
                        
                        # Добавляем все позиции из шаблона
                        for item in template.get("items", []):
                            item_data = {
                                "work_id": item["work_id"],
                                "quantity": item["quantity"],
                                "additional_coefficients": {}
                            }
                            st.session_state.estimate_items.append(item_data)
                        
                        # Устанавливаем параметры по умолчанию
                        default_params = template.get("default_params", {})
                        if "complexity" in default_params:
                            st.session_state.project_info["complexity"] = default_params["complexity"]
                        
                        st.success(f"✅ Шаблон «{template['name']}» применён! Перейдите на вкладку «Текущая смета».")
                        st.rerun()

with tab1:
    st.subheader("Добавление позиций в смету")
    
    # Выбор категории работ
    col1, col2 = st.columns([1, 2])
    
    with col1:
        work_category = st.radio(
            "Категория работ",
            options=["field", "laboratory", "office"],
            format_func=lambda x: {
                "field": "🔧 Полевые работы",
                "laboratory": "🔬 Лабораторные работы",
                "office": "📄 Камеральные работы"
            }[x]
        )
    
    with col2:
        # Получаем виды работ по категории
        work_types = calc.get_work_types_by_category(work_category)
        
        work_options = {w["id"]: f"{w['code']} - {w['name']}" for w in work_types}
        
        if work_options:
            selected_work_id = st.selectbox(
                "Вид работ",
                options=list(work_options.keys()),
                format_func=lambda x: work_options.get(x, x)
            )
            
            if selected_work_id:
                work_info = calc.get_work_type(selected_work_id)
                base_cost = calc.get_base_cost(selected_work_id)
                
                col_a, col_b, col_c = st.columns(3)
                
                with col_a:
                    quantity = st.number_input(
                        f"Количество ({work_info.get('unit', 'ед.')})",
                        min_value=0.0,
                        value=1.0,
                        step=1.0
                    )
                
                with col_b:
                    st.metric("Базовая цена", f"{float(base_cost):,.0f} ₽")
                
                with col_c:
                    # Расчёт предварительной стоимости
                    preliminary_cost = float(base_cost) * quantity
                    st.metric("Предв. стоимость", f"{preliminary_cost:,.0f} ₽")
                
                # Дополнительные коэффициенты для полевых работ
                if work_category == "field":
                    with st.expander("Дополнительные коэффициенты"):
                        additional_coefs = {}
                        
                        col_x, col_y = st.columns(2)
                        with col_x:
                            if st.checkbox("Работы в зимний период", key="winter_check"):
                                additional_coefs["winter"] = 1.15
                            
                            if st.checkbox("Работы в ночное время", key="night_check"):
                                additional_coefs["night"] = 1.25
                        
                        with col_y:
                            if st.checkbox("Труднодоступный участок", key="difficult_check"):
                                difficult_type = st.selectbox(
                                    "Тип",
                                    options=list(DIFFICULT_ACCESS_TYPES.keys()),
                                    format_func=lambda x: DIFFICULT_ACCESS_TYPES[x]
                                )
                                difficult_coefs = {
                                    "slope_10_20": 1.1,
                                    "slope_20_30": 1.2,
                                    "slope_30_plus": 1.35,
                                    "swamp": 1.25,
                                    "urban_dense": 1.15,
                                    "indoor": 1.30
                                }
                                additional_coefs["difficult"] = difficult_coefs.get(difficult_type, 1.0)
                else:
                    additional_coefs = {}
                
                # Кнопка добавления
                if st.button("➕ Добавить в смету", type="primary", use_container_width=True):
                    if quantity > 0:
                        item_data = {
                            "work_id": selected_work_id,
                            "quantity": quantity,
                            "additional_coefficients": additional_coefs
                        }
                        st.session_state.estimate_items.append(item_data)
                        st.success(f"Добавлено: {work_info.get('name', '')}")
                        st.rerun()
                    else:
                        st.error("Укажите количество больше 0")
        else:
            st.warning("Нет доступных видов работ в этой категории")


with tab2:
    st.subheader("Текущая смета")
    
    if not st.session_state.estimate_items:
        st.info("Смета пуста. Добавьте позиции на вкладке «Добавление работ» или выберите шаблон.")
    else:
        # Создаём смету для отображения
        estimate = calc.create_estimate(
            project_name=st.session_state.project_info["name"] or "Без названия",
            items_data=st.session_state.estimate_items,
            soil_category=st.session_state.project_info.get("soil_category", "II"),
            climate_zone="III",
            apply_price_index=True
        )
        estimate.project_code = st.session_state.project_info["code"]
        estimate.object_name = st.session_state.project_info["object"]
        estimate.customer = st.session_state.project_info["customer"]
        estimate.contractor = st.session_state.project_info["contractor"]
        estimate.price_index = Decimal(str(current_index))
        
        # Заголовок таблицы
        st.markdown("#### Локальная смета на работы по ИГИ")
        st.markdown(f"*Приказ Минстроя России №281/пр от 12.05.2025. Уровень цен: 01.01.2024*")
        st.divider()
        
        # Шапка таблицы
        header_cols = st.columns([0.5, 3, 0.7, 0.8, 2, 1.5, 1.2, 0.3])
        with header_cols[0]:
            st.markdown("**№**")
        with header_cols[1]:
            st.markdown("**Наименование работ и затрат**")
        with header_cols[2]:
            st.markdown("**Ед. изм.**")
        with header_cols[3]:
            st.markdown("**Кол-во**")
        with header_cols[4]:
            st.markdown("**Обоснование стоимости**")
        with header_cols[5]:
            st.markdown("**Расчёт стоимости**")
        with header_cols[6]:
            st.markdown("**Стоимость, руб.**")
        with header_cols[7]:
            st.markdown("")
        
        st.divider()
        
        # Группировка по категориям
        field_items = []
        lab_items = []
        office_items = []
        
        # 1. Сначала считаем сумму камеральных работ (без отчёта и программы)
        cameral_base_sum = 0
        report_item_index = -1
        
        for i, item_data in enumerate(st.session_state.estimate_items):
            work_info = calc.get_work_type(item_data["work_id"])
            grp = work_info.get("group", "")
            cat = work_info.get("category", "")
            
            # Считаем базу: это всё "офисное" (камеральное), включая Программу, но кроме самого Отчёта
            if cat == "office" and grp != "report":
                 base_c = float(calc.get_base_cost(item_data["work_id"]))
                 cameral_base_sum += base_c * item_data["quantity"]
            
            if grp == "report":
                report_item_index = i

        # 2. Если есть отчёт, пересчитываем его стоимость
        if report_item_index >= 0:
            # Ступенчатый выбор стоимости по Таблице 65 (без интерполяции, по требованию пользователя)
            # Значения для II категории сложности
            
            x_val = cameral_base_sum
            y_val = 0
            range_desc = ""
            
            if x_val <= 20000:
                y_val = 134685
                range_desc = "до 20 тыс. руб."
            elif x_val <= 50000:
                y_val = 203793
                range_desc = "20-50 тыс. руб."
            elif x_val <= 100000:
                y_val = 278776
                range_desc = "50-100 тыс. руб."
            elif x_val <= 250000:
                y_val = 421816
                range_desc = "100-250 тыс. руб."
            else:
                # Если свыше 250 000 - берем следующее значение (экстраполяция или макс)
                # Предположим следующее значение или оставим как есть пока
                y_val = 421816 
                range_desc = "свыше 100 тыс. руб."

            calculated_report_cost = y_val
        else:
            calculated_report_cost = 0
            range_desc = ""

        for i, item_data in enumerate(st.session_state.estimate_items):
            work_info = calc.get_work_type(item_data["work_id"])
            base_cost = calc.get_base_cost(item_data["work_id"])
            quantity = item_data["quantity"]
            
            # Если это отчёт - подменяем стоимость
            if work_info.get("group") == "report" and calculated_report_cost > 0:
                base_cost = calculated_report_cost
                
                # Сохраняем рассчитанную стоимость в сессию, чтобы она была доступна для экспорта и ДЗ
                st.session_state.estimate_items[i]["override_base_cost"] = float(calculated_report_cost)
                
                quantity = 1 # Отчет всегда 1
                total_cost = calculated_report_cost
                # Добавим пометку в название
                display_name = f"Технический отчёт (ИГИ, II кат.)"
                calc_formula = f"{calculated_report_cost:,.0f} (Таблица 65, {range_desc})"
            else:
                # Если это не отчет, убираем override (на случай если он был раньше)
                if "override_base_cost" in st.session_state.estimate_items[i] and work_info.get("group") != "report":
                    del st.session_state.estimate_items[i]["override_base_cost"]
                    
                total_cost = float(base_cost) * quantity
                display_name = work_info.get("name", item_data["work_id"])
                calc_formula = f"{float(base_cost):,.0f} × {quantity:.1f}"

            # Сохраняем формулу в сессию для экспорта
            st.session_state.estimate_items[i]["formula"] = calc_formula
            
            item_row = {
                "index": i,
                "work_id": item_data["work_id"],
                "name": display_name,
                "unit": work_info.get("unit", "ед."),
                "quantity": quantity,
                "base_cost": float(base_cost),
                "total_cost": total_cost,
                "table_ref": work_info.get("table_ref", ""),
                "code": work_info.get("code", ""),
                "category": work_info.get("category", "field"),
                "formula_display": calc_formula
            }
            
            if item_row["category"] == "field":
                field_items.append(item_row)
            elif item_row["category"] == "laboratory":
                lab_items.append(item_row)
            else:
                office_items.append(item_row)
        
        row_num = 1
        
        # Раздел I - Полевые работы
        if field_items:
            st.markdown("##### **Раздел I. Полевые работы**")
            field_total = 0
            for item in field_items:
                cols = st.columns([0.5, 3, 0.7, 0.8, 2, 1.5, 1.2, 0.3])
                with cols[0]:
                    st.write(f"{row_num}")
                with cols[1]:
                    st.write(item["name"])
                with cols[2]:
                    st.write(item["unit"])
                with cols[3]:
                    st.write(f"{item['quantity']:.1f}")
                with cols[4]:
                    # Обоснование: НЗ п.281/пр, глава, таблица
                    st.write(f"НЗ №281/пр, {item['table_ref']}")
                with cols[5]:
                    # Формула расчёта
                    st.write(item.get("formula_display", f"{item['base_cost']:,.0f} × {item['quantity']:.1f}"))
                with cols[6]:
                    st.write(f"**{item['total_cost']:,.0f}**")
                with cols[7]:
                    if st.button("🗑️", key=f"del_{item['index']}"):
                        st.session_state.estimate_items.pop(item['index'])
                        st.rerun()
                row_num += 1
                field_total += item["total_cost"]
            
            # Итого по разделу
            cols = st.columns([0.5, 3, 0.7, 0.8, 2, 1.5, 1.2, 0.3])
            with cols[1]:
                st.markdown("**Итого по разделу I (СПпз):**")
            with cols[6]:
                st.markdown(f"**{field_total:,.0f}**")
            st.divider()
        
        # Раздел II - Лабораторные работы
        if lab_items:
            st.markdown("##### **Раздел II. Лабораторные работы**")
            lab_total = 0
            for item in lab_items:
                cols = st.columns([0.5, 3, 0.7, 0.8, 2, 1.5, 1.2, 0.3])
                with cols[0]:
                    st.write(f"{row_num}")
                with cols[1]:
                    st.write(item["name"])
                with cols[2]:
                    st.write(item["unit"])
                with cols[3]:
                    st.write(f"{item['quantity']:.1f}")
                with cols[4]:
                    st.write(f"НЗ №281/пр, {item['table_ref']}")
                with cols[5]:
                    st.write(f"{item['base_cost']:,.0f} × {item['quantity']:.1f}")
                with cols[6]:
                    st.write(f"**{item['total_cost']:,.0f}**")
                with cols[7]:
                    if st.button("🗑️", key=f"del_{item['index']}"):
                        st.session_state.estimate_items.pop(item['index'])
                        st.rerun()
                row_num += 1
                lab_total += item["total_cost"]
            
            cols = st.columns([0.5, 3, 0.7, 0.8, 2, 1.5, 1.2, 0.3])
            with cols[1]:
                st.markdown("**Итого по разделу II (СЛпз):**")
            with cols[6]:
                st.markdown(f"**{lab_total:,.0f}**")
            st.divider()
        
        # Раздел III - Камеральные работы
        if office_items:
            st.markdown("##### **Раздел III. Камеральные работы**")
            office_total = 0
            for item in office_items:
                cols = st.columns([0.5, 3, 0.7, 0.8, 2, 1.5, 1.2, 0.3])
                with cols[0]:
                    st.write(f"{row_num}")
                with cols[1]:
                    st.write(item["name"])
                with cols[2]:
                    st.write(item["unit"])
                with cols[3]:
                    st.write(f"{item['quantity']:.1f}")
                with cols[4]:
                    st.write(f"НЗ №281/пр, {item['table_ref']}")
                with cols[5]:
                    st.write(f"{item['base_cost']:,.0f} × {item['quantity']:.1f}")
                with cols[6]:
                    st.write(f"**{item['total_cost']:,.0f}**")
                with cols[7]:
                    if st.button("🗑️", key=f"del_{item['index']}"):
                        st.session_state.estimate_items.pop(item['index'])
                        st.rerun()
                row_num += 1
                office_total += item["total_cost"]
            
            cols = st.columns([0.5, 3, 0.7, 0.8, 2, 1.5, 1.2, 0.3])
            with cols[1]:
                st.markdown("**Итого по разделу III (СКпз):**")
            with cols[6]:
                st.markdown(f"**{office_total:,.0f}**")
            st.divider()
        
        # Общие итоги по базовым затратам
        field_total = sum(i["total_cost"] for i in field_items)
        lab_total = sum(i["total_cost"] for i in lab_items)
        office_total = sum(i["total_cost"] for i in office_items)
        base_total = field_total + lab_total + office_total
        
        st.markdown("#### Итоги по базовым затратам (СП + СЛ + СК)")
        st.write("Дополнительные затраты (ДЗ) рассчитываются во вкладке **💰 Дополнительные затраты**")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("🔧 СП", f"{field_total:,.0f} ₽")
        with col2:
            st.metric("🔬 СЛ", f"{lab_total:,.0f} ₽")
        with col3:
            st.metric("📄 СК", f"{office_total:,.0f} ₽")
        with col4:
            st.metric("📊 Базовый итог", f"{base_total:,.0f} ₽")
        
        st.divider()
        
        # Расчёт и отображение дополнительных затрат
        temp_estimate = calc.create_estimate(
             project_name="Temp",
             items_data=st.session_state.estimate_items,
             soil_category=st.session_state.project_info.get("soil_category", "II"),
             climate_zone="III"
        )
        field_cost_base = float(temp_estimate.subtotal_field)
        
        dz_list = calculate_additional_costs(field_cost_base, st.session_state.project_info)
        dz_sum = sum(item["value"] for item in dz_list)
        final_total = base_total + dz_sum
        
        if dz_list:
            st.markdown("##### ➕ Дополнительные затраты")
            for dz in dz_list:
                d_col1, d_col2 = st.columns([3, 1])
                with d_col1:
                    st.write(f"{dz['name']}")
                    st.caption(f"Обоснование: {dz['basis']}")
                with d_col2:
                    st.write(f"**{dz['value']:,.0f} ₽**")
            st.divider()
            
        # Финальный итог крупно
        st.markdown(f"### 🏁 ИТОГО: {final_total:,.0f} ₽")
        
        st.divider()
        
        # Кнопка очистки
        if st.button("🗑️ Очистить смету", type="secondary"):
            st.session_state.estimate_items = []
            st.rerun()


with tab3:
    st.subheader("💰 Расчёт дополнительных затрат")
    
    if not st.session_state.estimate_items:
        st.info("Сначала добавьте позиции в смету.")
    else:
        # Создаём смету
        estimate = calc.create_estimate(
            project_name=st.session_state.project_info["name"] or "Без названия",
            items_data=st.session_state.estimate_items,
            soil_category=st.session_state.project_info.get("soil_category", "II"),
            climate_zone="III",
            apply_price_index=True
        )
        
        field_cost = float(estimate.subtotal_field)
        
        st.markdown(f"**Стоимость полевых работ:** {field_cost:,.0f} ₽")
        st.divider()
        
        # Расчёт дополнительных затрат
        additional_costs_list = calculate_additional_costs(field_cost, st.session_state.project_info)
        
        # Заголовок таблицы для правильного выравнивания
        cols = st.columns([3, 2, 2, 1])
        with cols[0]:
            st.markdown("**Наименование**")
        with cols[1]:
            st.markdown("**Обоснование**")
        with cols[2]:
            st.markdown("**Расчёт**")
        with cols[3]:
            st.markdown("**Стоимость**")
        
        st.divider()
        
        for cost in additional_costs_list:
            cols = st.columns([3, 2, 2, 1])
            with cols[0]:
                st.write(cost['name'])
            with cols[1]:
                st.write(cost['basis'])
            with cols[2]:
                st.write(cost['formula'])
            with cols[3]:
                st.write(f"**{cost['value']:,.0f}**")
            st.divider()
        
        if not additional_costs_list:
            st.info("Дополнительные затраты не начислены (или равны 0).")
        
        st.divider()
        
        # Итого дополнительных затрат
        total_dz = sum(item['value'] for item in additional_costs_list)
        estimate.additional_costs = additional_costs_list
        total_with_dz = estimate.total_with_dz
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("📊 Итого ДЗ", f"{total_dz:,.0f} ₽")
        with col2:
            st.metric("💰 ВСЕГО с ДЗ", f"{total_with_dz:,.0f} ₽")


with tab4:
    st.subheader("📥 Экспорт сметы")
    
    if not st.session_state.estimate_items:
        st.warning("Сначала добавьте позиции в смету.")
    else:
        # Создаём смету для экспорта
        estimate = calc.create_estimate(
            project_name=st.session_state.project_info["name"] or "Без названия",
            items_data=st.session_state.estimate_items,
            soil_category=st.session_state.project_info.get("soil_category", "II"),
            climate_zone="III",
            apply_price_index=True
        )
        
        # Добавляем ДЗ
        field_cost = float(estimate.subtotal_field)
        d_costs = calculate_additional_costs(field_cost, st.session_state.project_info)
        estimate.additional_costs = d_costs
        
        estimate.project_code = st.session_state.project_info["code"]
        estimate.object_name = st.session_state.project_info["object"]
        estimate.customer = st.session_state.project_info["customer"]
        estimate.contractor = st.session_state.project_info["contractor"]
        estimate.price_index = Decimal(str(current_index))
        estimate.contract_coefficient = Decimal(str(k_contract))
        
        # ВАЖНО: Добавляем дополнительные затраты в объект сметы перед экспортом!
        field_cost_for_dz = float(estimate.subtotal_field)
        dz_list = calculate_additional_costs(field_cost_for_dz, st.session_state.project_info)
        estimate.additional_costs = dz_list
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("### 📗 Excel")
            try:
                # В Windows нельзя открывать файл, если он уже открыт в NamedTemporaryFile
                # Поэтому создаем, получаем имя и сразу закрываем
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    tmp_name = tmp.name
                
                # Теперь файл закрыт, можно работать по пути
                try:
                    export_to_excel(estimate, tmp_name)
                    with open(tmp_name, "rb") as f:
                        excel_data = f.read()
                finally:
                    if os.path.exists(tmp_name):
                        os.unlink(tmp_name)
                
                st.download_button(
                    label="💾 Скачать .xlsx",
                    data=excel_data,
                    file_name=f"Смета_{estimate.project_name}_{estimate.date_created}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Ошибка экспорта: {e}")
        
        with col2:
            st.markdown("### 📕 PDF")
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                    tmp_name = tmp.name
                
                try:
                    export_to_pdf(estimate, tmp_name)
                    with open(tmp_name, "rb") as f:
                        pdf_data = f.read()
                finally:
                    if os.path.exists(tmp_name):
                        os.unlink(tmp_name)
                
                st.download_button(
                    label="💾 Скачать .pdf",
                    data=pdf_data,
                    file_name=f"Смета_{estimate.project_name}_{estimate.date_created}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Ошибка экспорта: {e}")
        
        with col3:
            st.markdown("### 📘 Word")
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                    tmp_name = tmp.name
                
                try:
                    export_to_word(estimate, tmp_name)
                    with open(tmp_name, "rb") as f:
                        word_data = f.read()
                finally:
                    if os.path.exists(tmp_name):
                        os.unlink(tmp_name)
                
                st.download_button(
                    label="💾 Скачать .docx",
                    data=word_data,
                    file_name=f"Смета_{estimate.project_name}_{estimate.date_created}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Ошибка экспорта: {e}")


# Футер
st.divider()
st.markdown("""
<div style="text-align: center; color: gray; font-size: 12px;">
    Расчёт по нормативным затратам (НЗ) в соответствии с Приказом Минстроя РФ №281/пр от 12.05.2025<br>
    Базовый уровень цен: 01.01.2024
</div>
""", unsafe_allow_html=True)
