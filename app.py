"""
Расчёт сметной стоимости ИГИ по НЗ
Приказ Минстроя РФ №281/пр от 12.05.2025

Главное приложение Streamlit
"""

import streamlit as st
import json
import uuid
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
    REGIONS, DIFFICULT_ACCESS_TYPES, CLIMATE_ZONES
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
def get_calculator_v8():
    return Calculator()

st.cache_resource.clear()
st.cache_resource.clear()
calc = get_calculator_v8()


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
    
    # Максимальная глубина исследования (для расчёта стоимости Программы, Таблица 66)
    DEPTH_OPTIONS = {
        "5": "до 5 м",
        "10": "от 5 до 10 м",
        "15": "от 10 до 15 м",
        "25": "от 15 до 25 м",
        "50": "от 25 до 50 м",
        "75": "от 50 до 75 м",
        "over": "свыше 75 м",
    }
    depth_keys = list(DEPTH_OPTIONS.keys())
    selected_depth = st.selectbox(
        "📏 Макс. глубина исследования (Табл. 66)",
        options=depth_keys,
        format_func=lambda x: DEPTH_OPTIONS[x],
        index=depth_keys.index(st.session_state.project_info.get("max_depth", "10")),
        help="Максимальная глубина бурения/зондирования на объекте. Определяет стоимость Программы ИГИ (Таблица 66)."
    )
    st.session_state.project_info["max_depth"] = selected_depth
    
    st.divider()
    st.subheader("📐 Параметры ДЗ")
    
    # Режимный объект (ДЗрежим, п.26-27)
    st.session_state.project_info["is_regime_object"] = st.checkbox(
        "Режимный объект (ДЗрежим +25%)",
        value=st.session_state.project_info.get("is_regime_object", False),
        help="п.26-27 НЗ: объекты военной инфраструктуры, ядерного комплекса, охранные зоны ЛЭП, полосы отвода ж/д, автодорог и т.п."
    )
    
    # Тип транспорта для проезда (Таблицы 4-7)
    transport_options = {"auto": "🚗 Автотранспорт", "non_auto": "🚂 Не автотранспорт (ж/д, авиа и т.п.)"}
    st.session_state.project_info["transport_type"] = st.radio(
        "Тип транспорта (проезд)",
        options=list(transport_options.keys()),
        format_func=lambda x: transport_options[x],
        index=0 if st.session_state.project_info.get("transport_type", "auto") == "auto" else 1,
        help="Таблицы 4-5 (авто) или 6-7 (не авто) НЗ"
    )
    
    # Статическое зондирование
    st.session_state.project_info["has_static_sounding"] = st.checkbox(
        "Со статическим зондированием",
        value=st.session_state.project_info.get("has_static_sounding", False),
        help="Влияет на таблицу ДЗ проезд: Таблица 16/17 (авто) или 6/7 (не авто)"
    )
    
    # Интерполяция коэффициентов проезда (п.160)
    st.session_state.project_info["use_interpolation"] = st.checkbox(
        "Интерполяция коэфф. проезда",
        value=st.session_state.project_info.get("use_interpolation", True),
        help="п.160, прим. 3: промежуточные значения определяются линейной интерполяцией"
    )
    
    # Работа по месту постоянной работы (п.12, п.38)
    st.session_state.project_info["is_local_work"] = st.checkbox(
        "Работа по месту постоянной работы",
        value=st.session_state.project_info.get("is_local_work", False),
        help="п.12 НЗ: применяется К1 (снижение 12-18%). п.38: ДЗорг не начисляется."
    )
    
    # Климатическая зона (К2, п.13)
    climate_options = list(CLIMATE_ZONES.keys())
    selected_climate = st.selectbox(
        "Климатическая зона (К2)",
        options=climate_options,
        format_func=lambda x: CLIMATE_ZONES[x],
        index=climate_options.index(st.session_state.project_info.get("climate_zone", "IV")),
        help="п.13 НЗ, Таблица 2: К2 учитывает эксплуатацию машин в зависимости от климата"
    )
    st.session_state.project_info["climate_zone"] = selected_climate
    
    # Показываем значение К2
    k2_val = calc.get_climate_coefficient(selected_climate)
    if float(k2_val) != 1.0:
        st.info(f"☃️ Коэфф. К2 = **{float(k2_val):.2f}** (климат. зона {selected_climate})")
    
    # Информация о районном коэффициенте
    pdz_r_value = calc.get_regional_coefficient(selected_region)
    if pdz_r_value > 1.0:
        st.info(f"📍 Районный коэффициент: **{pdz_r_value}** (ДЗрП будет начислено)")
    
    # Местоположение лаборатории
    lab_in_spb = st.checkbox(
        "🧪 Лаборатория в СПб (база)",
        value=True,
        help="Если лаборатория находится в СПб — районный коэффициент к лабораторным работам НЕ начисляется (п.47 НЗ). Снимите галочку, если лаборатория в регионе объекта."
    )
    st.session_state.project_info["lab_in_spb"] = lab_in_spb
    if pdz_r_value > 1.0 and lab_in_spb:
        st.caption("_ДЗрП начисляется только на полевые. Лаборатория в СПб → Крайон=1.0_")
    
    st.divider()
    
    # Индекс цен
    coefficients = load_coefficients()
    current_index = st.number_input(
        "Индекс пересчёта (к ценам 01.01.2024)",
        value=st.session_state.project_info.get("price_index", 1.0),
        min_value=0.01,
        step=0.01,
        format="%.2f",
        help="Индекс изменения сметной стоимости к уровню цен 01.01.2024"
    )
    st.session_state.project_info["price_index"] = current_index
    
    k_contract = st.number_input(
        "Коэффициент договорной цены",
        value=st.session_state.project_info.get("k_contract", 1.0),
        min_value=0.001,
        step=0.001,
        format="%.3f",
        help="Понижающий/повышающий коэффициент (тендерное снижение и т.п.)"
    )
    st.session_state.project_info["k_contract"] = k_contract


def load_templates():
    """Загрузка шаблонов смет"""
    data_path = Path(__file__).parent / "data" / "templates.json"
    with open(data_path, "r", encoding="utf-8") as f:
        return json.load(f)



def calculate_additional_costs(field_cost: float, project_info: dict, lab_cost: float = 0) -> list:
    """Расчет дополнительных затрат (п.20-48 НЗ №281/пр)
    
    Формула 3: ДЗП = ДЗНП + ДЗноч + ДЗрежим + ДЗпроезд + ДЗорг + ДЗрП + ДЗсП
    
    Args:
        field_cost: стоимость полевых работ (СПпз)
        project_info: информация о проекте (регион, расстояние, флаги)
        lab_cost: стоимость лабораторных работ (СЛпз) для расчёта ДЗрайонЛ
    """
    coefficients = load_coefficients()
    
    # === 1. ДЗ на неблагоприятный период (формула 4, п.21) ===
    if project_info.get("is_unfavorable_period_active", False):
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
    
    # === 2. ДЗ на неизбежные перерывы (формула 6, п.26-27) ===
    # ДЗрежим = СПрежим × ПДЗрежим
    # ПДЗрежим = 25% для объектов п.27
    if project_info.get("is_regime_object", False):
        regime_data = coefficients.get("intermittent_work", {})
        regime_percent = regime_data.get("pdz_regime_percent", 25)
        dz_regime = field_cost * regime_percent / 100
    else:
        dz_regime = 0
        regime_percent = 0
    
    # === 3. ДЗ на проезд (формулы 7-8, п.28-36) ===
    distance = project_info.get("distance_km", 50)
    transport_type = project_info.get("transport_type", "auto")  # auto / non_auto
    has_static_sounding = project_info.get("has_static_sounding", False)
    use_interpolation = project_info.get("use_interpolation", True)
    
    # Выбор таблицы коэффициентов по типу транспорта и зондирования
    if transport_type == "auto":
        if not has_static_sounding:
            travel_table_key = "travel_costs_IZ"  # Таблица 4 (авто, без зондирования)
            travel_table_name = "Таблица 4"
            travel_paragraph = "п.29"
        else:
            travel_table_key = "travel_costs_NZ"  # Таблица 5 (авто, с зондированием)
            travel_table_name = "Таблица 5"
            travel_paragraph = "п.30"
    else:
        if not has_static_sounding:
            travel_table_key = "travel_costs_table6"  # Таблица 6 (не авто, без зондирования)
            travel_table_name = "Таблица 6"
            travel_paragraph = "п.33"
        else:
            travel_table_key = "travel_costs_table7"  # Таблица 7 (не авто, с зондированием)
            travel_table_name = "Таблица 7"
            travel_paragraph = "п.34"
    
    travel_coefs = coefficients.get(travel_table_key, {}).get("coefficients_by_distance_km", {})
    
    # Определяем ключ стоимостного диапазона
    # Для travel_costs_NZ и table7 — другие диапазоны стоимости
    if travel_table_key in ("travel_costs_NZ", "travel_costs_table7"):
        cost_key = calc._get_travel_cost_range_key(field_cost)
    else:
        cost_key = calc._get_cost_range_key(field_cost)
    
    # Расчёт процента — с интерполяцией или без
    if use_interpolation and travel_coefs:
        travel_percent = calc.interpolate_coefficient(distance, travel_coefs, cost_key)
    else:
        travel_percent = 0
        for dist_key, percents in travel_coefs.items():
            if calc._check_distance_range(distance, dist_key):
                travel_percent = percents.get(cost_key, 0) or 0
                break
    
    dz_travel = field_cost * travel_percent / 100
    
    # === 4. ДЗ на организацию полевых работ (п.37-39, Таблица 20) ===
    # ДЗорг = СПпз × ПДЗорг / 100
    # Не применяется если работы по месту постоянной работы (п.38)
    is_local = project_info.get("is_local_work", False)
    
    if not is_local:
        org_coefs = coefficients.get("organization_costs", {}).get("coefficients_by_distance_km", {})
        org_percent = 0
        org_cost_key = calc._get_cost_range_key(field_cost)
        
        for dist_key, percents in org_coefs.items():
            if calc._check_distance_range(distance, dist_key):
                org_percent = percents.get(org_cost_key, 0) or 0
                break
                
        dz_org = field_cost * org_percent / 100
    else:
        dz_org = 0
        org_percent = 0
    
    # === 5. ДЗ на районные выплаты — полевые (формула 10, п.40) ===
    # ДЗрП = (СПпз + ДЗНП + ДЗрежим + ДЗноч + ДЗорг) × (ДЗП × ПДЗр + ДпрочП - 1)
    # где: ДЗП = доля ФОТ = 0.41 (labor_share_field)
    #       ДпрочП = доля прочих = 0.59 (other_share_field)
    #       ПДЗр = районный коэффициент
    region = project_info.get("region", "г. Москва")
    pdz_r = calc.get_regional_coefficient(region)
    
    reg_data = coefficients.get("regional_allowances", {})
    dzp_share = reg_data.get("labor_share_field", 0.41)
    dproch_field = reg_data.get("other_share_field", 0.59)
    
    dz_rp = 0
    rp_multiplier = 0
    
    if pdz_r > 1.0:
        # База для районных = СПпз + ДЗНП + ДЗрежим + ДЗноч + ДЗорг
        base_for_regional = field_cost + dz_unfav + dz_regime + dz_org
        # Множитель: (ДЗП × ПДЗр + ДпрочП - 1)
        rp_multiplier = dzp_share * pdz_r + dproch_field - 1
        dz_rp = base_for_regional * rp_multiplier
    
    # === 6. ДЗ на районные выплаты — лабораторные (формула 14, п.47) ===
    # ДЗрайонЛ = СЛпз × (ДЗПЛ × ПДЗрайон + ДпрочЛ - 1)
    # ВАЖНО: если лаборатория в СПб — районный коэффициент к лаб. НЕ начисляется (К=1.0 в СПб)
    dz_lab_regional = 0
    lab_rp_multiplier = 0
    lab_in_spb = project_info.get("lab_in_spb", True)
    
    if pdz_r > 1.0 and lab_cost > 0 and not lab_in_spb:
        dzpl_share = reg_data.get("labor_share_lab", 0.65)
        dproch_lab = reg_data.get("other_share_lab", 0.35)
        lab_rp_multiplier = dzpl_share * pdz_r + dproch_lab - 1
        dz_lab_regional = lab_cost * lab_rp_multiplier
    
    # === Формируем список дополнительных затрат ===
    additional_costs = []
    
    if dz_unfav > 0:
        additional_costs.append({
            "name": f"ДЗ на неблагоприятный период ({unfav_percent}%)",
            "value": dz_unfav,
            "percent": unfav_percent,
            "basis": f"НЗ №281/пр, п.21, формула 4",
            "formula": f"СПпз({field_cost:,.0f}) × {unfav_percent/100:.4f}"
        })
    
    if dz_regime > 0:
        additional_costs.append({
            "name": f"ДЗ на неизбежные перерывы ({regime_percent}%)",
            "value": dz_regime,
            "percent": regime_percent,
            "basis": f"НЗ №281/пр, п.26-27, формула 6",
            "formula": f"СПпз({field_cost:,.0f}) × {regime_percent/100:.2f}"
        })
    
    if dz_travel > 0:
        interp_note = " (интерп.)" if use_interpolation else ""
        additional_costs.append({
            "name": f"ДЗ на проезд ({travel_percent:.1f}%){interp_note}",
            "value": dz_travel,
            "percent": travel_percent,
            "basis": f"НЗ №281/пр, {travel_paragraph}, {travel_table_name} (расст. {distance} км, СПпз до {cost_key.replace('up_to_','').replace('k',' тыс.')})",
            "formula": f"СПпз({field_cost:,.0f}) × {travel_percent/100:.4f}"
        })
    
    if dz_org > 0:
        additional_costs.append({
            "name": f"ДЗ на организацию полевых работ ({org_percent}%)",
            "value": dz_org,
            "percent": org_percent,
            "basis": f"НЗ №281/пр, п.37, ф.(9), Таблица 8 (расст. {distance} км, СПпз до {org_cost_key.replace('up_to_','').replace('k',' тыс.')})",
            "formula": f"СПпз({field_cost:,.0f}) × {org_percent/100:.4f}"
        })
    
    if dz_rp > 0:
        additional_costs.append({
            "name": f"ДЗ на районные выплаты (полевые, Крайон={pdz_r})",
            "value": dz_rp,
            "percent": round(rp_multiplier * 100, 2),
            "basis": f"НЗ №281/пр, п.40, формула 10",
            "formula": f"({field_cost:,.0f} + {dz_unfav:,.0f} + {dz_regime:,.0f} + {dz_org:,.0f}) × {rp_multiplier:.4f}"
        })
    
    if dz_lab_regional > 0:
        additional_costs.append({
            "name": f"ДЗ на районные выплаты (лаб., Крайон={pdz_r})",
            "value": dz_lab_regional,
            "percent": round(lab_rp_multiplier * 100, 2),
            "basis": f"НЗ №281/пр, п.47, формула 14",
            "formula": f"{lab_cost:,.0f} × {lab_rp_multiplier:.4f}"
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
                    
                    # Множитель для per_support / per_km шаблонов
                    has_per_support = any(item.get("per_support") for item in template.get("items", []))
                    has_per_km = any(item.get("per_km") for item in template.get("items", []))
                    
                    multiplier = 1
                    if has_per_support:
                        mult_label = template.get("multiplier_label", "Количество опор")
                        st.markdown(f"**🔢 {mult_label}:**")
                        multiplier = st.number_input(
                            mult_label, 
                            value=3, min_value=1, max_value=50, step=1,
                            key=f"mult_{template['id']}",
                            help=f"Объемы бурения и лаборатории умножаются на {mult_label.lower()}. Программа и отчёт — 1 раз."
                        )
                        st.divider()
                    elif has_per_km:
                        st.markdown("**🔢 Протяженность трассы (км):**")
                        multiplier = st.number_input(
                            "Количество км", 
                            value=1, min_value=1, max_value=100, step=1,
                            key=f"mult_{template['id']}",
                            help="Объемы бурения умножаются на количество км. Программа и отчёт — 1 раз."
                        )
                        st.divider()
                    
                    # Состав работ с ссылками на НЗ
                    st.markdown("**📝 Состав работ:**")
                    for item in template.get("items", []):
                        work_info = calc.get_work_type(item["work_id"])
                        base_cost = calc.get_base_cost(item["work_id"])
                        
                        is_scalable = item.get("per_support") or item.get("per_km")
                        qty = item["quantity"] * multiplier if is_scalable else item["quantity"]
                        # Рекогносцировка — двухкомпонентная (п.49, ф.16)
                        if calc.is_reconnaissance(item["work_id"]):
                            pz1p, pz2p = calc.get_reconnaissance_components(item["work_id"])
                            item_cost = float(pz1p) + float(pz2p) * qty
                        else:
                            item_cost = float(base_cost) * qty
                        
                        # Название работы
                        work_name = work_info.get('name', item['work_id'])
                        table_ref = work_info.get('table_ref', item.get('nz_ref', ''))
                        
                        col_a, col_b = st.columns([3, 1])
                        with col_a:
                            st.markdown(f"**{work_name}**")
                            if item.get("description"):
                                st.caption(f"_{item['description']}_")
                            if table_ref:
                                st.caption(f"📖 _НЗ №281/пр, {table_ref}_")
                        with col_b:
                            qty_label = f"{qty} {work_info.get('unit', 'ед.')}"
                            if is_scalable and multiplier > 1:
                                qty_label += f" (×{multiplier})"
                            st.write(qty_label)
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
                        is_scalable = item.get("per_support") or item.get("per_km")
                        qty = item["quantity"] * multiplier if is_scalable else item["quantity"]
                        # Рекогносцировка — двухкомпонентная (п.49, ф.16)
                        if calc.is_reconnaissance(item["work_id"]):
                            pz1p, pz2p = calc.get_reconnaissance_components(item["work_id"])
                            total_cost += float(pz1p) + float(pz2p) * qty
                        else:
                            total_cost += float(base_cost) * qty
                    
                    # Учитываем ДЗрежим если есть
                    regime_surcharge = 0
                    for add_cost in template.get("additional_costs", []):
                        if add_cost.get("type") == "regime_surcharge":
                            regime_surcharge = total_cost * add_cost.get("percent", 0) / 100
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        label = "💰 Базовая стоимость"
                        if multiplier > 1:
                            label += f" (×{multiplier})"
                        st.metric(label, f"{total_cost:,.0f} ₽")
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
                            is_scalable = item.get("per_support") or item.get("per_km")
                            qty = item["quantity"] * multiplier if is_scalable else item["quantity"]
                            item_data = {
                                "work_id": item["work_id"],
                                "quantity": qty,
                                "additional_coefficients": {},
                                "uid": str(uuid.uuid4())[:8]
                            }
                            st.session_state.estimate_items.append(item_data)
                        
                        # Устанавливаем параметры по умолчанию
                        default_params = template.get("default_params", {})
                        if "complexity" in default_params:
                            st.session_state.project_info["complexity"] = default_params["complexity"]
                        
                        msg = f"✅ Шаблон «{template['name']}» применён!"
                        if multiplier > 1:
                            msg += f" (×{multiplier})"
                        msg += " Перейдите на вкладку «Текущая смета»."
                        st.session_state.project_info["template_id"] = template["id"]
                        st.success(msg)
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
                is_recon = calc.is_reconnaissance(selected_work_id)
                
                col_a, col_b, col_c = st.columns(3)
                
                with col_a:
                    if is_recon:
                        quantity = st.number_input(
                            f"Площадь ({work_info.get('unit', 'га')})",
                            min_value=0.1,
                            value=1.0,
                            step=0.5,
                            help="Площадь рекогносцировочного обследования (Sреког)"
                        )
                    else:
                        quantity = st.number_input(
                            f"Количество ({work_info.get('unit', 'ед.')})",
                            min_value=0.0,
                            value=1.0,
                            step=1.0
                        )
                
                with col_b:
                    if is_recon:
                        pz1p, pz2p = calc.get_reconnaissance_components(selected_work_id)
                        st.metric("ПЗ1п (пост.)", f"{float(pz1p):,.0f} ₽")
                        st.caption(f"ПЗ2п (уд.) = {float(pz2p):,.0f} ₽/га")
                    else:
                        st.metric("Базовая цена", f"{float(base_cost):,.0f} ₽")
                
                with col_c:
                    # Расчёт предварительной стоимости
                    if is_recon:
                        pz1p, pz2p = calc.get_reconnaissance_components(selected_work_id)
                        preliminary_cost = float(pz1p) + float(pz2p) * quantity
                        st.metric("Предв. стоимость", f"{preliminary_cost:,.0f} ₽")
                        st.caption(f"ПЗ1п + ПЗ2п × S = {float(pz1p):,.0f} + {float(pz2p):,.0f} × {quantity:.1f}")
                    else:
                        preliminary_cost = float(base_cost) * quantity
                        st.metric("Предв. стоимость", f"{preliminary_cost:,.0f} ₽")
                
                # Дополнительные коэффициенты для полевых работ
                if work_category == "field":
                    with st.expander("Дополнительные коэффициенты (К)"):
                        additional_coefs = {}
                        
                        col_x, col_y = st.columns(2)
                        with col_x:
                            k_winter = st.number_input(
                                "К (зимний)", value=1.0, min_value=1.0, step=0.05, 
                                help="Коэффициент на зимние условия"
                            )
                            k_night = st.number_input(
                                "К (ночной)", value=1.0, min_value=1.0, step=0.05,
                                help="Учитывается отдельно по НЗ, но если требуется множитель базовой цены"
                            )
                            if k_winter > 1.0: additional_coefs["K_winter"] = k_winter
                            if k_night > 1.0: additional_coefs["K_night"] = k_night
                            
                        with col_y:
                            k_diff = st.number_input(
                                "К (стесненность/уклон)", value=1.0, min_value=1.0, step=0.05,
                                help="Коэффициент на стесненность, уклон и пр."
                            )
                            k_pass = st.number_input(
                                "К (проходимость)", value=1.0, min_value=1.0, step=0.05,
                                help="Условия проходимости (болота, тайга и т.д.)"
                            )
                            if k_diff > 1.0: additional_coefs["K_difficult"] = k_diff
                            if k_pass > 1.0: additional_coefs["K_passability"] = k_pass
                else:
                    additional_coefs = {}
                
                # Кнопка добавления
                if st.button("➕ Добавить в смету", type="primary", use_container_width=True):
                    if quantity > 0:
                        item_data = {
                            "work_id": selected_work_id,
                            "quantity": quantity,
                            "additional_coefficients": additional_coefs,
                            "uid": str(uuid.uuid4())[:8]
                        }
                        st.session_state.estimate_items.append(item_data)
                        
                        auto_added = []
                        # Авто-добавление камералки при бурении
                        if "drill" in selected_work_id:
                            complexity = st.session_state.project_info.get("complexity", "II")
                            cat_suffix = "cat1" if complexity == "I" else ("cat3" if complexity == "III" else "cat2")
                            cameral_id = f"cameral_borehole_{cat_suffix}"
                            
                            # Проверяем, добавлена ли уже камералка для скважин
                            has_cameral_borehole = any("cameral_borehole" in i["work_id"] for i in st.session_state.estimate_items)
                            if not has_cameral_borehole:
                                st.session_state.estimate_items.append({
                                    "work_id": cameral_id,
                                    "quantity": quantity,
                                    "additional_coefficients": {},
                                    "uid": str(uuid.uuid4())[:8]
                                })
                                auto_added.append("Камеральная обработка скважин")
                            else:
                                # Если уже есть камералка, увеличиваем ее объем
                                for i in st.session_state.estimate_items:
                                    if "cameral_borehole" in i["work_id"]:
                                        i["quantity"] += quantity
                                        auto_added.append("обновлен объем камералки скважин")
                                        break
                                        
                            # Приемка образцов
                            has_lab = any("lab_" in i["work_id"] for i in st.session_state.estimate_items)
                            if not has_lab:
                                st.session_state.estimate_items.append({
                                    "work_id": "lab_sample_prep",
                                    "quantity": round(quantity / 2.0) or 1,
                                    "additional_coefficients": {},
                                    "uid": str(uuid.uuid4())[:8]
                                })
                                auto_added.append("Приёмка образцов (базово)")

                        # Авто-добавление камералки при зондировании
                        if "static_sounding" in selected_work_id or "cpt" in selected_work_id:
                            has_cameral_cpt = any("cameral_cpt" in i["work_id"] for i in st.session_state.estimate_items)
                            if not has_cameral_cpt:
                                st.session_state.estimate_items.append({
                                    "work_id": "cameral_cpt",
                                    "quantity": quantity,
                                    "additional_coefficients": {},
                                    "uid": str(uuid.uuid4())[:8]
                                })
                                auto_added.append("Камеральная обработка зондирования")
                            else:
                                for i in st.session_state.estimate_items:
                                    if "cameral_cpt" in i["work_id"]:
                                        i["quantity"] += quantity
                                        auto_added.append("обновлен объем камералки зондирования")
                                        break

                        # Авто-добавление камералки для иных полевых испытаний (если у них нет отдельного id, пропускаем, или можно добавить логику позже)
                        
                        msg = f"Добавлено: {work_info.get('name', '')}"
                        if auto_added:
                            msg += f"\n+ Автоматом добавлено/обновлено: {', '.join(auto_added)}"
                        
                        st.success(msg)
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
            climate_zone=st.session_state.project_info.get("climate_zone", "IV"),
            apply_price_index=True,
            is_local_work=st.session_state.project_info.get("is_local_work", False)
        )
        estimate.project_code = st.session_state.project_info["code"]
        estimate.object_name = st.session_state.project_info["object"]
        estimate.customer = st.session_state.project_info["customer"]
        estimate.contractor = st.session_state.project_info["contractor"]
        estimate.price_index = Decimal(str(current_index))
        estimate.contract_coefficient = Decimal(str(k_contract))
        
        # Заголовок таблицы
        st.markdown("#### Локальная смета на работы по ИГИ")
        st.markdown(f"*Приказ Минстроя России №281/пр от 12.05.2025. Уровень цен: 01.01.2024*")
        st.divider()
        
        # Шапка таблицы
        header_cols = st.columns([0.4, 2.5, 0.6, 0.9, 1.5, 1.2, 1.0, 0.3, 0.3, 0.3])
        with header_cols[0]:
            st.markdown("**№**")
        with header_cols[1]:
            st.markdown("**Наименование работ и затрат**")
        with header_cols[2]:
            st.markdown("**Ед.**")
        with header_cols[3]:
            st.markdown("**Кол-во**")
        with header_cols[4]:
            st.markdown("**Обоснование**")
        with header_cols[5]:
            st.markdown("**Расчёт**")
        with header_cols[6]:
            st.markdown("**Стоимость**")
        with header_cols[7]:
            st.markdown("")
        with header_cols[8]:
            st.markdown("")
        with header_cols[9]:
            st.markdown("")
        
        st.divider()
        
        # Обеспечиваем uid для всех позиций (совместимость со старыми данными)
        for item in st.session_state.estimate_items:
            if "uid" not in item:
                item["uid"] = str(uuid.uuid4())[:8]
        
        # Группировка по категориям
        field_items = []
        lab_items = []
        office_items = []
        
        # === Авто-подбор Программы ИГИ (Таблица 66) ===
        # Определяем площадку из рекогносцировки и глубину из параметров проекта
        recon_area_ha = 0
        for item_data in st.session_state.estimate_items:
            if 'recon' in item_data.get("work_id", ""):
                recon_area_ha = item_data.get("quantity", 1)
                break
        
        max_depth_key = st.session_state.project_info.get("max_depth", "10")
        
        # Определяем area_key для Таблицы 66
        if recon_area_ha <= 1:
            area_suffix = "lt1ha"
        elif recon_area_ha <= 10:
            area_suffix = "10ha"
        elif recon_area_ha <= 100:
            area_suffix = "100ha"
        else:
            area_suffix = "gt100ha"
        
        # Маппинг depth_key → suffix в ID
        depth_suffix_map = {"5": "5m", "10": "10m", "15": "15m", "25": "25m", "50": "50m", "75": "75m", "over": "over"}
        depth_suffix = depth_suffix_map.get(max_depth_key, "10m")
        
        auto_program_id = f"program_cat2_{area_suffix}_{depth_suffix}"
        
        # Убираем старую программу (если была) и вставляем новую
        st.session_state.estimate_items = [
            i for i in st.session_state.estimate_items 
            if 'program' not in i.get("work_id", "")
        ]
        
        # Вставляем программу ПЕРЕД отчётом (логичный порядок: программа → камеральные → отчёт)
        program_info = calc.get_work_type(auto_program_id)
        if program_info:
            program_item = {
                "work_id": auto_program_id,
                "quantity": 1,
                "additional_coefficients": {},
                "uid": "prog_auto"
            }
            
            # Ищем позицию отчёта
            report_idx = -1
            for idx, item in enumerate(st.session_state.estimate_items):
                wt = calc.get_work_type(item.get("work_id", ""))
                if wt and wt.get("group") == "report":
                    report_idx = idx
                    break
            
            if report_idx >= 0:
                st.session_state.estimate_items.insert(report_idx, program_item)
            else:
                st.session_state.estimate_items.append(program_item)
        
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
            complexity = st.session_state.project_info.get("complexity", "II")
            # Считаем стоимость отчёта с учётом интерполяции (по Таблице 65)
            calculated_report_cost, range_desc = calc.calculate_report_cost(cameral_base_sum, complexity)
        else:
            calculated_report_cost = 0
            range_desc = ""
            complexity = "II"

        for i, item_data in enumerate(st.session_state.estimate_items):
            work_info = calc.get_work_type(item_data["work_id"])
            base_cost = calc.get_base_cost(item_data["work_id"])
            report_ref = None  # Will be set if report cost is recalculated
            quantity = item_data["quantity"]
            
            # Если это отчёт - подменяем стоимость
            if work_info.get("group") == "report" and calculated_report_cost > 0:
                base_cost = calculated_report_cost
                
                # Попытка найти точную (табличную) расценку отчёта для замены work_id (сработает для крайних без интерполяции)
                correct_report_wt = None
                for wt in calc.work_types.get("work_types", []):
                    if wt.get("group") == "report" and wt.get("base_cost") == int(calculated_report_cost):
                        correct_report_wt = wt
                        break
                
                if correct_report_wt:
                    # Подменяем work_id на правильный
                    st.session_state.estimate_items[i]["work_id"] = correct_report_wt["id"]
                    display_name = correct_report_wt["name"]
                    report_ref = correct_report_wt.get("table_ref", "")
                else:
                    display_name = work_info.get("name", "Составление технического отчета по результатам выполнения работ по ИГИ")
                    report_ref = work_info.get("table_ref", "")
                
                # Сохраняем рассчитанную стоимость в сессию
                st.session_state.estimate_items[i]["override_base_cost"] = float(calculated_report_cost)
                
                quantity = 1 # Отчет всегда 1
                total_cost = calculated_report_cost
                calc_formula = f"{calculated_report_cost:,.0f} (Таблица 65, {complexity} кат., {range_desc})"
            else:
                # Если это не отчет, убираем override (на случай если он был раньше)
                if "override_base_cost" in st.session_state.estimate_items[i] and work_info.get("group") != "report":
                    del st.session_state.estimate_items[i]["override_base_cost"]
                
                # Рекогносцировка — двухкомпонентная формула (п.49, ф.16)
                if calc.is_reconnaissance(item_data["work_id"]):
                    pz1p, pz2p = calc.get_reconnaissance_components(item_data["work_id"])
                    total_cost = float(pz1p) + float(pz2p) * quantity
                    display_name = work_info.get("name", item_data["work_id"])
                    calc_formula = f"ПЗ1п({float(pz1p):,.0f}) + ПЗ2п({float(pz2p):,.0f}) × {quantity:.1f}"
                else:
                    total_cost = float(base_cost) * quantity
                    display_name = work_info.get("name", item_data["work_id"])
                    calc_formula = f"{float(base_cost):,.0f} × {quantity:.1f}"

            # Сохраняем формулу в сессию для экспорта
            st.session_state.estimate_items[i]["formula"] = calc_formula
            
            item_row = {
                "index": i,
                "uid": item_data.get("uid", str(i)),
                "work_id": item_data["work_id"],
                "name": display_name,
                "unit": work_info.get("unit", "ед."),
                "quantity": quantity,
                "base_cost": float(base_cost),
                "total_cost": total_cost,
                "table_ref": report_ref if report_ref else work_info.get("table_ref", ""),
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
        
        row_counter = [1]
        
        # --- Вспомогательная функция для рендера раздела ---
        def render_section(section_name, section_code, items_list):
            if not items_list:
                return 0
            st.markdown(f"##### **{section_name}**")
            section_total = 0
            
            for item in items_list:
                idx = item["index"]
                uid = item["uid"]
                cols = st.columns([0.4, 2.5, 0.6, 0.9, 1.5, 1.2, 1.0, 0.3, 0.3, 0.3])
                
                with cols[0]:
                    st.write(f"{row_counter[0]}")
                with cols[1]:
                    st.write(item["name"])
                with cols[2]:
                    st.write(item["unit"])
                with cols[3]:
                    new_qty = st.number_input(
                        "qty", value=float(item["quantity"]),
                        min_value=0.0, step=1.0, format="%.1f",
                        key=f"qty_{uid}", label_visibility="collapsed"
                    )
                    if new_qty != float(item["quantity"]):
                        st.session_state.estimate_items[idx]["quantity"] = new_qty
                        st.rerun()
                with cols[4]:
                    st.caption(f"НЗ №281/пр, {item['table_ref']}")
                with cols[5]:
                    st.caption(item.get("formula_display", f"{item['base_cost']:,.0f} × {new_qty:.1f}"))
                with cols[6]:
                    # Пересчитываем с учётом изменённого кол-ва
                    work_info_r = calc.get_work_type(item["work_id"])
                    if calc.is_reconnaissance(item["work_id"]):
                        pz1p, pz2p = calc.get_reconnaissance_components(item["work_id"])
                        actual_cost = float(pz1p) + float(pz2p) * new_qty
                    else:
                        actual_cost = item["base_cost"] * new_qty
                    st.write(f"**{actual_cost:,.0f}**")
                with cols[7]:
                    # Найти предыдущий элемент той же категории
                    prev_idx = None
                    for j in range(idx - 1, -1, -1):
                        wid = st.session_state.estimate_items[j]["work_id"]
                        wi = calc.get_work_type(wid)
                        if wi.get("category", "field") == item["category"]:
                            prev_idx = j
                            break
                    if prev_idx is not None:
                        if st.button("⬆", key=f"up_{uid}"):
                            items = st.session_state.estimate_items
                            items[idx], items[prev_idx] = items[prev_idx], items[idx]
                            st.rerun()
                with cols[8]:
                    # Найти следующий элемент той же категории
                    next_idx = None
                    for j in range(idx + 1, len(st.session_state.estimate_items)):
                        wid = st.session_state.estimate_items[j]["work_id"]
                        wi = calc.get_work_type(wid)
                        if wi.get("category", "field") == item["category"]:
                            next_idx = j
                            break
                    if next_idx is not None:
                        if st.button("⬇", key=f"dn_{uid}"):
                            items = st.session_state.estimate_items
                            items[idx], items[next_idx] = items[next_idx], items[idx]
                            st.rerun()
                with cols[9]:
                    if st.button("🗑️", key=f"del_{uid}"):
                        st.session_state.estimate_items.pop(idx)
                        st.rerun()
                
                row_counter[0] += 1
                # Пересчитываем с учётом изменённого кол-ва
                if calc.is_reconnaissance(item["work_id"]):
                    pz1p_s, pz2p_s = calc.get_reconnaissance_components(item["work_id"])
                    section_total += float(pz1p_s) + float(pz2p_s) * new_qty
                else:
                    section_total += item["base_cost"] * new_qty
            
            # Итого по разделу
            cols = st.columns([0.4, 2.5, 0.6, 0.9, 1.5, 1.2, 1.0, 0.3, 0.3, 0.3])
            with cols[1]:
                st.markdown(f"**Итого по {section_code}:**")
            with cols[6]:
                st.markdown(f"**{section_total:,.0f}**")
            st.divider()
            return section_total
        
        # Рендеринг разделов
        field_total = render_section("Раздел I. Полевые работы", "разделу I (СПпз)", field_items)
        lab_total = render_section("Раздел II. Лабораторные работы", "разделу II (СЛпз)", lab_items)
        office_total = render_section("Раздел III. Камеральные работы", "разделу III (СКпз)", office_items)
        
        # Общие итоги (field_total, lab_total, office_total уже посчитаны в render_section)
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
             climate_zone=st.session_state.project_info.get("climate_zone", "IV"),
             is_local_work=st.session_state.project_info.get("is_local_work", False)
        )
        field_cost_base = float(temp_estimate.subtotal_field)
        lab_cost_base = float(temp_estimate.subtotal_laboratory)
        
        dz_list = calculate_additional_costs(field_cost_base, st.session_state.project_info, lab_cost=lab_cost_base)
        dz_sum = sum(item["value"] for item in dz_list)
        final_total_base = base_total + dz_sum
        
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
            
        # Применяем индекс и коэффициент договорной цены
        pi = st.session_state.project_info.get("price_index", 1.0)
        kc = st.session_state.project_info.get("k_contract", 1.0)
        final_total = final_total_base * pi * kc
        
        # Финальный итог крупно
        st.markdown(f"### 🏁 ИТОГО: {final_total:,.0f} ₽")
        if pi != 1.0 or kc != 1.0:
            parts = [f"базовая: {final_total_base:,.0f} ₽"]
            if pi != 1.0:
                parts.append(f"× Индекс {pi:.2f}")
            if kc != 1.0:
                parts.append(f"× Кдог. {kc:.3f}")
            st.caption(" | ".join(parts))
        
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
            climate_zone=st.session_state.project_info.get("climate_zone", "IV"),
            apply_price_index=True,
            is_local_work=st.session_state.project_info.get("is_local_work", False)
        )
        
        field_cost = float(estimate.subtotal_field)
        lab_cost = float(estimate.subtotal_laboratory)
        
        st.markdown(f"**Стоимость полевых работ:** {field_cost:,.0f} ₽")
        if lab_cost > 0:
            st.markdown(f"**Стоимость лабораторных работ:** {lab_cost:,.0f} ₽")
        st.divider()
        
        # Расчёт дополнительных затрат
        additional_costs_list = calculate_additional_costs(field_cost, st.session_state.project_info, lab_cost=lab_cost)
        
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
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("📊 Итого ДЗ", f"{total_dz:,.0f} ₽")
        with col2:
            st.metric("💰 ВСЕГО (базовые цены)", f"{total_with_dz:,.0f} ₽")
        with col3:
            final_total = float(estimate.total)
            if float(estimate.price_index) != 1.0 or float(estimate.contract_coefficient) != 1.0:
                label = "💰 ИТОГО"
                notes = []
                if float(estimate.price_index) != 1.0:
                    notes.append(f"Инд.={float(estimate.price_index):.2f}")
                if float(estimate.contract_coefficient) != 1.0:
                    notes.append(f"Кдог.={float(estimate.contract_coefficient):.3f}")
                label += f" ({', '.join(notes)})"
                st.metric(label, f"{final_total:,.0f} ₽")
            else:
                st.metric("💰 ИТОГО", f"{final_total:,.0f} ₽")


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
            climate_zone=st.session_state.project_info.get("climate_zone", "IV"),
            apply_price_index=True,
            is_local_work=st.session_state.project_info.get("is_local_work", False)
        )
        
        # Добавляем ДЗ
        field_cost = float(estimate.subtotal_field)
        lab_cost = float(estimate.subtotal_laboratory)
        d_costs = calculate_additional_costs(field_cost, st.session_state.project_info, lab_cost=lab_cost)
        estimate.additional_costs = d_costs
        
        estimate.project_code = st.session_state.project_info["code"]
        estimate.object_name = st.session_state.project_info["object"]
        estimate.customer = st.session_state.project_info["customer"]
        estimate.contractor = st.session_state.project_info["contractor"]
        estimate.price_index = Decimal(str(current_index))
        estimate.contract_coefficient = Decimal(str(k_contract))
        estimate.base_city = "г. Санкт-Петербург"
        estimate.work_region = st.session_state.project_info.get("region", "")
        estimate.distance_km = st.session_state.project_info.get("distance_km", 0)
        estimate.template_id = st.session_state.project_info.get("template_id", "")
        
        # ДЗ уже добавлены выше; повторный вызов не нужен
        
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
