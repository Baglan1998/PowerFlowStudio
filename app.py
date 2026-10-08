import base64
import hmac
import json
import math
import re
import zlib
import pandas as pd
import pandapower as pp
import pandapower.shortcircuit as sc
import streamlit as st


# Полная копия справочных таблиц 35–1150 кВ из приложенного к источнику XLSX.
# Данные сжаты, чтобы приложение оставалось одним переносимым файлом app.py.
FULL_TRANSFORMER_REFERENCE_B64 = """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"""


def full_transformer_reference():
    raw = zlib.decompress(base64.b64decode(FULL_TRANSFORMER_REFERENCE_B64))
    return json.loads(raw.decode("utf-8"))


def first_number(value):
    if value is None:
        return None
    match = re.search(r"-?\d+(?:[.,]\d+)?", str(value))
    return float(match.group(0).replace(",", ".")) if match else None


def voltage_class(voltage):
    if voltage >= 700:
        return 1150 if voltage >= 1000 else 750
    if voltage >= 480:
        return 500
    if voltage >= 300:
        return 330
    if voltage >= 180:
        return 220
    if voltage >= 130:
        return 150
    if voltage >= 70:
        return 110
    return 35


def build_full_two_winding_catalog():
    result = {}
    for section in full_transformer_reference():
        if not section["kind"].startswith("2w"):
            continue
        regulated = section["kind"] == "2w_reg"
        offset = 1 if regulated else 0
        for row in section["rows"]:
            sn = first_number(row[1])
            hv = first_number(row[3 if regulated else 2])
            lv = first_number(row[4 if regulated else 3])
            vk = first_number(row[5 if regulated else 4])
            pk = first_number(row[6 if regulated else 5])
            pfe = first_number(row[7 if regulated else 6])
            i0 = first_number(row[8 if regulated else 7])
            if None in (sn, hv, lv, vk, pk, pfe, i0):
                continue
            name = f"{str(row[0]).strip()} ({hv:g}/{lv:g} кВ)"
            result[name] = {"class_kv": voltage_class(hv), "sn_mva": sn,
                "vn_hv_kv": hv, "vn_lv_kv": lv, "vk_percent": vk,
                "pk_kw": pk, "pfe_kw": pfe, "i0_percent": i0}
    return result


def build_three_winding_catalog():
    result = {}
    for section in full_transformer_reference():
        kind = section["kind"]
        if not kind.startswith("3w"):
            continue
        if kind == "3w_reg":
            hv_i, mv_i, lv_i, uk_i, pk_i, pfe_i, i0_i = 3, 4, 5, 6, 9, 12, 13
        elif kind == "3w_500":
            hv_i, mv_i, lv_i, uk_i, pk_i, pfe_i, i0_i = 3, 4, 5, 9, 12, 13, 14
        else:
            hv_i, mv_i, lv_i, uk_i, pk_i, pfe_i, i0_i = 2, 3, 4, 5, 8, 11, 12
        for row in section["rows"]:
            sn = first_number(row[1]); hv = first_number(row[hv_i])
            mv = first_number(row[mv_i]); lv = first_number(row[lv_i])
            uk_hm = first_number(row[uk_i]); uk_hl = first_number(row[uk_i + 1])
            uk_ml = first_number(row[uk_i + 2]); pk = first_number(row[pk_i])
            pfe = first_number(row[pfe_i]); i0 = first_number(row[i0_i])
            if None in (sn, hv, mv, lv, uk_hm, uk_hl, uk_ml, pk, pfe, i0):
                continue
            name = f"{str(row[0]).strip()} ({hv:g}/{mv:g}/{lv:g} кВ)"
            result[name] = {"class_kv": voltage_class(hv), "sn_mva": sn,
                "vn_hv_kv": hv, "vn_mv_kv": mv, "vn_lv_kv": lv,
                "vk_hv_percent": uk_hm, "vk_mv_percent": uk_ml,
                "vk_lv_percent": uk_hl, "pk_kw": pk,
                "pfe_kw": pfe, "i0_percent": i0}
    return result


# Справочник двухобмоточных трансформаторов. Каталожные значения:
# Sном, Uвн, Uнн, Uк, потери КЗ, потери ХХ и ток ХХ.
# Для класса 1150 кВ используется эквивалент трёхфазной группы ВН–СН.
TRANSFORMER_CATALOG = {
    "ТМ-100/35 (35/0.4 кВ)": {"class_kv": 35, "sn_mva": 0.1, "vn_hv_kv": 35.0, "vn_lv_kv": 0.4, "vk_percent": 6.5, "pk_kw": 1.9, "pfe_kw": 0.5, "i0_percent": 2.6},
    "ТМН-630/35 (35/0.4 кВ)": {"class_kv": 35, "sn_mva": 0.63, "vn_hv_kv": 35.0, "vn_lv_kv": 0.4, "vk_percent": 6.5, "pk_kw": 11.6, "pfe_kw": 2.7, "i0_percent": 1.5},
    "ТД-16000/35 (38.5/10.5 кВ)": {"class_kv": 35, "sn_mva": 16.0, "vn_hv_kv": 38.5, "vn_lv_kv": 10.5, "vk_percent": 8.0, "pk_kw": 90.0, "pfe_kw": 21.0, "i0_percent": 0.6},
    "ТМН-2500/110 (110/11 кВ)": {"class_kv": 110, "sn_mva": 2.5, "vn_hv_kv": 110.0, "vn_lv_kv": 11.0, "vk_percent": 10.5, "pk_kw": 22.0, "pfe_kw": 5.5, "i0_percent": 1.5},
    "ТДН-16000/110 (115/11 кВ)": {"class_kv": 110, "sn_mva": 16.0, "vn_hv_kv": 115.0, "vn_lv_kv": 11.0, "vk_percent": 10.5, "pk_kw": 85.0, "pfe_kw": 19.0, "i0_percent": 0.7},
    "ТДЦ-125000/110 (121/10.5 кВ)": {"class_kv": 110, "sn_mva": 125.0, "vn_hv_kv": 121.0, "vn_lv_kv": 10.5, "vk_percent": 10.5, "pk_kw": 400.0, "pfe_kw": 120.0, "i0_percent": 0.55},
    "ТДН-16000/150 (158/11 кВ)": {"class_kv": 150, "sn_mva": 16.0, "vn_hv_kv": 158.0, "vn_lv_kv": 11.0, "vk_percent": 11.0, "pk_kw": 85.0, "pfe_kw": 21.0, "i0_percent": 0.8},
    "ТЦ-250000/150 (165/15.75 кВ)": {"class_kv": 150, "sn_mva": 250.0, "vn_hv_kv": 165.0, "vn_lv_kv": 15.75, "vk_percent": 11.0, "pk_kw": 640.0, "pfe_kw": 190.0, "i0_percent": 0.5},
    "ТРДН-40000/220 (230/11 кВ)": {"class_kv": 220, "sn_mva": 40.0, "vn_hv_kv": 230.0, "vn_lv_kv": 11.0, "vk_percent": 12.0, "pk_kw": 170.0, "pfe_kw": 50.0, "i0_percent": 0.9},
    "ТДЦ-125000/220 (242/13.8 кВ)": {"class_kv": 220, "sn_mva": 125.0, "vn_hv_kv": 242.0, "vn_lv_kv": 13.8, "vk_percent": 11.0, "pk_kw": 380.0, "pfe_kw": 135.0, "i0_percent": 0.5},
    "ТДЦ-400000/220 (242/20 кВ)": {"class_kv": 220, "sn_mva": 400.0, "vn_hv_kv": 242.0, "vn_lv_kv": 20.0, "vk_percent": 11.0, "pk_kw": 880.0, "pfe_kw": 330.0, "i0_percent": 0.4},
    "ТДЦ-125000/330 (347/10.5 кВ)": {"class_kv": 330, "sn_mva": 125.0, "vn_hv_kv": 347.0, "vn_lv_kv": 10.5, "vk_percent": 11.0, "pk_kw": 360.0, "pfe_kw": 145.0, "i0_percent": 0.5},
    "ТЦС-400000/330 (347/20 кВ)": {"class_kv": 330, "sn_mva": 400.0, "vn_hv_kv": 347.0, "vn_lv_kv": 20.0, "vk_percent": 11.0, "pk_kw": 810.0, "pfe_kw": 365.0, "i0_percent": 0.4},
    "ТДЦ-250000/500 (525/15.75 кВ)": {"class_kv": 500, "sn_mva": 250.0, "vn_hv_kv": 525.0, "vn_lv_kv": 15.75, "vk_percent": 13.0, "pk_kw": 600.0, "pfe_kw": 250.0, "i0_percent": 0.45},
    "ТДЦ-400000/500 (525/20 кВ)": {"class_kv": 500, "sn_mva": 400.0, "vn_hv_kv": 525.0, "vn_lv_kv": 20.0, "vk_percent": 13.0, "pk_kw": 800.0, "pfe_kw": 350.0, "i0_percent": 0.4},
    "ТЦ-1000000/500 (525/24 кВ)": {"class_kv": 500, "sn_mva": 1000.0, "vn_hv_kv": 525.0, "vn_lv_kv": 24.0, "vk_percent": 14.5, "pk_kw": 2000.0, "pfe_kw": 600.0, "i0_percent": 0.38},
    "ОРЦ-417000/750 (787/24 кВ)": {"class_kv": 750, "sn_mva": 417.0, "vn_hv_kv": 787.0, "vn_lv_kv": 24.0, "vk_percent": 14.0, "pk_kw": 800.0, "pfe_kw": 400.0, "i0_percent": 0.3},
    "АОДЦТ-2001000/1150/500 экв. (1150/500 кВ)": {"class_kv": 1150, "sn_mva": 2001.0, "vn_hv_kv": 1150.0, "vn_lv_kv": 500.0, "vk_percent": 11.5, "pk_kw": 3750.0, "pfe_kw": 1050.0, "i0_percent": 0.35},
}
TRANSFORMER_CATALOG.update(build_full_two_winding_catalog())
THREE_WINDING_CATALOG = build_three_winding_catalog()


def pandapower_trafo_data(item):
    """Преобразует каталожные потери КЗ в vkr для pandapower."""
    return {
        "sn_mva": item["sn_mva"],
        "vn_hv_kv": item["vn_hv_kv"],
        "vn_lv_kv": item["vn_lv_kv"],
        "vk_percent": item["vk_percent"],
        "vkr_percent": item["pk_kw"] / (item["sn_mva"] * 10.0),
        "pfe_kw": item["pfe_kw"],
        "i0_percent": item["i0_percent"],
        "shift_degree": 0.0,
    }


def register_transformer_catalog(net):
    for name, item in TRANSFORMER_CATALOG.items():
        pp.create_std_type(net, pandapower_trafo_data(item), name, element="trafo", overwrite=True)


def transformer_catalog_frame(class_kv):
    rows = []
    for name, item in TRANSFORMER_CATALOG.items():
        if item["class_kv"] == class_kv:
            rows.append({
                "Тип": name,
                "Sном, МВА": item["sn_mva"],
                "U ВН, кВ": item["vn_hv_kv"],
                "U НН, кВ": item["vn_lv_kv"],
                "Uк, %": item["vk_percent"],
                "ΔPк, кВт": item["pk_kw"],
                "Pх, кВт": item["pfe_kw"],
                "Iх, %": item["i0_percent"],
            })
    return pd.DataFrame(rows)


def three_winding_catalog_frame(class_kv):
    rows = []
    for name, item in THREE_WINDING_CATALOG.items():
        if item["class_kv"] == class_kv:
            rows.append({
                "Тип": name, "Sном, МВА": item["sn_mva"],
                "U ВН, кВ": item["vn_hv_kv"], "U СН, кВ": item["vn_mv_kv"],
                "U НН, кВ": item["vn_lv_kv"], "Uк В-С, %": item["vk_hv_percent"],
                "Uк В-Н, %": item["vk_lv_percent"], "Uк С-Н, %": item["vk_mv_percent"],
                "ΔPк, кВт": item["pk_kw"], "Pх, кВт": item["pfe_kw"],
                "Iх, %": item["i0_percent"],
            })
    return pd.DataFrame(rows)


def example():
    return {
        "buses": [{"ID": 1, "Название": "Bus 1", "Uном, кВ": 20.0}, {"ID": 2, "Название": "Bus 2", "Uном, кВ": 0.4}, {"ID": 3, "Название": "Bus 3", "Uном, кВ": 0.4}],
        "grid": [{"Узел": 1, "U, о.е.": 1.02, "Sкз max, МВА": 500.0, "R/X сети": 0.1}],
        "lines": [{"Название": "Линия 1", "Начало": 2, "Конец": 3, "L, км": 0.1, "Режим": "Авто", "Материал": "Алюминий", "R, Ом/км": 0.642, "X, Ом/км": 0.083, "C, нФ/км": 210.0, "Imax, А": 142.0}],
        "trafos": [{"Название": "Трансформатор 1", "ВН": 1, "НН": 2, "Тип": "0.4 MVA 20/0.4 kV"}],
        "trafos3w": [],
        "loads": [{"Название": "Нагрузка 1", "Узел": 3, "P, кВт": 100.0, "Q, квар": 50.0}],
        "generation": [{"Название": "Генератор 1", "Узел": 3, "P, кВт": 0.0, "Q, квар": 0.0}],
    }

COLUMNS = {
    "buses": ["ID", "Название", "Uном, кВ"],
    "grid": ["Узел", "U, о.е.", "Sкз max, МВА", "R/X сети"],
    "lines": ["Название", "Начало", "Конец", "L, км", "Режим", "Материал", "R, Ом/км", "X, Ом/км", "C, нФ/км", "Imax, А"],
    "trafos": ["Название", "ВН", "НН", "Тип"],
    "trafos3w": ["Название", "ВН", "СН", "НН", "Тип"],
    "loads": ["Название", "Узел", "P, кВт", "Q, квар"],
    "generation": ["Название", "Узел", "P, кВт", "Q, квар"],
}
GRID_DEFAULTS = {"Sкз max, МВА": 500.0, "R/X сети": 0.1}
LINE_DEFAULTS = {"Режим": "Авто", "Материал": "Алюминий"}
LINE_MODES = ["Авто", "Вручную"]
LINE_MATERIALS = ["Алюминий", "Медь"]


# Расчётные данные КЛ 1–35 кВ на 1 км.
# Источник: справочные материалы для курсовых проектов С. С. Ананичевой,
# С. Н. Шелюга (приведены на powersystem.info). Значения R — при +20 °C.
CABLE_R_OHM_KM = {
    10: {"Медь": 1.790, "Алюминий": 2.940},
    16: {"Медь": 1.120, "Алюминий": 1.840},
    25: {"Медь": 0.720, "Алюминий": 1.700},
    35: {"Медь": 0.510, "Алюминий": 0.840},
    50: {"Медь": 0.360, "Алюминий": 0.590},
    70: {"Медь": 0.256, "Алюминий": 0.420},
    95: {"Медь": 0.190, "Алюминий": 0.310},
    120: {"Медь": 0.150, "Алюминий": 0.240},
    150: {"Медь": 0.120, "Алюминий": 0.200},
    185: {"Медь": 0.100, "Алюминий": 0.160},
    240: {"Медь": 0.070, "Алюминий": 0.120},
    300: {"Медь": 0.061, "Алюминий": 0.103},
    400: {"Медь": 0.046, "Алюминий": 0.077},
}
CABLE_X_OHM_KM = {
    1: {10: 0.073, 16: 0.068, 25: 0.066, 35: 0.064, 50: 0.063, 70: 0.061,
        95: 0.060, 120: 0.060, 150: 0.059, 185: 0.059, 240: 0.058},
    6: {10: 0.110, 16: 0.102, 25: 0.091, 35: 0.087, 50: 0.083, 70: 0.080,
        95: 0.078, 120: 0.076, 150: 0.074, 185: 0.073, 240: 0.071},
    10: {10: 0.122, 16: 0.113, 25: 0.099, 35: 0.095, 50: 0.090, 70: 0.086,
         95: 0.083, 120: 0.081, 150: 0.079, 185: 0.077, 240: 0.075},
    20: {25: 0.135, 35: 0.129, 50: 0.119, 70: 0.116, 95: 0.110, 120: 0.107,
         150: 0.104, 185: 0.101, 240: 0.098, 300: 0.095, 400: 0.092},
    35: {70: 0.137, 95: 0.126, 120: 0.120, 150: 0.116, 185: 0.113,
         240: 0.111, 300: 0.097, 400: 0.094},
}
CABLE_INSULATION_LOSS_KW_KM = {
    6: {10: 0.016, 16: 0.019, 25: 0.030, 35: 0.033, 50: 0.038, 70: 0.048,
        95: 0.063, 120: 0.068, 150: 0.076, 185: 0.084, 240: 0.095},
    10: {10: 0.038, 16: 0.042, 25: 0.063, 35: 0.078, 50: 0.087, 70: 0.098,
         95: 0.113, 120: 0.123, 150: 0.134, 185: 0.146, 240: 0.191},
    20: {25: 0.135, 35: 0.151, 50: 0.174, 70: 0.196, 95: 0.219, 120: 0.234,
         150: 0.257, 185: 0.279, 240: 0.320},
    35: {70: 0.461, 95: 0.508, 120: 0.532, 150: 0.600, 185: 0.623,
         240: 0.813},
}


def cable_voltage_class(voltage_kv):
    """Выбирает ближайший верхний класс напряжения из справочника КЛ."""
    for value in (1, 6, 10, 20, 35):
        if voltage_kv <= value:
            return value
    raise ValueError("Справочник кабелей действует для линий до 35 кВ.")


def cable_reference_frame(voltage_kv):
    """Таблица R, X и удельных потерь изоляции для отображения в приложении."""
    voltage = cable_voltage_class(voltage_kv)
    rows = []
    for section in sorted(CABLE_R_OHM_KM):
        reactance = CABLE_X_OHM_KM[voltage].get(section)
        if reactance is None:
            continue
        rows.append({
            "Сечение, мм²": section,
            "R Cu, Ом/км": CABLE_R_OHM_KM[section]["Медь"],
            "R Al, Ом/км": CABLE_R_OHM_KM[section]["Алюминий"],
            "X, Ом/км": reactance,
            "Потери изоляции, кВт/км": CABLE_INSULATION_LOSS_KW_KM.get(voltage, {}).get(section),
        })
    return pd.DataFrame(rows)


def cable_candidates(voltage_kv, material):
    """Справочник кабелей 1–35 кВ и проводов ВЛ выше 35 кВ."""
    rho = 29.4 if material == "Алюминий" else 18.1
    prefix = "Al" if material == "Алюминий" else "Cu"
    if voltage_kv <= 35.0:
        source_voltage = cable_voltage_class(voltage_kv)
        ratings = {
            "Алюминий": [(16, 65), (25, 85), (35, 105), (50, 130), (70, 165),
                         (95, 205), (120, 235), (150, 270), (185, 310),
                         (240, 370), (300, 425), (400, 500), (500, 670), (630, 760)],
            "Медь": [(16, 85), (25, 110), (35, 135), (50, 165), (70, 210),
                     (95, 260), (120, 300), (150, 340), (185, 390),
                     (240, 460), (300, 530), (400, 620), (500, 805), (630, 900)],
        }[material]
        result = []
        for section, ampacity in ratings:
            if section not in CABLE_X_OHM_KM[source_voltage]:
                continue
            result.append({
                "name": f"КЛ {source_voltage} кВ {prefix} {section} мм²",
                "r": CABLE_R_OHM_KM[section][material],
                "x": CABLE_X_OHM_KM[source_voltage][section],
                # В исходной справочной таблице C нет: оставлено типовое значение,
                # которое при необходимости можно заменить в ручном режиме.
                "c": 250.0 if source_voltage == 1 else 200.0,
                "imax": ampacity,
                "section": section,
                "p_iso_kw_km": CABLE_INSULATION_LOSS_KW_KM.get(source_voltage, {}).get(section),
            })
        return result

    if voltage_kv <= 150:
        bundle, minimum = 1, 70
    elif voltage_kv <= 220:
        bundle, minimum = 1, 240
    elif voltage_kv <= 330:
        bundle, minimum = 2, 300
    elif voltage_kv <= 500:
        bundle, minimum = 3, 300
    elif voltage_kv <= 750:
        bundle, minimum = 5, 300
    else:
        bundle, minimum = 8, 330
    base = [(70, 265), (95, 330), (120, 390), (150, 450), (185, 510),
            (240, 605), (300, 710), (330, 760), (400, 860),
            (500, 960), (600, 1050)]
    result = []
    for section, ampacity in base:
        if section < minimum:
            continue
        kind = "АС" if material == "Алюминий" else "М"
        name = f"{bundle}×{kind}-{section}" if bundle > 1 else f"{kind}-{section}"
        result.append({
            "name": f"Провод {name}",
            "r": (rho / section) / bundle,
            "x": 0.40 if bundle == 1 else 0.32,
            "c": 9.5 + 1.5 * (bundle - 1),
            "imax": ampacity * bundle * (1.10 if material == "Медь" else 1.0),
            "section": section * bundle,
            "p_iso_kw_km": None,
        })
    return result

def select_cable(voltage_kv, required_current_a, material, largest=False):
    if material not in LINE_MATERIALS:
        raise ValueError("Материал линии должен быть «Алюминий» или «Медь».")
    candidates = cable_candidates(voltage_kv, material)
    if not candidates:
        raise ValueError(f"Нет кабеля или провода для напряжения {voltage_kv:g} кВ.")
    if largest:
        return candidates[-1], False
    for item in candidates:
        if item["imax"] >= required_current_a:
            return item, False
    return candidates[-1], True


def apply_auto_cables(net, reserve=1.25):
    selected = []
    for i, line in net.line.iterrows():
        if not bool(line.get("auto_select", False)):
            selected.append("Вручную")
            continue
        current_a = float(net.res_line.at[i, "i_ka"]) * 1000
        voltage_kv = float(net.bus.at[int(line["from_bus"]), "vn_kv"])
        spec, insufficient = select_cable(
            voltage_kv, current_a * reserve, str(line["material"]))
        net.line.at[i, "r_ohm_per_km"] = spec["r"]
        net.line.at[i, "x_ohm_per_km"] = spec["x"]
        net.line.at[i, "c_nf_per_km"] = spec["c"]
        net.line.at[i, "max_i_ka"] = spec["imax"] / 1000
        net.line.at[i, "selected_conductor"] = spec["name"]
        net.line.at[i, "required_current_a"] = current_a * reserve
        net.line.at[i, "insulation_loss_kw_per_km"] = (
            spec["p_iso_kw_km"] if spec["p_iso_kw_km"] is not None else float("nan")
        )
        net.line.at[i, "auto_insufficient"] = bool(insufficient)
        selected.append(spec["name"])
    return tuple(selected)


def migrate_project(data):
    """Дополняет старые проекты полями новых версий."""
    if isinstance(data, dict):
        data.setdefault("trafos3w", [])
    if isinstance(data, dict) and isinstance(data.get("grid"), list):
        for row in data["grid"]:
            if isinstance(row, dict):
                for key, value in GRID_DEFAULTS.items():
                    row.setdefault(key, value)
    if isinstance(data, dict) and isinstance(data.get("lines"), list):
        for row in data["lines"]:
            if isinstance(row, dict):
                for key, value in LINE_DEFAULTS.items():
                    row.setdefault(key, value)
    return data


def number(row, key, minimum=None, positive=False):
    try:
        value = float(row[key])
    except (TypeError, ValueError, KeyError):
        raise ValueError(f"Заполните числом поле «{key}».")
    if not math.isfinite(value) or (minimum is not None and value < minimum) or (positive and value <= 0):
        raise ValueError(f"Недопустимое значение «{key}»: {value}.")
    return value


def integer(row, key):
    value = number(row, key, minimum=0)
    if value != int(value):
        raise ValueError(f"«{key}» должен быть целым номером узла.")
    return int(value)


def build(project):
    net = pp.create_empty_network()
    register_transformer_catalog(net)
    for r in project["buses"]:
        i = integer(r, "ID")
        if i in net.bus.index:
            raise ValueError(f"Номер узла {i} повторяется.")
        pp.create_bus(net, index=i, vn_kv=number(r, "Uном, кВ", positive=True), name=str(r["Название"]))
    if net.bus.empty:
        raise ValueError("Добавьте хотя бы один узел.")

    def bus(r, key):
        i = integer(r, key)
        if i not in net.bus.index:
            raise ValueError(f"Узел {i} из поля «{key}» отсутствует в таблице узлов.")
        return i

    if len(project["grid"]) != 1:
        raise ValueError("В этой версии задайте ровно одну внешнюю сеть.")
    r = project["grid"][0]
    root = bus(r, "Узел")
    pp.create_ext_grid(
        net, bus=root, vm_pu=number(r, "U, о.е.", positive=True),
        s_sc_max_mva=number(r, "Sкз max, МВА", positive=True),
        rx_max=number(r, "R/X сети", minimum=0),
    )
    links = {i: set() for i in net.bus.index}
    for r in project["lines"]:
        a, b = bus(r, "Начало"), bus(r, "Конец")
        if a == b:
            raise ValueError("Начало и конец линии должны отличаться.")
        if not math.isclose(net.bus.at[a, "vn_kv"], net.bus.at[b, "vn_kv"]):
            raise ValueError("Линия соединяет узлы разных напряжений. Используйте трансформатор.")
        mode = str(r.get("Режим", "Авто"))
        if mode not in LINE_MODES:
            raise ValueError("Режим линии должен быть «Авто» или «Вручную».")
        auto_select = mode == "Авто"
        material = str(r.get("Материал", "Алюминий"))
        if auto_select:
            spec, _ = select_cable(net.bus.at[a, "vn_kv"], 0, material, largest=True)
            resistance, reactance = spec["r"], spec["x"]
            capacitance, imax_a = spec["c"], spec["imax"]
            selected_name = spec["name"]
        else:
            resistance = number(r, "R, Ом/км", minimum=0)
            reactance = number(r, "X, Ом/км", minimum=0)
            capacitance = number(r, "C, нФ/км", minimum=0)
            imax_a = number(r, "Imax, А", positive=True)
            selected_name = "Вручную"
        if resistance == reactance == 0:
            raise ValueError("R и X линии не могут одновременно равняться нулю.")
        line_index = pp.create_line_from_parameters(
            net, from_bus=a, to_bus=b, length_km=number(r, "L, км", positive=True),
            r_ohm_per_km=resistance, x_ohm_per_km=reactance,
            c_nf_per_km=capacitance, max_i_ka=imax_a / 1000,
            name=str(r["Название"]), endtemp_degree=20)
        net.line.at[line_index, "auto_select"] = auto_select
        net.line.at[line_index, "material"] = material
        net.line.at[line_index, "selected_conductor"] = selected_name
        net.line.at[line_index, "required_current_a"] = float("nan")
        net.line.at[line_index, "insulation_loss_kw_per_km"] = (
            spec["p_iso_kw_km"] if auto_select and spec["p_iso_kw_km"] is not None else float("nan")
        )
        net.line.at[line_index, "auto_insufficient"] = False
        links[a].add(b); links[b].add(a)
    for r in project["trafos"]:
        a, b = bus(r, "ВН"), bus(r, "НН")
        typ = r["Тип"]
        if typ not in net.std_types["trafo"]:
            raise ValueError(f"Неизвестный тип трансформатора: {typ}.")
        spec = net.std_types["trafo"][typ]
        if a == b or not math.isclose(net.bus.at[a, "vn_kv"], spec["vn_hv_kv"]) or not math.isclose(net.bus.at[b, "vn_kv"], spec["vn_lv_kv"]):
            raise ValueError("Напряжения узлов ВН и НН должны соответствовать типу трансформатора.")
        pp.create_transformer(net, hv_bus=a, lv_bus=b, std_type=typ, name=str(r["Название"]))
        links[a].add(b); links[b].add(a)
    for r in project["trafos3w"]:
        a, b, c = bus(r, "ВН"), bus(r, "СН"), bus(r, "НН")
        typ = r["Тип"]
        if typ not in THREE_WINDING_CATALOG:
            raise ValueError(f"Неизвестный тип трёхобмоточного трансформатора: {typ}.")
        spec = THREE_WINDING_CATALOG[typ]
        expected = (spec["vn_hv_kv"], spec["vn_mv_kv"], spec["vn_lv_kv"])
        actual = (net.bus.at[a, "vn_kv"], net.bus.at[b, "vn_kv"], net.bus.at[c, "vn_kv"])
        if len({a, b, c}) != 3 or any(not math.isclose(x, y) for x, y in zip(actual, expected)):
            raise ValueError("Напряжения узлов ВН, СН и НН должны соответствовать типу трёхобмоточного трансформатора.")
        vkr = spec["pk_kw"] / (spec["sn_mva"] * 30.0)
        pp.create_transformer3w_from_parameters(
            net, hv_bus=a, mv_bus=b, lv_bus=c,
            sn_hv_mva=spec["sn_mva"], sn_mv_mva=spec["sn_mva"], sn_lv_mva=spec["sn_mva"],
            vn_hv_kv=spec["vn_hv_kv"], vn_mv_kv=spec["vn_mv_kv"], vn_lv_kv=spec["vn_lv_kv"],
            vk_hv_percent=spec["vk_hv_percent"], vk_mv_percent=spec["vk_mv_percent"],
            vk_lv_percent=spec["vk_lv_percent"], vkr_hv_percent=vkr,
            vkr_mv_percent=vkr, vkr_lv_percent=vkr, pfe_kw=spec["pfe_kw"],
            i0_percent=spec["i0_percent"], shift_mv_degree=0.0, shift_lv_degree=0.0,
            name=str(r["Название"]),
        )
        links[a].update((b, c)); links[b].update((a, c)); links[c].update((a, b))
    for r in project["loads"]:
        pp.create_load(net, bus=bus(r, "Узел"), p_mw=number(r, "P, кВт", minimum=0)/1000,
            q_mvar=number(r, "Q, квар")/1000, name=str(r["Название"]))
    for r in project["generation"]:
        p_mw = number(r, "P, кВт", minimum=0) / 1000
        q_mvar = number(r, "Q, квар") / 1000
        # Для IEC 60909 задаём номинал и коэффициент тока инвертора.
        pp.create_sgen(net, bus=bus(r, "Узел"), p_mw=p_mw, q_mvar=q_mvar,
            sn_mva=max(abs(p_mw), abs(q_mvar), 0.001), k=1.2,
            current_source=True, name=str(r["Название"]))
    reached, pending = set(), [root]
    while pending:
        i = pending.pop()
        if i not in reached:
            reached.add(i); pending.extend(links[i] - reached)
    missing = set(net.bus.index) - reached
    if missing:
        raise ValueError(f"Нет связи с внешней сетью у узлов: {sorted(missing)}.")
    return net


def solve(project):
    net = build(project)
    pp.runpp(net)
    previous = None
    for _ in range(6):
        selected = apply_auto_cables(net)
        if not any(bool(value) for value in net.line.get("auto_select", [])):
            break
        if selected == previous:
            break
        previous = selected
        pp.runpp(net)
    if not net.converged or not net.res_bus.vm_pu.map(math.isfinite).all():
        raise ValueError("Расчёт не сошёлся.")
    return net


def short_circuit(project):
    net = solve(project)
    sc.calc_sc(net, case="max", fault="3ph", ip=True)
    result = net.res_bus_sc[["ikss_ka", "ip_ka", "skss_mw"]].copy()
    result.insert(0, "Название", net.bus["name"])
    result.insert(0, "ID", net.bus.index)
    result.columns = ["ID", "Название", "Iк'', кА", "Iуд, кА", "Sкз, МВА"]
    return result


def normalize(data):
    data = migrate_project(data)
    if not isinstance(data, dict) or set(data) != set(COLUMNS):
        raise ValueError("Выберите файл проекта PowerFlow Studio. Старые проекты обновляются автоматически.")
    for key, columns in COLUMNS.items():
        if not isinstance(data[key], list) or any(not isinstance(r, dict) or set(r) != set(columns) for r in data[key]):
            raise ValueError(f"Неверный формат таблицы {key}.")
    build(data)
    return data



def require_login():
    """Показывает форму входа и разрешает доступ после проверки данных."""
    try:
        auth = st.secrets["auth"]
        expected_username = str(auth["username"])
        expected_password = str(auth["password"])
    except Exception:
        st.error("Вход ещё не настроен. Добавьте логин и пароль в Secrets приложения.")
        return False

    if st.session_state.get("authenticated", False):
        with st.sidebar:
            st.caption(f"Пользователь: {expected_username}")
            if st.button("Выйти", key="logout"):
                st.session_state.authenticated = False
                st.rerun()
        return True

    st.title("🔐 PowerFlow Studio")
    st.caption("Введите логин и пароль для доступа к приложению.")
    with st.form("login_form"):
        username = st.text_input("Логин")
        password = st.text_input("Пароль", type="password")
        submitted = st.form_submit_button("Войти", type="primary", width="stretch")

    if submitted:
        username_ok = hmac.compare_digest(username, expected_username)
        password_ok = hmac.compare_digest(password, expected_password)
        if username_ok and password_ok:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Неверный логин или пароль.")
    return False


def main():
    st.set_page_config(page_title="PowerFlow Studio", page_icon="⚡", layout="wide")
    if not require_login():
        return
    st.title("⚡ PowerFlow Studio")
    st.caption("Версия 4.2 • авто/ручной ввод R, X, C • автоподбор кабеля/провода • трансформаторы • КЗ")
    if "project_v2" not in st.session_state:
        st.session_state.project_v2 = example()
        st.session_state.revision = 0
    migrate_project(st.session_state.project_v2)
    with st.sidebar:
        st.header("Проект")
        upload = st.file_uploader("Открыть проект JSON", type=["json"])
        if st.button("Открыть файл", disabled=upload is None):
            try:
                data = normalize(json.loads(upload.getvalue().decode("utf-8-sig")))
                st.session_state.project_v2 = data
                st.session_state.revision += 1
                st.session_state.pop("result_v2", None)
                st.rerun()
            except (ValueError, TypeError, KeyError) as exc:
                st.error(str(exc))
        if st.button("Загрузить исходный пример"):
            st.session_state.project_v2 = example()
            st.session_state.revision += 1
            st.session_state.pop("result_v2", None)
            st.rerun()
        st.caption("Перед сбросом скачайте проект. Ввод сохраняется в текущей сессии; для хранения используйте JSON.")
    st.info("Добавляйте строки в таблицах. Соединения задаются по ID узлов. После ввода нажмите Enter, затем выберите вид расчёта.")
    labels = ["Узлы", "Внешняя сеть", "Линии", "Трансформаторы 2W", "Трансформаторы 3W", "Нагрузки", "Генерация P/Q"]
    edited = {}
    type_net = pp.create_empty_network()
    register_transformer_catalog(type_net)
    types_2w = sorted(type_net.std_types["trafo"])
    types_3w = sorted(THREE_WINDING_CATALOG)
    tabs = st.tabs(labels + ["Полный справочник"])
    for (key, columns), tab in zip(COLUMNS.items(), tabs[:-1]):
        with tab:
            if key == "lines":
                st.caption("Переключатель «Авто»: включён — программа подбирает кабель/провод и R, X, C, Imax по току с запасом 25%; выключен — вы вводите R, X, C и Imax вручную.")
                st.info("До 35 кВ выбирается кабель, выше 35 кВ — воздушный провод. Справочные токи зависят от прокладки, температуры и производителя; для рабочего проекта проверьте результат по каталогу изготовителя и требованиям ПУЭ.")
                with st.expander("Справочник КЛ 1–35 кВ (на 1 км)"):
                    cable_ref_voltage = st.selectbox(
                        "Номинальное напряжение КЛ, кВ", [1, 6, 10, 20, 35],
                        key="cable_reference_voltage"
                    )
                    st.dataframe(cable_reference_frame(cable_ref_voltage), hide_index=True, width="stretch")
                    st.caption("R приведено для +20 °C. Потери изоляции — справочные; для конкретной марки кабеля уточняйте их по данным изготовителя. Ёмкость C в этой таблице не приведена и остаётся редактируемой в ручном режиме.")
            if key == "generation":
                st.caption("Генерация с заданными P и Q. Положительные P и Q означают выдачу мощности в сеть.")
            if key == "grid":
                st.caption("Для КЗ задайте Sкз max внешней сети и отношение R/X. По умолчанию: 500 МВА и 0,1.")
            if key == "trafos":
                st.caption("Выберите тип — Sном, напряжения, Uк, потери и ток холостого хода подставятся автоматически.")
                with st.expander("Справочник двухобмоточных трансформаторов"):
                    selected_class = st.selectbox("Класс напряжения, кВ", [35, 110, 150, 220, 330, 500, 750, 1150])
                    st.dataframe(transformer_catalog_frame(selected_class), hide_index=True, width="stretch")
            if key == "trafos3w":
                st.caption("Трёхобмоточные трансформаторы и автотрансформаторы: задайте разные узлы ВН, СН и НН.")
                selected_class_3w = st.selectbox("Класс напряжения 3W, кВ", [110, 150, 220, 330, 500, 750, 1150])
                st.dataframe(three_winding_catalog_frame(selected_class_3w), hide_index=True, width="stretch")
            editor_frame = pd.DataFrame(st.session_state.project_v2[key], columns=columns)
            editor_columns = columns
            if key == "lines":
                # Чекбокс работает как тумблер: ✓ — автоматический подбор, пусто — ручной ввод.
                editor_frame["Авто"] = editor_frame["Режим"].fillna("Авто").astype(str).eq("Авто")
                editor_columns = ["Название", "Начало", "Конец", "L, км", "Авто", "Материал",
                                  "R, Ом/км", "X, Ом/км", "C, нФ/км", "Imax, А"]
                config = {
                    "Авто": st.column_config.CheckboxColumn(
                        "Авто", help="Вкл: автоподбор R, X, C и Imax. Выкл: ручной ввод.", default=True
                    ),
                    "Материал": st.column_config.SelectboxColumn(options=LINE_MATERIALS, required=True),
                }
            elif key == "trafos":
                config = {"Тип": st.column_config.SelectboxColumn(options=types_2w, required=True)}
            elif key == "trafos3w":
                config = {"Тип": st.column_config.SelectboxColumn(options=types_3w, required=True)}
            else:
                config = {}
            edited_frame = st.data_editor(
                editor_frame, key=f"{key}_{st.session_state.revision}", num_rows="dynamic",
                hide_index=True, width="stretch", column_config=config, column_order=editor_columns
            )
            if key == "lines":
                edited_frame["Режим"] = edited_frame["Авто"].fillna(True).map(
                    {True: "Авто", False: "Вручную"}
                )
                edited_frame = edited_frame.drop(columns=["Авто"])
            edited[key] = edited_frame[columns].dropna(how="all").to_dict("records")
    with tabs[-1]:
        st.caption("Все 104 строки из исходного Excel-файла: двухобмоточные, трёхобмоточные трансформаторы и автотрансформаторы.")
        reference = full_transformer_reference()
        section_title = st.selectbox("Раздел справочника", [item["title"] for item in reference])
        section = next(item for item in reference if item["title"] == section_title)
        st.dataframe(pd.DataFrame(section["rows"], columns=section["columns"]), hide_index=True, width="stretch")
    st.download_button("Скачать проект JSON", json.dumps(edited, ensure_ascii=False, indent=2, allow_nan=True), "powerflow_project.json", "application/json")
    steady_button, sc_button = st.columns(2)
    if steady_button.button("Рассчитать режим", type="primary", width="stretch"):
        st.session_state.pop("result_v2", None)
        try:
            with st.spinner("Расчёт…"):
                net = solve(edited)
            st.session_state.result_v2 = (json.dumps(edited, sort_keys=True, ensure_ascii=False), net)
        except (ValueError, TypeError, KeyError, pp.LoadflowNotConverged) as exc:
            st.error(str(exc))
    if sc_button.button("Рассчитать КЗ (3-ф)", width="stretch"):
        st.session_state.pop("short_circuit_v3", None)
        try:
            with st.spinner("Расчёт тока КЗ…"):
                result = short_circuit(edited)
            st.session_state.short_circuit_v3 = (json.dumps(edited, sort_keys=True, ensure_ascii=False), result)
        except (ValueError, TypeError, KeyError, pp.LoadflowNotConverged) as exc:
            st.error(str(exc))
    if "short_circuit_v3" in st.session_state:
        signature_sc, results_sc = st.session_state.short_circuit_v3
        if signature_sc == json.dumps(edited, sort_keys=True, ensure_ascii=False):
            st.subheader("Токи трёхфазного короткого замыкания")
            st.caption("Максимальный режим по IEC 60909. Iк'' — начальный симметричный ток, Iуд — ударный ток.")
            st.dataframe(results_sc, hide_index=True, width="stretch")
            st.download_button("Скачать результаты КЗ CSV", results_sc.to_csv(index=False).encode("utf-8-sig"), "short_circuit_results.csv", "text/csv", key="csv_sc")
        else:
            st.warning("Параметры изменены. Нажмите «Рассчитать КЗ (3-ф)», чтобы обновить результаты.")
    if "result_v2" not in st.session_state:
        return
    signature, net = st.session_state.result_v2
    if signature != json.dumps(edited, sort_keys=True, ensure_ascii=False):
        st.warning("Параметры изменены. Нажмите «Рассчитать», чтобы обновить результаты.")
        return
    st.success("Расчёт сошёлся")
    diagram = ['graph { rankdir=LR;']
    def node(identifier, label, shape="ellipse"):
        diagram.append(f"{json.dumps(identifier)} [label={json.dumps(label, ensure_ascii=False)}, shape={shape}];")
    def edge(first, last, label="", color="gray"):
        diagram.append(f"{json.dumps(first)} -- {json.dumps(last)} [label={json.dumps(label, ensure_ascii=False)}, color={color}];")
    for i, r in net.bus.iterrows():
        node(str(i), f'{r["name"]}\nID {i} • {r["vn_kv"]:g} кВ\nU = {net.res_bus.at[i, "vm_pu"]:.4f} о.е.')
    for key, first, last in (("line", "from_bus", "to_bus"), ("trafo", "hv_bus", "lv_bus")):
        for i, r in net[key].iterrows():
            loading = net["res_"+key].at[i, "loading_percent"]
            edge(str(r[first]), str(r[last]), label=f'{r["name"]}\n{loading:.2f} %', color="red" if loading > 100 else "gray")
    for i, r in net.trafo3w.iterrows():
        loading = net.res_trafo3w.at[i, "loading_percent"]
        label = f'{r["name"]}\n3W • {loading:.2f} %'
        color = "red" if loading > 100 else "gray"
        edge(str(r["hv_bus"]), str(r["mv_bus"]), label=label, color=color)
        edge(str(r["mv_bus"]), str(r["lv_bus"]), color=color)
    for i, r in net.load.iterrows():
        node(f"load{i}", f'{r["name"]}\n{r["p_mw"]*1000:g} кВт / {r["q_mvar"]*1000:g} квар', shape="box")
        edge(str(r["bus"]), f"load{i}")
    diagram.append("}")
    st.graphviz_chart("\n".join(diagram))
    a, b, c = st.columns(3)
    a.metric("Потребление P", f"{net.load.p_mw.sum()*1000:.2f} кВт")
    b.metric("Генерация P (заданная)", f"{net.sgen.p_mw.sum()*1000:.2f} кВт")
    losses_3w = net.res_trafo3w.pl_mw.sum()
    total_losses = net.res_line.pl_mw.sum() + net.res_trafo.pl_mw.sum() + losses_3w
    c.metric("Активные потери", f"{total_losses*1000:.3f} кВт")
    if ((net.res_line.loading_percent > 100).any() or
            (net.res_trafo.loading_percent > 100).any() or
            (net.res_trafo3w.loading_percent > 100).any()):
        st.warning("Есть оборудование с загрузкой выше 100%.")
    if "auto_insufficient" in net.line and net.line.auto_insufficient.astype(bool).any():
        st.warning("Для одной или нескольких линий ток выше предела справочника. Требуется параллельная линия или индивидуальный расчёт.")
    line_current_a = net.res_line.i_ka * 1000
    line_imax_a = net.line.max_i_ka * 1000
    line_reserve = ((line_imax_a / line_current_a) - 1) * 100
    line_reserve = line_reserve.where(line_current_a > 0)
    line_insulation_loss_kw = net.line.length_km * net.line.insulation_loss_kw_per_km
    tables = {
        "Напряжения узлов": pd.DataFrame({"ID": net.bus.index, "Название": net.bus.name, "U, кВ": net.res_bus.vm_pu * net.bus.vn_kv, "U, о.е.": net.res_bus.vm_pu, "Угол, °": net.res_bus.va_degree}),
        "Линии": pd.DataFrame({"Название": net.line.name, "Кабель/провод": net.line.selected_conductor, "Материал": net.line.material, "Ток, А": line_current_a, "Imax, А": line_imax_a, "Запас, %": line_reserve, "Потери P, кВт": net.res_line.pl_mw*1000, "Потери изоляции, кВт (справ.)": line_insulation_loss_kw, "Загрузка, %": net.res_line.loading_percent}),
        "Трансформаторы 2W": pd.DataFrame({"Название": net.trafo.name, "Потери P, кВт": net.res_trafo.pl_mw*1000, "Загрузка, %": net.res_trafo.loading_percent}),
        "Трансформаторы 3W": pd.DataFrame({"Название": net.trafo3w.name,
            "Потери P, кВт": net.res_trafo3w.pl_mw*1000,
            "Загрузка, %": net.res_trafo3w.loading_percent}),
        "Внешняя сеть": pd.DataFrame({"P, кВт": net.res_ext_grid.p_mw*1000, "Q, квар": net.res_ext_grid.q_mvar*1000})}
    for pos, (label, frame) in enumerate(tables.items()):
        st.subheader(label)
        st.dataframe(frame, hide_index=True, width="stretch")
        st.download_button("Скачать CSV", frame.to_csv(index=False).encode("utf-8-sig"), f"results_{pos}.csv", "text/csv", key=f"csv{pos}")
    st.bar_chart(tables["Напряжения узлов"].set_index("ID")[["U, о.е."]])


if __name__ == "__main__":
    main()
