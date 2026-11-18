"""
Профессиональный скринер облигаций Московской Биржи (ISS MOEX API) v10.
Все пользовательские настройки вынесены в САМОЕ НАЧАЛО ФАЙЛА (блок CONFIG).

Добавлено в v10:
1. Фильтрация по КРЕДИТНОМУ РЕЙТИНГУ эмитента и выпуска (АКРА, Эксперт РА, НКР, НРА):
   - `min_rating`    -> Минимальный допустимый рейтинг (например, "A-" или "BBB+")
   - `max_rating`    -> Максимальный рейтинг (например, "AA+" или "all")
   - `allow_unrated` -> Включать ли бумаги без опубликованного рейтинга (True / False)
2. Поиск рейтинга сразу по двум уровням MOEX ISS API:
   - по коду бумаги (SECID / ISIN)
   - по эмитенту (EMITENT_ID / ИНН) через справочник рейтингов эмитентов Мосбиржи
3. Сортировка по кредитному рейтингу (`sort_by = "rating"`).
"""

import sys
import re
import math
import argparse
import requests
import pandas as pd
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# 🎯 НАСТРОЙКИ СКРИНЕРА (РЕДАКТИРУЙТЕ ЗДЕСЬ)
# ==============================================================================
CONFIG = {
    # 1. КУДА ВЫВОДИТЬ РЕЗУЛЬТАТ:
    #    "both"     -> И в терминал на экран, и в Excel-файл (по умолчанию)
    #    "terminal" -> Только в терминал (быстрый просмотр без файлов)
    #    "file"     -> Только в Excel-файл (тихий режим)
    "output_mode": "terminal",
    "output_filename": "screened_bonds.xlsx",

    # 2. СОРТИРОВКА РЕЗУЛЬТАТА:
    #    Доступные варианты `sort_by`:
    #    - "ytm"          -> По доходности к погашению (YTM, %)
    #    - "tkd"          -> По текущей купонной доходности к цене покупки (ТКД, %)
    #    - "coupon_rate"  -> По ставке купона к номиналу (%)
    #    - "coupon_rub"   -> По сумме купонной выплаты в рублях (₽)
    #    - "maturity"     -> По сроку до погашения (дней)
    #    - "price"        -> По чистой цене (% от номинала)
    #    - "volume"       -> По объему торгов / ликвидности (₽)
    #    - "nkd"          -> По накопленному купонному доходу (НКД, ₽)
    #    - "freq"         -> По частоте выплат в год (12 -> 4 -> 2 -> 1)
    #    - "rating"       -> По кредитному рейтингу (AAA -> AA+ -> ... -> B-)
    #    - "smart"        -> Умный баланс: высокая доходность + высокая ликвидность + рейтинг
    "sort_by": "ytm",

    #    Направление сортировки `sort_order`:
    #    - "desc" -> По убыванию (от большего к меньшему: макс. доходность / высший рейтинг AAA)
    #    - "asc"  -> По возрастанию (от меньшего к большему: ближайший срок / мин. цена)
    "sort_order": "desc",

    # 3. КРЕДИТНЫЙ РЕЙТИНГ (АКРА / Эксперт РА / НКР / НРА):
    #    Шкала (от высшего к низшему):
    #    "AAA", "AA+", "AA", "AA-", "A+", "A", "A-", "BBB+", "BBB", "BBB-",
    #    "BB+", "BB", "BB-", "B+", "B", "B-", "CCC", "CC", "C", "D"
    #    Укажите "all" или None, чтобы отключить границу.
    "min_rating": "BBB+",       # Например: "A-" (отсечет все BBB+ и ниже) или "BBB+"
    "max_rating": "all",        # Например: "AA+" или "all" (без верхней границы)
    "allow_unrated": False,     # False -> исключать облигации без рейтинга; True -> оставлять

    # 4. ТИП КУПОНА:
    #    "all"     -> Все типы купонов (и фиксированные, и флоатеры)
    #    "fixed"   -> Только постоянный / фиксированный купон (прогнозируемый доход)
    #    "floater" -> Только плавающий купон / флоатер (привязан к ставке ЦБ / RUONIA)
    "coupon_type": "fixed",

    # 5. ПЕРИОДИЧНОСТЬ ВЫПЛАТ КУПОНОВ:
    #    ["monthly", "quarterly", "semi"] -> Месячные (12/год), квартальные (4/год) и полугодовые (2/год)
    #    "monthly"   -> Только ежемесячные (12 раз в год)
    #    "quarterly" -> Только раз в квартал (4 раза в год)
    #    "semi"      -> Только раз в полгода (2 раза в год)
    #    "annual"    -> Только раз в год (1 раз в год)
    #    "all"       -> Любая периодичность выплат
    "coupon_freq": ["monthly", "quarterly"],

    # 6. ТИП ЭМИТЕНТА:
    #    "all"        -> Все типы бумаг
    #    "corporate"  -> Только корпоративные облигации компаний
    #    "ofz"        -> Только государственные ОФЗ
    #    "municipal"  -> Только муниципальные и региональные (субъекты РФ)
    "issuer_type": "corporate",

    # 7. СРОК ДО ПОГАШЕНИЯ (в днях):
    #    550 дней ~ 18 месяцев под покупку квартиры в Москве
    "min_days_to_mat": 365,
    "max_days_to_mat": 1800,

    # 8. ЧИСТАЯ ЦЕНА (% от номинала):
    "max_price_pct": 101.0,
    "min_price_pct": 80.0,

    # 9. АМОРТИЗАЦИЯ И ОФЕРТА:
    #    "no"  -> СТРОГО БЕЗ амортизации (включая будущую по графику) и БЕЗ оферты
    #    "yes" -> Только с амортизацией / только с офертой
    #    "any" -> Не важно
    "has_amortization": "no",
    "has_offer": "no",

    # 10. ДОХОДНОСТЬ К ПОГАШЕНИЮ (YTM, % годовых) — применяется только для fixed:
    "min_ytm": 16.0,
    "max_ytm": 28.0,

    # 11. КУПОННАЯ ДОХОДНОСТЬ (% годовых):
    #     a) Ставка купона к номиналу (COUPONPERCENT), % годовых
    "min_coupon_rate": 16.0,
    "max_coupon_rate": 35.0,

    #     b) Текущая купонная доходность к рыночной цене покупки (ТКД), % годовых
    "min_current_coupon_yield": 16.0,
    "max_current_coupon_yield": 45.0,

    # 12. ЛИКВИДНОСТЬ (дневной объем торгов, ₽):
    "min_volume_rub": 200_000,
}
# ==============================================================================


BOARDS_MAPPING = {
    "TQOB": "ОФЗ",
    "TQCB": "Корпоративная",
    "TQIR": "Муниципальная",
    "TQOD": "Муниципальная"
}

RATING_SCALE = [
    "D", "SD", "RD", "C", "CC", "CCC-", "CCC", "CCC+",
    "B-", "B", "B+",
    "BB-", "BB", "BB+",
    "BBB-", "BBB", "BBB+",
    "A-", "A", "A+",
    "AA-", "AA", "AA+",
    "AAA"
]
RATING_SCORE_MAP = {r: idx + 1 for idx, r in enumerate(RATING_SCALE)}

SORT_OPTIONS_MAP = {
    "ytm": ("YIELD", "Доходность к погашению (YTM, %)"),
    "tkd": ("CURRENT_COUPON_YIELD", "Текущая купонная доходность к цене (ТКД, %)"),
    "coupon_rate": ("COUPONPERCENT", "Ставка купона к номиналу (%)"),
    "coupon_rub": ("COUPONVALUE", "Сумма купонной выплаты (₽)"),
    "maturity": ("DAYS_TO_MAT", "Срок до погашения (дней)"),
    "price": ("CLEAN_PRICE_PCT", "Чистая цена (% от номинала)"),
    "volume": ("VALTODAY", "Объем торгов за день (₽)"),
    "nkd": ("ACCRUEDINT", "Накопленный купонный доход (НКД, ₽)"),
    "freq": ("PAYMENTS_PER_YEAR", "Периодичность выплат купона"),
    "rating": ("RATING_SCORE", "Кредитный рейтинг (AAA -> D)"),
    "smart": ("SMART_SCORE", "Умный рейтинг (Доходность + Ликвидность + Рейтинг)"),
}


def normalize_rating_code(raw_rating):
    """
    Приводит строку рейтинга от АКРА / Эксперт РА / НКР / НРА (например, 'ruA+', 'A+(RU)', 'AA-.ru')
    к единому стандарту 'AAA', 'AA+', 'A-', 'BBB+' и т.д.
    """
    if raw_rating is None or (isinstance(raw_rating, float) and pd.isna(raw_rating)):
        return None
    s = str(raw_rating).upper().strip()
    if not s or s in ["NONE", "NAN", "NULL", "—", "-"]:
        return None
    s = re.sub(r"\(RU\)|\|RU\||\.RU\b|^RU", "", s)
    s = s.replace(" ", "")
    match = re.search(r"(AAA|AA[+-]?|A[+-]?|BBB[+-]?|BB[+-]?|B[+-]?|CCC[+-]?|CC|C|SD|RD|D)", s)
    if match:
        code = match.group(1)
        if code in RATING_SCORE_MAP:
            return code
    return None


def fetch_moex_bonds_data():
    """Загружает актуальный срез всех секций облигаций с Мосбиржи."""
    all_rows = []
    headers = {"User-Agent": "Mozilla/5.0"}

    for board in BOARDS_MAPPING.keys():
        url = (
            f"https://iss.moex.com/iss/engines/stock/markets/bonds/boards/{board}/securities.json"
            f"?iss.meta=off&iss.only=securities,marketdata"
        )
        try:
            r = requests.get(url, headers=headers, timeout=15)
            if r.status_code != 200:
                continue
            data = r.json()
            sec_block = data.get("securities", {})
            md_block = data.get("marketdata", {})
            if not sec_block.get("data") or not md_block.get("data"):
                continue
            sec_df = pd.DataFrame(sec_block["data"], columns=sec_block["columns"])
            md_df = pd.DataFrame(md_block["data"], columns=md_block["columns"])
            merged = pd.merge(sec_df, md_df, on="SECID", suffixes=("", "_md"))
            all_rows.append(merged)
        except Exception:
            continue

    if not all_rows:
        return pd.DataFrame()

    return pd.concat(all_rows, ignore_index=True)


def fetch_moex_ratings_bulk():
    """
    Загружает сводные рейтинги с MOEX ISS API (как по бумагам SECID/ISIN, так и по эмитентам EMITENT_ID/INN).
    """
    headers = {"User-Agent": "Mozilla/5.0"}
    sec_ratings = {}
    emitent_ratings = {}

    endpoints = [
        ("sec", "https://iss.moex.com/iss/securities/ratings.json?iss.meta=off"),
        ("sec", "https://iss.moex.com/iss/cci/rating/securities.json?iss.meta=off"),
        ("emi", "https://iss.moex.com/iss/cci/rating/emitents.json?iss.meta=off"),
    ]

    for kind, base_url in endpoints:
        start = 0
        while start < 15000:
            try:
                sep = "&" if "?" in base_url else "?"
                url = f"{base_url}{sep}start={start}"
                r = requests.get(url, headers=headers, timeout=10)
                if r.status_code != 200:
                    break
                data = r.json()
                block = data.get("ratings") or data.get("securities") or data.get("emitents") or {}
                cols = block.get("columns", [])
                rows = block.get("data", [])
                if not rows or not cols:
                    break

                cols_lower = [c.lower() for c in cols]
                secid_idx = next((i for i, c in enumerate(cols_lower) if c in ["secid", "isin"]), -1)
                emi_idx = next((i for i, c in enumerate(cols_lower) if c in ["emitent_id", "emitentid", "inn"]), -1)
                val_idx = next((i for i, c in enumerate(cols_lower) if c in ["rating", "value", "rating_value", "rate", "rating_level"]), -1)
                agency_idx = next((i for i, c in enumerate(cols_lower) if c in ["agency", "agency_name", "company", "agency_short_name"]), -1)

                if val_idx == -1 or (secid_idx == -1 and emi_idx == -1):
                    break

                for row in rows:
                    raw_val = row[val_idx]
                    norm_code = normalize_rating_code(raw_val)
                    if not norm_code:
                        continue
                    score = RATING_SCORE_MAP[norm_code]
                    agency = str(row[agency_idx]).strip() if (agency_idx != -1 and row[agency_idx]) else ""
                    label = f"{norm_code} ({agency})" if agency else norm_code

                    if secid_idx != -1 and row[secid_idx]:
                        key = str(row[secid_idx]).strip().upper()
                        if key not in sec_ratings or score > sec_ratings[key]["score"]:
                            sec_ratings[key] = {"rating": norm_code, "label": label, "score": score}

                    if emi_idx != -1 and row[emi_idx]:
                        ekey = str(row[emi_idx]).strip().upper()
                        if ekey not in emitent_ratings or score > emitent_ratings[ekey]["score"]:
                            emitent_ratings[ekey] = {"rating": norm_code, "label": label, "score": score}

                if len(rows) < 100:
                    break
                start += len(rows)
            except Exception:
                break

    return sec_ratings, emitent_ratings


def check_bond_details(secid):
    """
    Запрашивает официальный график амортизации, оферт и карточку рейтинга по конкретной бумаге (SECID).
    """
    headers = {"User-Agent": "Mozilla/5.0"}
    has_amort = False
    has_offer = False
    bond_rating = None
    bond_rating_label = None
    emitent_id = None

    url_bond = f"https://iss.moex.com/iss/statistics/engines/stock/markets/bonds/bondization/{secid}.json?iss.meta=off"
    try:
        r = requests.get(url_bond, headers=headers, timeout=8)
        if r.status_code == 200:
            data = r.json()

            amort_block = data.get("amortizations", {})
            amort_rows = amort_block.get("data", [])
            amort_cols = amort_block.get("columns", [])
            if len(amort_rows) > 1:
                has_amort = True
            elif len(amort_rows) == 1 and amort_cols:
                row_dict = dict(zip(amort_cols, amort_rows[0]))
                prc = row_dict.get("valueprc")
                if prc is not None and float(prc) < 100.0:
                    has_amort = True

            offers_block = data.get("offers", {})
            offers_rows = offers_block.get("data", [])
            offers_cols = offers_block.get("columns", [])
            if offers_rows and offers_cols:
                today_str = datetime.now().strftime("%Y-%m-%d")
                date_idx = offers_cols.index("offerdate") if "offerdate" in offers_cols else -1
                for o_row in offers_rows:
                    if date_idx != -1 and o_row[date_idx] and str(o_row[date_idx]) >= today_str:
                        has_offer = True
                        break
    except Exception:
        pass

    url_desc = f"https://iss.moex.com/iss/securities/{secid}.json?iss.meta=off&iss.only=description,ratings"
    try:
        r2 = requests.get(url_desc, headers=headers, timeout=8)
        if r2.status_code == 200:
            d2 = r2.json()
            r_block = d2.get("ratings", {})
            r_cols = r_block.get("columns", [])
            r_rows = r_block.get("data", [])
            best_score = 0
            if r_rows and r_cols:
                r_cols_l = [c.lower() for c in r_cols]
                val_i = next((i for i, c in enumerate(r_cols_l) if c in ["rating", "value", "rate"]), -1)
                ag_i = next((i for i, c in enumerate(r_cols_l) if c in ["agency", "company", "agency_name"]), -1)
                for rr in r_rows:
                    if val_i != -1:
                        nc = normalize_rating_code(rr[val_i])
                        if nc and RATING_SCORE_MAP[nc] > best_score:
                            best_score = RATING_SCORE_MAP[nc]
                            bond_rating = nc
                            ag = str(rr[ag_i]).strip() if (ag_i != -1 and rr[ag_i]) else ""
                            bond_rating_label = f"{nc} ({ag})" if ag else nc

            desc_block = d2.get("description", {})
            d_cols = desc_block.get("columns", [])
            d_rows = desc_block.get("data", [])
            if d_rows and "name" in d_cols and "value" in d_cols:
                n_i = d_cols.index("name")
                v_i = d_cols.index("value")
                for dr in d_rows:
                    field_name = str(dr[n_i] or "").upper()
                    field_val = dr[v_i]
                    if field_name in ["EMITENT_ID", "EMITENTID", "INN"] and field_val:
                        emitent_id = str(field_val).strip().upper()
                    if not bond_rating and ("RATING" in field_name or "РЕЙТИНГ" in field_name):
                        nc = normalize_rating_code(field_val)
                        if nc:
                            bond_rating = nc
                            bond_rating_label = nc
    except Exception:
        pass

    return secid, has_amort, has_offer, bond_rating, bond_rating_label, emitent_id


def normalize_freq_list(freq_arg):
    """Приводит аргумент частоты купонов к набору допустимых чисел выплат в год."""
    if freq_arg is None or freq_arg == "all":
        return None

    if isinstance(freq_arg, str):
        items = [x.strip().lower() for x in freq_arg.replace(",", " ").replace(";", " ").split() if x.strip()]
    elif isinstance(freq_arg, (list, tuple, set)):
        items = [str(x).strip().lower() for x in freq_arg]
    else:
        items = [str(freq_arg).strip().lower()]

    valid_counts = set()
    for item in items:
        if item in ["monthly", "12", "месячные", "ежемесячно", "ежемесячные", "месяц"]:
            valid_counts.add(12)
        elif item in ["quarterly", "4", "квартальные", "квартал", "раз в квартал"]:
            valid_counts.add(4)
        elif item in ["semi", "2", "полугодовые", "полгода", "раз в полгода", "полугодие"]:
            valid_counts.add(2)
        elif item in ["annual", "1", "годовые", "раз в год", "год"]:
            valid_counts.add(1)
        elif item.isdigit():
            valid_counts.add(int(item))
        elif item == "all":
            return None

    return valid_counts if valid_counts else None


def detect_coupon_type(row):
    """Определяет тип купона (Плавающий / Постоянный) по полям MOEX ISS API."""
    name = str(row.get("SHORTNAME", "")).upper()
    secid = str(row.get("SECID", "")).upper()
    full_name = str(row.get("NAME", "")).upper()

    markers = ["ПК", "-ПК", "ФЛ", "-ФЛ", "ФЛОАТ", "FLOATER", "RUONIA", "КС+", "КС +"]
    if any(m in name or m in full_name for m in markers) or secid.startswith("SU29"):
        return "Плавающий"

    cp = row.get("COUPONPERCENT")
    yld = row.get("YIELD")
    days = row.get("DAYS_TO_MAT")

    if pd.notnull(days) and days > 14:
        if pd.notnull(cp) and cp > 0:
            if pd.isna(yld) or yld == 0:
                return "Плавающий"
        elif pd.isna(cp) or cp == 0:
            if pd.isna(yld) or yld == 0:
                return "Плавающий"

    return "Постоянный / Фиксированный"


def apply_sorting(df, sort_by, sort_order):
    """Выполняет сортировку отфильтрованного датафрейма по выбранному критерию и направлению."""
    if df.empty:
        return df

    sb = str(sort_by).lower().strip()
    so = str(sort_order).lower().strip()
    ascending = so in ["asc", "up", "по возрастанию", "возрастание"]

    def calc_smart_score(row):
        y = row["YIELD"] if pd.notnull(row["YIELD"]) and row["YIELD"] > 0 else (row["CURRENT_COUPON_YIELD"] if pd.notnull(row["CURRENT_COUPON_YIELD"]) else 0.0)
        vol = max(float(row["VALTODAY"]) if pd.notnull(row["VALTODAY"]) else 1.0, 1.0)
        r_score = float(row["RATING_SCORE"]) if pd.notnull(row["RATING_SCORE"]) else 0.0
        return y * 0.70 + math.log10(vol) * 1.15 + r_score * 0.15

    df["SMART_SCORE"] = df.apply(calc_smart_score, axis=1)

    col_name, _ = SORT_OPTIONS_MAP.get(sb, ("YIELD", "Доходность к погашению (YTM, %)"))
    secondary_col = "YIELD" if col_name == "RATING_SCORE" else ("VALTODAY" if col_name != "VALTODAY" else "CURRENT_COUPON_YIELD")

    return df.sort_values(
        by=[col_name, secondary_col],
        ascending=[ascending, False],
        na_position="last"
    )


def screen_moex_bonds(
    issuer_type=CONFIG["issuer_type"],
    min_days_to_mat=CONFIG["min_days_to_mat"],
    max_days_to_mat=CONFIG["max_days_to_mat"],
    coupon_freq=CONFIG["coupon_freq"],
    coupon_type=CONFIG["coupon_type"],
    min_rating=CONFIG["min_rating"],
    max_rating=CONFIG["max_rating"],
    allow_unrated=CONFIG["allow_unrated"],
    has_amortization=CONFIG["has_amortization"],
    has_offer=CONFIG["has_offer"],
    max_price_pct=CONFIG["max_price_pct"],
    min_price_pct=CONFIG["min_price_pct"],
    min_ytm=CONFIG["min_ytm"],
    max_ytm=CONFIG["max_ytm"],
    min_coupon_rate=CONFIG["min_coupon_rate"],
    max_coupon_rate=CONFIG["max_coupon_rate"],
    min_current_coupon_yield=CONFIG["min_current_coupon_yield"],
    max_current_coupon_yield=CONFIG["max_current_coupon_yield"],
    min_volume_rub=CONFIG["min_volume_rub"],
    sort_by=CONFIG["sort_by"],
    sort_order=CONFIG["sort_order"],
    output_mode=CONFIG["output_mode"],
    output_filename=CONFIG["output_filename"]
):
    output_mode = str(output_mode).lower()
    show_terminal = output_mode in ["terminal", "both"]
    save_file = output_mode in ["file", "both"]
    ct = str(coupon_type).lower()
    is_floater_search = ct in ["floater", "плавающий", "флоатер"]

    min_rat_norm = None if (not min_rating or str(min_rating).lower() == "all") else normalize_rating_code(min_rating)
    max_rat_norm = None if (not max_rating or str(max_rating).lower() == "all") else normalize_rating_code(max_rating)
    min_rat_score = RATING_SCORE_MAP[min_rat_norm] if min_rat_norm else None
    max_rat_score = RATING_SCORE_MAP[max_rat_norm] if max_rat_norm else None

    rating_filter_active = (min_rat_score is not None) or (max_rat_score is not None) or (not allow_unrated)

    if not rating_filter_active:
        rating_desc = "ЛЮБОЙ (фильтр отключен)"
    else:
        low_txt = f"от {min_rat_norm}" if min_rat_norm else "от D"
        high_txt = f"до {max_rat_norm}" if max_rat_norm else "до AAA"
        unrated_txt = "включая без рейтинга" if allow_unrated else "строго с рейтингом"
        rating_desc = f"{low_txt} {high_txt} ({unrated_txt})"

    effective_sort_by = str(sort_by).lower()
    if is_floater_search and effective_sort_by == "ytm":
        effective_sort_by = "tkd"

    _, sort_desc = SORT_OPTIONS_MAP.get(effective_sort_by, ("YIELD", "Доходность к погашению (YTM, %)"))
    order_desc = "по возрастанию (от меньшего к меньшему)" if str(sort_order).lower() == "asc" else "по убыванию (от большего к меньшему)"

    freq_counts = normalize_freq_list(coupon_freq)
    if freq_counts is None:
        freq_desc = "ЛЮБЫЕ (месяц, квартал, полгода, год)"
    else:
        desc_map = {12: "ежемесячно (12)", 4: "квартал (4)", 2: "полгода (2)", 1: "раз в год (1)"}
        freq_desc = ", ".join([desc_map.get(c, f"{c} раз/год") for c in sorted(freq_counts, reverse=True)])

    if show_terminal:
        print("=" * 88)
        print("🔍 СКРИНЕР ОБЛИГАЦИЙ МОСБИРЖИ (ISS MOEX API) v10")
        print("Параметры фильтрации и сортировки:")
        print(f" • Сортировка:            {sort_desc} — {order_desc}")
        print(f" • Кредитный рейтинг:     {rating_desc}")
        print(f" • Эмитенты:              {str(issuer_type).upper()}")
        print(f" • Срок погашения:        от {min_days_to_mat} до {max_days_to_mat} дней (~{max_days_to_mat//30} мес.)")
        print(f" • Выплаты купона:        {freq_desc}")
        print(f" • Тип купона:            {str(coupon_type).upper()}")
        print(f" • Амортизация:           {str(has_amortization).upper()} | Оферта: {str(has_offer).upper()}")
        print(f" • Чистая цена:           от {min_price_pct}% до {max_price_pct}% от номинала")
        if is_floater_search:
            print(" • Режим флоатера:        YTM отключен (у флоатеров YTM на бирже = 0%)")
        else:
            print(f" • YTM (к погашению):     от {min_ytm}% до {max_ytm}%")
        print(f" • Ставка купона (ном.):  от {min_coupon_rate}% до {max_coupon_rate}%")
        print(f" • Текущая купон. дох-ть: от {min_current_coupon_yield}% до {max_current_coupon_yield}%")
        print(f" • Мин. объем торгов:     {min_volume_rub:,.0f} ₽")
        print("=" * 88)

    raw_df = fetch_moex_bonds_data()
    if raw_df.empty:
        if show_terminal:
            print("Не удалось получить данные с Мосбиржи. Проверьте интернет-соединение.")
        return pd.DataFrame()

    sec_ratings, emitent_ratings = fetch_moex_ratings_bulk()

    df = raw_df.copy()

    num_cols = [
        "YIELD", "VALTODAY", "PREVPRICE", "ACCRUEDINT", "LAST",
        "COUPONPERCENT", "COUPONVALUE", "COUPONPERIOD", "INITIALFACEVALUE", "FACEVALUE", "LOTVALUE"
    ]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    df["ISSUER_TYPE"] = df["BOARDID"].map(BOARDS_MAPPING).fillna("Корпоративная")
    df["MATDATE"] = pd.to_datetime(df["MATDATE"], errors="coerce")
    today = pd.Timestamp.today()
    df["DAYS_TO_MAT"] = (df["MATDATE"] - today).dt.days
    df["CLEAN_PRICE_PCT"] = df["LAST"].combine_first(df["PREVPRICE"])
    df["CURRENT_COUPON_YIELD"] = (df["COUPONPERCENT"] / df["CLEAN_PRICE_PCT"]) * 100.0

    def resolve_initial_rating(row):
        secid = str(row.get("SECID", "")).strip().upper()
        isin = str(row.get("ISIN", "")).strip().upper() if "ISIN" in row and pd.notnull(row["ISIN"]) else ""
        if row.get("ISSUER_TYPE") == "ОФЗ":
            return "AAA", "AAA (Суверенный РФ)", RATING_SCORE_MAP["AAA"]
        for k in [secid, isin]:
            if k and k in sec_ratings:
                info = sec_ratings[k]
                return info["rating"], info["label"], info["score"]
        for col_cand in ["EMITENT_ID", "EMITENTID", "INN"]:
            if col_cand in row and pd.notnull(row[col_cand]):
                ek = str(row[col_cand]).strip().upper()
                if ek in emitent_ratings:
                    info = emitent_ratings[ek]
                    return info["rating"], info["label"], info["score"]
        for col_cand in ["CREDITRATING", "RATING", "ISSUERRATING"]:
            if col_cand in row and pd.notnull(row[col_cand]):
                nc = normalize_rating_code(row[col_cand])
                if nc:
                    return nc, nc, RATING_SCORE_MAP[nc]
        return None, "Нет рейтинга", 0

    rat_tuples = df.apply(resolve_initial_rating, axis=1)
    df["RATING_CODE"] = [t[0] for t in rat_tuples]
    df["RATING_LABEL"] = [t[1] for t in rat_tuples]
    df["RATING_SCORE"] = [t[2] for t in rat_tuples]

    def calc_freq(period):
        if pd.isna(period) or period <= 0:
            return "—", 0
        freq = round(365 / period)
        if freq >= 11:
            return "12 (ежемес.)", 12
        elif freq == 4:
            return "4 (квартал)", 4
        elif freq == 2:
            return "2 (полугод.)", 2
        elif freq == 1:
            return "1 (годовой)", 1
        return f"{freq} раз/год", freq

    freq_tuples = df["COUPONPERIOD"].apply(calc_freq)
    df["FREQ_STR"] = [t[0] for t in freq_tuples]
    df["PAYMENTS_PER_YEAR"] = [t[1] for t in freq_tuples]

    def check_initial_amort(row):
        init_fv = row.get("INITIALFACEVALUE")
        curr_fv = row.get("FACEVALUE")
        if pd.notnull(init_fv) and pd.notnull(curr_fv) and curr_fv < init_fv:
            return True
        name = str(row.get("SHORTNAME", "")).upper()
        if "АМОРТ" in name or "-АМ" in name:
            return True
        return False

    df["HAS_AMORT"] = df.apply(check_initial_amort, axis=1)

    offer_cols = [c for c in ["OFFERDATE", "BUYBACKDATE"] if c in df.columns]
    if offer_cols:
        df["HAS_OFFER"] = df[offer_cols[0]].notna() & (df[offer_cols[0]] != "") & (df[offer_cols[0]] != "0000-00-00")
    else:
        df["HAS_OFFER"] = False

    df["COUPON_TYPE_NAME"] = df.apply(detect_coupon_type, axis=1)

    def format_term(days):
        if pd.isna(days) or days < 0:
            return "—"
        years = int(days // 365)
        months = int((days % 365) // 30)
        return f"{years} г. {months} мес." if years > 0 else f"{months} мес."

    df["TERM_STR"] = df["DAYS_TO_MAT"].apply(format_term)

    mask = pd.Series(True, index=df.index)

    it = str(issuer_type).lower()
    if it == "ofz":
        mask = mask & (df["ISSUER_TYPE"] == "ОФЗ")
    elif it in ["municipal", "mun", "муниципальные", "муниципальная"]:
        mask = mask & (df["ISSUER_TYPE"] == "Муниципальная")
    elif it in ["corporate", "corp", "корпоративные", "корпоративная"]:
        mask = mask & (df["ISSUER_TYPE"] == "Корпоративная")

    mask = mask & (df["DAYS_TO_MAT"] >= min_days_to_mat) & (df["DAYS_TO_MAT"] <= max_days_to_mat)

    if freq_counts is not None:
        mask = mask & (df["PAYMENTS_PER_YEAR"].isin(freq_counts))

    if ct in ["fixed", "постоянный", "фиксированный"]:
        mask = mask & (df["COUPON_TYPE_NAME"] == "Постоянный / Фиксированный")
    elif is_floater_search:
        mask = mask & (df["COUPON_TYPE_NAME"] == "Плавающий")

    mask = mask & (df["CLEAN_PRICE_PCT"] <= max_price_pct) & (df["CLEAN_PRICE_PCT"] >= min_price_pct)
    mask = mask & (df["VALTODAY"] >= min_volume_rub)

    if not is_floater_search:
        mask = mask & (df["YIELD"] >= min_ytm) & (df["YIELD"] <= max_ytm)

    if is_floater_search:
        has_known_coupon = df["COUPONPERCENT"].notna() & (df["COUPONPERCENT"] > 0)
        if min_coupon_rate is not None:
            mask = mask & (~has_known_coupon | (df["COUPONPERCENT"] >= min_coupon_rate))
        if max_coupon_rate is not None:
            mask = mask & (~has_known_coupon | (df["COUPONPERCENT"] <= max_coupon_rate))
        if min_current_coupon_yield is not None:
            mask = mask & (~has_known_coupon | (df["CURRENT_COUPON_YIELD"] >= min_current_coupon_yield))
        if max_current_coupon_yield is not None:
            mask = mask & (~has_known_coupon | (df["CURRENT_COUPON_YIELD"] <= max_current_coupon_yield))
    else:
        if min_coupon_rate is not None:
            mask = mask & (df["COUPONPERCENT"] >= min_coupon_rate)
        if max_coupon_rate is not None:
            mask = mask & (df["COUPONPERCENT"] <= max_coupon_rate)
        if min_current_coupon_yield is not None:
            mask = mask & (df["CURRENT_COUPON_YIELD"] >= min_current_coupon_yield)
        if max_current_coupon_yield is not None:
            mask = mask & (df["CURRENT_COUPON_YIELD"] <= max_current_coupon_yield)

    candidates = df[mask].copy()

    ha = str(has_amortization).lower()
    ho = str(has_offer).lower()

    if not candidates.empty:
        if show_terminal:
            print(f"Проверяем официальные графики и кредитные рейтинги по {len(candidates)} кандидатам...")

        secids = candidates["SECID"].unique().tolist()
        details_map = {}

        with ThreadPoolExecutor(max_workers=10) as executor:
            futures = {executor.submit(check_bond_details, s): s for s in secids}
            for future in as_completed(futures):
                s_id, fut_amort, fut_offer, b_rat, b_lbl, emi_id = future.result()
                details_map[s_id] = (fut_amort, fut_offer, b_rat, b_lbl, emi_id)

        for idx, row in candidates.iterrows():
            s_id = row["SECID"]
            if s_id in details_map:
                fut_amort, fut_offer, b_rat, b_lbl, emi_id = details_map[s_id]
                if fut_amort:
                    candidates.at[idx, "HAS_AMORT"] = True
                if fut_offer:
                    candidates.at[idx, "HAS_OFFER"] = True
                if b_rat and RATING_SCORE_MAP.get(b_rat, 0) > candidates.at[idx, "RATING_SCORE"]:
                    candidates.at[idx, "RATING_CODE"] = b_rat
                    candidates.at[idx, "RATING_LABEL"] = b_lbl
                    candidates.at[idx, "RATING_SCORE"] = RATING_SCORE_MAP[b_rat]
                elif emi_id and emi_id in emitent_ratings:
                    e_info = emitent_ratings[emi_id]
                    if e_info["score"] > candidates.at[idx, "RATING_SCORE"]:
                        candidates.at[idx, "RATING_CODE"] = e_info["rating"]
                        candidates.at[idx, "RATING_LABEL"] = e_info["label"]
                        candidates.at[idx, "RATING_SCORE"] = e_info["score"]

    final_mask = pd.Series(True, index=candidates.index)

    if ha in ["no", "false", "нет", "без"]:
        final_mask = final_mask & (~candidates["HAS_AMORT"])
    elif ha in ["yes", "true", "да", "с"]:
        final_mask = final_mask & (candidates["HAS_AMORT"])

    if ho in ["no", "false", "нет", "без"]:
        final_mask = final_mask & (~candidates["HAS_OFFER"])
    elif ho in ["yes", "true", "да", "с"]:
        final_mask = final_mask & (candidates["HAS_OFFER"])

    if rating_filter_active and not candidates.empty:
        has_rat = candidates["RATING_SCORE"] > 0
        rat_cond = pd.Series(True, index=candidates.index)

        if min_rat_score is not None:
            rat_cond = rat_cond & (candidates["RATING_SCORE"] >= min_rat_score)
        if max_rat_score is not None:
            rat_cond = rat_cond & (candidates["RATING_SCORE"] <= max_rat_score)

        if allow_unrated:
            final_mask = final_mask & (~has_rat | rat_cond)
        else:
            final_mask = final_mask & has_rat & rat_cond

    filtered = candidates[final_mask].copy()
    filtered = apply_sorting(filtered, effective_sort_by, sort_order)

    # ---------------------------------------------------------
    # ВЫВОД В ТЕРМИНАЛ
    # ---------------------------------------------------------
    if show_terminal:
        print(f"\nНайдено подходящих облигаций: {len(filtered)}\n" + "-" * 88)
        if filtered.empty:
            print("По заданным критериям ничего не найдено. Попробуйте смягчить фильтры в CONFIG.")
        else:
            for _, row in filtered.iterrows():
                icon = "🏛️" if row["ISSUER_TYPE"] == "ОФЗ" else ("🏙️" if row["ISSUER_TYPE"] == "Муниципальная" else "🏢")
                amort_tag = " [Аморт: ДА]" if row["HAS_AMORT"] else ""
                offer_tag = " [Оферта: ДА]" if row["HAS_OFFER"] else ""

                ytm_str = f"{row['YIELD']:>5.2f} %" if pd.notnull(row['YIELD']) and row['YIELD'] > 0 else "Флоатер"
                cp_str = f"{row['COUPONPERCENT']:>5.2f}%" if pd.notnull(row['COUPONPERCENT']) and row['COUPONPERCENT'] > 0 else "Расч. по КС"
                tkd_str = f"{row['CURRENT_COUPON_YIELD']:>5.2f}%" if pd.notnull(row['CURRENT_COUPON_YIELD']) and row['CURRENT_COUPON_YIELD'] > 0 else "—"
                val_str = f"{row['COUPONVALUE']:>6.2f} ₽" if pd.notnull(row['COUPONVALUE']) and row['COUPONVALUE'] > 0 else "—"
                rat_str = row.get("RATING_LABEL", "Нет рейтинга")

                print(f"{icon} {row['SHORTNAME']:<24} Рейтинг: {rat_str:<20} Цена: {row['CLEAN_PRICE_PCT']:>6.2f} %   YTM: {ytm_str}")
                mat_d = row['MATDATE'].strftime('%d.%m.%Y') if pd.notnull(row['MATDATE']) else '—'
                print(f"   {row['SECID']:<24} Срок: {row['TERM_STR']:<12} Погашение: {mat_d:<10} Объем: {row['VALTODAY']:>10,.0f} ₽")
                print(f"   Ставка купона: {cp_str} | ТКД (к цене): {tkd_str} | Выплата: {val_str} | Выплат: {row['FREQ_STR']} | НКД: {row['ACCRUEDINT']:.2f} ₽{amort_tag}{offer_tag}")
                print("-" * 88)

    # ---------------------------------------------------------
    # ВЫВОД В EXCEL
    # ---------------------------------------------------------
    if save_file:
        export_cols = {
            "ISSUER_TYPE": "Тип эмитента",
            "RATING_LABEL": "Кредитный рейтинг",
            "SECID": "ISIN / Код",
            "SHORTNAME": "Наименование",
            "CLEAN_PRICE_PCT": "Цена, % от ном.",
            "ACCRUEDINT": "НКД, ₽",
            "COUPONPERCENT": "Ставка купона (ном.), %",
            "CURRENT_COUPON_YIELD": "Текущая купон. дох-ть (ТКД), %",
            "COUPONVALUE": "Сумма купона, ₽",
            "FREQ_STR": "Периодичность",
            "COUPON_TYPE_NAME": "Тип купона",
            "HAS_AMORT": "Амортизация",
            "HAS_OFFER": "Оферта",
            "YIELD": "YTM (к пог.), %",
            "DAYS_TO_MAT": "Дней до пог.",
            "TERM_STR": "Срок",
            "MATDATE": "Дата погашения",
            "VALTODAY": "Объем торгов, ₽"
        }

        final_df = filtered[list(export_cols.keys())].rename(columns=export_cols)
        final_df["Амортизация"] = final_df["Амортизация"].apply(lambda x: "Да" if x else "Нет")
        final_df["Оферта"] = final_df["Оферта"].apply(lambda x: "Да" if x else "Нет")

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Отобранные облигации"
        ws.views.sheetView[0].showGridLines = True

        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        reg_font = Font(name="Calibri", size=10)
        thin_border = Border(
            left=Side(style='thin', color="D9D9D9"), right=Side(style='thin', color="D9D9D9"),
            top=Side(style='thin', color="D9D9D9"), bottom=Side(style='thin', color="D9D9D9")
        )
        zebra_fill = PatternFill(start_color="F2F5F8", end_color="F2F5F8", fill_type="solid")

        headers = list(final_df.columns)
        ws.append(headers)

        for cell in ws[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        ws.row_dimensions[1].height = 28

        for r_idx, row in enumerate(final_df.itertuples(index=False), start=2):
            row_vals = list(row)
            if pd.notnull(row_vals[16]) and hasattr(row_vals[16], 'strftime'):
                row_vals[16] = row_vals[16].strftime("%d.%m.%Y")
            else:
                row_vals[16] = str(row_vals[16]) if row_vals[16] else "—"
            ws.append(row_vals)
            is_even = (r_idx % 2 == 0)

            for c_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.font = reg_font
                cell.border = thin_border
                if is_even:
                    cell.fill = zebra_fill
                val = cell.value

                if c_idx in [5, 7, 8, 14]:
                    if isinstance(val, (int, float)):
                        cell.number_format = '0.00'
                        cell.alignment = Alignment(horizontal="right")
                elif c_idx in [6, 9, 18]:
                    if isinstance(val, (int, float)):
                        cell.number_format = '#,##0.00' if c_idx != 18 else '#,##0'
                        cell.alignment = Alignment(horizontal="right")
                elif c_idx == 15:
                    if isinstance(val, (int, float)):
                        cell.number_format = '#,##0'
                        cell.alignment = Alignment(horizontal="right")
                elif c_idx in [1, 2, 3, 10, 11, 12, 13, 16, 17]:
                    cell.alignment = Alignment(horizontal="center")
                else:
                    cell.alignment = Alignment(horizontal="left")

        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            max_len = max(len(str(cell.value or '')) for cell in col[:150])
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

        wb.save(output_filename)
        if show_terminal:
            print(f"Таблица сохранена в файл: {output_filename}")
        else:
            print(f"Готово! Результаты поиска сохранены в: {output_filename}")

    return filtered


def parse_cli_args():
    parser = argparse.ArgumentParser(description="Скринер облигаций Мосбиржи v10 (с фильтром по кредитному рейтингу)")
    parser.add_argument("-o", "--output", choices=["terminal", "file", "both"], default=CONFIG["output_mode"], help="Режим вывода")
    parser.add_argument(
        "-s", "--sort",
        choices=list(SORT_OPTIONS_MAP.keys()),
        default=CONFIG["sort_by"],
        help="Поле сортировки: ytm, tkd, coupon_rate, coupon_rub, maturity, price, volume, nkd, freq, rating, smart"
    )
    parser.add_argument(
        "--order",
        choices=["desc", "asc"],
        default=CONFIG["sort_order"],
        help="Направление сортировки: desc (по убыванию), asc (по возрастанию)"
    )
    parser.add_argument("--min-rating", type=str, default=CONFIG["min_rating"], help="Мин. кредитный рейтинг (например, A-, BBB+, AA) или all")
    parser.add_argument("--max-rating", type=str, default=CONFIG["max_rating"], help="Макс. кредитный рейтинг (например, AA+) или all")
    parser.add_argument("--allow-unrated", action="store_true", default=CONFIG["allow_unrated"], help="Разрешить бумаги без кредитного рейтинга")
    parser.add_argument("-t", "--type", choices=["all", "ofz", "municipal", "corporate"], default=CONFIG["issuer_type"], help="Тип эмитента")
    parser.add_argument("--min-days", type=int, default=CONFIG["min_days_to_mat"], help="Мин. дней до погашения")
    parser.add_argument("--max-days", type=int, default=CONFIG["max_days_to_mat"], help="Макс. дней до погашения")
    parser.add_argument("-p", "--max-price", type=float, default=CONFIG["max_price_pct"], help="Макс. цена в % от номинала")
    parser.add_argument("-f", "--freq", default=None, help="Периодичность купонов (monthly, quarterly, semi или через запятую)")
    parser.add_argument("-c", "--coupon-type", choices=["all", "fixed", "floater"], default=CONFIG["coupon_type"], help="Тип купона (fixed/floater/all)")
    parser.add_argument("--min-coupon", type=float, default=CONFIG["min_coupon_rate"], help="Мин. ставка купона к номиналу (%)")
    parser.add_argument("--min-tkd", type=float, default=CONFIG["min_current_coupon_yield"], help="Мин. текущая купонная доходность к цене (%)")
    parser.add_argument("--amort", choices=["any", "yes", "no"], default=CONFIG["has_amortization"], help="Амортизация (any/yes/no)")
    parser.add_argument("--offer", choices=["any", "yes", "no"], default=CONFIG["has_offer"], help="Оферта (any/yes/no)")
    parser.add_argument("--filename", type=str, default=CONFIG["output_filename"], help="Имя выходного Excel файла")

    args = parser.parse_args()
    freq = args.freq if args.freq is not None else CONFIG["coupon_freq"]

    screen_moex_bonds(
        issuer_type=args.type,
        min_days_to_mat=args.min_days,
        max_days_to_mat=args.max_days,
        max_price_pct=args.max_price,
        coupon_freq=freq,
        coupon_type=args.coupon_type,
        min_rating=args.min_rating,
        max_rating=args.max_rating,
        allow_unrated=args.allow_unrated,
        min_coupon_rate=args.min_coupon,
        min_current_coupon_yield=args.min_tkd,
        has_amortization=args.amort,
        has_offer=args.offer,
        sort_by=args.sort,
        sort_order=args.order,
        output_mode=args.output,
        output_filename=args.filename
    )


if __name__ == "__main__":
    if len(sys.argv) > 1:
        parse_cli_args()
    else:
        screen_moex_bonds()
