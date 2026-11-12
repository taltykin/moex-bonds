"""
Профессиональный скринер облигаций Московской Биржи (ISS MOEX API) v8.
Все пользовательские настройки вынесены в САМОЕ НАЧАЛО ФАЙЛА (блок CONFIG).

Исправления в v8:
1. Точная проверка БУДУЩЕЙ АМОРТИЗАЦИИ по официальному графику погашений Мосбиржи
   (/iss/statistics/engines/stock/markets/bonds/bondization/{secid}.json):
   - Теперь скринер проверяет не только уже прошедшую амортизацию (FACEVALUE < INITIALFACEVALUE),
     но и запланированные в будущем выплаты номинала до даты полного погашения!
   - Такие бумаги как RU000A1099P0 гарантированно определяются как бумаги С АМОРТИЗАЦИЕЙ
     и отсекаются при выставлении "has_amortization": "no".
2. Также по графику bondization проверяются будущие даты оферт (offers), если в общей таблице
   биржи поле OFFERDATE не заполнено.
"""

import sys
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

    # 2. ТИП КУПОНА:
    #    "all"     -> Все типы купонов (и фиксированные, и флоатеры)
    #    "fixed"   -> Только постоянный / фиксированный купон (прогнозируемый доход)
    #    "floater" -> Только плавающий купон / флоатер (привязан к ставке ЦБ / RUONIA)
    "coupon_type": "fixed",

    # 3. ПЕРИОДИЧНОСТЬ ВЫПЛАТ КУПОНОВ:
    #    ["monthly", "quarterly", "semi"] -> Месячные (12/год), квартальные (4/год) и полугодовые (2/год)
    #    "monthly"   -> Только ежемесячные (12 раз в год)
    #    "quarterly" -> Только раз в квартал (4 раза в год)
    #    "semi"      -> Только раз в полгода (2 раза в год)
    #    "annual"    -> Только раз в год (1 раз в год)
    #    "all"       -> Любая периодичность выплат
    "coupon_freq": ["monthly", "quarterly"],

    # 4. ТИП ЭМИТЕНТА:
    #    "all"        -> Все типы бумаг
    #    "corporate"  -> Только корпоративные облигации компаний
    #    "ofz"        -> Только государственные ОФЗ
    #    "municipal"  -> Только муниципальные и региональные (субъекты РФ)
    "issuer_type": "corporate",

    # 5. СРОК ДО ПОГАШЕНИЯ (в днях):
    #    550 дней ~ 18 месяцев под покупку квартиры в Москве
    "min_days_to_mat": 365,
    "max_days_to_mat": 550,

    # 6. ЧИСТАЯ ЦЕНА (% от номинала):
    "max_price_pct": 100.5,
    "min_price_pct": 70.0,

    # 7. АМОРТИЗАЦИЯ И ОФЕРТА:
    #    "no"  -> СТРОГО БЕЗ амортизации (включая будущую по графику) и БЕЗ оферты
    #    "yes" -> Только с амортизацией / только с офертой
    #    "any" -> Не важно
    "has_amortization": "no",
    "has_offer": "no",

    # 8. ДОХОДНОСТЬ К ПОГАШЕНИЮ (YTM, % годовых) — применяется только для fixed:
    "min_ytm": 16.0,
    "max_ytm": 28.0,

    # 9. КУПОННАЯ ДОХОДНОСТЬ (% годовых):
    #    a) Ставка купона к номиналу (COUPONPERCENT), % годовых
    "min_coupon_rate": 14.0,
    "max_coupon_rate": 35.0,

    #    b) Текущая купонная доходность к рыночной цене покупки (ТКД), % годовых
    "min_current_coupon_yield": 14.5,
    "max_current_coupon_yield": 35.0,

    # 10. ЛИКВИДНОСТЬ (дневной объем торгов, ₽):
    "min_volume_rub": 200_000,
}
# ==============================================================================


BOARDS_MAPPING = {
    "TQOB": "ОФЗ",
    "TQCB": "Корпоративная",
    "TQIR": "Муниципальная",
    "TQOD": "Муниципальная"
}

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
            sec_df = pd.DataFrame(data["securities"]["data"], columns=data["securities"]["columns"])
            md_df = pd.DataFrame(data["marketdata"]["data"], columns=data["marketdata"]["columns"])
            merged = pd.merge(sec_df, md_df, on="SECID", suffixes=("", "_md"))
            all_rows.append(merged)
        except Exception:
            continue

    if not all_rows:
        return pd.DataFrame()

    return pd.concat(all_rows, ignore_index=True)


def check_bond_schedule_details(secid):
    """
    Запрашивает официальный график купонов, амортизации и оферт по конкретной бумаге:
    https://iss.moex.com/iss/statistics/engines/stock/markets/bonds/bondization/{secid}.json
    Возвращает кортеж (secid, has_future_amortization, has_future_offer).
    """
    url = f"https://iss.moex.com/iss/statistics/engines/stock/markets/bonds/bondization/{secid}.json?iss.meta=off"
    headers = {"User-Agent": "Mozilla/5.0"}
    has_amort = False
    has_offer = False

    try:
        r = requests.get(url, headers=headers, timeout=8)
        if r.status_code == 200:
            data = r.json()

            # 1. Проверка графика амортизации (таблица amortizations)
            amort_block = data.get("amortizations", {})
            amort_rows = amort_block.get("data", [])
            amort_cols = amort_block.get("columns", [])
            # Если в графике амортизации более 1 даты погашения номинала (или есть промежуточная выплата тела)
            if len(amort_rows) > 1:
                has_amort = True
            elif len(amort_rows) == 1 and amort_cols:
                row_dict = dict(zip(amort_cols, amort_rows[0]))
                # Если доля погашения < 100% или data_source == 'amortization'
                prc = row_dict.get("valueprc")
                if prc is not None and float(prc) < 100.0:
                    has_amort = True

            # 2. Проверка графика оферт (таблица offers)
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

    return secid, has_amort, has_offer


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


def screen_moex_bonds(
    issuer_type=CONFIG["issuer_type"],
    min_days_to_mat=CONFIG["min_days_to_mat"],
    max_days_to_mat=CONFIG["max_days_to_mat"],
    coupon_freq=CONFIG["coupon_freq"],
    coupon_type=CONFIG["coupon_type"],
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
    output_mode=CONFIG["output_mode"],
    output_filename=CONFIG["output_filename"]
):
    output_mode = str(output_mode).lower()
    show_terminal = output_mode in ["terminal", "both"]
    save_file = output_mode in ["file", "both"]
    ct = str(coupon_type).lower()
    is_floater_search = ct in ["floater", "плавающий", "флоатер"]

    freq_counts = normalize_freq_list(coupon_freq)
    if freq_counts is None:
        freq_desc = "ЛЮБЫЕ (месяц, квартал, полгода, год)"
    else:
        desc_map = {12: "ежемесячно (12)", 4: "квартал (4)", 2: "полгода (2)", 1: "раз в год (1)"}
        freq_desc = ", ".join([desc_map.get(c, f"{c} раз/год") for c in sorted(freq_counts, reverse=True)])

    if show_terminal:
        print("=" * 88)
        print("🔍 СКРИНЕР ОБЛИГАЦИЙ МОСБИРЖИ (ISS MOEX API) v8")
        print("Параметры фильтрации:")
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

    # Базовая проверка уже начавшейся амортизации и оферты
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

    # =========================================================
    # ПЕРВИЧНАЯ ФИЛЬТРАЦИЯ (до глубокой проверки графика bondization)
    # =========================================================
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

    # =========================================================
    # ГЛУБОКАЯ ПРОВЕРКА БУДУЩЕЙ АМОРТИЗАЦИИ И ОФЕРТ ПО ГРАФИКУ
    # =========================================================
    ha = str(has_amortization).lower()
    ho = str(has_offer).lower()

    if not candidates.empty and (ha != "any" or ho != "any"):
        if show_terminal:
            print(f"Проверяем официальный график амортизации и оферт по {len(candidates)} кандидатам...")
        
        secids = candidates["SECID"].unique().tolist()
        schedule_map = {}

        # Запрашиваем график выплат параллельно в 10 потоков (занимает ~0.5 - 1.5 сек)
        with ThreadPoolExecutor(max_workers=10) as executor:
            futures = {executor.submit(check_bond_schedule_details, s): s for s in secids}
            for future in as_completed(futures):
                s_id, fut_amort, fut_offer = future.result()
                schedule_map[s_id] = (fut_amort, fut_offer)

        for idx, row in candidates.iterrows():
            s_id = row["SECID"]
            if s_id in schedule_map:
                fut_amort, fut_offer = schedule_map[s_id]
                if fut_amort:
                    candidates.at[idx, "HAS_AMORT"] = True
                if fut_offer:
                    candidates.at[idx, "HAS_OFFER"] = True

    # Применяем финальный фильтр амортизации и оферт
    final_mask = pd.Series(True, index=candidates.index)
    if ha in ["no", "false", "нет", "без"]:
        final_mask = final_mask & (~candidates["HAS_AMORT"])
    elif ha in ["yes", "true", "да", "с"]:
        final_mask = final_mask & (candidates["HAS_AMORT"])

    if ho in ["no", "false", "нет", "без"]:
        final_mask = final_mask & (~candidates["HAS_OFFER"])
    elif ho in ["yes", "true", "да", "с"]:
        final_mask = final_mask & (candidates["HAS_OFFER"])

    filtered = candidates[final_mask].copy()
    filtered = filtered.sort_values(by=["CURRENT_COUPON_YIELD", "VALTODAY"], ascending=[False, False], na_position="last")

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

                print(f"{icon} {row['SHORTNAME']:<24} {row['ISSUER_TYPE']:<14} Цена: {row['CLEAN_PRICE_PCT']:>6.2f} %   YTM: {ytm_str}")
                mat_d = row['MATDATE'].strftime('%d.%m.%Y') if pd.notnull(row['MATDATE']) else '—'
                print(f"   {row['SECID']:<24} Срок: {row['TERM_STR']:<12} Погашение: {mat_d:<10} Объем: {row['VALTODAY']:>10,.0f} ₽")
                print(f"   Ставка купона: {cp_str} | ТКД (к цене): {tkd_str} | Выплата: {val_str} | Выплат: {row['FREQ_STR']}{amort_tag}{offer_tag}")
                print("-" * 88)

    # ---------------------------------------------------------
    # ВЫВОД В EXCEL
    # ---------------------------------------------------------
    if save_file:
        export_cols = {
            "ISSUER_TYPE": "Тип эмитента",
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
            if pd.notnull(row_vals[15]) and hasattr(row_vals[15], 'strftime'):
                row_vals[15] = row_vals[15].strftime("%d.%m.%Y")
            else:
                row_vals[15] = str(row_vals[15]) if row_vals[15] else "—"
            ws.append(row_vals)
            is_even = (r_idx % 2 == 0)

            for c_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.font = reg_font
                cell.border = thin_border
                if is_even:
                    cell.fill = zebra_fill
                val = cell.value

                if c_idx in [4, 6, 7, 13]:
                    if isinstance(val, (int, float)):
                        cell.number_format = '0.00'
                        cell.alignment = Alignment(horizontal="right")
                elif c_idx in [5, 8, 17]:
                    if isinstance(val, (int, float)):
                        cell.number_format = '#,##0.00' if c_idx != 17 else '#,##0'
                        cell.alignment = Alignment(horizontal="right")
                elif c_idx == 14:
                    if isinstance(val, (int, float)):
                        cell.number_format = '#,##0'
                        cell.alignment = Alignment(horizontal="right")
                elif c_idx in [1, 2, 9, 10, 11, 12, 15, 16]:
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
    parser = argparse.ArgumentParser(description="Скринер облигаций Мосбиржи под покупку недвижимости v8")
    parser.add_argument("-o", "--output", choices=["terminal", "file", "both"], default=CONFIG["output_mode"], help="Режим вывода")
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
        min_coupon_rate=args.min_coupon,
        min_current_coupon_yield=args.min_tkd,
        has_amortization=args.amort,
        has_offer=args.offer,
        output_mode=args.output,
        output_filename=args.filename
    )


if __name__ == "__main__":
    if len(sys.argv) > 1:
        parse_cli_args()
    else:
        screen_moex_bonds()
