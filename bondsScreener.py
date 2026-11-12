"""
Профессиональный скринер облигаций Московской Биржи (ISS MOEX API) v5.
Все пользовательские настройки вынесены в САМОЕ НАЧАЛО ФАЙЛА (блок CONFIG).
"""

import sys
import argparse
import requests
import pandas as pd
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# 🎯 НАСТРОЙКИ СКРИНЕРА (РЕДАКТИРУЙТЕ ЗДЕСЬ)
# ==============================================================================
CONFIG = {
    # 1. КУДА ВЫВОДИТЬ РЕЗУЛЬТАТ:
    #    "both"     -> И в терминал на экран, и в Excel-файл (по умолчанию)
    #    "terminal" -> Только в терминал (быстрый просмотр без создания файлов)
    #    "file"     -> Только в Excel-файл (тихий режим)
    "output_mode": "terminal",
    "output_filename": "screened_bonds.xlsx",

    # 2. ПЕРИОДИЧНОСТЬ ВЫПЛАТ КУПОНОВ:
    #    Укажите список нужных частот или строку:
    #    ["monthly", "quarterly", "semi"] -> Месячные (12/год), квартальные (4/год) и полугодовые (2/год)
    #    "monthly"   -> Только ежемесячные (12 раз в год)
    #    "quarterly" -> Только раз в квартал (4 раза в год)
    #    "semi"      -> Только раз в полгода (2 раза в год)
    #    "annual"    -> Только раз в год (1 раз в год)
    #    "all"       -> Любая периодичность выплат
    "coupon_freq": ["monthly", "quarterly"],

    # 3. ТИП ЭМИТЕНТА:
    #    "all"        -> Все типы бумаг
    #    "corporate"  -> Только корпоративные облигации компаний
    #    "ofz"        -> Только государственные ОФЗ
    #    "municipal"  -> Только муниципальные и региональные (субъекты РФ)
    "issuer_type": "corporate",

    # 4. ТИП КУПОНА:
    #    "all"     -> Все типы купонов
    #    "fixed"   -> Только постоянный / фиксированный купон (прогнозируемый доход)
    #    "floater" -> Только плавающий купон / флоатер (привязан к ставке ЦБ / RUONIA)
    "coupon_type": "fixed",

    # 5. СРОК ДО ПОГАШЕНИЯ (в днях):
    #    min_days_to_mat: минимум дней (например, 60 дней, чтобы отсечь гасящиеся завтра)
    #    max_days_to_mat: максимум дней (550 дней ~ 18 месяцев под покупку квартиры)
    "min_days_to_mat": 365,
    "max_days_to_mat": 1095,

    # 6. ЧИСТАЯ ЦЕНА (% от номинала):
    #    max_price_pct: не выше скольки % от номинала покупать (например, 99.0% - строго с дисконтом)
    #    min_price_pct: отсечение проблемных бумаг и дефолтного дистресса (например, 70.0%)
    "max_price_pct": 101.0,
    "min_price_pct": 80.0,

    # 7. АМОРТИЗАЦИЯ И ОФЕРТА:
    #    "no"  -> СТРОГО БЕЗ амортизации / без оферты (рекомендуется для сохранения номинала)
    #    "yes" -> Только с амортизацией / только с офертой
    #    "any" -> Не важно
    "has_amortization": "no",
    "has_offer": "no",

    # 8. ДОХОДНОСТЬ И ЛИКВИДНОСТЬ:
    #    min_ytm: минимальная эффективная доходность YTM, % (для флоатеров игнорируется)
    #    max_ytm: максимальная доходность YTM, %
    #    min_volume_rub: минимальный дневной объем торгов на бирже, ₽ (отсекаем неликвид)
    "min_ytm": 16.0,
    "max_ytm": 28.0,
    "min_volume_rub": 200_000,
}
# ==============================================================================


# Секции Мосбиржи (boards)
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


def normalize_freq_list(freq_arg):
    """Приводит аргумент частоты купонов к набору допустимых чисел выплат в год."""
    if freq_arg is None or freq_arg == "all":
        return None
    
    # Если передана строка через запятую или пробел
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
    min_volume_rub=CONFIG["min_volume_rub"],
    output_mode=CONFIG["output_mode"],
    output_filename=CONFIG["output_filename"]
):
    output_mode = str(output_mode).lower()
    show_terminal = output_mode in ["terminal", "both"]
    save_file = output_mode in ["file", "both"]

    freq_counts = normalize_freq_list(coupon_freq)
    if freq_counts is None:
        freq_desc = "ЛЮБЫЕ (месяц, квартал, полгода, год)"
    else:
        desc_map = {12: "ежемесячно (12)", 4: "квартал (4)", 2: "полгода (2)", 1: "раз в год (1)"}
        freq_desc = ", ".join([desc_map.get(c, f"{c} раз/год") for c in sorted(freq_counts, reverse=True)])

    if show_terminal:
        print("=" * 86)
        print("🔍 СКРИНЕР ОБЛИГАЦИЙ МОСБИРЖИ (ISS MOEX API)")
        print("Параметры фильтрации:")
        print(f" • Эмитенты:         {str(issuer_type).upper()}")
        print(f" • Срок погашения:   от {min_days_to_mat} до {max_days_to_mat} дней (~{max_days_to_mat//30} мес.)")
        print(f" • Выплаты купона:   {freq_desc}")
        print(f" • Тип купона:       {str(coupon_type).upper()}")
        print(f" • Амортизация:      {str(has_amortization).upper()} | Оферта: {str(has_offer).upper()}")
        print(f" • Чистая цена:      от {min_price_pct}% до {max_price_pct}% от номинала")
        print(f" • Доходность YTM:   от {min_ytm}% до {max_ytm}% | Объем торгов: от {min_volume_rub:,.0f} ₽")
        print(f" • Режим вывода:     {'Терминал + Excel' if output_mode == 'both' else ('Только терминал' if output_mode == 'terminal' else 'Только Excel')}")
        print("=" * 86)

    raw_df = fetch_moex_bonds_data()
    if raw_df.empty:
        if show_terminal:
            print("Не удалось получить данные с Мосбиржи. Проверьте интернет-соединение.")
        return pd.DataFrame()

    df = raw_df.copy()

    # Числовые колонки
    num_cols = [
        "YIELD", "VALTODAY", "PREVPRICE", "ACCRUEDINT", "LAST", 
        "COUPONPERCENT", "COUPONVALUE", "COUPONPERIOD", "INITIALFACEVALUE", "FACEVALUE", "LOTVALUE"
    ]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # 1. Тип эмитента
    df["ISSUER_TYPE"] = df["BOARDID"].map(BOARDS_MAPPING).fillna("Корпоративная")

    # 2. Срок до погашения (дней)
    df["MATDATE"] = pd.to_datetime(df["MATDATE"], errors="coerce")
    today = pd.Timestamp.today()
    df["DAYS_TO_MAT"] = (df["MATDATE"] - today).dt.days

    # 3. Чистая цена (% от номинала)
    df["CLEAN_PRICE_PCT"] = df["LAST"].combine_first(df["PREVPRICE"])

    # 4. Периодичность купонов
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

    # 5. Амортизация
    def check_amortization(row):
        init_fv = row.get("INITIALFACEVALUE")
        curr_fv = row.get("FACEVALUE")
        if pd.notnull(init_fv) and pd.notnull(curr_fv) and curr_fv < init_fv:
            return True
        name = str(row.get("SHORTNAME", "")).upper()
        if "АМОРТ" in name or "-АМ" in name:
            return True
        return False

    df["HAS_AMORT"] = df.apply(check_amortization, axis=1)

    # 6. Оферта
    offer_cols = [c for c in ["OFFERDATE", "BUYBACKDATE"] if c in df.columns]
    if offer_cols:
        df["HAS_OFFER"] = df[offer_cols[0]].notna() & (df[offer_cols[0]] != "") & (df[offer_cols[0]] != "0000-00-00")
    else:
        df["HAS_OFFER"] = False

    # 7. Тип купона
    def detect_coupon_type(row):
        name = str(row.get("SHORTNAME", "")).upper()
        secid = str(row.get("SECID", "")).upper()
        if "ПК" in name or "ФЛОАТ" in name or secid.startswith("SU29") or "FLOATER" in name:
            return "Плавающий"
        if pd.isna(row.get("COUPONPERCENT")) or row.get("COUPONPERCENT") == 0:
            if row.get("YIELD") == 0 or pd.isna(row.get("YIELD")):
                return "Плавающий"
        return "Постоянный / Фиксированный"

    df["COUPON_TYPE_NAME"] = df.apply(detect_coupon_type, axis=1)

    # Форматирование срока строкой
    def format_term(days):
        if pd.isna(days) or days < 0:
            return "—"
        years = int(days // 365)
        months = int((days % 365) // 30)
        return f"{years} г. {months} мес." if years > 0 else f"{months} мес."

    df["TERM_STR"] = df["DAYS_TO_MAT"].apply(format_term)

    # =========================================================
    # ФИЛЬТРАЦИЯ
    # =========================================================
    mask = pd.Series(True, index=df.index)

    # 1. Эмитент
    it = str(issuer_type).lower()
    if it == "ofz":
        mask = mask & (df["ISSUER_TYPE"] == "ОФЗ")
    elif it in ["municipal", "mun", "муниципальные", "муниципальная"]:
        mask = mask & (df["ISSUER_TYPE"] == "Муниципальная")
    elif it in ["corporate", "corp", "корпоративные", "корпоративная"]:
        mask = mask & (df["ISSUER_TYPE"] == "Корпоративная")

    # 2. Срок
    mask = mask & (df["DAYS_TO_MAT"] >= min_days_to_mat) & (df["DAYS_TO_MAT"] <= max_days_to_mat)

    # 3. Периодичность купонов (поддержка нескольких частот сразу)
    if freq_counts is not None:
        mask = mask & (df["PAYMENTS_PER_YEAR"].isin(freq_counts))

    # 4. Тип купона
    ct = str(coupon_type).lower()
    if ct in ["fixed", "постоянный", "фиксированный"]:
        mask = mask & (df["COUPON_TYPE_NAME"] == "Постоянный / Фиксированный")
    elif ct in ["floater", "плавающий", "флоатер"]:
        mask = mask & (df["COUPON_TYPE_NAME"] == "Плавающий")

    # 5. Амортизация
    ha = str(has_amortization).lower()
    if ha in ["no", "false", "нет", "без"]:
        mask = mask & (~df["HAS_AMORT"])
    elif ha in ["yes", "true", "да", "с"]:
        mask = mask & (df["HAS_AMORT"])

    # 6. Оферта
    ho = str(has_offer).lower()
    if ho in ["no", "false", "нет", "без"]:
        mask = mask & (~df["HAS_OFFER"])
    elif ho in ["yes", "true", "да", "с"]:
        mask = mask & (df["HAS_OFFER"])

    # 7. Цена, доходность и объем торгов
    mask = mask & (df["CLEAN_PRICE_PCT"] <= max_price_pct) & (df["CLEAN_PRICE_PCT"] >= min_price_pct)
    if ct not in ["floater", "плавающий"]:
        mask = mask & (df["YIELD"] >= min_ytm) & (df["YIELD"] <= max_ytm)
    mask = mask & (df["VALTODAY"] >= min_volume_rub)

    filtered = df[mask].copy()
    filtered = filtered.sort_values(by=["YIELD", "VALTODAY"], ascending=[False, False])

    # ---------------------------------------------------------
    # ВЫВОД В ТЕРМИНАЛ
    # ---------------------------------------------------------
    if show_terminal:
        print(f"\nНайдено подходящих облигаций: {len(filtered)}\n" + "-" * 86)
        if filtered.empty:
            print("По заданным критериям ничего не найдено. Попробуйте смягчить фильтры в CONFIG.")
        else:
            for _, row in filtered.iterrows():
                icon = "🏛️" if row["ISSUER_TYPE"] == "ОФЗ" else ("🏙️" if row["ISSUER_TYPE"] == "Муниципальная" else "🏢")
                amort_tag = " [Аморт: ДА]" if row["HAS_AMORT"] else ""
                offer_tag = " [Оферта: ДА]" if row["HAS_OFFER"] else ""
                
                print(f"{icon} {row['SHORTNAME']:<24} {row['ISSUER_TYPE']:<14} Цена: {row['CLEAN_PRICE_PCT']:>6.2f} %   YTM: {row['YIELD']:>5.2f} %")
                mat_d = row['MATDATE'].strftime('%d.%m.%Y') if pd.notnull(row['MATDATE']) else '—'
                print(f"   {row['SECID']:<24} Срок: {row['TERM_STR']:<12} Погашение: {mat_d:<10} Объем: {row['VALTODAY']:>10,.0f} ₽")
                print(f"   Купон: {row['COUPONPERCENT']:>5.2f}% ({row['COUPON_TYPE_NAME']}) | Выплата: {row['COUPONVALUE']:>6.2f} ₽ | Выплат: {row['FREQ_STR']}{amort_tag}{offer_tag}")
                print("-" * 86)

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
            "COUPONPERCENT": "Купон, % годовых",
            "COUPONVALUE": "Сумма купона, ₽",
            "FREQ_STR": "Периодичность",
            "COUPON_TYPE_NAME": "Тип купона",
            "HAS_AMORT": "Амортизация",
            "HAS_OFFER": "Оферта",
            "YIELD": "YTM, %",
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
            if pd.notnull(row_vals[14]) and hasattr(row_vals[14], 'strftime'):
                row_vals[14] = row_vals[14].strftime("%d.%m.%Y")
            else:
                row_vals[14] = str(row_vals[14]) if row_vals[14] else "—"
            ws.append(row_vals)
            is_even = (r_idx % 2 == 0)

            for c_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.font = reg_font
                cell.border = thin_border
                if is_even:
                    cell.fill = zebra_fill
                val = cell.value

                if c_idx in [4, 6, 12]:  # Проценты
                    if isinstance(val, (int, float)):
                        cell.number_format = '0.00'
                        cell.alignment = Alignment(horizontal="right")
                elif c_idx in [5, 7, 16]: # Рубли
                    if isinstance(val, (int, float)):
                        cell.number_format = '#,##0.00' if c_idx != 16 else '#,##0'
                        cell.alignment = Alignment(horizontal="right")
                elif c_idx == 13: # Дней
                    if isinstance(val, (int, float)):
                        cell.number_format = '#,##0'
                        cell.alignment = Alignment(horizontal="right")
                elif c_idx in [1, 2, 8, 9, 10, 11, 14, 15]:
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
    parser = argparse.ArgumentParser(description="Скринер облигаций Мосбиржи под покупку недвижимости")
    parser.add_argument("-o", "--output", choices=["terminal", "file", "both"], default=CONFIG["output_mode"], help="Режим вывода")
    parser.add_argument("-t", "--type", choices=["all", "ofz", "municipal", "corporate"], default=CONFIG["issuer_type"], help="Тип эмитента")
    parser.add_argument("--min-days", type=int, default=CONFIG["min_days_to_mat"], help="Мин. дней до погашения")
    parser.add_argument("--max-days", type=int, default=CONFIG["max_days_to_mat"], help="Макс. дней до погашения (~18 мес)")
    parser.add_argument("-p", "--max-price", type=float, default=CONFIG["max_price_pct"], help="Макс. цена в % от номинала")
    parser.add_argument("-f", "--freq", default=None, help="Периодичность купонов (monthly, quarterly, semi или через запятую)")
    parser.add_argument("-c", "--coupon-type", choices=["all", "fixed", "floater"], default=CONFIG["coupon_type"], help="Тип купона (fixed/floater/all)")
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
        has_amortization=args.amort,
        has_offer=args.offer,
        output_mode=args.output,
        output_filename=args.filename
    )


if __name__ == "__main__":
    if len(sys.argv) > 1:
        parse_cli_args()
    else:
        # Запуск со всеми параметрами из блока CONFIG в начале файла
        screen_moex_bonds()
