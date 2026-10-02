import os
import threading
import sys
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Button, Entry, Frame, Label, Progressbar, Checkbutton

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side


def sanitize_tp_name(raw_name: str) -> str:
    text = (raw_name or "").strip()
    lower = text.lower()
    if lower.startswith("тп"):
        i = 2
        while i < len(text) and text[i] in [":", "-", " "]:
            i += 1
        return text[i:].strip()
    return text


# Цвета заголовков по колонке F (RGB без альфы)
COLOR_TP = (198, 226, 255)
COLOR_CLIENT = (220, 241, 255)
COLOR_CONSIGNEE = (240, 255, 255)

RGB_TP = "C6E2FF"
RGB_CLIENT = "DCF1FF"
RGB_CONSIGNEE = "F0FFFF"


@dataclass
class Shipment:
    amount: float
    row_idx: int
    days_due: float
    consignee_key: str


def rgb_from_fill(cell: Cell) -> Optional[str]:
    fill = cell.fill
    if fill is None or fill.start_color is None:
        return None
    color = fill.start_color.rgb or fill.start_color.indexed
    if not color:
        return None
    if isinstance(color, str):
        if len(color) == 8:  # ARGB
            return color[2:].upper()
        return color.upper()
    return None


def build_merged_map(ws: Worksheet) -> Dict[Tuple[int, int], Tuple[int, int]]:
    merged_top_left_by_cell: Dict[Tuple[int, int], Tuple[int, int]] = {}
    for rng in ws.merged_cells.ranges:
        min_row, min_col, max_row, max_col = rng.min_row, rng.min_col, rng.max_row, rng.max_col
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                merged_top_left_by_cell[(r, c)] = (min_row, min_col)
    return merged_top_left_by_cell


def get_effective_cell(ws: Worksheet, merged_map: Dict[Tuple[int, int], Tuple[int, int]], row: int, col: int) -> Cell:
    key = (row, col)
    if key in merged_map:
        top_left = merged_map[key]
        return ws.cell(row=top_left[0], column=top_left[1])
    return ws.cell(row=row, column=col)


def parse_header_type(cell_f: Cell, cell_a_text: str) -> Optional[str]:
    rgb = rgb_from_fill(cell_f)
    if rgb is not None:
        if rgb == RGB_TP:
            # Игнорируем шапку без ФИО (в A только «ТП» / «ТП:» и т.п.)
            from_text = sanitize_tp_name(cell_a_text)
            if from_text:
                return "TP"
            return None
        if rgb == RGB_CLIENT:
            return "CLIENT"
        if rgb == RGB_CONSIGNEE:
            return "CONSIGNEE"
    # Фолбэк по тексту из колонки A
    text = (cell_a_text or "").strip().lower()
    if text.startswith("тп"):
        # Только если после «ТП» есть имя
        if sanitize_tp_name(cell_a_text):
            return "TP"
        return None
    if text.startswith("клиент"):
        return "CLIENT"
    if text.startswith("груз") or "грузополуч" in text:
        return "CONSIGNEE"
    return None


def safe_float(value) -> float:
    if value is None:
        return 0.0
    try:
        if isinstance(value, str):
            value = value.replace(" ", "").replace("\xa0", "").replace(',', '.')
        return float(value)
    except Exception:
        return 0.0


def process_workbook(path: str, threshold_k_days: float, threshold_sv_days: float, pdz_percent_threshold: float = 0.0, omit_balance_columns: bool = False, progress_cb=None) -> str:
    wb = load_workbook(filename=path)
    # Берем исходный активный лист и создаем его копию для расчетов,
    # чтобы сохранить палитру/формат исходника нетронутыми
    ws_src: Worksheet = wb.active
    ws: Worksheet = wb.copy_worksheet(ws_src)
    # Копия будет финальным первым листом, сохраним исходное имя
    ws.title = ws_src.title

    # Перед расчётами раскроем все уровни/строки исходного листа (если были скрыты)
    try:
        for r in range(1, ws.max_row + 1):
            ws.row_dimensions[r].hidden = False
    except Exception:
        pass

    merged_map = build_merged_map(ws)

    # Индексы колонок (1-based): F=6, G=7, H=8, I=9, K=11, A=1
    COL_A, COL_F, COL_G, COL_H, COL_I, COL_K = 1, 6, 7, 8, 9, 11

    current_tp = None
    current_client = None
    current_consignee = None

    # Открывающиеся сальдо по клиенту (сумма G по его грузополучателям)
    client_opening: Dict[str, float] = {}
    # Список отгрузок по клиенту для FIFO (рабочая очередь)
    client_shipments: Dict[str, List[Shipment]] = {}
    # Хронология операций по клиенту (для детерминированного пересчета FIFO)
    client_ops: Dict[str, List[Tuple[str, float, int, float, str]]] = {}
    # Остаток «предшествующего» долга по клиенту (отк. сальдо), который гасится в первую очередь платежами
    client_prior_debt: Dict[str, float] = {}
    # Суммы для второго листа (агрегация по ТП)
    tp_sum_k1: Dict[str, float] = {}
    tp_sum_k2: Dict[str, float] = {}
    # Добавочные суммы из положительного начального остатка (прошлый период)
    tp_prior_k1: Dict[str, float] = {}
    tp_prior_k2: Dict[str, float] = {}
    # Индексы строк заголовков ТП для чтения значений G/J в «Итоги ТП»
    tp_header_row_for_summary: Dict[str, int] = {}
    # Новые агрегаты для «Итоги ТП»
    tp_opening: Dict[str, float] = {}
    tp_ship_total: Dict[str, float] = {}
    tp_pay_total: Dict[str, float] = {}

    # Для суммирования K по грузополучателю -> значение в его строке
    consignee_sum_k: Dict[Tuple[str, int], float] = {}
    # Карта: индекс строки грузополучателя -> его ключи
    consignee_row_keys: Dict[int, Tuple[str, str, str]] = {}

    max_row = ws.max_row

    # Первый проход: определить структуру и собрать открывающиеся сальдо
    for r in range(1, max_row + 1):
        if progress_cb and r % 200 == 0:
            progress_cb(min(20, int(r / max_row * 20)))

        cell_f = get_effective_cell(ws, merged_map, r, COL_F)
        cell_a = get_effective_cell(ws, merged_map, r, COL_A)
        header_type = parse_header_type(cell_f, str(cell_a.value) if cell_a.value is not None else "")

        if header_type == "TP":
            current_tp = sanitize_tp_name(str(cell_a.value or f"TP@{r}").strip())
            if not current_tp:
                # пустое имя ТП — игнорируем как шапку
                continue
            tp_header_row_for_summary[current_tp] = r
            tp_sum_k1.setdefault(current_tp, 0.0)
            tp_sum_k2.setdefault(current_tp, 0.0)
            tp_prior_k1.setdefault(current_tp, 0.0)
            tp_prior_k2.setdefault(current_tp, 0.0)
            current_client = None
            current_consignee = None
            continue
        if header_type == "CLIENT":
            current_client = str(cell_a.value or f"CLIENT@{r}").strip()
            client_opening.setdefault(current_client, 0.0)
            client_shipments.setdefault(current_client, [])
            client_prior_debt.setdefault(current_client, 0.0)
            client_ops.setdefault(current_client, [])
            current_consignee = None
            continue
        if header_type == "CONSIGNEE":
            current_consignee = str(cell_a.value or f"CONSIGNEE@{r}").strip()
            # Сохраняем связь для будущей записи суммы K в строку грузополучателя
            consignee_row_keys[r] = (current_tp or "", current_client or "", current_consignee)

            # Открывающееся сальдо из G
            cell_g = get_effective_cell(ws, merged_map, r, COL_G)
            opening = safe_float(cell_g.value)
            if current_client is not None:
                client_opening[current_client] = client_opening.get(current_client, 0.0) + opening
            if current_tp is not None:
                tp_opening[current_tp] = tp_opening.get(current_tp, 0.0) + opening
            continue

        # Иначе — потенциально строка операций
        # Ничего не делаем на первом проходе

    # Инициализируем приоритетный долг из открывающихся сальдо
    for client, opening in client_opening.items():
        client_prior_debt[client] = float(opening)

    # Второй проход: FIFO распределение платежей по клиентам и расчет K
    # Для расчета K нам нужно знать: для каждой отгрузки остаток после гашения платежами и prior_debt
    # Мы модифицируем client_shipments по мере чтения строк
    for r in range(1, max_row + 1):
        if progress_cb and r % 200 == 0:
            progress_cb(20 + min(50, int(r / max_row * 50)))

        cell_f = get_effective_cell(ws, merged_map, r, COL_F)
        cell_a = get_effective_cell(ws, merged_map, r, COL_A)
        header_type = parse_header_type(cell_f, str(cell_a.value) if cell_a.value is not None else "")

        if header_type in {"TP", "CLIENT", "CONSIGNEE"}:
            # Обновление текущих контекстов
            if header_type == "TP":
                current_tp = sanitize_tp_name(str(cell_a.value or f"TP@{r}").strip())
                if not current_tp:
                    continue
            elif header_type == "CLIENT":
                current_client = str(cell_a.value or f"CLIENT@{r}").strip()
            elif header_type == "CONSIGNEE":
                current_consignee = str(cell_a.value or f"CONSIGNEE@{r}").strip()
            continue

        # Строка операции: учитываем только значения в H (отгрузка) и I (оплата)
        if not current_client:
            continue

        cell_h = ws.cell(row=r, column=COL_H)
        cell_i = ws.cell(row=r, column=COL_I)
        ship = safe_float(cell_h.value)
        pay = safe_float(cell_i.value)
        days_due = safe_float(get_effective_cell(ws, merged_map, r, COL_F).value)

        # Отгрузка — перед постановкой в очередь уменьшаем её за счёт аванса (отрицательного prior_debt)
        if ship > 0.0:
            client_ops[current_client].append(("SHIP", ship, r, days_due, current_consignee or ""))
            if current_tp is not None:
                tp_ship_total[current_tp] = tp_ship_total.get(current_tp, 0.0) + ship
            if client_prior_debt[current_client] < 0:
                advance_available = -client_prior_debt[current_client]
                use_from_advance = min(advance_available, ship)
                ship -= use_from_advance
                client_prior_debt[current_client] += use_from_advance  # двигаем аванс к нулю
            if ship > 1e-9:
                client_shipments[current_client].append(
                    Shipment(amount=ship, row_idx=r, days_due=days_due, consignee_key=(current_consignee or ""))
                )

        # Платёж: сначала гасим положительный prior_debt, затем отгрузки FIFO, остаток увеличивает аванс
        if pay > 0.0:
            client_ops[current_client].append(("PAY", pay, r, days_due, current_consignee or ""))
            remain = pay
            if current_tp is not None:
                tp_pay_total[current_tp] = tp_pay_total.get(current_tp, 0.0) + pay
            # Гашение предшествующего долга
            if client_prior_debt[current_client] > 0:
                use = min(client_prior_debt[current_client], remain)
                client_prior_debt[current_client] -= use
                remain -= use

            # Гасим очередные отгрузки FIFO
            shipments = client_shipments[current_client]
            idx = 0
            while remain > 0 and idx < len(shipments):
                sh = shipments[idx]
                use = min(sh.amount, remain)
                sh.amount -= use
                remain -= use
                if sh.amount <= 1e-9:
                    # Полностью погашено — удаляем
                    shipments.pop(idx)
                else:
                    idx += 1

            # Если остался излишек платежа — это новый аванс (увеличиваем отрицательное prior_debt)
            if remain > 0:
                client_prior_debt[current_client] -= remain
                remain = 0.0

    # Третий проход: FIFO строго по каждому грузополучателю (боксу)
    # Для каждого заголовка грузополучателя:
    # - берём начальный остаток G
    # - собираем отгрузки (H>0) и оплаты (I>0) в его блоке по порядку строк
    # - сначала гасим отгрузки начальным остатком (FIFO), затем платежами (FIFO)
    # - непогашенные остатки отгрузок записываем в shipment_remaining_by_row
    shipment_remaining_by_row: Dict[int, float] = {}

    merged_map2 = build_merged_map(ws)
    max_row2 = ws.max_row
    rpos = 1
    while rpos <= max_row2:
        cell_f = get_effective_cell(ws, merged_map2, rpos, COL_F)
        cell_a = get_effective_cell(ws, merged_map2, rpos, COL_A)
        ht = parse_header_type(cell_f, str(cell_a.value) if cell_a.value is not None else "")
        if ht != "CONSIGNEE":
            rpos += 1
            continue

        # Найти границы бокса грузополучателя
        start = rpos
        rr = rpos + 1
        while rr <= max_row2:
            cell_f2 = get_effective_cell(ws, merged_map2, rr, COL_F)
            cell_a2 = get_effective_cell(ws, merged_map2, rr, COL_A)
            ht2 = parse_header_type(cell_f2, str(cell_a2.value) if cell_a2.value is not None else "")
            if ht2 in {"CONSIGNEE", "CLIENT", "TP"} and rr != start:
                break
            rr += 1
        end = rr - 1

        opening = safe_float(get_effective_cell(ws, merged_map2, start, COL_G).value)
        # События в хронологии бокса (PAY/SHIP)
        events: List[Tuple[str, float, int, float]] = []  # (kind, amount, row, days_due)
        for rline in range(start + 1, end + 1):
            cell_f3 = get_effective_cell(ws, merged_map2, rline, COL_F)
            cell_a3 = get_effective_cell(ws, merged_map2, rline, COL_A)
            ht3 = parse_header_type(cell_f3, str(cell_a3.value) if cell_a3.value is not None else "")
            if ht3 is not None:
                continue
            h = safe_float(ws.cell(row=rline, column=COL_H).value)
            i = safe_float(ws.cell(row=rline, column=COL_I).value)
            if h > 0:
                events.append(("SHIP", h, rline, safe_float(get_effective_cell(ws, merged_map2, rline, COL_F).value)))
            if i > 0:
                events.append(("PAY", i, rline, 0.0))

        # FIFO расчёт по боксу с учётом начального остатка:
        # положительный G = долг (prior_debt), отрицательный G = аванс (advance_credit)
        prior_debt = opening if opening > 0 else 0.0
        advance_credit = -opening if opening < 0 else 0.0
        shipments_box: List[Shipment] = []

        # Проходим события по порядку
        for kind, amount, row_idx, days_due in events:
            if kind == "SHIP":
                # Перед постановкой в очередь погасим часть отгрузки авансом
                if advance_credit > 0:
                    use = min(advance_credit, amount)
                    amount -= use
                    advance_credit -= use
                if amount > 1e-9:
                    shipments_box.append(Shipment(amount=amount, row_idx=row_idx, days_due=days_due, consignee_key=""))
            elif kind == "PAY":
                remain = amount
                if prior_debt > 0:
                    use = min(prior_debt, remain)
                    prior_debt -= use
                    remain -= use
                si = 0
                while remain > 0 and si < len(shipments_box):
                    sh = shipments_box[si]
                    if sh.amount <= 1e-9:
                        si += 1
                        continue
                    use = min(sh.amount, remain)
                    sh.amount -= use
                    remain -= use
                    if sh.amount <= 1e-9:
                        si += 1
                # Остаток платежа становится авансом для будущих отгрузок
                if remain > 0:
                    advance_credit += remain

        # Зафиксировать остатки и сразу заполнить K внутри бокса, а также сумму в строку заголовка бокса и агрегаты по ТП
        sum_ops_k = 0.0  # сумма операционных долгов (по отгрузкам с F > threshold_k_days)
        sum_ops_sv = 0.0  # суммы для СВ (по отгрузкам с F > threshold_sv_days)
        sum_prior = 0.0  # долг прошлого периода из колонки L
        sum_prior_sv = 0.0
        # Определим ФИО ТП для агрегации: идём вверх до ближайшего TP
        tp_name = None
        seek = start
        while seek >= 1:
            cell_f_up = get_effective_cell(ws, merged_map2, seek, COL_F)
            cell_a_up = get_effective_cell(ws, merged_map2, seek, COL_A)
            ht_up = parse_header_type(cell_f_up, str(cell_a_up.value) if cell_a_up.value is not None else "")
            if ht_up == "TP":
                tp_name = str(cell_a_up.value or f"TP@{seek}").strip()
                break
            seek -= 1

        for sh in shipments_box:
            if sh.amount > 1e-9:
                shipment_remaining_by_row[sh.row_idx] = sh.amount
                if sh.days_due >= float(threshold_k_days):
                    cell_k = ws.cell(row=sh.row_idx, column=COL_K, value=round(sh.amount, 2))
                    # мягкая жёлтая заливка для операционных K
                    cell_k.fill = PatternFill(start_color="FFFFFF99", end_color="FFFFFF99", fill_type="solid")
                    sum_ops_k += sh.amount
                if sh.days_due >= float(threshold_sv_days):
                    sum_ops_sv += sh.amount

        # Долг прошлого периода: если L есть — берём его; иначе используем рассчитанный положительный prior_debt
        header_days_due = safe_float(get_effective_cell(ws, merged_map2, start, COL_F).value)
        l_val = safe_float(get_effective_cell(ws, merged_map2, start, 12).value)
        eff_prior = l_val if l_val > 1e-9 else (prior_debt if prior_debt > 1e-9 else 0.0)
        if eff_prior > 1e-9:
            sum_prior = eff_prior
            # По требованию: колонка «СВ с прошлого периода» = сумма долга ТП с прошлого периода
            sum_prior_sv = eff_prior
            # Запишем L для наглядности, если он отсутствовал
            if l_val <= 1e-9:
                cell_L = ws.cell(row=start, column=12, value=round(eff_prior, 2))
                cell_L.fill = PatternFill(start_color="FFFF5353", end_color="FFFF5353", fill_type="solid")

        # Рыжая ячейка K у грузополучателя: сумма операционных K + L (если L != 0)
        total_box_debt = sum_ops_k + (sum_prior if abs(sum_prior) > 1e-9 else 0.0)
        if total_box_debt > 0:
            cell_k_hdr = ws.cell(row=start, column=COL_K, value=round(total_box_debt, 2))
            # RGB 255,200,100 для суммы грузополучателя
            cell_k_hdr.fill = PatternFill(start_color="FFFFC864", end_color="FFFFC864", fill_type="solid")
        if tp_name is not None:
            # В сводку по ТП учитываем только видимые боксы (не скрытые строкой фильтра)
            if not ws.row_dimensions[start].hidden:
                tp_sum_k1[tp_name] = tp_sum_k1.get(tp_name, 0.0) + total_box_debt
                tp_prior_k1[tp_name] = tp_prior_k1.get(tp_name, 0.0) + sum_prior
                # Для СВ: долг по порогу СВ; колонка «СВ с прошлого периода» = долг ТП с прошлого периода
                # Долг СВ берём только по фильтру 2 (только операционные суммы по порогу 2)
                tp_sum_k2[tp_name] = tp_sum_k2.get(tp_name, 0.0) + sum_ops_sv
                # Колонка «Сумма долга ТП с прошлого периода» — копия «Из него: с прошлого периода»
                tp_prior_k2[tp_name] = tp_prior_k1.get(tp_name, 0.0)

        # Пометить остаток долга прошлого периода в колонке L и залить цветом 255,83,83
        if prior_debt > 1e-9:
            cell_L = ws.cell(row=start, column=12, value=round(prior_debt, 2))
            cell_L.fill = PatternFill(start_color="FFFF5353", end_color="FFFF5353", fill_type="solid")

        rpos = end + 1

    # Чтобы корректно отразить частично погашенные отгрузки, нам нужно исходную сумму каждой отгрузки.
    # Повторно пройдемся, собирая исходные суммы и вычисляя (исходная - погашенная) через разницу
    original_ship_amount: Dict[int, float] = {}
    current_tp = None
    current_client = None
    current_consignee = None
    for r in range(1, max_row + 1):
        cell_f = get_effective_cell(ws, merged_map, r, COL_F)
        cell_a = get_effective_cell(ws, merged_map, r, COL_A)
        header_type = parse_header_type(cell_f, str(cell_a.value) if cell_a.value is not None else "")

        if header_type in {"TP", "CLIENT", "CONSIGNEE"}:
            if header_type == "TP":
                current_tp = sanitize_tp_name(str(cell_a.value or f"TP@{r}").strip())
                if not current_tp:
                    continue
            elif header_type == "CLIENT":
                current_client = str(cell_a.value or f"CLIENT@{r}").strip()
            elif header_type == "CONSIGNEE":
                current_consignee = str(cell_a.value or f"CONSIGNEE@{r}").strip()
            continue

        if not current_client:
            continue
        ship = safe_float(ws.cell(row=r, column=COL_H).value)
        if ship > 0:
            original_ship_amount[r] = ship

    # Теперь вычислим фактический остаток по каждой отгрузке (с учетом того, что часть могла быть погашена):
    # Если отгрузка отсутствует в shipment_remaining_by_row, это означает, что она полностью погашена.
    # Остаток = shipment_remaining_by_row.get(row, 0)
    # Затем фильтрация по дням просрочки для K1 и K2
    current_tp = None
    current_client = None
    current_consignee = None
    consignee_running_sum: Dict[int, float] = {}

    for r in range(1, max_row + 1):
        if progress_cb and r % 200 == 0:
            progress_cb(70 + min(20, int(r / max_row * 20)))

        cell_f = get_effective_cell(ws, merged_map, r, COL_F)
        cell_a = get_effective_cell(ws, merged_map, r, COL_A)
        header_type = parse_header_type(cell_f, str(cell_a.value) if cell_a.value is not None else "")

        if header_type in {"TP", "CLIENT", "CONSIGNEE"}:
            if header_type == "TP":
                current_tp = str(cell_a.value or f"TP@{r}").strip()
            elif header_type == "CLIENT":
                current_client = str(cell_a.value or f"CLIENT@{r}").strip()
            elif header_type == "CONSIGNEE":
                current_consignee = str(cell_a.value or f"CONSIGNEE@{r}").strip()
            continue

        if not current_client:
            continue

        ship = safe_float(ws.cell(row=r, column=COL_H).value)
        if ship <= 0:
            continue

        days_due = safe_float(get_effective_cell(ws, merged_map, r, COL_F).value)
        remain = shipment_remaining_by_row.get(r, 0.0)

        k1 = remain if days_due >= float(threshold_k_days) else 0.0
        k2 = remain if days_due >= float(threshold_sv_days) else 0.0

        if k1 > 0:
            cell_k2 = ws.cell(row=r, column=COL_K, value=round(k1, 2))
            cell_k2.fill = PatternFill(start_color="FFFFFF99", end_color="FFFFFF99", fill_type="solid")

        # Сумма по текущему грузополучателю: находим его заголовочную строку для накопления
        # Ищем ближайшую вышестоящую строку CONSIGNEE — уже запомнили их индексы в первом проходе
        # Упростим: найдем последнюю встреченную строку CONSIGNEE (current_consignee) — её индекс можно найти через обратный поиск
        # Для точности пройдем назад до ближайшего CONSIGNEE
        rr = r
        header_row_for_consignee = None
        while rr >= 1:
            cell_f2 = get_effective_cell(ws, merged_map, rr, COL_F)
            cell_a2 = get_effective_cell(ws, merged_map, rr, COL_A)
            ht2 = parse_header_type(cell_f2, str(cell_a2.value) if cell_a2.value is not None else "")
            if ht2 == "CONSIGNEE":
                header_row_for_consignee = rr
                break
            rr -= 1
        if header_row_for_consignee is not None and k1 > 0:
            consignee_running_sum[header_row_for_consignee] = consignee_running_sum.get(header_row_for_consignee, 0.0) + k1

        # Агрегацию по ТП считаем только по сумме в строке грузополучателя (см. расчёт боксов выше),
        # чтобы исключить двойной учёт. Здесь не добавляем в tp_sum_k1/tp_sum_k2.

    # Записать сумму K в строки грузополучателей (если ранее не записана)
    for header_row, ops_sum in consignee_running_sum.items():
        # если уже выставлена «рыжая» K ранее, не перезаписываем
        existing_k = ws.cell(row=header_row, column=COL_K).value
        if existing_k is not None and existing_k != "":
            continue
        # добавим долг из L (если есть) к сумме операций
        l_val_hdr = safe_float(get_effective_cell(ws, merged_map, header_row, 12).value)
        total_k = ops_sum + (l_val_hdr if abs(l_val_hdr) > 1e-9 else 0.0)
        if total_k > 0:
            cell_hdr2 = ws.cell(row=header_row, column=COL_K, value=round(total_k, 2))
            cell_hdr2.fill = PatternFill(start_color="FFFFC864", end_color="FFFFC864", fill_type="solid")

    # Сформировать второй лист
    ws2 = wb.create_sheet(title="Итоги ТП")
    if omit_balance_columns:
        ws2.append(["ФИО ТП", "Общий долг", "Из него: с прошлого периода", "Долг СВ", "ПДЗ %"]) 
    else:
        ws2.append(["ФИО ТП", "Начальный остаток", "Конечный остаток", "Общий долг", "Из него: с прошлого периода", "Долг СВ", "ПДЗ %"]) 
    total_opening = 0.0
    total_ending = 0.0
    total_all = 0.0
    total_prior = 0.0
    total_sv = 0.0
    all_tp_names = sorted(set(list(tp_sum_k1.keys()) + list(tp_prior_k1.keys()) + list(tp_sum_k2.keys()) + list(tp_header_row_for_summary.keys())))
    for tp_name in all_tp_names:
        # Берём напрямую из строки ТП: G (7) и J (10)
        opening_val = 0.0
        ending_val = 0.0
        tp_row_idx = tp_header_row_for_summary.get(tp_name)
        if tp_row_idx is not None:
            opening_val = round(safe_float(get_effective_cell(ws, merged_map, tp_row_idx, 7).value), 2)
            ending_val = round(safe_float(get_effective_cell(ws, merged_map, tp_row_idx, 10).value), 2)
        total_val = round(tp_sum_k1.get(tp_name, 0.0), 2)
        prior_val = round(tp_prior_k1.get(tp_name, 0.0), 2)
        sv_val = round(tp_sum_k2.get(tp_name, 0.0), 2)
        perc_val = (total_val / ending_val) if ending_val not in (0, 0.0) else 0.0
        if not omit_balance_columns:
            total_opening += opening_val
            total_ending += ending_val
        total_all += total_val
        total_prior += prior_val
        total_sv += sv_val
        if omit_balance_columns:
            ws2.append([tp_name, total_val, prior_val, sv_val, perc_val])
        else:
            ws2.append([tp_name, opening_val, ending_val, total_val, prior_val, sv_val, perc_val])
    total_perc = 0.0
    if omit_balance_columns:
        # Пересчитываем по скрытой сумме ending (не показываем, но считаем):
        # sum(ending_val) мы не копили видимо, так как не нужно в итоговой строке без колонок остатков.
        # Тогда посчитаем как (Итог по D) / (сумма J по всем ТП)
        sum_all_ending = 0.0
        for tp_name in all_tp_names:
            tp_row_idx = tp_header_row_for_summary.get(tp_name)
            if tp_row_idx is not None:
                sum_all_ending += safe_float(get_effective_cell(ws, merged_map, tp_row_idx, 10).value)
        total_perc = (total_all / sum_all_ending) if sum_all_ending not in (0, 0.0) else 0.0
        ws2.append(["ИТОГО:", round(total_all, 2), round(total_prior, 2), round(total_sv, 2), total_perc])
    else:
        total_perc = (total_all / total_ending) if total_ending not in (0, 0.0) else 0.0
        ws2.append(["ИТОГО:", round(total_opening, 2), round(total_ending, 2), round(total_all, 2), round(total_prior, 2), round(total_sv, 2), total_perc])

    # Оформление «Итоги ТП» как таблицы с зеброй, шапкой и итогом
    header_fill = PatternFill(start_color="FFDCE6F1", end_color="FFDCE6F1", fill_type="solid")
    total_fill = PatternFill(start_color="FFCFE2F3", end_color="FFCFE2F3", fill_type="solid")
    zebra_fill = PatternFill(start_color="FFF5F5F5", end_color="FFF5F5F5", fill_type="solid")
    bold_font = Font(bold=True)
    center = Alignment(vertical="center")
    right = Alignment(horizontal="right", vertical="center")
    left = Alignment(horizontal="left", vertical="center")
    thin = Side(style="thin", color="FFB7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    max_row2 = ws2.max_row
    max_col2 = ws2.max_column

    # Freeze header
    ws2.freeze_panes = "A2"

    # Header styling
    for c in range(1, max_col2 + 1):
        cell = ws2.cell(row=1, column=c)
        cell.fill = header_fill
        cell.font = bold_font
        cell.alignment = center
        cell.border = border

    # Data rows zebra and number formats for numeric columns
    for r in range(2, max_row2):
        if (r - 2) % 2 == 0:
            for c in range(1, max_col2 + 1):
                ws2.cell(row=r, column=c).fill = zebra_fill
        # Alignments and borders
        ws2.cell(row=r, column=1).alignment = left
        ws2.cell(row=r, column=1).border = border
        for c in range(2, max_col2 + 1):
            cell = ws2.cell(row=r, column=c)
            cell.alignment = right
            # Определим индекс колонки ПДЗ %
            pdz_col_idx = 7 if not omit_balance_columns else 5
            if c == pdz_col_idx:  # "ПДЗ %"
                cell.number_format = '0.00%'
            else:
                cell.number_format = '#,##0.00'
            cell.border = border
        # Подсветка ПДЗ % по порогу
        pdz_thr = (pdz_percent_threshold or 0.0) / 100.0
        pdz_cell = ws2.cell(row=r, column=pdz_col_idx)
        try:
            if pdz_cell.value is not None and float(pdz_cell.value) >= pdz_thr and r < max_row2:
                pdz_cell.fill = PatternFill(start_color="FFC86414", end_color="FFC86414", fill_type="solid")
        except Exception:
            pass

    # Total row styling
    for c in range(1, max_col2 + 1):
        cell = ws2.cell(row=max_row2, column=c)
        cell.fill = total_fill
        cell.font = bold_font
        cell.alignment = right if c > 1 else left
        cell.border = border
        if c > 1:
            pdz_col_idx = 7 if not omit_balance_columns else 5
            if c == pdz_col_idx:  # процент в итоговой строке
                cell.number_format = '0.00%'
            else:
                cell.number_format = '#,##0.00'

    # Auto column widths
    for c in range(1, max_col2 + 1):
        max_len = 0
        for r in range(1, max_row2 + 1):
            val = ws2.cell(row=r, column=c).value
            s = str(val) if val is not None else ""
            if len(s) > max_len:
                max_len = len(s)
        # add padding
        ws2.column_dimensions[chr(64 + c)].width = min(max_len + 2, 60)

    # Фильтрация: оставить полностью те «боксы» грузополучателей, где есть
    # хотя бы одна операционная строка с F > threshold_k_days. Остальные скрыть.
    merged_map = build_merged_map(ws)
    max_row = ws.max_row

    # Собираем границы боксов грузополучателей и клиентов
    # consignee_boxes: (start_row, end_row, client_header_row)
    consignee_boxes: List[Tuple[int, int, int]] = []
    # client_blocks: (start_row, end_row)
    client_blocks: List[Tuple[int, int]] = []
    # TP -> header row index (для записи итогов по ТП в K и группировки уровней)
    tp_header_row: Dict[str, int] = {}

    r = 1
    current_client_row: Optional[int] = None
    while r <= max_row:
        cell_f = get_effective_cell(ws, merged_map, r, COL_F)
        cell_a = get_effective_cell(ws, merged_map, r, COL_A)
        header_type = parse_header_type(cell_f, str(cell_a.value) if cell_a.value is not None else "")
        if header_type == "TP":
            # Закрываем предыдущий client блок, если он не закрыт явно
            current_client_row = None
            # Запомним строку заголовка ТП
            tp_name_scan = sanitize_tp_name(str(cell_a.value or f"TP@{r}").strip())
            if tp_name_scan:
                tp_header_row[tp_name_scan] = r
            r += 1
            continue
        if header_type == "CLIENT":
            # Определим конец блока клиента до следующего CLIENT или TP
            start_client = r
            rr = r + 1
            while rr <= max_row:
                cell_f2 = get_effective_cell(ws, merged_map, rr, COL_F)
                cell_a2 = get_effective_cell(ws, merged_map, rr, COL_A)
                ht2 = parse_header_type(cell_f2, str(cell_a2.value) if cell_a2.value is not None else "")
                if ht2 in {"CLIENT", "TP"} and rr != start_client:
                    break
                rr += 1
            end_client = rr - 1
            client_blocks.append((start_client, end_client))
            current_client_row = start_client
            r += 1
            continue
        if header_type == "CONSIGNEE":
            start = r
            rr = r + 1
            while rr <= max_row:
                cell_f2 = get_effective_cell(ws, merged_map, rr, COL_F)
                cell_a2 = get_effective_cell(ws, merged_map, rr, COL_A)
                ht2 = parse_header_type(cell_f2, str(cell_a2.value) if cell_a2.value is not None else "")
                if ht2 in {"CONSIGNEE", "CLIENT", "TP"} and rr != start:
                    break
                rr += 1
            end = rr - 1
            consignee_boxes.append((start, end, current_client_row or start))
            r = rr
            continue
        r += 1

    # Определяем, какие боксы оставить
    boxes_to_keep: List[Tuple[int, int]] = []
    kept_clients: set[int] = set()
    # Режим показа бокса: full — показывать весь бокс, header_only — только строку грузополучателя
    box_show_mode: Dict[Tuple[int, int], str] = {}
    for start, end, client_row in consignee_boxes:
        keep_full = False
        # Критерий 1: есть операции с F > порога => показываем весь бокс
        for rr in range(start + 1, end + 1):
            cell_f = get_effective_cell(ws, merged_map, rr, COL_F)
            cell_a = get_effective_cell(ws, merged_map, rr, COL_A)
            ht = parse_header_type(cell_f, str(cell_a.value) if cell_a.value is not None else "")
            if ht is None:
                days_due = safe_float(get_effective_cell(ws, merged_map, rr, COL_F).value)
                if days_due >= float(threshold_k_days):
                    keep_full = True
                    break
        # Критерий 2: в колонке L (у заголовка) есть долг => показываем только строку заголовка
        l_val_hdr = safe_float(get_effective_cell(ws, merged_map, start, 12).value)
        if keep_full or l_val_hdr > 1e-9:
            boxes_to_keep.append((start, end))
            kept_clients.add(client_row)
            box_show_mode[(start, end)] = 'full' if keep_full else 'header_only'

    boxes_to_delete: List[Tuple[int, int]] = []
    keep_set = set(boxes_to_keep)
    for start, end, _client_row in consignee_boxes:
        if (start, end) not in keep_set:
            boxes_to_delete.append((start, end))

    # Скрываем ненужные боксы (не удаляем), чтобы сохранить формат и палитру
    for start, end in boxes_to_delete:
        for rr in range(start, end + 1):
            ws.row_dimensions[rr].hidden = True
    # Для боксов с режимом header_only скрываем все строки, кроме заголовка
    for (start, end), mode in box_show_mode.items():
        if mode == 'header_only':
            for rr in range(start + 1, end + 1):
                ws.row_dimensions[rr].hidden = True

    # Скрыть полностью блоки клиентов, у которых нет ни одного оставленного «бокса»
    
    for start_client, end_client in client_blocks:
        if start_client not in kept_clients:
            for rr in range(start_client, end_client + 1):
                ws.row_dimensions[rr].hidden = True

    # Удаляем исходный лист и гарантируем порядок: 1) копия, 2) Итоги ТП
    wb.remove(ws_src)
    # Проставим уровни группировки (outline) для удобной навигации:
    # Уровень 0 — строки ТП (итоги в K/L)
    # Уровень 1 — заголовки клиентов
    # Уровень 2 — заголовки грузополучателей
    # Уровень 3 — операционные строки
    try:
        ws.sheet_properties.outlinePr.summaryBelow = False
    except Exception:
        pass
    merged_map_outline = build_merged_map(ws)
    for r in range(1, ws.max_row + 1):
        cell_f = get_effective_cell(ws, merged_map_outline, r, COL_F)
        cell_a = get_effective_cell(ws, merged_map_outline, r, COL_A)
        ht = parse_header_type(cell_f, str(cell_a.value) if cell_a.value is not None else "")
        if ht == "TP":
            ws.row_dimensions[r].outlineLevel = 0
        elif ht == "CLIENT":
            ws.row_dimensions[r].outlineLevel = 1
        elif ht == "CONSIGNEE":
            ws.row_dimensions[r].outlineLevel = 2
        else:
            ws.row_dimensions[r].outlineLevel = 3

    # Запишем итоги по ТП в их заголовки (K и L) и подсветим
    for tp_name, row_idx in tp_header_row.items():
        total_tp = round(tp_sum_k1.get(tp_name, 0.0), 2)
        if total_tp > 0:
            cell_tp_k = ws.cell(row=row_idx, column=COL_K, value=total_tp)
            cell_tp_k.fill = PatternFill(start_color="FFFFC864", end_color="FFFFC864", fill_type="solid")
        total_prior_tp = round(tp_prior_k1.get(tp_name, 0.0), 2)
        if total_prior_tp > 0:
            cell_tp_l = ws.cell(row=row_idx, column=12, value=total_prior_tp)
            cell_tp_l.fill = PatternFill(start_color="FFFF5353", end_color="FFFF5353", fill_type="solid")
    
    # Сохранение рядом с исходником
    base, ext = os.path.splitext(path)
    out_path = f"{base}_out{ext}"
    try:
        wb.save(out_path)
    except PermissionError:
        out_path = f"{base}_out_new{ext}"
        wb.save(out_path)

    if progress_cb:   
        progress_cb(100)
    return out_path


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("PD FIFO Processor")
        self.geometry("560x260")

        self.file_path_var = tk.StringVar()
        self.threshold_k_var = tk.StringVar(value="0")
        self.threshold_sv_var = tk.StringVar(value="0")
        self.threshold_pdz_var = tk.StringVar(value="0")
        # По имени исполняемого файла включаем режим без колонок остатков
        exe_name = os.path.basename(sys.argv[0]).lower()
        default_omit = 'pdzonly' in exe_name
        self.omit_balance_var = tk.BooleanVar(value=default_omit)

        root = Frame(self)
        root.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

        row0 = Frame(root)
        row0.pack(fill=tk.X, pady=4)
        Label(row0, text="Файл XLSX:").pack(side=tk.LEFT)
        Entry(row0, textvariable=self.file_path_var, width=50).pack(side=tk.LEFT, padx=6)
        Button(row0, text="Обзор", command=self.on_browse).pack(side=tk.LEFT)

        row1 = Frame(root)
        row1.pack(fill=tk.X, pady=4)
        Label(row1, text="Фильтр K: дни просрочки >=").pack(side=tk.LEFT)
        Entry(row1, textvariable=self.threshold_k_var, width=8).pack(side=tk.LEFT, padx=6)

        row2 = Frame(root)
        row2.pack(fill=tk.X, pady=4)
        Label(row2, text="Фильтр СВ: дни просрочки >=").pack(side=tk.LEFT)
        Entry(row2, textvariable=self.threshold_sv_var, width=8).pack(side=tk.LEFT, padx=6)

        row2b = Frame(root)
        row2b.pack(fill=tk.X, pady=4)
        Label(row2b, text="Порог ПДЗ % (>=)").pack(side=tk.LEFT)
        Entry(row2b, textvariable=self.threshold_pdz_var, width=8).pack(side=tk.LEFT, padx=6)
        Checkbutton(row2b, text="Без колонок остатков в Итоги ТП", variable=self.omit_balance_var).pack(side=tk.LEFT, padx=12)

        row3 = Frame(root)
        row3.pack(fill=tk.X, pady=10)
        Button(row3, text="Обработать", command=self.on_process).pack(side=tk.LEFT)

        row4 = Frame(root)
        row4.pack(fill=tk.X, pady=4)
        self.progress = Progressbar(row4, orient="horizontal", length=420, mode="determinate")
        self.progress.pack(side=tk.LEFT)

    def on_browse(self) -> None:
        path = filedialog.askopenfilename(
            title="Выберите XLSX файл",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.file_path_var.set(path)

    def set_progress(self, value: int) -> None:
        value = max(0, min(100, int(value)))
        self.progress["value"] = value
        self.update_idletasks()

    def on_process(self) -> None:
        path = self.file_path_var.get().strip()
        if not path:
            messagebox.showwarning("Внимание", "Выберите файл XLSX")
            return
        try:
            th_k = float(self.threshold_k_var.get().strip())
            th_sv = float(self.threshold_sv_var.get().strip())
            th_pdz = float(self.threshold_pdz_var.get().strip())
            omit_bal = bool(self.omit_balance_var.get())
        except Exception:
            messagebox.showerror("Ошибка", "Пороговые значения должны быть числом")
            return

        def run():
            try:
                out_path = process_workbook(path, th_k, th_sv, th_pdz, omit_bal, progress_cb=self.set_progress)
                messagebox.showinfo("Готово", f"Файл сохранён:\n{out_path}")
            except Exception as e:
                messagebox.showerror("Ошибка", str(e))
            finally:
                self.set_progress(0)

        threading.Thread(target=run, daemon=True).start()


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()


