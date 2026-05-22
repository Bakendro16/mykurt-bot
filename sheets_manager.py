import json
import logging
import sqlite3
from datetime import datetime, timedelta
from collections import defaultdict
import gspread
from google.oauth2.service_account import Credentials


def _sm_get_conn(db_file, database_url=""):
    if database_url:
        import psycopg2
        url = database_url
        if url.startswith("postgres://"):
            url = url.replace("postgres://", "postgresql://", 1)
        return psycopg2.connect(url)
    return sqlite3.connect(db_file)

logger = logging.getLogger(__name__)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

KURT_ORDER = ["ordinary", "butter", "rhombus", "tablet", "spiderweb", "smoked", "mix"]
KURT_NAMES = {
    "ordinary":  "Обычный",
    "butter":    "С маслом",
    "rhombus":   "Ромбик",
    "tablet":    "Таблетка",
    "spiderweb": "Паутинка",
    "smoked":    "Копчёный",
    "mix":       "Микс",
}
COST_PER_KG = {
    "ordinary": 3100, "butter": 3100, "rhombus": 1600,
    "spiderweb": 1600, "tablet": 1600, "smoked": 1700, "mix": 2200,
}
WEIGHT_KG         = 0.040
PACK_COST         = 60
LABEL_COST        = 20
DEFECT_RATE       = 0.03
DELIVERY_PER_10KG = 500
PRICE_PER_UNIT    = 700

SHEET_ORDERS    = "Заказы"
SHEET_DASHBOARD = "Дашборд"
SHEET_COSTS     = "Расходы"
SHEET_ANALYTICS = "Аналитика"
SHEET_CLIENTS   = "Клиенты"

# ── Цвета ─────────────────────────────────────
C_DARK_BLUE  = {"red": 0.10, "green": 0.24, "blue": 0.41}
C_MID_BLUE   = {"red": 0.18, "green": 0.46, "blue": 0.71}
C_LIGHT_BLUE = {"red": 0.87, "green": 0.92, "blue": 0.97}
C_GREEN      = {"red": 0.20, "green": 0.66, "blue": 0.33}
C_DARK_GREEN = {"red": 0.09, "green": 0.44, "blue": 0.21}
C_LIGHT_GRN  = {"red": 0.88, "green": 0.96, "blue": 0.87}
C_ORANGE     = {"red": 0.95, "green": 0.61, "blue": 0.07}
C_DARK_ORG   = {"red": 0.80, "green": 0.44, "blue": 0.00}
C_LIGHT_ORG  = {"red": 1.00, "green": 0.95, "blue": 0.80}
C_RED        = {"red": 0.86, "green": 0.20, "blue": 0.18}
C_LIGHT_RED  = {"red": 1.00, "green": 0.86, "blue": 0.87}
C_PURPLE     = {"red": 0.42, "green": 0.19, "blue": 0.71}
C_LIGHT_PRP  = {"red": 0.93, "green": 0.88, "blue": 0.99}
C_TEAL       = {"red": 0.00, "green": 0.59, "blue": 0.53}
C_DARK_TEAL  = {"red": 0.00, "green": 0.40, "blue": 0.36}
C_LIGHT_TL   = {"red": 0.88, "green": 0.97, "blue": 0.96}
C_YELLOW     = {"red": 1.00, "green": 0.92, "blue": 0.23}
C_GREY_BG    = {"red": 0.95, "green": 0.95, "blue": 0.95}
C_GREY_DARK  = {"red": 0.85, "green": 0.85, "blue": 0.85}
C_WHITE      = {"red": 1.00, "green": 1.00, "blue": 1.00}
C_FG_WHITE   = {"red": 1.00, "green": 1.00, "blue": 1.00}
C_FG_DARK    = {"red": 0.13, "green": 0.13, "blue": 0.13}
C_FG_RED     = {"red": 0.72, "green": 0.11, "blue": 0.11}
C_FG_GREEN   = {"red": 0.09, "green": 0.44, "blue": 0.21}
C_FG_ORANGE  = {"red": 0.60, "green": 0.30, "blue": 0.00}


def _fmt(bold=False, size=10, fg=None, italic=False):
    f = {"bold": bold, "fontSize": size, "italic": italic}
    if fg:
        f["foregroundColor"] = fg
    return f


def _cell_fmt(bg=None, bold=False, size=10, fg=None, halign="CENTER", italic=False, wrap=False):
    f = {
        "textFormat": _fmt(bold=bold, size=size, fg=fg, italic=italic),
        "horizontalAlignment": halign,
        "verticalAlignment": "MIDDLE",
    }
    if bg:
        f["backgroundColor"] = bg
    if wrap:
        f["wrapStrategy"] = "WRAP"
    return f


class SheetsManager:
    def __init__(self, credentials_source, spreadsheet_id: str):
        if isinstance(credentials_source, dict):
            creds = Credentials.from_service_account_info(credentials_source, scopes=SCOPES)
        else:
            creds = Credentials.from_service_account_file(credentials_source, scopes=SCOPES)
        self.gc = gspread.authorize(creds)
        self.spreadsheet_id = spreadsheet_id
        self._sh = None

    @property
    def sh(self):
        if self._sh is None:
            self._sh = self.gc.open_by_key(self.spreadsheet_id)
        return self._sh

    def _get_or_create_sheet(self, title: str, rows=1000, cols=25):
        try:
            return self.sh.worksheet(title)
        except gspread.WorksheetNotFound:
            ws = self.sh.add_worksheet(title=title, rows=rows, cols=cols)
            logger.info(f"Создан лист: {title}")
            return ws

    def _col_widths(self, sheet_id, widths):
        reqs = []
        for i, w in enumerate(widths):
            reqs.append({"updateDimensionProperties": {
                "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                          "startIndex": i, "endIndex": i+1},
                "properties": {"pixelSize": w}, "fields": "pixelSize"
            }})
        return reqs

    def _row_height(self, sheet_id, start, end, px):
        return {"updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "ROWS",
                      "startIndex": start, "endIndex": end},
            "properties": {"pixelSize": px}, "fields": "pixelSize"
        }}

    def _merge(self, sheet_id, r1, c1, r2, c2):
        return {"mergeCells": {
            "range": {"sheetId": sheet_id,
                      "startRowIndex": r1, "endRowIndex": r2,
                      "startColumnIndex": c1, "endColumnIndex": c2},
            "mergeType": "MERGE_ALL"
        }}

    def _border_req(self, sheet_id, r1, c1, r2, c2, style="SOLID", width=1, color=None):
        if color is None:
            color = {"red": 0.8, "green": 0.8, "blue": 0.8}
        side = {"style": style, "width": width, "color": color}
        return {"updateBorders": {
            "range": {"sheetId": sheet_id,
                      "startRowIndex": r1, "endRowIndex": r2,
                      "startColumnIndex": c1, "endColumnIndex": c2},
            "top": side, "bottom": side, "left": side, "right": side,
            "innerHorizontal": side, "innerVertical": side
        }}

    def _freeze(self, sheet_id, rows=0, cols=0):
        return {"updateSheetProperties": {
            "properties": {
                "sheetId": sheet_id,
                "gridProperties": {"frozenRowCount": rows, "frozenColumnCount": cols}
            },
            "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"
        }}

    # ══════════════════════════════════════════
    # ИНИЦИАЛИЗАЦИЯ
    # ══════════════════════════════════════════
    def init_sheet(self):
        logger.info("Инициализация Google Sheets...")

        ws_orders = self._get_or_create_sheet(SHEET_ORDERS, rows=2000, cols=20)
        try:
            val = ws_orders.acell("A1").value
        except Exception:
            val = None
        if val != "ЖУРНАЛ ЗАКАЗОВ — КУРТ":
            logger.info("Настраиваю лист Заказы...")
            self._setup_orders_sheet(ws_orders)

        ws_dash = self._get_or_create_sheet(SHEET_DASHBOARD, rows=80, cols=20)
        try:
            val2 = ws_dash.acell("A1").value
        except Exception:
            val2 = None
        if val2 != "ДАШБОРД":
            logger.info("Настраиваю лист Дашборд...")
            self._setup_dashboard_shell(ws_dash)

        ws_costs = self._get_or_create_sheet(SHEET_COSTS, rows=30, cols=10)
        try:
            val3 = ws_costs.acell("A1").value
        except Exception:
            val3 = None
        if val3 != "СЕБЕСТОИМОСТЬ":
            logger.info("Настраиваю лист Расходы...")
            self._setup_costs_sheet(ws_costs)

        ws_analytics = self._get_or_create_sheet(SHEET_ANALYTICS, rows=100, cols=20)
        try:
            val4 = ws_analytics.acell("A1").value
        except Exception:
            val4 = None
        if val4 != "АНАЛИТИКА":
            logger.info("Настраиваю лист Аналитика...")
            self._setup_analytics_shell(ws_analytics)

        ws_clients = self._get_or_create_sheet(SHEET_CLIENTS, rows=500, cols=12)
        try:
            val5 = ws_clients.acell("A1").value
        except Exception:
            val5 = None
        if val5 != "БАЗА КЛИЕНТОВ":
            logger.info("Настраиваю лист Клиенты...")
            self._setup_clients_sheet(ws_clients)

        logger.info("Google Sheets инициализированы ✅")

    # ══════════════════════════════════════════
    # ЛИСТ: ЗАКАЗЫ
    # ══════════════════════════════════════════
    ORDERS_HEADERS = [
        "№", "Дата", "Клиент", "Username", "Телефон", "Адрес",
        "Обычный", "С маслом", "Ромбик", "Таблетка", "Паутинка", "Копчёный", "Микс",
        "Итого шт", "Бонус", "Выручка", "Себест.", "Прибыль", "Статус"
    ]

    def _setup_orders_sheet(self, ws):
        ws.clear()
        ws.update([["ЖУРНАЛ ЗАКАЗОВ — КУРТ"], self.ORDERS_HEADERS], "A1")

        sid = ws.id
        reqs = []
        reqs.append(self._merge(sid, 0, 0, 1, 19))
        reqs += self._col_widths(sid, [55, 145, 185, 130, 125, 215, 75, 80, 70, 80, 75, 80, 70, 85, 65, 110, 110, 105, 110])
        reqs.append(self._row_height(sid, 0, 1, 36))
        reqs.append(self._row_height(sid, 1, 2, 32))
        reqs.append(self._border_req(sid, 1, 0, 2, 19))
        reqs.append(self._freeze(sid, rows=2))
        self.sh.batch_update({"requests": reqs})

        ws.format("A1:S1", _cell_fmt(bg=C_DARK_BLUE, bold=True, size=14, fg=C_FG_WHITE))
        ws.format("A2:S2", _cell_fmt(bg=C_MID_BLUE, bold=True, size=10, fg=C_FG_WHITE))

    # ══════════════════════════════════════════
    # ЛИСТ: ДАШБОРД — структура
    # ══════════════════════════════════════════
    def _setup_dashboard_shell(self, ws):
        ws.clear()
        sid = ws.id
        reqs = []
        reqs += self._col_widths(sid, [18, 220, 150, 25, 220, 150, 25, 220, 150])
        reqs.append(self._row_height(sid, 0, 1, 40))
        reqs.append(self._merge(sid, 0, 0, 1, 9))
        self.sh.batch_update({"requests": reqs})
        ws.update([["ДАШБОРД"]], "A1")
        ws.format("A1:I1", _cell_fmt(bg=C_DARK_BLUE, bold=True, size=16, fg=C_FG_WHITE))

    # ══════════════════════════════════════════
    # ЛИСТ: АНАЛИТИКА — структура
    # ══════════════════════════════════════════
    def _setup_analytics_shell(self, ws):
        ws.clear()
        sid = ws.id
        reqs = []
        reqs += self._col_widths(sid, [18, 120, 120, 120, 120, 120, 120, 120, 120, 120])
        reqs.append(self._row_height(sid, 0, 1, 40))
        reqs.append(self._merge(sid, 0, 0, 1, 10))
        self.sh.batch_update({"requests": reqs})
        ws.update([["АНАЛИТИКА"]], "A1")
        ws.format("A1:J1", _cell_fmt(bg=C_DARK_BLUE, bold=True, size=16, fg=C_FG_WHITE))

    # ══════════════════════════════════════════
    # ЛИСТ: КЛИЕНТЫ — структура
    # ══════════════════════════════════════════
    def _setup_clients_sheet(self, ws):
        ws.clear()
        sid = ws.id
        reqs = []
        reqs += self._col_widths(sid, [55, 185, 145, 125, 100, 100, 130, 130, 130, 130, 110])
        reqs.append(self._row_height(sid, 0, 1, 36))
        reqs.append(self._row_height(sid, 1, 2, 32))
        reqs.append(self._merge(sid, 0, 0, 1, 11))
        reqs.append(self._freeze(sid, rows=2))
        self.sh.batch_update({"requests": reqs})

        headers = ["№", "Клиент", "Username", "Телефон", "Заказов", "Куртов",
                   "Выручка (тг)", "Расходы (тг)", "Прибыль (тг)", "Средний чек", "Последний заказ"]
        ws.update([["БАЗА КЛИЕНТОВ"], headers], "A1")
        ws.format("A1:K1", _cell_fmt(bg=C_DARK_BLUE, bold=True, size=14, fg=C_FG_WHITE))
        ws.format("A2:K2", _cell_fmt(bg=C_MID_BLUE, bold=True, size=10, fg=C_FG_WHITE))

    # ══════════════════════════════════════════
    # ЛИСТ: РАСХОДЫ
    # ══════════════════════════════════════════
    def _setup_costs_sheet(self, ws):
        ws.clear()
        sid = ws.id
        reqs = []
        reqs += self._col_widths(sid, [180, 130, 110, 110, 110, 130])
        reqs.append(self._row_height(sid, 0, 1, 36))
        reqs.append(self._merge(sid, 0, 0, 1, 6))
        self.sh.batch_update({"requests": reqs})

        headers = ["Вид курта", "Цена за кг (тг)", "Вес 1 шт (кг)", "Упаковка (тг)", "Наклейка (тг)", "Себест./шт (тг)"]
        data = [["СЕБЕСТОИМОСТЬ"], headers]
        for key in KURT_ORDER:
            cost = round(COST_PER_KG[key] * WEIGHT_KG + PACK_COST + LABEL_COST, 1)
            data.append([KURT_NAMES[key], COST_PER_KG[key], WEIGHT_KG, PACK_COST, LABEL_COST, cost])
        ws.update(data, "A1")

        ws.format("A1:F1", _cell_fmt(bg=C_DARK_BLUE, bold=True, size=13, fg=C_FG_WHITE))
        ws.format("A2:F2", _cell_fmt(bg=C_MID_BLUE, bold=True, fg=C_FG_WHITE))
        for i in range(len(KURT_ORDER)):
            r = i + 3
            bg = C_LIGHT_BLUE if i % 2 == 0 else C_WHITE
            ws.format(f"A{r}:F{r}", _cell_fmt(bg=bg))

        extra_row = len(KURT_ORDER) + 4
        extra = [
            [],
            ["ДОПОЛНИТЕЛЬНЫЕ РАСХОДЫ"],
            ["Доставка (за 10 кг)", f"{DELIVERY_PER_10KG} тг"],
            ["Брак", f"{int(DEFECT_RATE * 100)}% от себестоимости"],
            ["Цена продажи", f"{PRICE_PER_UNIT} тг/шт"],
            [],
            ["МАРЖА ПО ВИДАМ КУРТА"],
            ["Вид курта", "Себест./шт", "Цена", "Прибыль/шт", "Маржа %"],
        ]
        ws.update(extra, f"A{extra_row}")
        ws.format(f"A{extra_row+1}:F{extra_row+1}", _cell_fmt(bg=C_MID_BLUE, bold=True, fg=C_FG_WHITE))
        ws.format(f"A{extra_row+6}:F{extra_row+6}", _cell_fmt(bg=C_TEAL, bold=True, fg=C_FG_WHITE))
        ws.format(f"A{extra_row+7}:E{extra_row+7}", _cell_fmt(bg=C_MID_BLUE, bold=True, fg=C_FG_WHITE))

        for i, key in enumerate(KURT_ORDER):
            cost_per = round(COST_PER_KG[key] * WEIGHT_KG + PACK_COST + LABEL_COST, 1)
            cost_with_defect = round(cost_per * (1 + DEFECT_RATE), 1)
            profit_per = PRICE_PER_UNIT - cost_with_defect
            margin_pct = round(profit_per / PRICE_PER_UNIT * 100, 1)
            r = extra_row + 8 + i
            ws.update([[KURT_NAMES[key], cost_with_defect, PRICE_PER_UNIT, profit_per, f"{margin_pct}%"]], f"A{r}")
            bg = C_LIGHT_GRN if margin_pct >= 30 else C_LIGHT_ORG if margin_pct >= 20 else C_LIGHT_RED
            ws.format(f"A{r}:E{r}", _cell_fmt(bg=bg))

    # ══════════════════════════════════════════
    # ДОБАВИТЬ ЗАКАЗ
    # ══════════════════════════════════════════
    def _calc_cost(self, items, total_units):
        total_cost = 0.0
        for key, qty in items.items():
            if qty > 0:
                uc = COST_PER_KG.get(key, 2200) * WEIGHT_KG + PACK_COST + LABEL_COST
                uc *= (1 + DEFECT_RATE)
                total_cost += uc * qty
        total_cost += (total_units * WEIGHT_KG / 10) * DELIVERY_PER_10KG
        return total_cost

    def add_order(self, order_id, user_id, username, full_name,
                  phone, address, items, total_units, bonus_units, total_price):
        ws = self._get_or_create_sheet(SHEET_ORDERS)

        total_cost = self._calc_cost(items, total_units)
        profit = total_price - total_cost

        row = [order_id, datetime.now().strftime("%d.%m.%Y %H:%M"),
               full_name, f"@{username}" if username else "—",
               str(phone), address]
        for key in KURT_ORDER:
            row.append(items.get(key, 0))
        row += [total_units, bonus_units, total_price, round(total_cost), round(profit), "новый"]

        ws.append_row(row, value_input_option="USER_ENTERED")
        logger.info(f"Заказ #{order_id} записан в Sheets")

        try:
            all_ids = ws.col_values(1)
            for i, v in enumerate(all_ids):
                if str(v) == str(order_id):
                    ri = i + 1
                    sid = ws.id
                    reqs = [self._border_req(sid, ri - 1, 0, ri, 19)]
                    self.sh.batch_update({"requests": reqs})
                    ws.format(f"A{ri}:S{ri}", _cell_fmt(bg=C_WHITE, halign="CENTER"))
                    break
        except Exception as e:
            logger.warning(f"Форматирование строки: {e}")

    # ══════════════════════════════════════════
    # ОБНОВИТЬ СТАТУС
    # ══════════════════════════════════════════
    def update_order_status(self, order_id, status):
        ws = self._get_or_create_sheet(SHEET_ORDERS)
        all_ids = ws.col_values(1)
        row_idx = None
        for i, v in enumerate(all_ids):
            if str(v) == str(order_id):
                row_idx = i + 1
                break
        if row_idx is None:
            logger.warning(f"Заказ #{order_id} не найден в Sheets")
            return

        ws.update_cell(row_idx, 19, status)

        s = status.lower()
        if "отмен" in s:
            color = C_LIGHT_RED
        elif "выполнен" in s:
            color = C_LIGHT_GRN
        elif "пути" in s:
            color = C_LIGHT_ORG
        elif "принят" in s:
            color = C_LIGHT_BLUE
        else:
            color = C_WHITE

        ws.format(f"A{row_idx}:S{row_idx}", _cell_fmt(bg=color, halign="CENTER"))
        logger.info(f"Статус заказа #{order_id} → {status}")

    # ══════════════════════════════════════════
    # СБОР СТАТИСТИКИ ИЗ БД
    # ══════════════════════════════════════════
    def _collect_stats(self, db_file, database_url=""):
        conn = _sm_get_conn(db_file, database_url)
        c = conn.cursor()
        c.execute("SELECT * FROM orders ORDER BY id")
        rows = c.fetchall()
        c.execute("SELECT * FROM users")
        users = c.fetchall()
        conn.close()

        now = datetime.now()
        today = now.date()
        week_ago  = today - timedelta(days=7)
        month_ago = today - timedelta(days=30)

        total_orders = total_units = total_revenue = total_cost = total_profit = 0
        today_orders = today_rev = 0
        week_orders  = week_rev  = 0
        month_orders = month_rev = 0
        cancelled = pending = done = in_progress = 0

        kurt_sales   = defaultdict(int)
        kurt_revenue = defaultdict(int)
        kurt_profit  = defaultdict(float)
        daily_rev    = defaultdict(int)
        daily_orders = defaultdict(int)
        hour_counts  = defaultdict(int)
        weekday_counts = defaultdict(int)

        # Для клиентской аналитики
        client_data = {}  # user_id -> {orders, units, revenue, cost, last_date}

        for row in rows:
            oid, uid, uname, fname, phone, addr, items_j, tu, bu, tp, status, created_at = row
            items = json.loads(items_j)
            s = status.lower()

            try:
                dt = datetime.strptime(created_at, "%Y-%m-%d %H:%M:%S")
                d  = dt.date()
                h  = dt.hour
                wd = dt.weekday()
            except Exception:
                dt = now; d = today; h = 0; wd = 0

            hour_counts[h] += 1
            weekday_counts[wd] += 1

            if "отмен" in s:
                cancelled += 1
                continue

            cost = self._calc_cost(items, tu)
            profit = tp - cost

            total_orders  += 1
            total_units   += tu
            total_revenue += tp
            total_cost    += cost
            total_profit  += profit

            if d == today:
                today_orders += 1; today_rev += tp
            if d >= week_ago:
                week_orders  += 1; week_rev  += tp
            if d >= month_ago:
                month_orders += 1; month_rev += tp

            if "выполнен" in s:
                done += 1
            elif "пути" in s or "принят" in s:
                in_progress += 1
            else:
                pending += 1

            for key, qty in items.items():
                if qty > 0:
                    item_cost = COST_PER_KG.get(key, 2200) * WEIGHT_KG + PACK_COST + LABEL_COST
                    item_cost *= (1 + DEFECT_RATE)
                    kurt_sales[key]   += qty
                    kurt_revenue[key] += qty * PRICE_PER_UNIT
                    kurt_profit[key]  += (PRICE_PER_UNIT - item_cost) * qty

            if d >= today - timedelta(days=14):
                daily_rev[str(d)]    += tp
                daily_orders[str(d)] += 1

            # Клиентские данные
            if uid not in client_data:
                client_data[uid] = {
                    "username": uname, "name": fname, "phone": phone,
                    "orders": 0, "units": 0, "revenue": 0, "cost": 0.0, "last_date": d
                }
            client_data[uid]["orders"]  += 1
            client_data[uid]["units"]   += tu
            client_data[uid]["revenue"] += tp
            client_data[uid]["cost"]    += cost
            if d > client_data[uid]["last_date"]:
                client_data[uid]["last_date"] = d

        avg_check = round(total_revenue / total_orders) if total_orders else 0
        margin    = round(total_profit / total_revenue * 100, 1) if total_revenue else 0

        # Топ клиенты (по выручке)
        top_clients = sorted(client_data.values(), key=lambda u: u["revenue"], reverse=True)[:10]

        # Популярность видов
        kurt_popularity = sorted(
            [(KURT_NAMES[k], kurt_sales[k], kurt_revenue[k], round(kurt_profit[k])) for k in KURT_ORDER],
            key=lambda x: x[1], reverse=True
        )

        # Пиковые часы (все)
        peak_hours = sorted(hour_counts.items(), key=lambda x: x[1], reverse=True)[:5]

        # По дням недели
        day_names = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]
        weekday_stats = [(day_names[wd], cnt) for wd, cnt in sorted(weekday_counts.items())]

        # Все заказы за 30 дней по датам
        last_30 = {}
        for i in range(30):
            d_str = str(today - timedelta(days=29 - i))
            last_30[d_str] = {"rev": daily_rev.get(d_str, 0), "orders": daily_orders.get(d_str, 0)}

        # Конверсия и удержание
        repeat_clients = sum(1 for cl in client_data.values() if cl["orders"] > 1)
        retention_rate = round(repeat_clients / len(client_data) * 100, 1) if client_data else 0

        # Лучший день
        best_day = max(daily_rev.items(), key=lambda x: x[1]) if daily_rev else (str(today), 0)

        return {
            "total_orders": total_orders,
            "total_units": total_units,
            "total_revenue": total_revenue,
            "total_cost": round(total_cost),
            "total_profit": round(total_profit),
            "avg_check": avg_check,
            "margin": margin,
            "today_orders": today_orders, "today_rev": today_rev,
            "week_orders":  week_orders,  "week_rev":  week_rev,
            "month_orders": month_orders, "month_rev": month_rev,
            "cancelled": cancelled, "pending": pending,
            "done": done, "in_progress": in_progress,
            "total_clients": len(users),
            "repeat_clients": repeat_clients,
            "retention_rate": retention_rate,
            "kurt_popularity": kurt_popularity,
            "top_clients": top_clients,
            "client_data": client_data,
            "peak_hours": peak_hours,
            "weekday_stats": weekday_stats,
            "daily_rev": daily_rev,
            "daily_orders": daily_orders,
            "last_30": last_30,
            "best_day": best_day,
            "updated_at": now.strftime("%d.%m.%Y %H:%M"),
        }

    # ══════════════════════════════════════════
    # ОБНОВИТЬ ВСЕ ЛИСТЫ (вызывается после каждого заказа)
    # ══════════════════════════════════════════
    def refresh_dashboard(self, db_file="kurt_orders.db", database_url=""):
        try:
            stats = self._collect_stats(db_file, database_url)
            ws_dash = self._get_or_create_sheet(SHEET_DASHBOARD)
            self._write_dashboard(ws_dash, stats)

            ws_analytics = self._get_or_create_sheet(SHEET_ANALYTICS)
            self._write_analytics(ws_analytics, stats)

            ws_clients = self._get_or_create_sheet(SHEET_CLIENTS)
            self._write_clients(ws_clients, stats)

            logger.info("Дашборд и аналитика обновлены ✅")
        except Exception as e:
            logger.error(f"Ошибка обновления дашборда: {e}", exc_info=True)

    # ══════════════════════════════════════════
    # ЗАПИСЬ ДАШБОРДА
    # ══════════════════════════════════════════
    def _write_dashboard(self, ws, s):
        ws.clear()
        sid = ws.id
        reqs = []
        reqs += self._col_widths(sid, [18, 220, 150, 25, 220, 150, 25, 220, 150])
        reqs.append(self._row_height(sid, 0, 1, 44))
        reqs.append(self._merge(sid, 0, 0, 1, 9))
        self.sh.batch_update({"requests": reqs})

        all_data = []

        def blank():
            all_data.append([""] * 9)

        def kv3(l1, v1, l2="", v2="", l3="", v3=""):
            row = [""] * 9
            row[1] = l1; row[2] = v1
            row[4] = l2; row[5] = v2
            row[7] = l3; row[8] = v3
            all_data.append(row)

        def section(title):
            all_data.append(["", title, "", "", "", "", "", "", ""])

        # ── Заголовок ──
        all_data.append(["ДАШБОРД"] + [""] * 8)

        # ─── БЛОК 1: Ключевые метрики ───────────
        blank()
        section("📊 КЛЮЧЕВЫЕ МЕТРИКИ")
        kv3("Всего заказов",    s["total_orders"],
            "Всего куртов",     s["total_units"],
            "Клиентов",         s["total_clients"])
        kv3("Выручка (тг)",     f"{s['total_revenue']:,}",
            "Расходы (тг)",     f"{s['total_cost']:,}",
            "Прибыль (тг)",     f"{s['total_profit']:,}")
        kv3("Средний чек (тг)", s["avg_check"],
            "Маржа",            f"{s['margin']}%",
            "Повторных клиентов",f"{s['retention_rate']}%")
        kv3("Выполнено",        s["done"],
            "В работе",         s["in_progress"],
            "Отменено",         s["cancelled"])

        # ─── БЛОК 2: Периоды ───────────────────
        blank()
        section("📅 ПО ПЕРИОДАМ")
        all_data.append(["", "Период", "Заказов", "", "Выручка (тг)", "", "", "Прибыль (тг)", ""])
        today_profit = round(s["today_rev"] * (s["margin"] / 100)) if s["margin"] else 0
        week_profit  = round(s["week_rev"] * (s["margin"] / 100)) if s["margin"] else 0
        month_profit = round(s["month_rev"] * (s["margin"] / 100)) if s["margin"] else 0
        kv3("Сегодня",  s["today_orders"],  f"  {s['today_rev']:,} тг",  "", "~прибыль", f"{today_profit:,} тг")
        kv3("7 дней",   s["week_orders"],   f"  {s['week_rev']:,} тг",   "", "~прибыль", f"{week_profit:,} тг")
        kv3("30 дней",  s["month_orders"],  f"  {s['month_rev']:,} тг",  "", "~прибыль", f"{month_profit:,} тг")
        kv3("Лучший день", s["best_day"][0], f"  {s['best_day'][1]:,} тг", "", "", "")

        # ─── БЛОК 3: Статусы ───────────────────
        blank()
        section("🚦 СТАТУСЫ ЗАКАЗОВ")
        total_active = s["total_orders"] + s["cancelled"]
        kv3("Новые / ожидают",  s["pending"],
            "В работе / в пути", s["in_progress"],
            "Выполнено",         s["done"])
        cancel_rate = round(s["cancelled"] / total_active * 100, 1) if total_active else 0
        done_rate   = round(s["done"] / total_active * 100, 1) if total_active else 0
        kv3("Отменено",         s["cancelled"],
            "Процент отмен",    f"{cancel_rate}%",
            "Процент выполнения", f"{done_rate}%")

        # ─── БЛОК 4: Топ виды курта ────────────
        blank()
        section("🧀 ПРОДАЖИ ПО ВИДАМ КУРТА")
        all_data.append(["", "Вид", "Продано (шт)", "", "Выручка (тг)", "", "", "Прибыль (тг)", ""])
        for name, qty, rev, profit in s["kurt_popularity"]:
            share = round(qty / s["total_units"] * 100, 1) if s["total_units"] else 0
            kv3(f"{name}", f"{qty} шт ({share}%)", f"  {rev:,} тг", "", f"  {profit:,} тг", "")

        # ─── БЛОК 5: Топ клиенты ───────────────
        blank()
        section("🏆 ТОП-5 КЛИЕНТОВ (по выручке)")
        all_data.append(["", "Клиент", "Заказов", "", "Куртов", "", "", "Выручка (тг)", ""])
        for i, cl in enumerate(s["top_clients"][:5]):
            display = (cl.get("name") or cl.get("username") or "?")[:28]
            kv3(f"{i+1}. {display}", cl["orders"], f"  {cl['units']} шт", "", f"  {cl['revenue']:,} тг", "")

        # ─── БЛОК 6: Пиковые часы ──────────────
        blank()
        section("⏰ АКТИВНОСТЬ ПО ЧАСАМ (топ-5)")
        for rank, (hour, cnt) in enumerate(s["peak_hours"], 1):
            bar = "█" * min(cnt, 20)
            kv3(f"#{rank}  {hour:02d}:00 – {hour+1:02d}:00", f"{cnt} заказов", bar, "", "", "")

        # ─── БЛОК 7: По дням недели ─────────────
        blank()
        section("📆 ЗАКАЗЫ ПО ДНЯМ НЕДЕЛИ")
        wd_row1 = [""] * 9
        wd_row2 = [""] * 9
        for i, (day, cnt) in enumerate(s["weekday_stats"][:7]):
            col = 1 + i  # B..H
            if col < 9:
                wd_row1[col] = day
                wd_row2[col] = cnt
        all_data.append(wd_row1)
        all_data.append(wd_row2)

        # ─── БЛОК 8: Последние 14 дней ─────────
        blank()
        section("📈 ВЫРУЧКА ЗА ПОСЛЕДНИЕ 14 ДНЕЙ")
        all_data.append(["", "Дата", "Заказов", "", "Выручка (тг)", "", "", "", ""])
        sorted_days = sorted(s["daily_rev"].items())[-14:]
        for d_str, rev in sorted_days:
            d_fmt = datetime.strptime(d_str, "%Y-%m-%d").strftime("%d.%m.%Y")
            orders_cnt = s["daily_orders"].get(d_str, 0)
            kv3(d_fmt, f"{orders_cnt} заказ(ов)", f"  {rev:,} тг", "", "", "")

        # ─── Футер ─────────────────────────────
        blank()
        all_data.append(["", f"🔄 Обновлено: {s['updated_at']}", "", "", "⚡ Автоматически при каждом заказе", "", "", "", ""])

        # ── Записать всё ──────────────────────────
        ws.update(all_data, "A1")

        # ── Форматирование ────────────────────────
        section_titles = [
            "📊 КЛЮЧЕВЫЕ МЕТРИКИ", "📅 ПО ПЕРИОДАМ",
            "🚦 СТАТУСЫ ЗАКАЗОВ", "🧀 ПРОДАЖИ ПО ВИДАМ КУРТА",
            "🏆 ТОП-5 КЛИЕНТОВ", "⏰ АКТИВНОСТЬ ПО ЧАСАМ",
            "📆 ЗАКАЗЫ ПО ДНЯМ НЕДЕЛИ", "📈 ВЫРУЧКА ЗА ПОСЛЕДНИЕ 14 ДНЕЙ"
        ]
        section_colors = [C_DARK_BLUE, C_TEAL, C_RED, C_GREEN,
                          C_PURPLE, C_MID_BLUE, C_DARK_TEAL, C_TEAL]
        subheader_markers = ["Период", "Вид", "Клиент", "Дата", "Пн"]

        sec_idx = 0
        batch_fmt = []

        ws.format("A1:I1", _cell_fmt(bg=C_DARK_BLUE, bold=True, size=16, fg=C_FG_WHITE))

        for ri, row in enumerate(all_data, 1):
            b_val = str(row[1]) if len(row) > 1 else ""
            is_section = any(t in b_val for t in section_titles)
            is_sub = any(m in b_val for m in subheader_markers)

            if is_section:
                color = section_colors[sec_idx % len(section_colors)]
                sec_idx += 1
                batch_fmt.append((f"A{ri}:I{ri}", _cell_fmt(bg=color, bold=True, size=11, fg=C_FG_WHITE)))
            elif is_sub:
                batch_fmt.append((f"A{ri}:I{ri}", _cell_fmt(bg=C_GREY_DARK, bold=True, size=10)))
            elif "Обновлено" in b_val:
                batch_fmt.append((f"A{ri}:I{ri}", _cell_fmt(bg=C_GREY_BG, italic=True,
                                                              fg={"red": 0.5, "green": 0.5, "blue": 0.5})))
            elif b_val and b_val != "ДАШБОРД":
                bg = C_LIGHT_BLUE if ri % 2 == 0 else C_WHITE
                batch_fmt.append((f"B{ri}:C{ri}", _cell_fmt(bg=bg, halign="LEFT")))
                batch_fmt.append((f"E{ri}:F{ri}", _cell_fmt(bg=bg, halign="LEFT")))
                batch_fmt.append((f"H{ri}:I{ri}", _cell_fmt(bg=bg, halign="LEFT")))
                c_val = str(row[2]) if len(row) > 2 else ""
                if c_val:
                    batch_fmt.append((f"C{ri}", _cell_fmt(bg=bg, bold=True, size=11, halign="LEFT")))
                    batch_fmt.append((f"F{ri}", _cell_fmt(bg=bg, bold=True, size=11, halign="LEFT")))
                    batch_fmt.append((f"I{ri}", _cell_fmt(bg=bg, bold=True, size=11, halign="LEFT")))

        # Высоты строк
        reqs2 = []
        for ri, row in enumerate(all_data, 1):
            b_val = str(row[1]) if len(row) > 1 else ""
            if any(t in b_val for t in section_titles) or b_val == "ДАШБОРД":
                reqs2.append(self._row_height(sid, ri - 1, ri, 34))
            else:
                reqs2.append(self._row_height(sid, ri - 1, ri, 26))
        if reqs2:
            self.sh.batch_update({"requests": reqs2})

        for rng, fmt in batch_fmt:
            try:
                ws.format(rng, fmt)
            except Exception:
                pass

    # ══════════════════════════════════════════
    # ЗАПИСЬ АНАЛИТИКИ
    # ══════════════════════════════════════════
    def _write_analytics(self, ws, s):
        ws.clear()
        sid = ws.id
        reqs = []
        reqs += self._col_widths(sid, [18, 155, 110, 110, 115, 110, 110, 110, 110, 110])
        reqs.append(self._row_height(sid, 0, 1, 44))
        reqs.append(self._merge(sid, 0, 0, 1, 10))
        self.sh.batch_update({"requests": reqs})

        all_data = []
        all_data.append(["АНАЛИТИКА"] + [""] * 9)

        def section(title):
            all_data.append(["", title] + [""] * 8)

        def blank():
            all_data.append([""] * 10)

        # ─── Секция 1: Обзор продукта ──────────
        blank()
        section("🧀 ДЕТАЛЬНАЯ АНАЛИТИКА ПО ВИДАМ КУРТА")
        all_data.append(["", "Вид курта", "Продано", "Доля %", "Выручка (тг)",
                         "Расходы (тг)", "Прибыль (тг)", "Маржа %", "Ср. прибыль/шт", ""])

        total_units = s["total_units"] or 1
        for name, qty, rev, profit in s["kurt_popularity"]:
            share = round(qty / total_units * 100, 1) if total_units else 0
            cost_total = rev - profit
            margin = round(profit / rev * 100, 1) if rev else 0
            avg_profit = round(profit / qty) if qty else 0
            all_data.append(["", name, qty, f"{share}%", f"{rev:,}",
                              f"{cost_total:,}", f"{profit:,}", f"{margin}%", avg_profit, ""])

        # Итоговая строка
        total_rev = s["total_revenue"]
        total_cost = s["total_cost"]
        total_profit = s["total_profit"]
        total_margin = s["margin"]
        all_data.append(["", "ИТОГО", total_units, "100%",
                         f"{total_rev:,}", f"{total_cost:,}", f"{total_profit:,}",
                         f"{total_margin}%", s["avg_check"], ""])

        # ─── Секция 2: Динамика 30 дней ────────
        blank()
        section("📅 ДИНАМИКА ЗА 30 ДНЕЙ")
        all_data.append(["", "Дата", "День нед.", "Заказов", "Куртов (прим.)",
                         "Выручка (тг)", "Прибыль (~тг)", "", "", ""])

        day_names_full = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]
        today = datetime.now().date()
        for d_str in sorted(s["last_30"].keys()):
            info = s["last_30"][d_str]
            d_obj = datetime.strptime(d_str, "%Y-%m-%d").date()
            wd = day_names_full[d_obj.weekday()]
            rev = info["rev"]
            ord_cnt = info["orders"]
            units_est = round(rev / PRICE_PER_UNIT) if rev else 0
            profit_est = round(rev * s["margin"] / 100) if s["margin"] and rev else 0
            d_fmt = d_obj.strftime("%d.%m.%Y")
            all_data.append(["", d_fmt, wd, ord_cnt, units_est,
                              f"{rev:,}" if rev else "—",
                              f"{profit_est:,}" if profit_est else "—", "", "", ""])

        # ─── Секция 3: Часовой анализ ──────────
        blank()
        section("⏰ ЗАКАЗЫ ПО ЧАСАМ (все время)")
        all_data.append(["", "Час", "Заказов", "Доля %", "Рейтинг", "", "", "", "", ""])

        hour_data = s["peak_hours"]
        max_hour_orders = hour_data[0][1] if hour_data else 1
        for hour, cnt in sorted(hour_data, key=lambda x: x[0]):
            share = round(cnt / (s["total_orders"] + s["cancelled"] or 1) * 100, 1)
            stars = "★" * min(5, round(cnt / max_hour_orders * 5))
            all_data.append(["", f"{hour:02d}:00", cnt, f"{share}%", stars, "", "", "", "", ""])

        # ─── Секция 4: KPI сводка ──────────────
        blank()
        section("🎯 KPI СВОДКА")
        kpi_rows = [
            ["", "Показатель", "Значение", "Описание", "", "", "", "", "", ""],
            ["", "Выручка всего", f"{s['total_revenue']:,} тг", "Сумма всех оплаченных заказов"],
            ["", "Чистая прибыль", f"{s['total_profit']:,} тг", "После расходов на сырьё, упаковку, доставку"],
            ["", "Маржинальность", f"{s['margin']}%", "Процент прибыли от выручки"],
            ["", "Средний чек", f"{s['avg_check']} тг", "Средняя сумма одного заказа"],
            ["", "Средний заказ", f"{round(s['total_units']/s['total_orders'])} шт" if s['total_orders'] else "—", "Среднее кол-во куртов в заказе"],
            ["", "Удержание клиентов", f"{s['retention_rate']}%", "Доля клиентов с 2+ заказами"],
            ["", "Отмены", f"{s['cancelled']} / {round(s['cancelled']/(s['total_orders']+s['cancelled'])*100, 1) if s['total_orders']+s['cancelled'] else 0}%", "Кол-во и процент отменённых заказов"],
            ["", "Бест-день", f"{s['best_day'][0]}", f"{s['best_day'][1]:,} тг выручки"],
        ]
        all_data.extend(kpi_rows)

        blank()
        all_data.append(["", f"🔄 Обновлено: {s['updated_at']}", "", "", "", "", "", "", "", ""])

        ws.update(all_data, "A1")

        # Форматирование
        section_titles = [
            "🧀 ДЕТАЛЬНАЯ АНАЛИТИКА", "📅 ДИНАМИКА ЗА 30 ДНЕЙ",
            "⏰ ЗАКАЗЫ ПО ЧАСАМ", "🎯 KPI СВОДКА"
        ]
        section_colors = [C_GREEN, C_TEAL, C_PURPLE, C_DARK_BLUE]
        subheader_check = ["Вид курта", "Дата", "Час", "Показатель"]

        ws.format("A1:J1", _cell_fmt(bg=C_DARK_BLUE, bold=True, size=16, fg=C_FG_WHITE))

        sec_idx = 0
        reqs3 = []
        batch_fmt = []

        for ri, row in enumerate(all_data, 1):
            b_val = str(row[1]) if len(row) > 1 else ""
            is_section = any(t in b_val for t in section_titles)
            is_sub = any(m == b_val for m in subheader_check)
            is_total = b_val == "ИТОГО"
            is_footer = "Обновлено" in b_val

            if is_section:
                color = section_colors[sec_idx % len(section_colors)]
                sec_idx += 1
                batch_fmt.append((f"A{ri}:J{ri}", _cell_fmt(bg=color, bold=True, size=11, fg=C_FG_WHITE)))
                reqs3.append(self._row_height(sid, ri - 1, ri, 32))
            elif is_sub:
                batch_fmt.append((f"A{ri}:J{ri}", _cell_fmt(bg=C_GREY_DARK, bold=True)))
                reqs3.append(self._row_height(sid, ri - 1, ri, 28))
            elif is_total:
                batch_fmt.append((f"A{ri}:J{ri}", _cell_fmt(bg=C_MID_BLUE, bold=True, size=10, fg=C_FG_WHITE)))
                reqs3.append(self._row_height(sid, ri - 1, ri, 28))
            elif is_footer:
                batch_fmt.append((f"A{ri}:J{ri}", _cell_fmt(bg=C_GREY_BG, italic=True,
                                                               fg={"red": 0.5, "green": 0.5, "blue": 0.5})))
                reqs3.append(self._row_height(sid, ri - 1, ri, 26))
            elif b_val and b_val != "АНАЛИТИКА":
                bg = C_LIGHT_BLUE if ri % 2 == 0 else C_WHITE
                batch_fmt.append((f"B{ri}:J{ri}", _cell_fmt(bg=bg, halign="LEFT")))
                reqs3.append(self._row_height(sid, ri - 1, ri, 24))

        if reqs3:
            self.sh.batch_update({"requests": reqs3})

        for rng, fmt in batch_fmt:
            try:
                ws.format(rng, fmt)
            except Exception:
                pass

    # ══════════════════════════════════════════
    # ЗАПИСЬ ЛИСТА КЛИЕНТОВ
    # ══════════════════════════════════════════
    def _write_clients(self, ws, s):
        ws.clear()
        sid = ws.id
        reqs = []
        reqs += self._col_widths(sid, [55, 185, 145, 125, 100, 100, 130, 130, 130, 130, 130])
        reqs.append(self._row_height(sid, 0, 1, 36))
        reqs.append(self._row_height(sid, 1, 2, 32))
        reqs.append(self._merge(sid, 0, 0, 1, 11))
        reqs.append(self._freeze(sid, rows=2))
        self.sh.batch_update({"requests": reqs})

        headers = ["№", "Клиент", "Username", "Телефон", "Заказов", "Куртов",
                   "Выручка (тг)", "Расходы (тг)", "Прибыль (тг)", "Средний чек", "Последний заказ"]
        ws.update([["БАЗА КЛИЕНТОВ"], headers], "A1")
        ws.format("A1:K1", _cell_fmt(bg=C_DARK_BLUE, bold=True, size=14, fg=C_FG_WHITE))
        ws.format("A2:K2", _cell_fmt(bg=C_MID_BLUE, bold=True, size=10, fg=C_FG_WHITE))

        clients_sorted = sorted(s["client_data"].values(), key=lambda c: c["revenue"], reverse=True)

        rows_data = []
        for i, cl in enumerate(clients_sorted):
            name = (cl.get("name") or "—")[:30]
            uname = f"@{cl['username']}" if cl.get("username") else "—"
            phone = cl.get("phone") or "—"
            orders = cl["orders"]
            units = cl["units"]
            revenue = cl["revenue"]
            cost = round(cl["cost"])
            profit = revenue - cost
            avg_check = round(revenue / orders) if orders else 0
            last_d = cl["last_date"].strftime("%d.%m.%Y") if hasattr(cl["last_date"], "strftime") else str(cl["last_date"])
            rows_data.append([i + 1, name, uname, phone, orders, units,
                               revenue, cost, profit, avg_check, last_d])

        if rows_data:
            ws.update(rows_data, "A3", value_input_option="USER_ENTERED")

            for i, row_d in enumerate(rows_data):
                ri = i + 3
                orders_cnt = row_d[4]
                if orders_cnt >= 5:
                    bg = C_LIGHT_GRN  # VIP
                elif orders_cnt >= 3:
                    bg = C_LIGHT_BLUE  # постоянный
                elif orders_cnt >= 2:
                    bg = C_LIGHT_ORG  # повторный
                else:
                    bg = C_WHITE  # новый
                try:
                    ws.format(f"A{ri}:K{ri}", _cell_fmt(bg=bg, halign="CENTER"))
                except Exception:
                    pass

        # Легенда
        legend_row = len(rows_data) + 5
        legend = [
            [],
            ["", "ЛЕГЕНДА"],
            ["", "🟩 Зелёный",  "VIP — 5+ заказов"],
            ["", "🟦 Синий",    "Постоянный — 3-4 заказа"],
            ["", "🟧 Оранжевый", "Повторный — 2 заказа"],
            ["", "⬜ Белый",    "Новый — 1 заказ"],
        ]
        ws.update(legend, f"A{legend_row}")
        ws.format(f"A{legend_row+1}:C{legend_row+1}", _cell_fmt(bg=C_MID_BLUE, bold=True, fg=C_FG_WHITE))

    # ══════════════════════════════════════════
    # REBUILD FROM DB
    # ══════════════════════════════════════════
    def rebuild_from_db(self, db_file="kurt_orders.db", database_url=""):
        logger.info("Перенос данных из БД в Google Sheets...")

        ws = self._get_or_create_sheet(SHEET_ORDERS)
        ws.clear()
        self._setup_orders_sheet(ws)

        conn = _sm_get_conn(db_file, database_url)
        c = conn.cursor()
        c.execute("SELECT * FROM orders ORDER BY id")
        rows = c.fetchall()
        conn.close()

        for row in rows:
            oid, uid, uname, fname, phone, addr, items_j, tu, bu, tp, status, created = row
            items = json.loads(items_j)
            self.add_order(oid, uid, uname, fname, phone, addr, items, tu, bu, tp)
            self.update_order_status(oid, status)

        self.refresh_dashboard(db_file, database_url)

        # Пересоздать структуру клиентов
        ws_clients = self._get_or_create_sheet(SHEET_CLIENTS)
        self._setup_clients_sheet(ws_clients)

        # Пересоздать аналитику
        ws_analytics = self._get_or_create_sheet(SHEET_ANALYTICS)
        self._setup_analytics_shell(ws_analytics)

        logger.info(f"Перенесено {len(rows)} заказов ✅")
