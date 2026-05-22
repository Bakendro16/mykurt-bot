"""
КуртБот - Telegram бот для приёма заказов на курт
Установка: pip install -r requirements.txt
Запуск: python bot.py
"""

import asyncio
import hashlib
import logging
import signal
import sqlite3
import json
import os
import re
from datetime import datetime

from dotenv import load_dotenv
load_dotenv()

from telegram import (
    Update, InlineKeyboardButton, InlineKeyboardMarkup, BotCommand,
    BotCommandScopeAllPrivateChats, BotCommandScopeChat,
)
from telegram.ext import (
    Application, CommandHandler, CallbackQueryHandler,
    MessageHandler, filters, ContextTypes,
)
from telegram.error import BadRequest, TelegramError
from sheets_manager import SheetsManager

BOT_TOKEN    = os.environ["BOT_TOKEN"]
BOT_USERNAME = os.environ.get("BOT_USERNAME", "Kurtmy_bot")
ADMIN_IDS    = [int(x.strip()) for x in os.environ.get("ADMIN_IDS", "1224332440").split(",") if x.strip()]
DB_FILE      = os.environ.get("DB_FILE", "kurt_orders.db")
DATABASE_URL = os.environ.get("DATABASE_URL", "")

SPREADSHEET_ID          = os.environ.get("SPREADSHEET_ID", "")
CREDENTIALS_FILE        = os.environ.get("CREDENTIALS_FILE", "credentials.json")
GOOGLE_CREDENTIALS_JSON = os.environ.get("GOOGLE_CREDENTIALS_JSON", "")

WEBHOOK_URL  = os.environ.get("WEBHOOK_URL", "").rstrip("/")
WEBHOOK_PATH = hashlib.sha256(BOT_TOKEN.encode()).hexdigest()[:32]
PORT         = int(os.environ.get("PORT", 8080))

_CREDS_SOURCE = None

KURT_TYPES = {
    "ordinary":  {"name": "Обычный",      "price": 700, "emoji": "🟡"},
    "butter":    {"name": "С маслом",     "price": 700, "emoji": "🧈"},
    "rhombus":   {"name": "Ромбик",       "price": 700, "emoji": "💠"},
    "tablet":    {"name": "Таблетка",     "price": 700, "emoji": "⬜"},
    "spiderweb": {"name": "Паутинка",     "price": 700, "emoji": "🕸️"},
    "smoked":    {"name": "Копчёный",     "price": 700, "emoji": "🔥"},
    "mix":       {"name": "Микс из всех", "price": 700, "emoji": "🎉"},
}

PHOTO_KEY_ALIASES = {
    "обычный":       "ordinary",
    "ordinary":      "ordinary",
    "масло":         "butter",
    "маслом":        "butter",
    "с маслом":      "butter",
    "butter":        "butter",
    "ромбик":        "rhombus",
    "rhombus":       "rhombus",
    "таблетка":      "tablet",
    "tablet":        "tablet",
    "паутинка":      "spiderweb",
    "spiderweb":     "spiderweb",
    "копченый":      "smoked",
    "копчёный":      "smoked",
    "smoked":        "smoked",
    "микс":          "mix",
    "mix":           "mix",
    "микс из всех":  "mix",
    "mix of all":    "mix",
}

PRICE_PER_UNIT  = 700
BONUS_THRESHOLD = 20
BONUS_UNITS     = 1
MIN_ORDER_QTY   = 10

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)
logger = logging.getLogger(__name__)


def escape_md(text: str) -> str:
    return re.sub(r'([_*\[\]()~`>#+\-=|{}.!\\])', r'\\\1', str(text))


# ──────────────────────────────────────────────
# АБСТРАКЦИЯ БАЗЫ ДАННЫХ
# ──────────────────────────────────────────────
def _get_conn():
    if DATABASE_URL:
        import psycopg2
        url = DATABASE_URL
        if url.startswith("postgres://"):
            url = url.replace("postgres://", "postgresql://", 1)
        return psycopg2.connect(url)
    return sqlite3.connect(DB_FILE)


def _ph() -> str:
    return "%s" if DATABASE_URL else "?"


def _id_col() -> str:
    return "SERIAL PRIMARY KEY" if DATABASE_URL else "INTEGER PRIMARY KEY AUTOINCREMENT"


# ──────────────────────────────────────────────
# БАЗА ДАННЫХ
# ──────────────────────────────────────────────
def init_db():
    id_col      = _id_col()
    uid_type    = "BIGINT" if DATABASE_URL else "INTEGER"
    conn = _get_conn()
    c = conn.cursor()
    c.execute(f"""
        CREATE TABLE IF NOT EXISTS orders (
            id          {id_col},
            user_id     {uid_type},
            username    TEXT,
            full_name   TEXT,
            phone       TEXT,
            address     TEXT,
            items       TEXT,
            total_units INTEGER,
            bonus_units INTEGER,
            total_price INTEGER,
            status      TEXT DEFAULT 'new',
            created_at  TEXT
        )
    """)
    c.execute(f"""
        CREATE TABLE IF NOT EXISTS users (
            user_id      {uid_type} PRIMARY KEY,
            username     TEXT,
            full_name    TEXT,
            phone        TEXT,
            address      TEXT,
            total_orders INTEGER DEFAULT 0,
            total_units  INTEGER DEFAULT 0
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS kurt_photos (
            kurt_key  TEXT PRIMARY KEY,
            file_id   TEXT NOT NULL
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS kurt_stock (
            kurt_key  TEXT PRIMARY KEY,
            quantity  INTEGER DEFAULT -1
        )
    """)
    conn.commit()
    # Миграции существующих таблиц
    for migration in [
        "ALTER TABLE users ADD COLUMN address TEXT",
        "ALTER TABLE orders ALTER COLUMN user_id TYPE BIGINT" if DATABASE_URL else "",
        "ALTER TABLE users  ALTER COLUMN user_id TYPE BIGINT" if DATABASE_URL else "",
    ]:
        if not migration:
            continue
        try:
            c.execute(migration)
            conn.commit()
        except Exception:
            conn.rollback()
    conn.close()


def get_photo(kurt_key):
    ph = _ph()
    conn = _get_conn()
    c = conn.cursor()
    c.execute(f"SELECT file_id FROM kurt_photos WHERE kurt_key={ph}", (kurt_key,))
    row = c.fetchone()
    conn.close()
    return row[0] if row else None


def set_photo(kurt_key, file_id):
    ph = _ph()
    conn = _get_conn()
    c = conn.cursor()
    c.execute(f"""
        INSERT INTO kurt_photos (kurt_key, file_id) VALUES ({ph},{ph})
        ON CONFLICT(kurt_key) DO UPDATE SET file_id=excluded.file_id
    """, (kurt_key, file_id))
    conn.commit()
    conn.close()


def get_all_photos():
    conn = _get_conn()
    c = conn.cursor()
    c.execute("SELECT kurt_key, file_id FROM kurt_photos")
    rows = c.fetchall()
    conn.close()
    return {r[0]: r[1] for r in rows}


def get_user_saved_data(user_id):
    ph = _ph()
    conn = _get_conn()
    c = conn.cursor()
    c.execute(f"SELECT phone, address FROM users WHERE user_id={ph}", (user_id,))
    row = c.fetchone()
    conn.close()
    if row:
        return row[0] or "", row[1] or ""
    return "", ""


def save_order(user_id, username, full_name, phone, address, items,
               total_units, bonus_units, total_price):
    ph = _ph()
    conn = _get_conn()
    c = conn.cursor()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    items_json = json.dumps(items, ensure_ascii=False)
    params = (user_id, username, full_name, phone, address, items_json,
              total_units, bonus_units, total_price, now)
    if DATABASE_URL:
        c.execute("""
            INSERT INTO orders (user_id, username, full_name, phone, address, items,
                                total_units, bonus_units, total_price, created_at)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) RETURNING id
        """, params)
        order_id = c.fetchone()[0]
    else:
        c.execute("""
            INSERT INTO orders (user_id, username, full_name, phone, address, items,
                                total_units, bonus_units, total_price, created_at)
            VALUES (?,?,?,?,?,?,?,?,?,?)
        """, params)
        order_id = c.lastrowid
    c.execute(f"""
        INSERT INTO users (user_id, username, full_name, phone, address, total_orders, total_units)
        VALUES ({ph},{ph},{ph},{ph},{ph},1,{ph})
        ON CONFLICT(user_id) DO UPDATE SET
            total_orders = users.total_orders + 1,
            total_units  = users.total_units + {ph},
            username     = excluded.username,
            full_name    = excluded.full_name,
            phone        = COALESCE(excluded.phone, users.phone),
            address      = COALESCE(excluded.address, users.address)
    """, (user_id, username, full_name, phone, address, total_units, total_units))
    conn.commit()
    conn.close()
    return order_id


def get_order(order_id):
    ph = _ph()
    conn = _get_conn()
    c = conn.cursor()
    c.execute(f"SELECT * FROM orders WHERE id={ph}", (order_id,))
    row = c.fetchone()
    conn.close()
    return row


def update_order_status(order_id, status):
    ph = _ph()
    conn = _get_conn()
    c = conn.cursor()
    c.execute(f"UPDATE orders SET status={ph} WHERE id={ph}", (status, order_id))
    conn.commit()
    conn.close()


def get_all_orders():
    conn = _get_conn()
    c = conn.cursor()
    c.execute("SELECT * FROM orders ORDER BY created_at DESC")
    rows = c.fetchall()
    conn.close()
    return rows


def get_stock(kurt_key: str) -> int:
    ph = _ph()
    conn = _get_conn()
    c = conn.cursor()
    c.execute(f"SELECT quantity FROM kurt_stock WHERE kurt_key={ph}", (kurt_key,))
    row = c.fetchone()
    conn.close()
    return row[0] if row else -1


def set_stock(kurt_key: str, quantity: int):
    ph = _ph()
    conn = _get_conn()
    c = conn.cursor()
    c.execute(f"""
        INSERT INTO kurt_stock (kurt_key, quantity) VALUES ({ph},{ph})
        ON CONFLICT(kurt_key) DO UPDATE SET quantity=excluded.quantity
    """, (kurt_key, quantity))
    conn.commit()
    conn.close()


def get_all_stock() -> dict:
    conn = _get_conn()
    c = conn.cursor()
    c.execute("SELECT kurt_key, quantity FROM kurt_stock")
    rows = c.fetchall()
    conn.close()
    return {r[0]: r[1] for r in rows}


def get_all_user_ids() -> list:
    conn = _get_conn()
    c = conn.cursor()
    c.execute("SELECT user_id FROM users")
    rows = c.fetchall()
    conn.close()
    return [r[0] for r in rows]


def calc_bonus(total_units):
    return (total_units // BONUS_THRESHOLD) * BONUS_UNITS


def format_admin_order(order_id, user_id, username, full_name, phone, address,
                       items_json, total_units, bonus_units, total_price, status, created_at):
    items = json.loads(items_json)
    lines = [
        f"🔔 Новый заказ #{order_id}",
        f"🕐 {created_at}",
        f"👤 {full_name} (@{username or '—'})",
        f"📞 {phone}",
        f"📍 {address}",
        "",
        "📦 Состав:",
    ]
    for key, qty in items.items():
        if qty > 0:
            info = KURT_TYPES[key]
            lines.append(f"  {info['emoji']} {info['name']}: {qty} шт")
    lines.append(f"\n🧮 Всего: {total_units} шт")
    if bonus_units > 0:
        lines.append(f"🎁 Бонус: +{bonus_units} шт")
    lines.append(f"💰 Сумма: {total_price} тг")
    return "\n".join(lines)


# ──────────────────────────────────────────────
# КЛАВИАТУРЫ
# ──────────────────────────────────────────────
def main_menu_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🛒 Сделать заказ", callback_data="new_order")],
        [InlineKeyboardButton("🧀 Ассортимент",   callback_data="catalog"),
         InlineKeyboardButton("📋 Мои заказы",    callback_data="my_orders")],
        [InlineKeyboardButton("🚚 Доставка и оплата", callback_data="delivery_info")],
    ])


def catalog_keyboard():
    stock = get_all_stock()
    buttons = []
    for key, info in KURT_TYPES.items():
        qty = stock.get(key, -1)
        if qty == 0:
            label = f"{info['emoji']} {info['name']} — нет в наличии ❌"
        else:
            label = f"{info['emoji']} {info['name']} — 700 тг"
        buttons.append([InlineKeyboardButton(label, callback_data=f"info_{key}")])
    buttons.append([InlineKeyboardButton("◀️ Назад", callback_data="back_main")])
    return InlineKeyboardMarkup(buttons)


def select_type_keyboard(cart):
    stock = get_all_stock()
    buttons = []
    for key, info in KURT_TYPES.items():
        qty_stock = stock.get(key, -1)
        qty_cart  = cart.get(key, 0)
        if qty_stock == 0:
            label = f"{info['emoji']} {info['name']} ❌"
        else:
            label = f"{info['emoji']} {info['name']}"
            if qty_cart > 0:
                label += f" ✅ ({qty_cart})"
        buttons.append([InlineKeyboardButton(label, callback_data=f"type_{key}")])
    total = sum(cart.values())
    row = []
    if total >= MIN_ORDER_QTY:
        row.append(InlineKeyboardButton(f"✅ Оформить ({total} шт)", callback_data="checkout"))
    elif total > 0:
        row.append(InlineKeyboardButton(
            f"⚠️ Мин. {MIN_ORDER_QTY} шт (сейчас {total})", callback_data="min_order_warn"
        ))
    row.append(InlineKeyboardButton("❌ Отмена", callback_data="cancel"))
    buttons.append(row)
    return InlineKeyboardMarkup(buttons)


def qty_keyboard(kurt_key, current_qty):
    buttons = [
        [
            InlineKeyboardButton("➖", callback_data=f"qty_{kurt_key}_minus"),
            InlineKeyboardButton(f"  {current_qty}  ", callback_data="noop"),
            InlineKeyboardButton("➕", callback_data=f"qty_{kurt_key}_plus"),
        ],
        [
            InlineKeyboardButton("5",  callback_data=f"qty_{kurt_key}_set_5"),
            InlineKeyboardButton("10", callback_data=f"qty_{kurt_key}_set_10"),
            InlineKeyboardButton("20", callback_data=f"qty_{kurt_key}_set_20"),
            InlineKeyboardButton("50", callback_data=f"qty_{kurt_key}_set_50"),
        ],
        [InlineKeyboardButton("◀️ Назад к списку", callback_data="back_types")],
    ]
    return InlineKeyboardMarkup(buttons)


def confirm_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ Подтвердить заказ", callback_data="confirm_yes")],
        [InlineKeyboardButton("✏️ Изменить",          callback_data="back_types")],
        [InlineKeyboardButton("❌ Отмена",             callback_data="cancel")],
    ])


def contact_keyboard(saved_phone=""):
    buttons = []
    if saved_phone:
        buttons.append([InlineKeyboardButton(
            f"📞 Использовать {saved_phone}", callback_data="use_saved_phone"
        )])
    buttons.append([InlineKeyboardButton("❌ Отмена", callback_data="cancel")])
    return InlineKeyboardMarkup(buttons)


def address_keyboard(saved_address=""):
    buttons = []
    if saved_address:
        short = saved_address[:30] + ("…" if len(saved_address) > 30 else "")
        buttons.append([InlineKeyboardButton(
            f"📍 Использовать: {short}", callback_data="use_saved_address"
        )])
    buttons.append([InlineKeyboardButton("❌ Отмена", callback_data="cancel")])
    return InlineKeyboardMarkup(buttons)


def admin_order_keyboard(order_id):
    return InlineKeyboardMarkup([
        [
            InlineKeyboardButton("✅ Принят",   callback_data=f"admin_accepted_{order_id}"),
            InlineKeyboardButton("🚚 В пути",   callback_data=f"admin_shipping_{order_id}"),
        ],
        [
            InlineKeyboardButton("✔️ Выполнен", callback_data=f"admin_done_{order_id}"),
            InlineKeyboardButton("❌ Отменён",  callback_data=f"admin_cancel_{order_id}"),
        ],
    ])


# ──────────────────────────────────────────────
# БЕЗОПАСНЫЕ ОБЁРТКИ
# ──────────────────────────────────────────────
async def safe_answer(query, text="", show_alert=False):
    try:
        await query.answer(text, show_alert=show_alert)
    except BadRequest as e:
        err = str(e).lower()
        if "query is too old" in err or "query id is invalid" in err:
            logger.warning(f"Устаревший callback от {query.from_user.id}: {e}")
        else:
            raise


async def safe_edit(query, text, parse_mode=None, reply_markup=None):
    if query.message and query.message.photo:
        try:
            await query.message.delete()
        except Exception:
            pass
        try:
            await query.message.chat.send_message(
                text, parse_mode=parse_mode, reply_markup=reply_markup
            )
        except Exception as e:
            logger.warning(f"Не удалось отправить сообщение после удаления фото: {e}")
        return
    try:
        await query.edit_message_text(text, parse_mode=parse_mode, reply_markup=reply_markup)
    except BadRequest as e:
        err = str(e).lower()
        if "message is not modified" in err:
            pass
        elif "query is too old" in err or "message to edit not found" in err:
            logger.warning(f"Не удалось отредактировать: {e}")
            try:
                await query.message.reply_text(
                    text, parse_mode=parse_mode, reply_markup=reply_markup
                )
            except Exception:
                pass
        else:
            raise


# ──────────────────────────────────────────────
# ФОРМИРОВАНИЕ ТЕКСТА ЗАКАЗА
# ──────────────────────────────────────────────
def build_order_text(cart, name, phone, address):
    lines = ["📦 *Ваш заказ:*"]
    total = 0
    for key, qty in cart.items():
        if qty > 0:
            info = KURT_TYPES[key]
            subtotal = qty * PRICE_PER_UNIT
            total += subtotal
            lines.append(
                f"  {info['emoji']} {escape_md(info['name'])}: "
                f"{qty} шт × 700 \\= {subtotal} тг"
            )
    total_units = sum(v for v in cart.values() if v > 0)
    bonus = calc_bonus(total_units)
    lines.append(f"\n🧮 Итого: *{total_units} шт*")
    if bonus > 0:
        lines.append(f"🎁 Бонус 20\\+1: *\\+{bonus} шт бесплатно\\!*")
    lines.append(f"💰 Сумма: *{total} тг*")
    lines.append(f"\n👤 Имя: {escape_md(name)}")
    lines.append(f"📞 Телефон: {escape_md(phone)}")
    lines.append(f"📍 Адрес: {escape_md(address)}")
    return "\n".join(lines)


# ──────────────────────────────────────────────
# ОБРАБОТЧИКИ
# ──────────────────────────────────────────────
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    text = (
        f"🧀 *Добро пожаловать в MyKurt\\!*\n\n"
        f"Настоящий домашний курт из Казахстана\\.\n"
        f"Готовим с душой — доставляем к вашей двери\\.\n\n"
        f"💰 *Цена:* 700 тг/шт\n"
        f"📦 *Минимальный заказ:* {MIN_ORDER_QTY} шт\n"
        f"🎁 *Акция:* каждые 20 шт — \\+1 в подарок\\!\n\n"
        f"👇 Выберите действие:"
    )
    if update.message:
        await update.message.reply_text(
            text, parse_mode="MarkdownV2", reply_markup=main_menu_keyboard()
        )
    else:
        query = update.callback_query
        await safe_answer(query)
        await safe_edit(query, text, parse_mode="MarkdownV2", reply_markup=main_menu_keyboard())


async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    user  = update.effective_user
    data  = query.data

    if query.message.chat.type != "private" and not data.startswith("admin_"):
        await safe_answer(query)
        await query.message.reply_text(
            "👋 Продолжите в личных сообщениях с ботом:",
            reply_markup=InlineKeyboardMarkup([[
                InlineKeyboardButton(
                    "📩 Открыть главное меню",
                    url=f"https://t.me/{BOT_USERNAME}?start=menu"
                )
            ]])
        )
        return

    await safe_answer(query)

    # ── Навигация ─────────────────────────────
    if data == "back_main":
        await start(update, context)
        return

    if data == "delivery_info":
        await safe_edit(
            query,
            "🚚 *Доставка и оплата*\n\n"
            "📍 Доставляем по городу и пригороду\n"
            "🕐 Срок: 1\\-2 дня после подтверждения заказа\n"
            "💳 Оплата: наличными или Kaspi при получении\n\n"
            "📞 Есть вопросы? Просто напишите нам — ответим быстро\\!",
            parse_mode="MarkdownV2",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🛒 Сделать заказ", callback_data="new_order")],
                [InlineKeyboardButton("◀️ Главное меню",  callback_data="back_main")],
            ])
        )
        return

    if data == "catalog":
        await safe_edit(
            query,
            f"🧀 *Наш ассортимент*\n\n"
            f"Весь курт — *700 тг/шт*\n"
            f"📦 Минимальный заказ: *{MIN_ORDER_QTY} шт*\n"
            f"🎁 При заказе от 20 шт — один в подарок\\!\n\n"
            f"Нажмите на вид чтобы узнать подробнее 👇",
            parse_mode="MarkdownV2",
            reply_markup=catalog_keyboard()
        )
        return

    if data.startswith("info_"):
        key  = data[5:]
        info = KURT_TYPES[key]
        back_kb  = InlineKeyboardMarkup([[InlineKeyboardButton("◀️ Назад", callback_data="catalog")]])
        photo_id = get_photo(key)
        caption  = f"{info['emoji']} *{escape_md(info['name'])}*\nЦена: 700 тг/шт"
        if photo_id:
            try:
                await query.message.delete()
            except Exception:
                pass
            await query.message.chat.send_photo(
                photo=photo_id, caption=caption,
                parse_mode="MarkdownV2", reply_markup=back_kb
            )
        else:
            await safe_edit(query, caption, parse_mode="MarkdownV2", reply_markup=back_kb)
        return

    if data == "my_orders":
        await show_my_orders(update, context)
        return

    if data == "new_order":
        context.user_data["cart"] = {k: 0 for k in KURT_TYPES}
        await safe_edit(
            query,
            f"🛒 *Выберите вид курта*\n\n_Добавьте нужное количество каждого вида\\.\nМинимальный заказ: {MIN_ORDER_QTY} шт_",
            parse_mode="MarkdownV2",
            reply_markup=select_type_keyboard(context.user_data["cart"])
        )
        return

    if data == "back_types":
        cart = context.user_data.get("cart", {k: 0 for k in KURT_TYPES})
        await safe_edit(
            query,
            f"🛒 *Выберите вид курта*\n\n_Добавьте нужное количество каждого вида\\.\nМинимальный заказ: {MIN_ORDER_QTY} шт_",
            parse_mode="MarkdownV2",
            reply_markup=select_type_keyboard(cart)
        )
        return

    if data == "min_order_warn":
        total = sum(context.user_data.get("cart", {}).values())
        await safe_answer(
            query,
            f"⚠️ Минимальный заказ — {MIN_ORDER_QTY} шт. У вас: {total}.",
            show_alert=True
        )
        return

    if data.startswith("type_"):
        key  = data[5:]
        if get_stock(key) == 0:
            await safe_answer(
                query,
                f"❌ {KURT_TYPES[key]['name']} сейчас нет в наличии.",
                show_alert=True
            )
            return
        context.user_data["current_type"] = key
        qty  = context.user_data.get("cart", {}).get(key, 0)
        info = KURT_TYPES[key]
        caption  = f"{info['emoji']} *{escape_md(info['name'])}*\nКоличество: {qty} шт"
        kb       = qty_keyboard(key, qty)
        photo_id = get_photo(key)
        if photo_id:
            try:
                await query.message.delete()
            except Exception:
                pass
            await query.message.chat.send_photo(
                photo=photo_id, caption=caption,
                parse_mode="MarkdownV2", reply_markup=kb
            )
        else:
            await safe_edit(query, caption, parse_mode="MarkdownV2", reply_markup=kb)
        return

    if data.startswith("qty_"):
        parts   = data.split("_")
        key     = parts[1]
        action  = parts[2]
        cart    = context.user_data.get("cart", {k: 0 for k in KURT_TYPES})
        current = cart.get(key, 0)
        if action == "plus":
            current = min(current + 1, 999)
        elif action == "minus":
            current = max(current - 1, 0)
        elif action == "set":
            current = int(parts[3])
        cart[key] = current
        context.user_data["cart"] = cart
        info    = KURT_TYPES[key]
        caption = f"{info['emoji']} *{escape_md(info['name'])}*\nКоличество: {current} шт"
        kb      = qty_keyboard(key, current)
        if query.message.photo:
            try:
                await query.edit_message_caption(
                    caption=caption, parse_mode="MarkdownV2", reply_markup=kb
                )
            except BadRequest as e:
                if "message is not modified" not in str(e).lower():
                    logger.warning(f"edit_message_caption: {e}")
        else:
            await safe_edit(query, caption, parse_mode="MarkdownV2", reply_markup=kb)
        return

    if data == "noop":
        return

    if data == "checkout":
        cart        = context.user_data.get("cart", {})
        total_units = sum(cart.values())
        if total_units < MIN_ORDER_QTY:
            await safe_answer(
                query,
                f"⚠️ Минимальный заказ — {MIN_ORDER_QTY} шт! У вас: {total_units}.",
                show_alert=True
            )
            return
        saved_phone, saved_address = get_user_saved_data(user.id)
        context.user_data["saved_phone"]   = saved_phone
        context.user_data["saved_address"] = saved_address

        phone = context.user_data.get("phone", "")
        if not phone:
            context.user_data["awaiting"] = "phone"
            hint = f"\n\nПрошлый номер: {escape_md(saved_phone)}" if saved_phone else ""
            await safe_edit(
                query,
                f"📞 *Введите номер телефона*\n\nМы позвоним для подтверждения доставки\\.{hint}",
                parse_mode="MarkdownV2",
                reply_markup=contact_keyboard(saved_phone)
            )
            return

        address = context.user_data.get("address", "")
        if not address:
            context.user_data["awaiting"] = "address"
            hint = f"\n\nПрошлый адрес: {escape_md(saved_address)}" if saved_address else ""
            await safe_edit(
                query,
                f"📍 *Введите адрес доставки*\n\nУкажите улицу, дом, квартиру\\.{hint}",
                parse_mode="MarkdownV2",
                reply_markup=address_keyboard(saved_address)
            )
            return

        await show_confirm(update, context)
        return

    if data == "use_saved_phone":
        saved_phone = context.user_data.get("saved_phone", "")
        if not saved_phone:
            await safe_answer(query, "Сохранённый номер не найден", show_alert=True)
            return
        context.user_data["phone"]    = saved_phone
        context.user_data["awaiting"] = None
        saved_address = context.user_data.get("saved_address", "")
        address = context.user_data.get("address", "")
        if not address:
            context.user_data["awaiting"] = "address"
            hint = f"\nПрошлый адрес: {saved_address}" if saved_address else ""
            await safe_edit(
                query,
                f"📍 Введите адрес доставки:{hint}",
                reply_markup=address_keyboard(saved_address)
            )
        else:
            await show_confirm(update, context)
        return

    if data == "use_saved_address":
        saved_address = context.user_data.get("saved_address", "")
        if not saved_address:
            await safe_answer(query, "Сохранённый адрес не найден", show_alert=True)
            return
        context.user_data["address"]  = saved_address
        context.user_data["awaiting"] = None
        await show_confirm(update, context)
        return

    if data == "confirm_yes":
        await finalize_order(update, context)
        return

    if data == "cancel":
        context.user_data.clear()
        await safe_edit(
            query,
            "Заказ отменён\\. Если передумаете — мы всегда здесь 😊",
            parse_mode="MarkdownV2",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🏠 Главное меню", callback_data="back_main")]
            ])
        )
        return

    # ── Админские кнопки ──────────────────────
    if data.startswith("admin_"):
        if user.id not in ADMIN_IDS:
            await safe_answer(query, "Нет доступа", show_alert=True)
            return
        parts    = data.split("_")
        action   = parts[1]
        order_id = int(parts[2])
        status_map = {
            "accepted": ("принят ✅",   "✅ Принят"),
            "shipping": ("в пути 🚚",   "🚚 В пути"),
            "done":     ("выполнен ✔️", "✔️ Выполнен"),
            "cancel":   ("отменён ❌",  "❌ Отменён"),
        }
        client_notify = {
            "accepted": (
                f"✅ Ваш заказ #{order_id} подтверждён!\n\n"
                f"Мы уже приступили к сборке 🧀\n"
                f"Скоро свяжемся с вами для уточнения деталей доставки."
            ),
            "shipping": (
                f"🚚 Ваш заказ #{order_id} в пути!\n\n"
                f"Курт уже едет к вам — совсем скоро будет 🧀\n"
                f"Ориентировочно: в течение дня.\n\n"
                f"Спасибо, что выбрали MyKurt! 🙏"
            ),
            "done": (
                f"🎉 Ваш заказ #{order_id} доставлен!\n\n"
                f"Надеемся, курт вам понравился 🧀\n"
                f"Будем рады видеть вас снова!\n\n"
                f"Нажмите /start чтобы сделать новый заказ."
            ),
            "cancel": (
                f"😔 Ваш заказ #{order_id} был отменён.\n\n"
                f"Если возникли вопросы — просто напишите нам.\n"
                f"Нажмите /start чтобы оформить новый заказ."
            ),
        }

        if action in status_map:
            db_status, label = status_map[action]
            update_order_status(order_id, db_status)
            try:
                sm = SheetsManager(_CREDS_SOURCE, SPREADSHEET_ID)
                sm.update_order_status(order_id, db_status)
            except Exception as e:
                logger.error(f"Sheets update error: {e}")

            order_row = get_order(order_id)
            if order_row and action in client_notify:
                client_user_id = order_row[1]
                try:
                    await context.bot.send_message(
                        chat_id=client_user_id,
                        text=client_notify[action]
                    )
                except TelegramError as e:
                    logger.error(f"Не удалось уведомить клиента {client_user_id}: {e}")

            new_text = (query.message.text or "") + f"\n\nСтатус изменён: {label}"
            try:
                await query.edit_message_text(new_text)
            except Exception:
                pass
        return


async def message_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    awaiting = context.user_data.get("awaiting")
    if not awaiting:
        if update.message and update.message.chat.type == "private":
            await update.message.reply_text(
                "👋 Нажмите /start чтобы открыть меню и сделать заказ."
            )
        return
    text = update.message.text.strip()
    if awaiting == "phone":
        context.user_data["phone"]    = text
        context.user_data["awaiting"] = "address"
        saved_address = context.user_data.get("saved_address", "")
        hint = f"\n\nПрошлый адрес: {saved_address}" if saved_address else ""
        await update.message.reply_text(
            f"📍 Отлично! Теперь укажите адрес доставки.\n"
            f"Улица, дом, квартира, ориентир.{hint}",
            reply_markup=address_keyboard(saved_address)
        )
        return
    if awaiting == "address":
        context.user_data["address"]  = text
        context.user_data["awaiting"] = None
        try:
            await show_confirm_msg(update, context)
        except Exception as e:
            logger.error(f"show_confirm_msg error: {e}", exc_info=True)
            await update.message.reply_text(
                "⚠️ Произошла ошибка при формировании заказа.\n"
                "Попробуйте начать заново — /start"
            )
        return


async def show_confirm_msg(update: Update, context: ContextTypes.DEFAULT_TYPE):
    cart    = context.user_data.get("cart", {})
    phone   = context.user_data.get("phone", "")
    address = context.user_data.get("address", "")
    user    = update.effective_user
    name    = f"{user.first_name} {user.last_name or ''}".strip()
    text    = build_order_text(cart, name, phone, address)
    await update.message.reply_text(
        text + "\n\n_Подтвердите заказ:_",
        parse_mode="MarkdownV2",
        reply_markup=confirm_keyboard()
    )


async def show_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query   = update.callback_query
    cart    = context.user_data.get("cart", {})
    phone   = context.user_data.get("phone", "")
    address = context.user_data.get("address", "")
    user    = update.effective_user
    name    = f"{user.first_name} {user.last_name or ''}".strip()
    text    = build_order_text(cart, name, phone, address)
    await safe_edit(
        query,
        text + "\n\n_Подтвердите заказ:_",
        parse_mode="MarkdownV2",
        reply_markup=confirm_keyboard()
    )


async def finalize_order(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query   = update.callback_query
    user    = update.effective_user
    cart    = context.user_data.get("cart", {})
    phone   = context.user_data.get("phone", "—")
    address = context.user_data.get("address", "—")
    name    = f"{user.first_name} {user.last_name or ''}".strip()

    total_units = sum(v for v in cart.values())
    if total_units < MIN_ORDER_QTY:
        await safe_answer(
            query,
            f"⚠️ Минимальный заказ — {MIN_ORDER_QTY} шт! У вас: {total_units}.",
            show_alert=True
        )
        return

    bonus       = calc_bonus(total_units)
    total_price = total_units * PRICE_PER_UNIT
    filtered    = {k: v for k, v in cart.items() if v > 0}

    try:
        order_id = save_order(
            user_id=user.id, username=user.username or "",
            full_name=name, phone=phone, address=address,
            items=filtered, total_units=total_units,
            bonus_units=bonus, total_price=total_price
        )
    except Exception as e:
        logger.error(f"save_order error: {e}", exc_info=True)
        await safe_edit(
            query,
            "⚠️ Не удалось сохранить заказ\\. Попробуйте ещё раз или напишите нам\\.",
            parse_mode="MarkdownV2",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🔄 Попробовать снова", callback_data="confirm_yes")],
                [InlineKeyboardButton("🏠 Главное меню",      callback_data="back_main")],
            ])
        )
        return

    # ── Сразу показываем успех пользователю ──────────────────────────
    bonus_text = f"\n🎁 *Бонус:* \\+{bonus} шт бесплатно\\!" if bonus > 0 else ""
    await safe_edit(
        query,
        f"✅ *Заказ \\#{escape_md(str(order_id))} оформлен\\!*\n\n"
        f"💰 Сумма к оплате: *{total_price} тг*{bonus_text}\n\n"
        f"📞 Мы позвоним вам для подтверждения времени доставки\\.\n\n"
        f"Спасибо, что выбрали MyKurt\\! 🧀",
        parse_mode="MarkdownV2",
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("🏠 Главное меню", callback_data="back_main")]
        ])
    )
    context.user_data.clear()

    # ── Уведомление админу (async, не блокирует) ──────────────────────
    try:
        order_row = get_order(order_id)
        if order_row:
            admin_text = format_admin_order(*order_row)
            for admin_id in ADMIN_IDS:
                try:
                    await context.bot.send_message(
                        admin_id, admin_text,
                        reply_markup=admin_order_keyboard(order_id)
                    )
                except Exception as e:
                    logger.error(f"Admin notify error (id={admin_id}): {e}")
    except Exception as e:
        logger.error(f"Admin notify block error: {e}", exc_info=True)

    # ── Google Sheets — в отдельном потоке, не блокирует event loop ───
    _oid = order_id
    _uid = user.id
    _uname = user.username or ""
    _filtered = dict(filtered)

    def _sheets_sync():
        try:
            sm = SheetsManager(_CREDS_SOURCE, SPREADSHEET_ID)
            sm.add_order(_oid, _uid, _uname, name,
                         phone, address, _filtered, total_units, bonus, total_price)
            sm.refresh_dashboard(DB_FILE, DATABASE_URL)
            logger.info(f"Sheets обновлены для заказа #{_oid}")
        except Exception as e:
            logger.error(f"Sheets save error (заказ #{_oid}): {e}", exc_info=True)

    asyncio.create_task(asyncio.to_thread(_sheets_sync))


async def show_my_orders(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    user  = update.effective_user
    ph    = _ph()
    conn  = _get_conn()
    c     = conn.cursor()
    c.execute(
        f"SELECT id, items, total_units, bonus_units, total_price, status, created_at "
        f"FROM orders WHERE user_id={ph} ORDER BY created_at DESC LIMIT 5",
        (user.id,)
    )
    rows = c.fetchall()
    conn.close()
    if not rows:
        await safe_edit(
            query,
            "📋 У вас пока нет заказов\\.\n\nСделайте первый — это быстро\\! 🛒",
            parse_mode="MarkdownV2",
            reply_markup=InlineKeyboardMarkup([
                [InlineKeyboardButton("🛒 Сделать заказ", callback_data="new_order")],
                [InlineKeyboardButton("◀️ Назад",         callback_data="back_main")],
            ])
        )
        return

    STATUS_ICON = {
        "новый": "🆕", "принят ✅": "✅", "в пути 🚚": "🚚",
        "выполнен ✔️": "✔️", "отменён ❌": "❌",
    }
    lines = ["📋 *Ваши последние заказы:*\n"]
    for row in rows:
        oid, items_json, total_u, bonus_u, price, status, created = row
        items = json.loads(items_json)
        parts = []
        for k, v in items.items():
            if v > 0:
                parts.append(f"{KURT_TYPES[k]['emoji']} {KURT_TYPES[k]['name']} × {v}")
        icon = STATUS_ICON.get(status, "📦")
        bonus_line = f" \\+ {bonus_u} бонус" if bonus_u > 0 else ""
        lines.append(
            f"*Заказ \\#{oid}* — {escape_md(created[:10])}\n"
            f"{escape_md(', '.join(parts))}\n"
            f"💰 {total_u} шт{bonus_line} \\= *{price} тг*\n"
            f"{icon} {escape_md(status)}\n"
        )
    await safe_edit(
        query,
        "\n".join(lines),
        parse_mode="MarkdownV2",
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("🛒 Новый заказ", callback_data="new_order")],
            [InlineKeyboardButton("◀️ Главное меню", callback_data="back_main")],
        ])
    )


# ──────────────────────────────────────────────
# КОМАНДЫ АДМИНИСТРАТОРА
# ──────────────────────────────────────────────
async def admin_orders(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        return
    orders = get_all_orders()
    if not orders:
        await update.message.reply_text("Заказов пока нет.")
        return
    for row in orders[:10]:
        oid, uid, uname, fname, phone, addr, items_j, tu, bu, price, status, created = row
        items = json.loads(items_j)
        parts = [f"{KURT_TYPES[k]['name']}×{v}" for k, v in items.items() if v > 0]
        text = (
            f"#{oid} [{created[:10]}] — {status}\n"
            f"👤 {fname} (@{uname or '—'}) | 📞 {phone}\n"
            f"📦 {', '.join(parts)}\n"
            f"💰 {price} тг"
        )
        await update.message.reply_text(
            text,
            reply_markup=admin_order_keyboard(oid)
        )


async def admin_stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        return
    conn = _get_conn()
    c    = conn.cursor()
    if DATABASE_URL:
        today_f = "created_at::date = CURRENT_DATE"
        week_f  = "created_at::date >= CURRENT_DATE - INTERVAL '7 days'"
    else:
        today_f = "DATE(created_at) = DATE('now','localtime')"
        week_f  = "DATE(created_at) >= DATE('now','-7 days','localtime')"
    c.execute("SELECT COUNT(*), SUM(total_units), SUM(total_price) FROM orders WHERE status NOT LIKE '%отмен%'")
    row = c.fetchone()
    c.execute(f"SELECT COUNT(*), SUM(total_units), SUM(total_price) FROM orders WHERE {today_f} AND status NOT LIKE '%отмен%'")
    today_row = c.fetchone()
    c.execute(f"SELECT COUNT(*), SUM(total_units), SUM(total_price) FROM orders WHERE {week_f} AND status NOT LIKE '%отмен%'")
    week_row = c.fetchone()
    c.execute("SELECT COUNT(*) FROM users")
    users_count = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM users WHERE total_orders > 1")
    repeat_count = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM orders WHERE status LIKE '%отмен%'")
    cancelled = c.fetchone()[0]
    c.execute("SELECT COUNT(*) FROM orders WHERE status LIKE '%пути%' OR status LIKE '%принят%'")
    in_progress = c.fetchone()[0]
    conn.close()

    count, total_u, total_p = row
    count = count or 0; total_u = total_u or 0; total_p = total_p or 0
    t_ord, t_u, t_p = (today_row[0] or 0, today_row[1] or 0, today_row[2] or 0)
    w_ord, w_u, w_p = (week_row[0] or 0, week_row[1] or 0, week_row[2] or 0)
    avg_check = round(total_p / count) if count else 0
    retention = round(repeat_count / users_count * 100) if users_count else 0
    url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}"
    await update.message.reply_text(
        f"Статистика магазина Курт\n\n"
        f"Клиентов: {users_count} (повторных: {repeat_count} / {retention}%)\n"
        f"Заказов всего: {count} | Отменено: {cancelled}\n"
        f"Всего куртов: {total_u}\n"
        f"Выручка: {total_p:,} тг\n"
        f"Средний чек: {avg_check:,} тг\n\n"
        f"Сегодня: {t_ord} заказов · {t_u} шт · {t_p:,} тг\n"
        f"7 дней: {w_ord} заказов · {w_u} шт · {w_p:,} тг\n"
        f"В работе сейчас: {in_progress}\n\n"
        f"Таблица: {url}"
    )


async def admin_sheets(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        return
    url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}"
    await update.message.reply_text(
        f"Google Sheets — КуртБот\n\n{url}\n\n"
        f"Таблица обновляется при каждом заказе.\n"
        f"Если устарела — /syncsheets"
    )


async def admin_syncsheets(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        return
    msg = await update.message.reply_text("⏳ Синхронизирую с Google Sheets...")
    url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}"

    def _sync():
        sm = SheetsManager(_CREDS_SOURCE, SPREADSHEET_ID)
        sm.rebuild_from_db(DB_FILE, DATABASE_URL)

    try:
        await asyncio.to_thread(_sync)
        await msg.edit_text(f"✅ Готово! Все данные синхронизированы.\n\n{url}")
    except Exception as e:
        logger.error(f"Syncsheets error: {e}", exc_info=True)
        await msg.edit_text(f"❌ Ошибка: {e}")


async def admin_broadcast(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        return
    if not context.args:
        await update.message.reply_text(
            "📢 Рассылка всем клиентам\n\n"
            "Использование:\n/broadcast Ваше сообщение\n\n"
            "Пример:\n/broadcast Привет! Появились новые вкусы курта 🧀"
        )
        return
    text = " ".join(context.args)
    user_ids = get_all_user_ids()
    if not user_ids:
        await update.message.reply_text("Нет пользователей для рассылки.")
        return
    msg = await update.message.reply_text(f"⏳ Отправляю {len(user_ids)} пользователям...")
    sent = 0
    failed = 0
    for uid in user_ids:
        try:
            await context.bot.send_message(chat_id=uid, text=text)
            sent += 1
        except TelegramError:
            failed += 1
    await msg.edit_text(
        f"✅ Рассылка завершена!\n\n"
        f"📤 Отправлено: {sent}\n"
        f"❌ Ошибок (заблокировали бота): {failed}"
    )


async def admin_stock(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        return
    stock = get_all_stock()
    lines = ["📦 Остатки на складе:\n"]
    for key, info in KURT_TYPES.items():
        qty = stock.get(key, -1)
        if qty == -1:
            status = "♾️ без лимита"
        elif qty == 0:
            status = "❌ нет в наличии"
        else:
            status = f"✅ {qty} шт"
        lines.append(f"{info['emoji']} {info['name']}: {status}")
    lines.append("\n✏️ Изменить: /setstock <вид> <количество>")
    lines.append("Примеры:")
    lines.append("  /setstock ромбик 50")
    lines.append("  /setstock микс 0  (закончилось)")
    lines.append("  /setstock обычный -1  (без лимита)")
    await update.message.reply_text("\n".join(lines))


async def admin_setstock(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        return
    if len(context.args) < 2:
        await update.message.reply_text(
            "Использование: /setstock <вид> <количество>\n\n"
            "Примеры:\n"
            "/setstock ромбик 50\n"
            "/setstock микс из всех 0  (закончилось)\n"
            "/setstock обычный -1  (без лимита)\n\n"
            "Список видов: /stock"
        )
        return
    # Последний аргумент — число, всё остальное — название вида
    try:
        qty = int(context.args[-1])
    except ValueError:
        await update.message.reply_text(
            "❌ Последним словом должно быть количество (целое число).\n\n"
            "Пример: /setstock микс из всех 10"
        )
        return
    key_raw = " ".join(context.args[:-1]).lower()
    kurt_key = PHOTO_KEY_ALIASES.get(key_raw)
    if not kurt_key:
        await update.message.reply_text(
            f"❌ Неизвестный вид: {key_raw}\n\n"
            "Используйте русские или английские названия.\n"
            "Список: /stock"
        )
        return
    set_stock(kurt_key, qty)
    info = KURT_TYPES[kurt_key]
    if qty == -1:
        status = "♾️ без лимита"
    elif qty == 0:
        status = "❌ нет в наличии"
    else:
        status = f"✅ {qty} шт"
    await update.message.reply_text(
        f"✅ Обновлено!\n{info['emoji']} {info['name']}: {status}"
    )


async def admin_setphoto(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        return
    keys_list = "\n".join(
        f"  • {info['name'].lower()} или {key}"
        for key, info in KURT_TYPES.items()
    )
    await update.message.reply_text(
        f"📸 Загрузка фото куртов\n\n"
        f"Отправьте фото с подписью — названием вида курта.\n"
        f"Допустимые названия:\n{keys_list}\n\n"
        f"Пример: отправьте фото и в подписи напишите 'ромбик'\n\n"
        f"Текущие фото: /photos"
    )


async def admin_photos(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        return
    loaded = get_all_photos()
    lines = ["📸 Статус фото куртов:\n"]
    for key, info in KURT_TYPES.items():
        status = "✅ загружено" if key in loaded else "❌ нет фото"
        lines.append(f"{info['emoji']} {info['name']}: {status}")
    lines.append("\nДля загрузки: /setphoto")
    await update.message.reply_text("\n".join(lines))


async def photo_upload_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        return
    msg = update.message
    if not msg.photo:
        return
    caption  = (msg.caption or "").strip().lower()
    if not caption:
        await msg.reply_text(
            "⚠️ Укажите в подписи к фото название вида курта.\n"
            "Например: обычный или ромбик\n\n"
            "Список названий: /setphoto"
        )
        return
    kurt_key = PHOTO_KEY_ALIASES.get(caption)
    if not kurt_key:
        await msg.reply_text(
            f"❌ Неизвестный вид: {caption}\n"
            f"Допустимые названия: /setphoto"
        )
        return
    file_id = msg.photo[-1].file_id
    set_photo(kurt_key, file_id)
    info = KURT_TYPES[kurt_key]
    await msg.reply_text(
        f"✅ Фото для {info['emoji']} {info['name']} сохранено!\n"
        f"Теперь оно будет показываться в каталоге и при выборе вида."
    )


# ──────────────────────────────────────────────
# ЗАПУСК
# ──────────────────────────────────────────────
async def _run_webhook(app: Application) -> None:
    from aiohttp import web

    async def handle_update(request: web.Request) -> web.Response:
        try:
            data = await request.json()
            update = Update.de_json(data, app.bot)
            await app.process_update(update)
        except Exception as e:
            logger.error(f"Webhook error: {e}")
        return web.Response(text="OK")

    async def handle_health(request: web.Request) -> web.Response:
        return web.Response(text="OK")

    await app.initialize()
    await app.bot.set_webhook(
        url=f"{WEBHOOK_URL}/{WEBHOOK_PATH}",
        allowed_updates=Update.ALL_TYPES,
    )
    await app.start()
    await setup_commands(app)

    web_app = web.Application()
    web_app.router.add_get("/", handle_health)
    web_app.router.add_get("/health", handle_health)
    web_app.router.add_post(f"/{WEBHOOK_PATH}", handle_update)

    runner = web.AppRunner(web_app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", PORT)
    await site.start()
    logger.info(f"Webhook запущен: {WEBHOOK_URL}/{WEBHOOK_PATH} (порт {PORT})")

    stop_event = asyncio.Event()
    loop = asyncio.get_running_loop()
    for sig in (signal.SIGTERM, signal.SIGINT):
        try:
            loop.add_signal_handler(sig, stop_event.set)
        except NotImplementedError:
            pass

    await stop_event.wait()
    await app.stop()
    await app.shutdown()
    await runner.cleanup()


async def setup_commands(app: Application) -> None:
    """Устанавливает подсказки команд для пользователей и администраторов."""
    user_commands = [
        BotCommand("start", "🏠 Главное меню"),
    ]
    admin_commands = [
        BotCommand("start",      "🏠 Главное меню"),
        BotCommand("orders",     "📋 Последние заказы"),
        BotCommand("stats",      "📊 Статистика продаж"),
        BotCommand("broadcast",  "📢 Рассылка всем клиентам"),
        BotCommand("stock",      "📦 Остатки на складе"),
        BotCommand("setstock",   "✏️ Изменить остаток"),
        BotCommand("sheets",     "📄 Ссылка на Google Sheets"),
        BotCommand("syncsheets", "🔄 Синхронизировать таблицу"),
        BotCommand("setphoto",   "📸 Загрузить фото курта"),
        BotCommand("photos",     "🖼 Статус фото куртов"),
    ]

    try:
        await app.bot.set_my_commands(user_commands, scope=BotCommandScopeAllPrivateChats())
        logger.info("✅ Команды для пользователей установлены")
    except Exception as e:
        logger.error(f"❌ Команды для пользователей: {e}", exc_info=True)

    for admin_id in ADMIN_IDS:
        try:
            await app.bot.set_my_commands(
                admin_commands,
                scope=BotCommandScopeChat(chat_id=admin_id)
            )
            logger.info(f"✅ Команды для admin {admin_id} установлены")
        except Exception as e:
            logger.error(f"❌ Команды для admin {admin_id}: {e}", exc_info=True)

    logger.info("✅ setup_commands завершён")


def main():
    global _CREDS_SOURCE
    init_db()
    _CREDS_SOURCE = json.loads(GOOGLE_CREDENTIALS_JSON) if GOOGLE_CREDENTIALS_JSON else CREDENTIALS_FILE

    if SPREADSHEET_ID:
        try:
            sm = SheetsManager(_CREDS_SOURCE, SPREADSHEET_ID)
            sm.init_sheet()
            logger.info("✅ Google Sheets подключены")
        except Exception as e:
            logger.error(f"⚠️ Google Sheets ошибка: {e}", exc_info=True)
            logger.warning("Бот работает без Sheets.")

    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start",      start))
    app.add_handler(CommandHandler("orders",     admin_orders))
    app.add_handler(CommandHandler("stats",      admin_stats))
    app.add_handler(CommandHandler("sheets",     admin_sheets))
    app.add_handler(CommandHandler("syncsheets", admin_syncsheets))
    app.add_handler(CommandHandler("broadcast",  admin_broadcast))
    app.add_handler(CommandHandler("stock",      admin_stock))
    app.add_handler(CommandHandler("setstock",   admin_setstock))
    app.add_handler(CommandHandler("setphoto",   admin_setphoto))
    app.add_handler(CommandHandler("photos",     admin_photos))
    app.add_handler(CallbackQueryHandler(button_handler))
    app.add_handler(MessageHandler(filters.PHOTO, photo_upload_handler))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, message_handler))

    logger.info("🤖 Бот запущен...")
    if WEBHOOK_URL:
        asyncio.run(_run_webhook(app))
    else:
        app.run_polling()


if __name__ == "__main__":
    main()
