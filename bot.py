"""Telegram-бот: запись ников в Google-таблицы по ссылкам учебных заведений.

Схема:
  * у каждого человека (например, Умида) своя Google-таблица;
  * в таблице — отдельный лист на каждое учебное заведение (УЗ);
  * у каждого УЗ своя ссылка вида https://t.me/<бот>?start=umid__school1;
  * при переходе по ссылке ник сразу записывается в лист этого УЗ,
    ещё ДО того, как пользователь поделился номером;
  * когда пользователь делится номером, номер дописывается в ту же строку.
"""

import asyncio
import json
import logging
import os
import re
import threading
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import gspread
from aiogram import Bot, Dispatcher, F
from aiogram.client.default import DefaultBotProperties
from aiogram.filters import Command, CommandObject, CommandStart
from aiogram.types import (
    KeyboardButton,
    Message,
    ReplyKeyboardMarkup,
    ReplyKeyboardRemove,
    User,
)

BOT_TOKEN = os.environ["BOT_TOKEN"]
CONFIG_PATH = os.getenv("CONFIG_PATH", "config.json")
GOOGLE_CREDENTIALS = os.getenv("GOOGLE_CREDENTIALS", "credentials.json")
STATE_PATH = os.getenv("STATE_PATH", "state.json")
TIMEZONE = ZoneInfo(os.getenv("TIMEZONE", "Asia/Tashkent"))
ADMIN_IDS = {int(x) for x in os.getenv("ADMIN_IDS", "").replace(" ", "").split(",") if x}

SEP = "__"  # разделитель в ссылке: <человек>__<УЗ>
HEADER = ["Дата перехода", "Telegram ID", "Ник", "Имя", "Телефон", "Дата номера"]
COL_ID, COL_NICK, COL_PHONE = 2, 3, 5

log = logging.getLogger("bot")


# ---------------------------------------------------------------- конфиг

def load_config(path: str) -> dict:
    with open(path, encoding="utf-8") as f:
        persons = json.load(f)["persons"]
    key_re = re.compile(r"^[A-Za-z0-9_-]+$")
    for p_key, person in persons.items():
        for i_key in person["institutions"]:
            payload = f"{p_key}{SEP}{i_key}"
            if not key_re.match(payload) or SEP in p_key or SEP in i_key or len(payload) > 64:
                raise ValueError(
                    f"Недопустимый ключ «{payload}»: только латиница, цифры, _ и -, "
                    f"без «{SEP}» внутри ключа, до 64 символов вместе"
                )
    return persons


def sheet_title(name: str) -> str:
    # Google запрещает в названиях листов символы [ ] : * ? / \
    return re.sub(r"[\[\]:*?/\\]", " ", name).strip()[:100] or "Без названия"


def now() -> str:
    return datetime.now(TIMEZONE).strftime("%d.%m.%Y %H:%M")


def nick(user: User) -> str:
    return f"@{user.username}" if user.username else "—"


# ---------------------------------------------------------------- таблицы

class Sheets:
    def __init__(self, creds_path: str, persons: dict):
        self.gc = gspread.service_account(filename=creds_path)
        self.persons = persons
        self._books: dict[str, gspread.Spreadsheet] = {}
        self._sheets: dict[tuple[str, str], gspread.Worksheet] = {}
        self._lock = threading.Lock()

    def _worksheet(self, person: str, inst: str) -> gspread.Worksheet:
        key = (person, inst)
        if key in self._sheets:
            return self._sheets[key]
        cfg = self.persons[person]
        book = self._books.get(person)
        if book is None:
            book = self._books[person] = self.gc.open_by_key(cfg["spreadsheet_id"])
        title = sheet_title(cfg["institutions"][inst])
        try:
            ws = book.worksheet(title)
        except gspread.WorksheetNotFound:
            ws = book.add_worksheet(title=title, rows=1000, cols=len(HEADER))
            ws.append_row(HEADER)
            ws.freeze(rows=1)
        self._sheets[key] = ws
        return ws

    def ensure_all(self) -> None:
        """Создаёт листы для всех УЗ из конфига (вызывается при старте)."""
        with self._lock:
            for person, cfg in self.persons.items():
                for inst in cfg["institutions"]:
                    self._worksheet(person, inst)

    def record(self, person: str, inst: str, user: User, phone: str | None = None) -> None:
        """Добавляет пользователя в лист УЗ или обновляет его строку."""
        with self._lock:
            ws = self._worksheet(person, inst)
            ids = ws.col_values(COL_ID)
            uid = str(user.id)
            if uid in ids:
                row = ids.index(uid) + 1
                ws.update(range_name=f"C{row}:D{row}", values=[[nick(user), user.full_name]])
                if phone:
                    ws.update(range_name=f"E{row}:F{row}", values=[[phone, now()]])
            else:
                ws.append_row(
                    [now(), uid, nick(user), user.full_name, phone or "", now() if phone else ""],
                    value_input_option="RAW",
                )


# ------------------------------------------- откуда пришёл пользователь

class State:
    """Помнит, по какой ссылке пришёл пользователь (переживает перезапуск)."""

    def __init__(self, path: str):
        self.path = Path(path)
        try:
            self.data: dict[str, list[str]] = json.loads(self.path.read_text("utf-8"))
        except (FileNotFoundError, json.JSONDecodeError):
            self.data = {}

    def get(self, user_id: int) -> tuple[str, str] | None:
        v = self.data.get(str(user_id))
        return (v[0], v[1]) if v else None

    def set(self, user_id: int, person: str, inst: str) -> None:
        self.data[str(user_id)] = [person, inst]
        tmp = self.path.with_suffix(".tmp")
        tmp.write_text(json.dumps(self.data, ensure_ascii=False), "utf-8")
        tmp.replace(self.path)


# ---------------------------------------------------------------- бот

persons = load_config(CONFIG_PATH)
sheets = Sheets(GOOGLE_CREDENTIALS, persons)
state = State(STATE_PATH)
dp = Dispatcher()

PHONE_KB = ReplyKeyboardMarkup(
    keyboard=[[KeyboardButton(text="📱 Поделиться номером", request_contact=True)]],
    resize_keyboard=True,
    one_time_keyboard=True,
)


def parse_payload(payload: str | None) -> tuple[str, str] | None:
    if not payload or SEP not in payload:
        return None
    person, inst = payload.split(SEP, 1)
    if person in persons and inst in persons[person]["institutions"]:
        return person, inst
    return None


@dp.message(CommandStart())
async def on_start(message: Message, command: CommandObject) -> None:
    user = message.from_user
    source = parse_payload(command.args) or state.get(user.id)
    if source is None:
        await message.answer(
            "Пожалуйста, перейдите по ссылке своего учебного заведения, "
            "которую вам отправили."
        )
        return

    person, inst = source
    state.set(user.id, person, inst)
    try:
        # Ник записывается сразу, до того как человек поделился номером
        await asyncio.to_thread(sheets.record, person, inst, user)
    except Exception:
        log.exception("Не удалось записать ник %s в таблицу", user.id)

    await message.answer(
        "Здравствуйте! Чтобы продолжить, поделитесь, пожалуйста, своим номером "
        "телефона — нажмите кнопку ниже.",
        reply_markup=PHONE_KB,
    )


@dp.message(F.contact)
async def on_contact(message: Message) -> None:
    user = message.from_user
    contact = message.contact
    if contact.user_id != user.id:
        await message.answer("Отправьте, пожалуйста, свой номер кнопкой ниже.", reply_markup=PHONE_KB)
        return

    source = state.get(user.id)
    if source is None:
        await message.answer(
            "Не удалось определить учебное заведение. Перейдите, пожалуйста, "
            "по ссылке ещё раз.",
            reply_markup=ReplyKeyboardRemove(),
        )
        return

    phone = contact.phone_number
    if not phone.startswith("+"):
        phone = "+" + phone
    try:
        await asyncio.to_thread(sheets.record, *source, user, phone)
    except Exception:
        log.exception("Не удалось записать номер %s в таблицу", user.id)
        await message.answer("Произошла ошибка, попробуйте ещё раз чуть позже.")
        return

    await message.answer("Спасибо, номер получен! ✅", reply_markup=ReplyKeyboardRemove())


@dp.message(Command("links"))
async def on_links(message: Message, bot: Bot) -> None:
    """Админ получает ссылки всех людей, человек из конфига — только свои."""
    uid = message.from_user.id
    is_admin = uid in ADMIN_IDS
    visible = {
        p: cfg for p, cfg in persons.items()
        if is_admin or uid in cfg.get("telegram_ids", [])
    }
    if not visible:
        return
    me = await bot.get_me()
    for person, cfg in visible.items():
        lines = [f"<b>{cfg['title']}</b>"]
        for inst, title in cfg["institutions"].items():
            lines.append(f"{title}:\nhttps://t.me/{me.username}?start={person}{SEP}{inst}")
        await message.answer("\n\n".join(lines), disable_web_page_preview=True)


async def main() -> None:
    logging.basicConfig(level=logging.INFO)
    await asyncio.to_thread(sheets.ensure_all)
    bot = Bot(BOT_TOKEN, default=DefaultBotProperties(parse_mode="HTML"))
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
