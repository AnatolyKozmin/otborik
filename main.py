import json
from aiogram import Bot, Dispatcher, Router, types
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardMarkup, InlineKeyboardButton, InputFile
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.filters.state import StateFilter
from aiogram.filters import Command
import datetime
import os
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv, find_dotenv
load_dotenv()
API_TOKEN = os.getenv('BOT_TOKEN')
ADMIN_ID = 922109605

SLOTS_FILE = 'slots.json'

bot = Bot(token=API_TOKEN)
storage = MemoryStorage()
dp = Dispatcher(storage=storage)
router = Router()

DIRECTIONS = ['ЦТ', 'Фото', 'СМИ', 'Дизайн', 'F&U prod.']

class Form(StatesGroup):
    name = State()
    vk = State()
    direction = State()
    date = State()
    time = State()

# Файл для хранения опубликованного сообщения с направлениями
PUBLISHED_FILE = 'published.json'
LOG_FILE = 'bot.log'

# Загрузка слотов

def load_slots():
    with open(SLOTS_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_slots(data):
    with open(SLOTS_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_published():
    if not os.path.exists(PUBLISHED_FILE):
        return {}
    with open(PUBLISHED_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_published(data):
    with open(PUBLISHED_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def log_action(text: str):
    ts = datetime.datetime.now().isoformat()
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(f"[{ts}] {text}\n")


def export_registrations_csv():
    slots = load_slots()
    regs = slots.get('registrations', [])
    import io, csv
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(['user_id', 'full_name', 'vk_link', 'direction', 'date', 'time', 'registered_at'])
    for r in regs:
        writer.writerow([r.get('user_id'), r.get('full_name'), r.get('vk_link'), r.get('direction'), r.get('date'), r.get('time'), r.get('registered_at')])
    buf.seek(0)
    return buf


def build_sheets_service():
    # заглушка — при использовании gspread функция не нужна, но оставлена для совместимости
    return None


@router.message(Command('get_slots'))
async def cmd_get_slots(message: types.Message):
    if message.from_user.id != ADMIN_ID:
        await message.answer('Только админ может обновлять слоты из Google Sheets')
        return
    await message.answer('Начинаю парсинг Google Sheets...')
    # Параметры: ожидаем, что админ предварительно подставит ID таблицы в переменную окружения SHEET_ID
    sheet_id = os.getenv('SHEET_ID')
    if not sheet_id:
        await message.answer('Переменная окружения SHEET_ID не задана.')
        return
    # Открываем gspread через service account
    try:
        # Prefer gspread helper which configures scopes automatically from service account file
        try:
            gc = gspread.service_account(filename='credentials.json')
        except Exception:
            # Fallback: explicitly set scopes for google oauth credentials
            creds = Credentials.from_service_account_file(
                'credentials.json',
                scopes=['https://www.googleapis.com/auth/spreadsheets',
                        'https://www.googleapis.com/auth/drive']
            )
            gc = gspread.authorize(creds)
        sh = gc.open_by_key(sheet_id)
        # Выбираем конкретный лист: сначала пробуем имя из SHEET_NAME, затем индекс из SHEET_INDEX (0-based).
        # Если ни одна переменная не задана или открытие не удалось — используем первый лист (sheet1).
        sheet_name = os.getenv('SHEET_NAME')
        sheet_index = os.getenv('SHEET_INDEX')
        if sheet_name:
            try:
                ws = sh.worksheet(sheet_name)
            except Exception as e:
                await message.answer(f'Не удалось открыть лист с именем "{sheet_name}": {e}. Будет использован первый лист.')
                ws = sh.sheet1
        elif sheet_index:
            try:
                idx = int(sheet_index)
                ws = sh.get_worksheet(idx)
            except Exception as e:
                await message.answer(f'Не удалось открыть лист по индексу {sheet_index}: {e}. Будет использован первый лист.')
                ws = sh.sheet1
        else:
            ws = sh.sheet1
    except Exception as e:
        await message.answer(f'Ошибка подключения к Google Sheets: {e}')
        return

    try:
        # Собираем все листы, названия которых совпадают с нашими направлениями.
        all_ws = sh.worksheets()
        direction_sheets = [w for w in all_ws if w.title in DIRECTIONS]
        if not direction_sheets:
            await message.answer(f'Не найдено листов с названиями направлений. Ожидаемые имена: {DIRECTIONS}')
            return

        # Берём даты и времена с первого листа-направления (ожидается одинаковая структура на всех листах)
        first_ws = direction_sheets[0]
        dates = (first_ws.get('B1:G1') or [['']])[0]
        times = [r[0] for r in (first_ws.get('A2:A13') or [['']])]

    except Exception as e:
        await message.answer(f'Ошибка чтения диапазонов/листов: {e}')
        return

    # Preserve existing registrations if possible
    old_slots = {}
    try:
        old_slots = load_slots()
    except Exception:
        old_slots = {'slots': {}, 'registrations': []}

    # Инициализация структуры слотов для всех дат/времён и всех направлений
    new_slots = {'slots': {}, 'registrations': []}
    for date in dates:
        date_key = date.strip()
        new_slots['slots'][date_key] = {}
        for time_slot in times:
            new_slots['slots'][date_key][time_slot] = {d: None for d in DIRECTIONS}

    # Проходим по каждому листу-направлению и заполняем blocked/available по B2:G13
    for ws_dir in direction_sheets:
        direction = ws_dir.title
        try:
            values = ws_dir.get('B2:G13') or []
        except Exception:
            values = []
        for i, date in enumerate(dates):
            date_key = date.strip()
            for j, time_slot in enumerate(times):
                cell_val = ''
                try:
                    cell_val = values[j][i]
                except Exception:
                    cell_val = ''
                # Если в ячейке явно указано 'могу' (регистр игнорируем) — считаем доступным
                # Во всех остальных случаях (пусто или 'не могу' и т.д.) — помечаем как blocked
                if isinstance(cell_val, str) and cell_val.strip().lower() == 'могу':
                    # оставляем None — доступно
                    pass
                else:
                    # помечаем как blocked
                    if date_key in new_slots['slots'] and time_slot in new_slots['slots'][date_key]:
                        new_slots['slots'][date_key][time_slot][direction] = 'blocked'

    # Подробный отчёт по каждому листу, который мы парсим
    report_lines = []
    rows_count = len(times)  # ожидаем 12
    cols_count = len(dates)  # ожидаем 6
    total_cells_expected = rows_count * cols_count
    for ws_dir in direction_sheets:
        try:
            vals = ws_dir.get('B2:G13') or []
            # Считаем только ячейки, где явно встречается слово 'могу'
            mogu_cells = []
            for r_idx, row in enumerate(vals):
                for c_idx, cell in enumerate(row):
                    try:
                        if isinstance(cell, str) and cell.strip().lower() == 'могу':
                            mogu_cells.append((r_idx + 2, c_idx + 2))  # координаты в таблице (начиная с 1)
                    except Exception:
                        pass
            mogu_count = len(mogu_cells)
            report = f"{ws_dir.title}: прочитано {len(vals)} строк, всего ячеек {total_cells_expected}, 'могу' = {mogu_count}"
            if mogu_count:
                # включаем пару примеров координат для отладки
                sample = ', '.join([f"({r},{c})" for r, c in mogu_cells[:6]])
                report += f", примеры 'могу' в ячейках: {sample}"
            report_lines.append(report)
        except Exception as e:
            report_lines.append(f"{ws_dir.title}: ошибка чтения ({e})")
    try:
        await message.answer('Отчёт парсинга: ' + '; '.join(report_lines))
    except Exception:
        pass

    # Try to reattach previous registrations to the new_slots when date/time still exists
    for reg in old_slots.get('registrations', []):
        d = reg.get('date')
        t = reg.get('time')
        dirn = reg.get('direction')
        if d in new_slots['slots'] and t in new_slots['slots'][d] and new_slots['slots'][d][t].get(dirn) is None:
            # slot exists and free in new layout -> keep registration
            new_slots['slots'][d][t][dirn] = reg.get('user_id')
            new_slots['registrations'].append(reg)
        else:
            # otherwise drop the registration (admin will need to contact user)
            pass

    # Backup old slots.json
    try:
        import shutil
        shutil.copyfile(SLOTS_FILE, SLOTS_FILE + '.bak')
    except Exception:
        pass

    save_slots(new_slots)
    await message.answer('Слоты обновлены из Google Sheets и файл slots.json перезаписан.')
    await update_published_message()


@router.message(Command('list_sheets'))
async def cmd_list_sheets(message: types.Message):
    if message.from_user.id != ADMIN_ID:
        await message.answer('Только админ может посмотреть список листов.')
        return
    sheet_id = os.getenv('SHEET_ID')
    if not sheet_id:
        await message.answer('Переменная окружения SHEET_ID не задана.')
        return
    try:
        try:
            gc = gspread.service_account(filename='credentials.json')
        except Exception:
            creds = Credentials.from_service_account_file(
                'credentials.json',
                scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
            )
            gc = gspread.authorize(creds)
        sh = gc.open_by_key(sheet_id)
        sheets = sh.worksheets()
        if not sheets:
            await message.answer('В таблице нет листов.')
            return
        lines = [f"{i}: {ws.title}" for i, ws in enumerate(sheets)]
        # Если сообщение слишком длинное, можно отправить как документ. Пока отправляем текстом.
        await message.answer('Листы таблицы:\n' + '\n'.join(lines))
    except Exception as e:
        await message.answer(f'Ошибка при получении списка листов: {e}')


def direction_has_free_slots(direction, slots):
    # проверяем, есть ли хотя бы один свободный слот для направления
    for date, times in slots.get('slots', {}).items():
        for t, dirs in times.items():
            if dirs.get(direction) is None:
                return True
    return False


def parse_slot_datetime(date_str: str, time_str: str) -> datetime.datetime:
    """Parse date and time strings into datetime. Strips weekday notes like '(сб)'.
    Expected date format: 'DD.MM.YYYY' possibly with trailing '(...')
    Expected time format: 'HH:MM' or similar; if time contains extra text, take first 5 chars.
    """
    # remove parenthesis with weekday e.g. '11.10.2025(сб)' -> '11.10.2025'
    import re
    d = re.sub(r"\s*\(.*\)$", "", date_str).strip()
    t = time_str.strip()[:5]
    return datetime.datetime.strptime(f"{d} {t}", "%d.%m.%Y %H:%M")


# Пояснение по хранилищам:
# - Слоты и регистрации хранятся в файле `slots.json`. Там структуру вы можете редактировать вручную.
# - Информация о опубликованном сообщении (chat_id и message_id) хранится в `published.json`.
# - MemoryStorage (aiogram.fsm.storage.memory.MemoryStorage) хранит временные состояния пользователей в памяти процесса.
#   Это включает: текущие значения состояний FSM для каждого пользователя (какий шаг заполнения формы),
#   и временные данные (data), которые мы сохраняем через `state.update_data()` — например, name, vk, direction, date.
#   MemoryStorage не предназначен для долгосрочного хранения: при перезапуске бота все состояния будут утеряны.


async def update_published_message():
    pub = load_published()
    if not pub:
        return
    chat_id = pub.get('chat_id')
    message_id = pub.get('message_id')
    if not chat_id or not message_id:
        return
    slots = load_slots()
    kb_buttons = [[InlineKeyboardButton(text=d, callback_data=f'dir:{d}')] for d in DIRECTIONS if direction_has_free_slots(d, slots)]
    kb = InlineKeyboardMarkup(inline_keyboard=kb_buttons)
    try:
        # Редактируем текст и клавиатуру
        await bot.edit_message_text(chat_id=chat_id, message_id=message_id, text='Выберите направление:', reply_markup=kb)
    except Exception:
        try:
            await bot.edit_message_reply_markup(chat_id=chat_id, message_id=message_id, reply_markup=kb)
        except Exception:
            pass


# Старт
@router.message(Command('start'))
async def cmd_start(message: types.Message, state: FSMContext):
    # Инструкция для пользователя
    text = (
        'Здравствуйте! Чтобы записаться на собеседование, пройдите 4 шага:\n'
        '1) Введите ваше ФИО.\n'
        '2) Пришлите ссылку на ваш VK.\n'
        '3) Выберите направление \n'
        '4) Выберите дату и время из доступных слотов.\n\n'
        'Команды:\n'
        '/my — показать вашу текущую запись.\n'
        '/cancel — отменить запись (не позднее чем за 24 часа).\n\n'
        'Начнём: введите ваше Имя и Фамилию:'
    )
    # set initial FSM state for this user (aiogram v3)
    await state.set_state(Form.name)
    await message.answer(text)

@router.message(StateFilter(Form.name))
async def process_name(message: types.Message, state: FSMContext):
    await state.update_data(name=message.text)
    # move to next state: vk
    await state.set_state(Form.vk)
    await message.answer('Пришлите ссылку на ваш VK:')

@router.message(StateFilter(Form.vk))
async def process_vk(message: types.Message, state: FSMContext):
    await state.update_data(vk=message.text)
    # Кнопки направлений
    # Предлагаем варианты: либо через reply-клавиатуру, либо через опубликованное сообщение (inline)
    kb = ReplyKeyboardMarkup(keyboard=[[KeyboardButton(text=d)] for d in DIRECTIONS], resize_keyboard=True)
    # move to next state: direction
    await state.set_state(Form.direction)
    await message.answer('Выберите направление (или нажмите кнопку в опубликованном сообщении):', reply_markup=kb)


@router.message(Command('directions'))
async def cmd_directions(message: types.Message):
    # Показываем публицированный набор направлений, если он есть, иначе отправляем кнопки
    pub = load_published()
    if pub:
        try:
            await bot.forward_message(chat_id=message.chat.id, from_chat_id=pub['chat_id'], message_id=pub['message_id'])
            return
        except Exception:
            pass
    slots = load_slots()
    kb_buttons = [[InlineKeyboardButton(text=d, callback_data=f'dir:{d}')] for d in DIRECTIONS if direction_has_free_slots(d, slots)]
    kb = InlineKeyboardMarkup(inline_keyboard=kb_buttons)
    await message.answer('Выберите направление:', reply_markup=kb)


# Admin: публикуем сообщение с кнопками направлений и фотографией
@router.message(Command('publish'))
async def cmd_publish(message: types.Message):
    if message.from_user.id != ADMIN_ID:
        await message.answer('Только админ может публиковать меню направлений.')
        return
    slots = load_slots()
    kb_buttons = [[InlineKeyboardButton(text=d, callback_data=f'dir:{d}')] for d in DIRECTIONS if direction_has_free_slots(d, slots)]
    kb = InlineKeyboardMarkup(inline_keyboard=kb_buttons)
    sent = await message.answer('Выберите направление:', reply_markup=kb)
    save_published({'chat_id': sent.chat.id, 'message_id': sent.message_id})
    # Добавляем кнопку Export для админа
    try:
        kb2 = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text='Export CSV', callback_data='export:csv')]])
        await bot.edit_message_reply_markup(chat_id=sent.chat.id, message_id=sent.message_id, reply_markup=kb)
    except Exception:
        pass
    await message.answer('Опубликовано.')
    log_action(f'published by {message.from_user.id}')


# Выбор направления
@router.message(StateFilter(Form.direction))
async def process_direction(message: types.Message, state: FSMContext):
    if message.text not in DIRECTIONS:
        await message.answer('Пожалуйста, выберите направление с помощью кнопок.')
        return
    await state.update_data(direction=message.text)
    # Кнопки с датами (только с доступными слотами)
    slots = load_slots()
    available_dates = []
    for date, times in slots['slots'].items():
        for t, dirs in times.items():
            if dirs[message.text] is None:
                available_dates.append(date)
                break
    if not available_dates:
        await message.answer('Нет доступных дат для этого направления.')
        return
    kb = ReplyKeyboardMarkup(keyboard=[[KeyboardButton(text=d)] for d in available_dates], resize_keyboard=True)
    # Отправляем фото и сообщение с выбором даты
    # Переводим состояние в Form.date (aiogram v3)
    await state.set_state(Form.date)
    await message.answer('Выберите дату:', reply_markup=kb)


# Обработка нажатия inline-кнопки с направлением
@router.callback_query(lambda c: c.data and c.data.startswith('dir:'))
async def callback_dir(callback: types.CallbackQuery, state: FSMContext):
    direction = callback.data.split(':', 1)[1]
    await bot.answer_callback_query(callback.id)
    # Сохраняем направление в state и просим ФИО (если ещё не было)
    # Если пользователь уже начал ввод ФИО или VK — не сбрасываем состояние, просто сохраняем направление
    st = state
    data = await st.get_data()
    # Если пользователь ещё не начал через /start — просим сначала нажать /start
    if not data.get('name'):
        await bot.answer_callback_query(callback.id, 'Пожалуйста, сначала нажмите /start в боте и введите ФИО, затем вернитесь и выберите направление.')
        return
    # Сохраняем направление в state
    await st.update_data(direction=direction)
    # Если VK ещё нет — просим ссылку
    if not data.get('vk'):
        await st.set_state(Form.vk)
        await bot.send_message(callback.from_user.id, f'Вы выбрали направление: {direction}.\nПожалуйста, пришлите ссылку на ваш VK:')
        return
    # Иначе продолжаем процесс выбора даты (имитируем переход)
    await st.set_state(Form.date)
    # Показываем доступные даты
    slots = load_slots()
    available_dates = []
    for date, times in slots['slots'].items():
        for t, dirs in times.items():
            if dirs.get(direction) is None:
                available_dates.append(date)
                break
    if not available_dates:
        await bot.send_message(callback.from_user.id, 'Нет доступных дат для этого направления.')
        return
    kb = ReplyKeyboardMarkup(keyboard=[[KeyboardButton(text=d)] for d in available_dates], resize_keyboard=True)
    await bot.send_message(callback.from_user.id, 'Выберите дату:', reply_markup=kb)


@router.callback_query(lambda c: c.data == 'export:csv')
async def callback_export(callback: types.CallbackQuery):
    if callback.from_user.id != ADMIN_ID:
        await bot.answer_callback_query(callback.id, 'Только админ может экспортировать')
        return
    await bot.answer_callback_query(callback.id, 'Генерирую CSV...')
    buf = export_registrations_csv()
    await bot.send_document(ADMIN_ID, ("registrations.csv", buf.getvalue().encode('utf-8')))
    log_action(f'export_csv by {callback.from_user.id}')


@router.message(Command('export'))
async def cmd_export(message: types.Message):
    if message.from_user.id != ADMIN_ID:
        await message.answer('Только админ может экспортировать')
        return
    buf = export_registrations_csv()
    await message.answer_document(("registrations.csv", buf.getvalue().encode('utf-8')))
    log_action(f'export_csv by {message.from_user.id}')


@router.message(Command('my'))
async def cmd_my(message: types.Message):
    slots = load_slots()
    user_id = message.from_user.id
    for reg in slots.get('registrations', []):
        if reg['user_id'] == user_id:
            text = (
                f"Ваша запись:\nФИО: {reg['full_name']}\nVK: {reg['vk_link']}\nНаправление: {reg['direction']}\nДата: {reg['date']}\nВремя: {reg['time']}"
            )
            await message.answer(text)
            return
    await message.answer('У вас нет активной записи.')


# Выбор даты
@router.message(StateFilter(Form.date))
async def process_date(message: types.Message, state: FSMContext):
    slots = load_slots()
    data = await state.get_data()
    direction = data['direction']
    date = message.text
    if date not in slots['slots']:
        await message.answer('Пожалуйста, выберите дату из списка.')
        return
    # Кнопки с доступным временем
    available_times = []
    for t, dirs in slots['slots'][date].items():
        if dirs[direction] is None:
            # Проверка на 12 часов до слота
            slot_dt = parse_slot_datetime(date, t)
            if slot_dt - datetime.datetime.now() > datetime.timedelta(hours=12):
                available_times.append(t)
    if not available_times:
        await message.answer('Нет доступного времени на эту дату.')
        return
    kb = ReplyKeyboardMarkup(keyboard=[[KeyboardButton(text=t)] for t in available_times], resize_keyboard=True)
    await state.update_data(date=date)
    # Переводим состояние в Form.time (aiogram v3)
    await state.set_state(Form.time)
    await message.answer('Выберите время:', reply_markup=kb)

# Выбор времени и запись
@router.message(StateFilter(Form.time))
async def process_time(message: types.Message, state: FSMContext):
    slots = load_slots()
    data = await state.get_data()
    direction = data['direction']
    date = data['date']
    time = message.text
    if time not in slots['slots'][date] or slots['slots'][date][time][direction] is not None:
        await message.answer('Это время уже занято или неверно выбрано.')
        return
    # Проверка на 12 часов до слота
    slot_dt = parse_slot_datetime(date, time)
    if slot_dt - datetime.datetime.now() < datetime.timedelta(hours=12):
        await message.answer('Записаться можно не позднее чем за 12 часов до собеседования.')
        return
    # Запись
    user_id = message.from_user.id
    full_name = data['name']
    vk_link = data['vk']
    reg = {
        "user_id": user_id,
        "full_name": full_name,
        "vk_link": vk_link,
        "direction": direction,
        "date": date,
        "time": time,
        "registered_at": datetime.datetime.now().isoformat()
    }
    slots['slots'][date][time][direction] = user_id
    slots['registrations'].append(reg)
    save_slots(slots)
    # Уведомление админу
    text = f"Новая запись!\nФИО: {full_name}\nVK: {vk_link}\nНаправление: {direction}\nДата: {date}\nВремя: {time}\nTG: @{message.from_user.username} ({user_id})"
    await bot.send_message(ADMIN_ID, text)
    log_action(f'registration: {reg}')
    await message.answer('Вы успешно записаны! Если хотите отменить запись, напишите /cancel')
    await state.clear()


# Отмена записи
@router.message(Command('cancel'))
async def cancel_registration(message: types.Message):
    slots = load_slots()
    user_id = message.from_user.id
    found = None
    for reg in slots['registrations']:
        if reg['user_id'] == user_id:
            found = reg
            break
    if not found:
        await message.answer('У вас нет активной записи.')
        return
    # Проверка ограничения 24 часа
    slot_dt = parse_slot_datetime(found['date'], found['time'])
    if slot_dt - datetime.datetime.now() < datetime.timedelta(hours=24):
        await message.answer('Отменить запись можно не позднее чем за 24 часа до собеседования. Если нужно отменить позже — напишите в группу.')
        return
    # Удаляем запись
    direction = found['direction']
    date = found['date']
    time = found['time']
    if slots['slots'][date][time][direction] == user_id:
        slots['slots'][date][time][direction] = None
    slots['registrations'] = [r for r in slots['registrations'] if r['user_id'] != user_id]
    save_slots(slots)
    # Уведомление админу об отмене
    admin_text = (
        f"Отмена записи:\nФИО: {found.get('full_name')}\nVK: {found.get('vk_link')}\n"
        f"Направление: {direction}\nДата: {date}\nВремя: {time}\nTG: @{message.from_user.username} ({user_id})"
    )
    await bot.send_message(ADMIN_ID, admin_text)
    log_action(f'cancellation: {found}')
    await message.answer('Ваша запись успешно отменена.')


if __name__ == '__main__':
    import asyncio, logging
    logging.basicConfig(level=logging.INFO)
    # Регистрируем роутер и запускаем polling
    dp.include_router(router)
    asyncio.run(dp.start_polling(bot))