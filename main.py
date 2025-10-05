import json
from aiogram import Bot, Dispatcher, types
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardMarkup, InlineKeyboardButton, InputFile
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram import executor
import datetime
import os
import gspread
from google.oauth2.service_account import Credentials

API_TOKEN = os.getenv('BOT_TOKEN', '8047664560:AAFRqtU2ATsG9Cw9cdojFrDfQmjXFUWtgFs')
ADMIN_ID = 922109605
#PHOTO_PATH = 'photo.jpg'  # Картинку больше не используем
SLOTS_FILE = 'slots.json'

bot = Bot(token=API_TOKEN)
dp = Dispatcher(bot, storage=MemoryStorage())

DIRECTIONS = ['ЦТ', 'Фото', 'СМИ', 'Дизайн', 'F&U prod']

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


@dp.message_handler(commands=['get_slots'])
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
        creds = Credentials.from_service_account_file('credentials.json')
        gc = gspread.authorize(creds)
        sh = gc.open_by_key(sheet_id)
        ws = sh.sheet1
    except Exception as e:
        await message.answer(f'Ошибка подключения к Google Sheets: {e}')
        return

    try:
        # Новый диапазон доступности: B17:G28 (ячейки с B17 по G28)
        values = ws.get('B17:G28') or []
        dates = ws.get('B1:G1')[0]
        times = [r[0] for r in ws.get('A2:A13')]
        # ID направления теперь в A15
        dir_id_val = (ws.get('A15') or [['']])[0][0]
    except Exception as e:
        await message.answer(f'Ошибка чтения диапазонов: {e}')
        return

    new_slots = {'slots': {}, 'registrations': []}
    for i, date in enumerate(dates):
        date_key = date.strip()
        new_slots['slots'][date_key] = {}
        for j, time_slot in enumerate(times):
            cell_val = ''
            try:
                cell_val = values[j][i]
            except Exception:
                cell_val = ''
            new_slots['slots'][date_key][time_slot] = {
                'ЦТ': None,
                'Фото': None,
                'СМИ': None,
                'Дизайн': None,
                'F&U prod': None
            }
            # По новой логике: если ячейка пустая => НЕ может (blocked); если есть текст => может
            if not str(cell_val).strip():
                try:
                    dir_id = int(dir_id_val.strip())
                except Exception:
                    dir_id = None
                id_map = {1: 'ЦТ', 2: 'СМИ', 3: 'Дизайн', 4: 'F&U prod', 5: 'Фото'}
                if dir_id and dir_id in id_map:
                    new_slots['slots'][date_key][time_slot][id_map[dir_id]] = 'blocked'

    save_slots(new_slots)
    await message.answer('Слоты обновлены из Google Sheets и файл slots.json перезаписан.')
    await update_published_message()


def direction_has_free_slots(direction, slots):
    # проверяем, есть ли хотя бы один свободный слот для направления
    for date, times in slots.get('slots', {}).items():
        for t, dirs in times.items():
            if dirs.get(direction) is None:
                return True
    return False


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
    kb = InlineKeyboardMarkup(row_width=2)
    for d in DIRECTIONS:
        if direction_has_free_slots(d, slots):
            kb.insert(InlineKeyboardButton(d, callback_data=f'dir:{d}'))
    try:
        # Редактируем текст и клавиатуру
        await bot.edit_message_text(chat_id=chat_id, message_id=message_id, text='Выберите направление:', reply_markup=kb)
    except Exception:
        try:
            await bot.edit_message_reply_markup(chat_id=chat_id, message_id=message_id, reply_markup=kb)
        except Exception:
            pass


# Старт
@dp.message_handler(commands=['start'])
async def cmd_start(message: types.Message):
    # Инструкция для пользователя
    text = (
        'Здравствуйте! Чтобы записаться на собеседование, пройдите 4 шага:\n'
        '1) Введите ваше ФИО.\n'
        '2) Пришлите ссылку на ваш VK.\n'
        '3) Выберите направление (через команду /directions или через опубликованное сообщение).\n'
        '4) Выберите дату и время из доступных слотов.\n\n'
        'Команды:\n'
        '/directions — получить меню направлений.\n'
        '/my — показать вашу текущую запись.\n'
        '/cancel — отменить запись (не позднее чем за 24 часа).\n\n'
        'Начнём: введите ваше Имя и Фамилию:'
    )
    await Form.name.set()
    await message.answer(text)

@dp.message_handler(state=Form.name)
async def process_name(message: types.Message, state: FSMContext):
    await state.update_data(name=message.text)
    await Form.next()
    await message.answer('Пришлите ссылку на ваш VK:')

@dp.message_handler(state=Form.vk)
async def process_vk(message: types.Message, state: FSMContext):
    await state.update_data(vk=message.text)
    # Кнопки направлений
    # Предлагаем варианты: либо через reply-клавиатуру, либо через опубликованное сообщение (inline)
    kb = ReplyKeyboardMarkup(resize_keyboard=True)
    for d in DIRECTIONS:
        kb.add(KeyboardButton(d))
    await Form.next()
    await message.answer('Выберите направление (или нажмите кнопку в опубликованном сообщении):', reply_markup=kb)


@dp.message_handler(commands=['directions'])
async def cmd_directions(message: types.Message):
    # Показываем публицированный набор направлений, если он есть, иначе отправляем кнопки
    pub = load_published()
    if pub:
        try:
            await bot.forward_message(chat_id=message.chat.id, from_chat_id=pub['chat_id'], message_id=pub['message_id'])
            return
        except Exception:
            pass
    kb = InlineKeyboardMarkup(row_width=2)
    slots = load_slots()
    for d in DIRECTIONS:
        if direction_has_free_slots(d, slots):
            kb.insert(InlineKeyboardButton(d, callback_data=f'dir:{d}'))
    await message.answer('Выберите направление:', reply_markup=kb)


# Admin: публикуем сообщение с кнопками направлений и фотографией
@dp.message_handler(commands=['publish'])
async def cmd_publish(message: types.Message):
    if message.from_user.id != ADMIN_ID:
        await message.answer('Только админ может публиковать меню направлений.')
        return
    slots = load_slots()
    kb = InlineKeyboardMarkup(row_width=2)
    for d in DIRECTIONS:
        if direction_has_free_slots(d, slots):
            kb.insert(InlineKeyboardButton(d, callback_data=f'dir:{d}'))
    sent = await message.answer('Выберите направление:', reply_markup=kb)
    save_published({'chat_id': sent.chat.id, 'message_id': sent.message_id})
    # Добавляем кнопку Export для админа
    try:
        kb2 = InlineKeyboardMarkup(row_width=2)
        kb2.add(InlineKeyboardButton('Export CSV', callback_data='export:csv'))
        await bot.edit_message_reply_markup(chat_id=sent.chat.id, message_id=sent.message_id, reply_markup=kb)
    except Exception:
        pass
    await message.answer('Опубликовано.')
    log_action(f'published by {message.from_user.id}')


# Выбор направления
@dp.message_handler(state=Form.direction)
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
    kb = ReplyKeyboardMarkup(resize_keyboard=True)
    for d in available_dates:
        kb.add(KeyboardButton(d))
    # Отправляем фото и сообщение с выбором даты
    await Form.next()
    await message.answer('Выберите дату:', reply_markup=kb)


# Обработка нажатия inline-кнопки с направлением
@dp.callback_query_handler(lambda c: c.data and c.data.startswith('dir:'))
async def callback_dir(callback: types.CallbackQuery, state: FSMContext):
    direction = callback.data.split(':', 1)[1]
    await bot.answer_callback_query(callback.id)
    # Сохраняем направление в state и просим ФИО (если ещё не было)
    # Если пользователь уже начал ввод ФИО или VK — не сбрасываем состояние, просто сохраняем направление
    st = dp.current_state(user=callback.from_user.id)
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
    kb = ReplyKeyboardMarkup(resize_keyboard=True)
    for d in available_dates:
        kb.add(KeyboardButton(d))
    await bot.send_message(callback.from_user.id, 'Выберите дату:', reply_markup=kb)


@dp.callback_query_handler(lambda c: c.data and c.data.startswith('export:'))
async def callback_export(callback: types.CallbackQuery):
    if callback.from_user.id != ADMIN_ID:
        await bot.answer_callback_query(callback.id, 'Только админ может экспортировать')
        return
    await bot.answer_callback_query(callback.id, 'Генерирую CSV...')
    buf = export_registrations_csv()
    await bot.send_document(ADMIN_ID, ("registrations.csv", buf.getvalue().encode('utf-8')))
    log_action(f'export_csv by {callback.from_user.id}')


@dp.message_handler(commands=['export'])
async def cmd_export(message: types.Message):
    if message.from_user.id != ADMIN_ID:
        await message.answer('Только админ может экспортировать')
        return
    buf = export_registrations_csv()
    await message.answer_document(("registrations.csv", buf.getvalue().encode('utf-8')))
    log_action(f'export_csv by {message.from_user.id}')


@dp.message_handler(commands=['my'])
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
@dp.message_handler(state=Form.date)
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
            slot_dt = datetime.datetime.strptime(f"{date} {t[:5]}", "%d.%m.%Y %H:%M")
            if slot_dt - datetime.datetime.now() > datetime.timedelta(hours=12):
                available_times.append(t)
    if not available_times:
        await message.answer('Нет доступного времени на эту дату.')
        return
    kb = ReplyKeyboardMarkup(resize_keyboard=True)
    for t in available_times:
        kb.add(KeyboardButton(t))
    await state.update_data(date=date)
    await Form.next()
    await message.answer('Выберите время:', reply_markup=kb)

# Выбор времени и запись
@dp.message_handler(state=Form.time)
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
    slot_dt = datetime.datetime.strptime(f"{date} {time[:5]}", "%d.%m.%Y %H:%M")
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
    await state.finish()


# Отмена записи
@dp.message_handler(commands=['cancel'])
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
    slot_dt = datetime.datetime.strptime(f"{found['date']} {found['time'][:5]}", "%d.%m.%Y %H:%M")
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