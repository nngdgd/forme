import telebot
from telebot import types
import gspread
import os
import time
import re

# ================== ⚙️ НАСТРОЙКИ ==================
TOKEN = '8431099507:AAH9qaaYkPRwiEx3S7KPcVbRTiIF6Xj9vJo'
bot = telebot.TeleBot(TOKEN)

ADMIN_ID = 495646038
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, 'practice-493511-5f937c9370f5.json')
SHEET_NAME = 'ИСП(9) и (11)'

CONTRACTS_FOLDER = 'договоры'
os.makedirs(CONTRACTS_FOLDER, exist_ok=True)
user_data = {}

# Названия колонок для записи (должны быть в таблице)
COL_FIO = "Фамилия, имя, отчество"
COL_PLACE = "Место прохождения практики (ООО 'Ромашка', ИП Иванов И.И.)"
COL_ADDR = "Адрес прохождения практики (659635, г. Иркутск, ул. Новая, д. 7)"
COL_BOSS = "Фамилия и имя руководителя"
COL_PHONE = "Номер телефона руководителя практики от организации"
COL_INN = "ИНН"
COL_DOC = "Договор"


# ================== 🛠 ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ==================

def clean_string(text):
    """Удаляет невидимые символы, лишние пробелы и приводит к нижнему регистру."""
    if not text: return ""
    # Удаляем неразрывные пробелы \xa0 и прочий мусор
    text = text.replace('\xa0', ' ').replace('\r', '').replace('\n', '')
    return " ".join(text.split()).lower()


# ================== 📊 РАБОТА С GOOGLE TABLES ==================

def get_spreadsheet(chat_id=None):
    try:
        if not os.path.exists(JSON_FILE):
            error_msg = f"❌ Файл не найден по пути: {JSON_FILE}"
            print(error_msg)
            if chat_id: bot.send_message(chat_id, error_msg)
            return None

        gc = gspread.service_account(filename=JSON_FILE)
        return gc.open(SHEET_NAME)
    except Exception as e:
        error_msg = f"❌ Ошибка Google API: {str(e)}"
        print(error_msg)
        if chat_id: bot.send_message(chat_id, error_msg)
        return None


def load_all_students(chat_id=None):
    sh = get_spreadsheet(chat_id)
    if not sh: return []

    all_students = []
    for worksheet in sh.worksheets():
        try:
            data = worksheet.get_all_values()
            if len(data) < 2: continue

            headers = [h.strip() for h in data[0]]
            rows = data[1:]

            for idx, row in enumerate(rows):
                if not row or not any(row): continue

                # Создаем словарь студента
                student_dict = {}
                # ВАЖНО: Всегда считаем первый столбец как ФИО (COL_FIO)
                student_dict[COL_FIO] = row[0].strip()

                for col_idx, cell_value in enumerate(row):
                    if col_idx < len(headers):
                        student_dict[headers[col_idx]] = cell_value.strip()

                student_dict['sheet_name'] = worksheet.title
                student_dict['row_idx'] = idx + 2
                all_students.append(student_dict)
        except Exception as e:
            print(f"⚠️ Ошибка листа {worksheet.title}: {e}")
    return all_students


def save_student_to_sheet(sheet_name, row_idx, new_data):
    try:
        sh = get_spreadsheet()
        if not sh: return False
        worksheet = sh.worksheet(sheet_name)
        headers = [h.strip() for h in worksheet.row_values(1)]

        for col_name, value in new_data.items():
            if col_name in headers:
                col_idx = headers.index(col_name) + 1
                worksheet.update_cell(row_idx, col_idx, value)
            else:
                # Добавление новой колонки, если её нет
                new_col_idx = len(headers) + 1
                worksheet.update_cell(1, new_col_idx, col_name)
                worksheet.update_cell(row_idx, new_col_idx, value)
                headers.append(col_name)
        return True
    except Exception as e:
        print(f"❌ Ошибка записи: {e}")
        return False


# ================== ОБРАБОТКА КОМАНД ==================

@bot.message_handler(commands=['start', 'admin'])
def start(message):
    # При старте всегда очищаем данные пользователя
    user_data.pop(message.chat.id, None)

    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    markup.add('📝 Заполнить всё', '📄 Только отправить договор')
    if message.from_user.id == ADMIN_ID:
        markup.add('🔗 Ссылка', '📂 Договоры')
    bot.send_message(message.chat.id, "Бот готов. Выберите действие или введите фамилию.", reply_markup=markup)


@bot.message_handler(func=lambda message: message.text in ['📝 Заполнить всё', '📄 Только отправить договор'])
def search_student_init(message):
    mode = 'full' if message.text == '📝 Заполнить всё' else 'contract_only'
    user_data[message.chat.id] = {'mode': mode}
    bot.send_message(message.chat.id, "Введите вашу Фамилию (как в списке):")


@bot.message_handler(func=lambda message: True)
def handle_all_messages(message):
    cid = message.chat.id

    # Если мы уже в процессе заполнения анкеты, не реагируем на поиск
    if cid in user_data and 'row_idx' in user_data[cid]:
        return

    query = clean_string(message.text)
    if len(query) < 3:
        bot.send_message(cid, "Введите хотя бы 3 буквы фамилии.")
        return

    bot.send_message(cid, "🔍 Ищу в базе Google...")
    students = load_all_students(cid)

    if not students:
        bot.send_message(cid, "❌ База данных недоступна или пуста. Проверьте JSON файл на хостинге.")
        return

    # Поиск с глубокой очисткой строк
    results = [s for s in students if query in clean_string(s.get(COL_FIO, ''))]

    if not results:
        # Для отладки пишем, сколько всего людей бот увидел в таблице
        bot.send_message(cid, f"❌ Студент не найден.\n(В базе всего строк: {len(students)})\nПроверьте фамилию.")
    elif len(results) > 1:
        msg = "Найдено несколько совпадений. Уточните фамилию:\n"
        for s in results[:5]: msg += f"- {s[COL_FIO]}\n"
        bot.send_message(cid, msg)
    else:
        student = results[0]
        if cid not in user_data: user_data[cid] = {'mode': 'full'}

        user_data[cid].update({
            'row_idx': student['row_idx'],
            'sheet_name': student['sheet_name'],
            'fio': student[COL_FIO]
        })

        if user_data[cid]['mode'] == 'full':
            bot.send_message(cid, f"✅ Выбран: {student[COL_FIO]}\n\nВведите место прохождения практики:")
            bot.register_next_step_handler(message, process_practice_place)
        else:
            bot.send_message(cid, f"✅ Выбран: {student[cid]['fio']}\n\nПришлите файл договора:")
            bot.register_next_step_handler(message, process_contract_file)


# --- Цепочка опроса (без изменений) ---
def process_practice_place(message):
    user_data[message.chat.id][COL_PLACE] = message.text
    bot.send_message(message.chat.id, "Введите адрес организации:")
    bot.register_next_step_handler(message, process_address)


def process_address(message):
    user_data[message.chat.id][COL_ADDR] = message.text
    bot.send_message(message.chat.id, "Введите ФИО руководителя практики:")
    bot.register_next_step_handler(message, process_boss)


def process_boss(message):
    user_data[message.chat.id][COL_BOSS] = message.text
    bot.send_message(message.chat.id, "Введите телефон руководителя:")
    bot.register_next_step_handler(message, process_phone)


def process_phone(message):
    user_data[message.chat.id][COL_PHONE] = message.text
    bot.send_message(message.chat.id, "Введите ИНН организации:")
    bot.register_next_step_handler(message, process_inn)


def process_inn(message):
    user_data[message.chat.id][COL_INN] = message.text
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    markup.add('Да', 'Нет')
    bot.send_message(message.chat.id, "Загрузить файл договора сейчас?", reply_markup=markup)
    bot.register_next_step_handler(message, process_contract_decision)


def process_contract_decision(message):
    if message.text.lower() == 'да':
        bot.send_message(message.chat.id, "Пришлите файл (документом):")
        bot.register_next_step_handler(message, process_contract_file)
    else:
        finish_saving(message.chat.id)


def process_contract_file(message):
    if not message.document:
        bot.send_message(message.chat.id, "Ошибка! Отправьте именно файл.")
        bot.register_next_step_handler(message, process_contract_file)
        return

    file_info = bot.get_file(message.document.file_id)
    downloaded_file = bot.download_file(file_info.file_path)
    ext = os.path.splitext(message.document.file_name)[1]
    filename = f"{user_data[message.chat.id]['fio']}_{int(time.time())}{ext}".replace(' ', '_')
    path = os.path.join(CONTRACTS_FOLDER, filename)
    with open(path, 'wb') as f: f.write(downloaded_file)
    user_data[message.chat.id][COL_DOC] = filename
    finish_saving(message.chat.id)


def finish_saving(chat_id):
    data = user_data[chat_id]
    row_idx = data.pop('row_idx');
    sheet_name = data.pop('sheet_name')
    data.pop('mode', None);
    data.pop('fio', None)

    bot.send_message(chat_id, "⏳ Сохраняю...")
    if save_student_to_sheet(sheet_name, row_idx, data):
        bot.send_message(chat_id, "✅ Данные успешно обновлены.")
    else:
        bot.send_message(chat_id, "❌ Ошибка записи.")
    user_data.pop(chat_id, None)


print("🚀 Бот запущен...")
bot.infinity_polling()
