import telebot
from telebot import types
import gspread
import os
import time

# ================== ⚙️ НАСТРОЙКИ ==================
TOKEN = '8431099507:AAHFrDJUEnduHO1TMBuMcTSeNBbygDTjgBE'
bot = telebot.TeleBot(TOKEN)

ADMIN_ID = 495646038
JSON_FILE = 'Practice-493511-73467575d295.json'

# СТРОГО КАК В БРАУЗЕРЕ:
SHEET_NAME = 'ИСП(9) и (11)'

CONTRACTS_FOLDER = 'договоры'
os.makedirs(CONTRACTS_FOLDER, exist_ok=True)
user_data = {}

# Заголовки (как в вашей таблице)
COL_FIO = "Фамилия, имя, отчество"
COL_PLACE = "Место прохождения практики (ООО 'Ромашка', ИП Иванов И.И.)"
COL_ADDR = "Адрес прохождения практики (659635, г. Иркутск, ул. Новая, д. 7)"
COL_BOSS = "Фамилия и имя руководителя"
COL_PHONE = "Номер телефона руководителя практики от организации"
COL_INN = "ИНН"
COL_DOC = "Договор"


# ================== 📊 РАБОТА С GOOGLE TABLES ==================

def get_spreadsheet():
    try:
        gc = gspread.service_account(filename=JSON_FILE)
        # Пробуем открыть по названию
        return gc.open(SHEET_NAME)
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"❌ ОШИБКА: Таблица с названием '{SHEET_NAME}' не найдена!")
        print("Проверьте, что вы дали доступ (Editor) для email из JSON-файла.")
        return None
    except Exception as e:
        print(f"❌ Ошибка доступа: {e}")
        return None


def load_all_students():
    sh = get_spreadsheet()
    if not sh: return []

    all_students = []
    for worksheet in sh.worksheets():
        try:
            # Читаем все данные листа как список списков
            data = worksheet.get_all_values()
            if not data: continue

            headers = [h.strip() for h in data[0]]  # Первая строка - заголовки
            rows = data[1:]  # Остальные строки - данные

            for idx, row in enumerate(rows):
                # Создаем словарь, только если заголовок не пустой
                student_dict = {}
                for col_idx, cell_value in enumerate(row):
                    if col_idx < len(headers) and headers[col_idx]:  # Если заголовок есть
                        student_dict[headers[col_idx]] = cell_value

                if student_dict:
                    student_dict['sheet_name'] = worksheet.title
                    student_dict['row_idx'] = idx + 2
                    all_students.append(student_dict)
        except Exception as e:
            print(f"Ошибка чтения листа {worksheet.title}: {e}")
    return all_students


def save_student_to_sheet(sheet_name, row_idx, new_data):
    try:
        sh = get_spreadsheet()
        if not sh: return False
        worksheet = sh.worksheet(sheet_name)
        headers = [h.strip() for h in worksheet.row_values(1)]

        for col_name, value in new_data.items():
            if col_name not in headers:
                new_col_idx = len(headers) + 1
                worksheet.update_cell(1, new_col_idx, col_name)
                headers.append(col_name)

            col_idx = headers.index(col_name) + 1
            worksheet.update_cell(row_idx, col_idx, value)
        return True
    except Exception as e:
        print(f"Ошибка сохранения: {e}")
        return False


# ================== ОБРАБОТКА КОМАНД ==================

@bot.message_handler(commands=['start', 'admin'])
def start(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    markup.add('📝 Заполнить всё', '📄 Только отправить договор')
    if message.from_user.id == ADMIN_ID:
        markup.add('🔗 Ссылка', '📂 Договоры')
    bot.send_message(message.chat.id, "Бот готов. Введите фамилию для поиска.", reply_markup=markup)


@bot.message_handler(func=lambda message: message.text in ['📝 Заполнить всё', '📄 Только отправить договор'])
def search_student_init(message):
    mode = 'full' if message.text == '📝 Заполнить всё' else 'contract_only'
    user_data[message.chat.id] = {'mode': mode}
    bot.send_message(message.chat.id, "Введите вашу Фамилию:")


@bot.message_handler(func=lambda message: message.chat.id in user_data and 'row_idx' not in user_data[message.chat.id])
def search_student_process(message):
    query = message.text.strip().lower()
    bot.send_message(message.chat.id, "🔍 Ищу в базе Google...")

    students = load_all_students()
    if not students:
        bot.send_message(message.chat.id, "❌ База данных недоступна или пуста. Проверьте консоль бота.")
        return

    results = [s for s in students if query in str(s.get(COL_FIO, '')).lower().strip()]

    if not results:
        bot.send_message(message.chat.id, "Студент не найден. Проверьте правильность написания.")
    elif len(results) > 1:
        msg = "Найдено несколько совпадений:\n"
        for s in results[:5]: msg += f"- {s[COL_FIO]}\n"
        bot.send_message(message.chat.id, msg)
    else:
        student = results[0]
        user_data[message.chat.id].update({
            'row_idx': student['row_idx'],
            'sheet_name': student['sheet_name'],
            'fio': student[COL_FIO]
        })

        if user_data[message.chat.id]['mode'] == 'full':
            bot.send_message(message.chat.id, f"Выбран: {student[COL_FIO]}\nВведите место прохождения практики:")
            bot.register_next_step_handler(message, process_practice_place)
        else:
            bot.send_message(message.chat.id, f"Выбран: {student[COL_FIO]}\nПришлите договор:")
            bot.register_next_step_handler(message, process_contract_file)


# --- Цепочка опроса ---
def process_practice_place(message):
    user_data[message.chat.id][COL_PLACE] = message.text
    bot.send_message(message.chat.id, "Введите адрес организации:")
    bot.register_next_step_handler(message, process_address)


def process_address(message):
    user_data[message.chat.id][COL_ADDR] = message.text
    bot.send_message(message.chat.id, "Введите ФИО руководителя:")
    bot.register_next_step_handler(message, process_boss)


def process_boss(message):
    user_data[message.chat.id][COL_BOSS] = message.text
    bot.send_message(message.chat.id, "Введите телефон руководителя:")
    bot.register_next_step_handler(message, process_phone)


def process_phone(message):
    user_data[message.chat.id][COL_PHONE] = message.text
    bot.send_message(message.chat.id, "Введите ИНН:")
    bot.register_next_step_handler(message, process_inn)


def process_inn(message):
    user_data[message.chat.id][COL_INN] = message.text
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    markup.add('Да', 'Нет')
    bot.send_message(message.chat.id, "Загрузить договор сейчас?", reply_markup=markup)
    bot.register_next_step_handler(message, process_contract_decision)


def process_contract_decision(message):
    if message.text.lower() == 'да':
        bot.send_message(message.chat.id, "Пришлите файл:")
        bot.register_next_step_handler(message, process_contract_file)
    else:
        finish_saving(message.chat.id)


def process_contract_file(message):
    if not message.document:
        bot.send_message(message.chat.id, "Пожалуйста, отправьте файл.")
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

    if save_student_to_sheet(sheet_name, row_idx, data):
        bot.send_message(chat_id, "✅ Данные успешно занесены в таблицу!")
    else:
        bot.send_message(chat_id, "❌ Ошибка записи. Проверьте консоль.")
    user_data.pop(chat_id, None)


bot.infinity_polling()
