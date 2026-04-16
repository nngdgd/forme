import telebot
from telebot import types
import gspread
import os
import time

# ================== ⚙️ НАСТРОЙКИ ==================
TOKEN = '8431099507:AAGxVNREWudOXK-I5uhlL_CQaGFU-7zkRIA'
bot = telebot.TeleBot(TOKEN)

ADMIN_ID = 495646038
JSON_FILE = 'practice-493511-54adf2cc8263.json'
SHEET_NAME = 'ИСП(9) и (11)'

CONTRACTS_FOLDER = 'договоры'
os.makedirs(CONTRACTS_FOLDER, exist_ok=True)
user_data = {}

# Заголовки (Бот будет искать их в первой строке таблицы)
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
        return gc.open(SHEET_NAME)
    except Exception as e:
        print(f"❌ Ошибка подключения: {e}")
        return None


def load_all_students():
    sh = get_spreadsheet()
    if not sh: return []

    all_students = []
    for worksheet in sh.worksheets():
        try:
            data = worksheet.get_all_values()
            if len(data) < 2: continue

            # Чистим заголовки от пробелов по краям
            headers = [h.strip() for h in data[0]]
            rows = data[1:]

            for idx, row in enumerate(rows):
                if not row or not any(row): continue  # Пропуск пустых строк

                student_dict = {}
                for col_idx, cell_value in enumerate(row):
                    if col_idx < len(headers):
                        student_dict[headers[col_idx]] = cell_value.strip()

                # Добавляем служебную информацию
                student_dict['sheet_name'] = worksheet.title
                student_dict['row_idx'] = idx + 2

                # Принудительно проверяем наличие ФИО
                if student_dict.get(COL_FIO):
                    all_students.append(student_dict)

        except Exception as e:
            print(f"⚠️ Ошибка чтения листа {worksheet.title}: {e}")

    print(f"✅ Загружено студентов: {len(all_students)}")
    return all_students


def save_student_to_sheet(sheet_name, row_idx, new_data):
    try:
        sh = get_spreadsheet()
        if not sh: return False
        worksheet = sh.worksheet(sheet_name)

        # Получаем текущие заголовки для поиска нужных колонок
        headers = [h.strip() for h in worksheet.row_values(1)]

        for col_name, value in new_data.items():
            if col_name in headers:
                col_idx = headers.index(col_name) + 1
                worksheet.update_cell(row_idx, col_idx, value)
            else:
                # Если такой колонки нет, создаем её в конце
                new_col_idx = len(headers) + 1
                worksheet.update_cell(1, new_col_idx, col_name)
                worksheet.update_cell(row_idx, new_col_idx, value)
                headers.append(col_name)
        return True
    except Exception as e:
        print(f"❌ Ошибка сохранения: {e}")
        return False


# ================== ОБРАБОТКА КОМАНД ==================

@bot.message_handler(commands=['start', 'admin'])
def start(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    markup.add('📝 Заполнить всё', '📄 Только отправить договор')
    if message.from_user.id == ADMIN_ID:
        markup.add('🔗 Ссылка', '📂 Договоры')
    bot.send_message(message.chat.id, "Бот готов. Выберите действие или просто введите фамилию для поиска.",
                     reply_markup=markup)


@bot.message_handler(func=lambda message: message.text in ['📝 Заполнить всё', '📄 Только отправить договор'])
def search_student_init(message):
    mode = 'full' if message.text == '📝 Заполнить всё' else 'contract_only'
    user_data[message.chat.id] = {'mode': mode}
    bot.send_message(message.chat.id, "Введите вашу Фамилию (как в списке):")


@bot.message_handler(func=lambda message: True)
def handle_all_messages(message):
    chat_id = message.chat.id

    # Если мы уже в процессе заполнения данных, этот хэндлер не должен мешать
    if chat_id in user_data and 'row_idx' in user_data[chat_id]:
        return

    query = message.text.strip().lower()
    if len(query) < 3:
        bot.send_message(chat_id, "Введите хотя бы 3 буквы фамилии.")
        return

    bot.send_message(chat_id, "🔍 Ищу в базе Google...")
    students = load_all_students()

    # Ищем совпадения (без учета регистра и лишних пробелов)
    results = [s for s in students if query in str(s.get(COL_FIO, '')).lower()]

    if not results:
        bot.send_message(chat_id, "❌ Студент не найден. Проверьте, что вы есть в таблице и фамилия написана верно.")
    elif len(results) > 1:
        msg = "Найдено несколько совпадений. Уточните фамилию:\n"
        for s in results[:5]: msg += f"- {s[COL_FIO]}\n"
        bot.send_message(chat_id, msg)
    else:
        student = results[0]
        # Инициализируем данные, если пользователь просто ввел фамилию без нажатия кнопки
        if chat_id not in user_data:
            user_data[chat_id] = {'mode': 'full'}

        user_data[chat_id].update({
            'row_idx': student['row_idx'],
            'sheet_name': student['sheet_name'],
            'fio': student[COL_FIO]
        })

        if user_data[chat_id]['mode'] == 'full':
            bot.send_message(chat_id, f"✅ Выбран: {student[COL_FIO]}\n\nВведите место прохождения практики (ООО/ИП):")
            bot.register_next_step_handler(message, process_practice_place)
        else:
            bot.send_message(chat_id, f"✅ Выбран: {student[COL_FIO]}\n\nПришлите файл договора:")
            bot.register_next_step_handler(message, process_contract_file)


# --- Цепочка опроса (осталась без изменений в логике) ---
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
        bot.send_message(message.chat.id, "Ошибка! Пожалуйста, отправьте именно файл (как документ).")
        bot.register_next_step_handler(message, process_contract_file)
        return

    file_info = bot.get_file(message.document.file_id)
    downloaded_file = bot.download_file(file_info.file_path)
    ext = os.path.splitext(message.document.file_name)[1]

    filename = f"{user_data[message.chat.id]['fio']}_{int(time.time())}{ext}".replace(' ', '_')
    path = os.path.join(CONTRACTS_FOLDER, filename)

    with open(path, 'wb') as f:
        f.write(downloaded_file)

    user_data[message.chat.id][COL_DOC] = filename
    finish_saving(message.chat.id)


def finish_saving(chat_id):
    data = user_data[chat_id]
    row_idx = data.pop('row_idx')
    sheet_name = data.pop('sheet_name')
    data.pop('mode', None)
    data.pop('fio', None)

    bot.send_message(chat_id, "⏳ Сохраняю данные в таблицу...")
    if save_student_to_sheet(sheet_name, row_idx, data):
        bot.send_message(chat_id, "✅ Всё готово! Данные успешно обновлены.")
    else:
        bot.send_message(chat_id, "❌ Произошла ошибка при записи в таблицу. Свяжитесь с администратором.")

    user_data.pop(chat_id, None)


print("🚀 Бот запущен...")
bot.infinity_polling()
