import telebot
from telebot import types
import pandas as pd
import os
from datetime import datetime
import shutil
import time

# ================== ⚙️ НАСТРОЙКИ ==================
TOKEN = '8431099507:AAHFrDJUEnduHO1TMBuMcTSeNBbygDTjgBE'
bot = telebot.TeleBot(TOKEN)

ADMIN_ID = 495646038

EXCEL_FILE = 'ИСП(9) и (11).xlsx'
CONTRACTS_FOLDER = 'договоры'
BACKUP_FOLDER = 'backups'

os.makedirs(CONTRACTS_FOLDER, exist_ok=True)
os.makedirs(BACKUP_FOLDER, exist_ok=True)

user_data = {}
PRACTICE_COLUMNS = ['Место прохождения практики', 'Адрес', 'Руководитель', 'Номер телефона', 'ИНН', 'Договор']


# ================== 📊 РАБОТА С ДАННЫМИ ==================

def load_all_students():
    try:
        if not os.path.exists(EXCEL_FILE): return pd.DataFrame()
        xl = pd.ExcelFile(EXCEL_FILE)
        all_dfs = []
        for sheet_name in xl.sheet_names:
            df = pd.read_excel(xl, sheet_name=sheet_name, dtype=object)
            df = df.fillna('')
            if not df.empty:
                if df.columns[0] != 'ФИО':
                    df = df.rename(columns={df.columns[0]: 'ФИО'})
                df['Группа'] = str(sheet_name)
                for col in PRACTICE_COLUMNS:
                    if col not in df.columns: df[col] = ''
                all_dfs.append(df)
        return pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()
    except Exception as e:
        print(f"❌ Ошибка чтения Excel: {e}")
        return pd.DataFrame()


def save_students(df):
    try:
        shutil.copy2(EXCEL_FILE, os.path.join(BACKUP_FOLDER, f"backup_{int(time.time())}.xlsx"))
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
            for group in df['Группа'].unique():
                group_df = df[df['Группа'] == group].copy()
                save_df = group_df.drop(columns=['Группа'])
                save_df.to_excel(writer, sheet_name=str(group), index=False)
        return True
    except Exception as e:
        return False


# ================== 🛠 КЛАВИАТУРЫ ==================

def get_keyboard(user_id):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    markup.add('📝 Заполнить данные практики')
    if user_id == ADMIN_ID:
        markup.add('📊 Статистика', '📋 Список заполнивших')
        markup.add('📥 Скачать таблицу', '🔄 Сброс данных студента')
    return markup


# ================== 🤖 ОБРАБОТЧИКИ (ОБЩИЕ) ==================

@bot.message_handler(commands=['start', 'cancel'])
def start(message):
    user_data.pop(message.from_user.id, None)
    bot.send_message(
        message.chat.id,
        f"👋 Привет! Я бот для сбора данных о практике.\n"
        f"{'🛠 Вы вошли как администратор' if message.from_user.id == ADMIN_ID else ''}",
        reply_markup=get_keyboard(message.from_user.id)
    )


# ================== 👑 АДМИН-КОНСОЛЬ ==================

@bot.message_handler(func=lambda m: m.text == '📊 Статистика' and m.from_user.id == ADMIN_ID)
def admin_stats(message):
    df = load_all_students()
    if df.empty: return bot.send_message(message.chat.id, "Таблица пуста.")

    total = len(df)
    filled = len(df[df['Место прохождения практики'] != ''])

    bot.send_message(
        message.chat.id,
        f"📈 **Статистика:**\n\n"
        f"👥 Всего студентов: {total}\n"
        f"✅ Заполнили данные: {filled}\n"
        f"⏳ Осталось: {total - filled}",
        parse_mode="Markdown"
    )


@bot.message_handler(func=lambda m: m.text == '📋 Список заполнивших' and m.from_user.id == ADMIN_ID)
def admin_list(message):
    df = load_all_students()
    filled = df[df['Место прохождения практики'] != '']
    if filled.empty:
        return bot.send_message(message.chat.id, "Еще никто не заполнил данные.")

    text = "📝 **Студенты, внесшие данные:**\n\n"
    for i, row in filled.iterrows():
        text += f"🔹 {row['ФИО']} ({row['Группа']})\n"
        if len(text) > 3800:  # Лимит сообщения Телеграм
            bot.send_message(message.chat.id, text, parse_mode="Markdown")
            text = ""
    bot.send_message(message.chat.id, text, parse_mode="Markdown")


@bot.message_handler(func=lambda m: m.text == '📥 Скачать таблицу' and m.from_user.id == ADMIN_ID)
def admin_download(message):
    if os.path.exists(EXCEL_FILE):
        with open(EXCEL_FILE, 'rb') as f:
            bot.send_document(message.chat.id, f, caption="📂 Актуальный файл со всеми данными")
    else:
        bot.send_message(message.chat.id, "Файл не найден.")


@bot.message_handler(func=lambda m: m.text == '🔄 Сброс данных студента' and m.from_user.id == ADMIN_ID)
def admin_reset_start(message):
    user_data[message.from_user.id] = {'step': 'admin_wait_reset'}
    bot.send_message(message.chat.id, "🔄 Введите полное ФИО студента, чьи данные нужно очистить:")


@bot.message_handler(
    func=lambda m: user_data.get(m.from_user.id, {}).get('step') == 'admin_wait_reset' and m.from_user.id == ADMIN_ID)
def admin_reset_process(message):
    fio = message.text.strip().lower()
    df = load_all_students()

    found = False
    for idx, row in df.iterrows():
        if str(row['ФИО']).lower().strip() == fio:
            for col in PRACTICE_COLUMNS:
                df.at[idx, col] = ''
            found = True
            break

    if found and save_students(df):
        bot.send_message(message.chat.id, "✅ Данные студента успешно очищены. Он может заполнить их снова.")
    else:
        bot.send_message(message.chat.id, "❌ Студент не найден.")
    user_data.pop(message.from_user.id, None)

#===========боть=============


@bot.message_handler(commands=['start', 'cancel'])
def start(message):
    user_data.pop(message.from_user.id, None)
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    markup.add('📝 Заполнить данные практики')
    bot.send_message(
        message.chat.id,
        "👋 Привет! Я помогу тебе внести данные о практике в таблицу.\n\n"
        "Нажми кнопку ниже, чтобы начать 👇",
        reply_markup=markup
    )


@bot.message_handler(func=lambda m: m.text == '📝 Заполнить данные практики')
def ask_fio(message):
    user_id = message.from_user.id
    user_data[user_id] = {'step': 'wait_fio'}
    bot.send_message(
        user_id,
        "👤 Введите ваше ФИО полностью (как в списке группы):",
        reply_markup=types.ReplyKeyboardRemove()
    )


@bot.message_handler(func=lambda m: user_data.get(m.from_user.id, {}).get('step') == 'wait_fio')
def process_fio(message):
    user_id = message.from_user.id
    fio_input = message.text.strip().lower()
    df = load_all_students()

    student_idx = None
    for idx, row in df.iterrows():
        if str(row['ФИО']).lower().strip() == fio_input:
            student_idx = idx
            break

    if student_idx is not None:
        user_data[user_id].update({'idx': student_idx, 'step': 'wait_place'})
        bot.send_message(
            user_id,
            "🏢 Введите место прохождения (название организации) практики:\n\n"
            "*Пример: ООО «Ромашка», ИП Иванов И.И.*",
            parse_mode="Markdown"
        )
    else:
        bot.send_message(user_id, "🔍 К сожалению, ФИО не найдено в списках. Попробуйте еще раз или напишите /cancel")


@bot.message_handler(func=lambda m: user_data.get(m.from_user.id, {}).get('step') == 'wait_place')
def process_place(message):
    user_id = message.from_user.id
    user_data[user_id]['Место прохождения практики'] = message.text
    user_data[user_id]['step'] = 'wait_address'
    bot.send_message(user_id, "Введите адрес организации(Пример: 659635, г. Иркутск, ул. Новая, д. 7):")


@bot.message_handler(func=lambda m: user_data.get(m.from_user.id, {}).get('step') == 'wait_address')
def process_address(message):
    user_id = message.from_user.id
    user_data[user_id]['Адрес'] = message.text
    user_data[user_id]['step'] = 'wait_manager'
    bot.send_message(user_id, "👨‍💼 Введите **ФИО руководителя** практики от организации:")


@bot.message_handler(func=lambda m: user_data.get(m.from_user.id, {}).get('step') == 'wait_manager')
def process_manager(message):
    user_id = message.from_user.id
    user_data[user_id]['Руководитель'] = message.text
    user_data[user_id]['step'] = 'wait_phone'
    bot.send_message(user_id, "📞 Введите номер телефона руководителя практики от организации:")


@bot.message_handler(func=lambda m: user_data.get(m.from_user.id, {}).get('step') == 'wait_phone')
def process_phone(message):
    user_id = message.from_user.id
    user_data[user_id]['Номер телефона'] = message.text
    user_data[user_id]['step'] = 'wait_inn'
    bot.send_message(user_id, "🔢 Введите ИНН организации(10 символов):")


@bot.message_handler(func=lambda m: user_data.get(m.from_user.id, {}).get('step') == 'wait_inn')
def process_inn(message):
    user_id = message.from_user.id
    user_data[user_id]['ИНН'] = message.text
    user_data[user_id]['step'] = 'wait_file'
    bot.send_message(user_id, "📎 Отправьте скан или фото договора (файл или фото):")


@bot.message_handler(content_types=['document', 'photo'],
                     func=lambda m: user_data.get(m.from_user.id, {}).get('step') == 'wait_file')
def process_file(message):
    user_id = message.from_user.id
    data = user_data[user_id]

    try:
        if message.document:
            file_id = message.document.file_id
            ext = os.path.splitext(message.document.file_name)[1]
        else:
            file_id = message.photo[-1].file_id
            ext = '.jpg'

        file_info = bot.get_file(file_id)
        downloaded = bot.download_file(file_info.file_path)

        filename = f"contract_{data['idx']}_{int(time.time())}{ext}"
        path = os.path.join(CONTRACTS_FOLDER, filename)

        with open(path, 'wb') as f:
            f.write(downloaded)

        # 📥 Запись в таблицу
        df = load_all_students()
        idx = data['idx']

        df.at[idx, 'Место прохождения практики'] = data['Место прохождения практики']
        df.at[idx, 'Адрес'] = data['Адрес']
        df.at[idx, 'Руководитель'] = data['Руководитель']
        df.at[idx, 'Номер телефона'] = data['Номер телефона']
        df.at[idx, 'ИНН'] = data['ИНН']
        df.at[idx, 'Договор'] = filename

        if save_students(df):
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            markup.add('📝 Заполнить данные практики')
            bot.send_message(user_id, "✨ **Готово!** Данные успешно внесены в таблицу. Спасибо!", reply_markup=markup,
                             parse_mode="Markdown")
            user_data.pop(user_id, None)
    except Exception as e:
        bot.send_message(user_id, f"❌ Произошла ошибка при сохранении: {e}\nПопробуйте отправить файл еще раз.")


# ================== 🚀 ЗАПУСК ==================
if __name__ == '__main__':
    print(f"🚀 Бот запущен. Админ ID: {ADMIN_ID}")
    while True:
        try:
            bot.polling(none_stop=True, interval=1, timeout=60)
        except Exception as e:
            time.sleep(5)
