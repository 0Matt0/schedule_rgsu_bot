import os
import json
import pandas as pd
from aiogram import Bot, Dispatcher, types
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton
from aiogram import F
from datetime import datetime, timedelta
import asyncio

# Токен бота  @schedule_RGSU_bot
TOKEN = "7511135107:AAFA8yrW1vhMEH23UKKvCZH7-OG1-KCfsRo"
bot = Bot(token=TOKEN)
dp = Dispatcher()

# Путь к файлу с расписанием
file_path = "data\\Sched2.xlsx"
user_data_file = "data\\users_data.json"  # Файл для хранения данных о пользователях
schedule_dir = "data\\"

# Время пар
class_time = {
    '1 пара': '8:30',
    '2 пара': '10:10',
    '3 пара': '12:10',
    '4 пара': '13:50',
    '5 пара': '15:30',
    '6 пара': '17:10'
}

# Дни недели
day_names = {
    'Monday': 'ПН',
    'Tuesday': 'ВТ',
    'Wednesday': 'СР',
    'Thursday': 'ЧТ',
    'Friday': 'ПТ',
    'Saturday': 'СБ',
    'Sunday': 'ВС'
}

# клавиатура для выбора группы
def generate_group_keyboard():
    df = pd.read_excel(file_path)
    groups = df['Учебная группа'].unique()
    
    sorted_groups = sorted(groups) 
    
    keyboard = ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text=group)] for group in sorted_groups],
        resize_keyboard=True
    )
    return keyboard

# Функция для сохранения данных о пользователе
def save_user_data(user_id, username, group):
    users_data = {}
    if os.path.exists(user_data_file):
        with open(user_data_file, 'r') as f:
            users_data = json.load(f)
    users_data[user_id] = {"username": username, "group": group}
    with open(user_data_file, 'w') as f:
        json.dump(users_data, f)

# Функция для загрузки данных о пользователе
def load_user_data(user_id):
    if os.path.exists(user_data_file):
        with open(user_data_file, 'r') as f:
            users_data = json.load(f)
        return users_data.get(str(user_id))
    return None

# Обработчик команды /start
@dp.message(F.text == "/start")
async def send_welcome(message: types.Message):
    user_data = load_user_data(message.from_user.id)
    if user_data:
        group = user_data['group']
        await message.answer(f"Привет, {message.from_user.username}! Ваша группа: {group}. Выберите действие.", reply_markup=generate_schedule_keyboard())
    else:
        await message.answer("Привет! Пожалуйста, выберите свою учебную группу.", reply_markup=generate_group_keyboard())

# клавиатура с кнопками
def generate_schedule_keyboard():
    keyboard = ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="На сегодня")],
            [KeyboardButton(text="На завтра")],
            [KeyboardButton(text="На 5 дней")],
            [KeyboardButton(text="На следующую неделю")],
            [KeyboardButton(text="Удалить Данные")]
        ],
        resize_keyboard=True
    )
    return keyboard

# Обработчик выбора группы
@dp.message(lambda message: message.text in pd.read_excel(file_path)['Учебная группа'].unique())
async def choose_group(message: types.Message):
    user_data = load_user_data(message.from_user.id)
    if user_data:
        group = user_data['group']
        await message.answer(f"Вы уже выбрали группу: {group}. Теперь вы можете запросить расписание.", reply_markup=generate_schedule_keyboard())
    else:
        group = message.text
        save_user_data(message.from_user.id, message.from_user.username, group)
        await message.answer(f"Вы выбрали группу: {group}. Теперь вы можете запросить расписание.", reply_markup=generate_schedule_keyboard())

# Функция для создания файла расписания по группе
def generate_schedule_file(group):
    output_file = f"data\\{group}_schedule.xlsx"

    try:
        df = pd.read_excel(file_path, sheet_name=0)
        df['День'] = pd.to_datetime(df['День'], errors='coerce', format='%d.%m.%Y')
        columns_to_drop = ["Специальность", "Факультет уч группы", "Иностранный язык", "Форма обучения"]
        df = df.drop(columns=columns_to_drop, errors='ignore')
        df['Номер пары'] = df['Учебная пара'].str.extract(r'(\d+)').astype(float)
        df = df.sort_values(by=['День', 'Номер пары'])

        filtered_df = df[df['Учебная группа'] == group]

        if os.path.exists(output_file):
            existing_df = pd.read_excel(output_file)
            if filtered_df.equals(existing_df):
                return output_file

        rows_with_empty = []
        prev_day = None

        for index, row in filtered_df.iterrows():
            if prev_day is not None and row['День'] != prev_day:
                rows_with_empty.append(pd.Series(dtype='object'))
            rows_with_empty.append(row)
            prev_day = row['День']

        final_df = pd.DataFrame(rows_with_empty)
        final_df.to_excel(output_file, index=False)

        return output_file
    except Exception as e:
        return f"Ошибка при обработке файла: {e}"

################################################   НАЧАЛО РАБОТЫ С 5 КНОПКАМИ  ##################################################

# Функция для получения расписания на завтра
def get_schedule_for_next_day(group):
    try:
        file_path = generate_schedule_file(group)
        if not os.path.exists(file_path):
            return "Расписание для вашей группы не найдено."

        df = pd.read_excel(file_path)
        df['День'] = pd.to_datetime(df['День'], errors='coerce')

        next_day = datetime.now().date() + timedelta(days=1)

        next_day_schedule = df[df['День'] == pd.Timestamp(next_day)]

        if next_day_schedule.empty:
            return f"Расписание на {next_day.strftime('%d.%m.%Y')} отсутствует."

        next_day_of_week = day_names[next_day.strftime('%A')]
        schedule_text = f"*-----{next_day.strftime('%d.%m.%Y')} ({next_day_of_week})------*\n\n"

        for _, row in next_day_schedule.iterrows():
            schedule_text += (
                f"  *Пара:* {row['Учебная пара']} ({class_time.get(row['Учебная пара'])})\n"
                f"  *Аудитория:* {row['Аудитория']}\n"
                f"  *Дисциплина:* {row['Дисциплина']}\n\n"
            )
        return schedule_text.strip()
    except Exception as e:
        return f"Ошибка при обработке файла: {e}"
    
# Обработчик нажатия на кнопку "На завтра"
@dp.message(lambda message: message.text == "На завтра")
async def send_tomorrow_schedule(message: types.Message):
    user_data = load_user_data(message.from_user.id)
    if user_data:
        group = user_data['group']
        
        schedule_text = get_schedule_for_next_day(group)
        
        await message.answer(schedule_text, parse_mode="Markdown", protect_content=True)
    else:
        await message.answer("Пожалуйста, выберите свою учебную группу.", reply_markup=generate_group_keyboard())

# Функция для получения расписания на 5 дней
def get_schedule_for_next_5_days(group):
    try:
        file_path = generate_schedule_file(group)
        if not os.path.exists(file_path):
            return "Расписание для вашей группы не найдено."

        df = pd.read_excel(file_path)
        df['День'] = pd.to_datetime(df['День'], errors='coerce')

        start_date = datetime.now().date() + timedelta(days=1)
        end_date = start_date + timedelta(days=4)

        next_5_days_schedule = df[(df['День'] >= pd.Timestamp(start_date)) & (df['День'] <= pd.Timestamp(end_date))]

        if next_5_days_schedule.empty:
            return "Расписание на ближайшие 5 дней отсутствует."

        schedule_text = f"Расписание на ближайшие 5 дней:\n\n"

        for day in pd.date_range(start_date, end_date):
            day_of_week = day_names[day.strftime('%A')]
            day_schedule = next_5_days_schedule[next_5_days_schedule['День'] == pd.Timestamp(day)]

            if not day_schedule.empty:
                schedule_text += f"*-----{day.strftime('%d.%m.%Y')} ({day_of_week})------*\n\n"
                for _, row in day_schedule.iterrows():
                    schedule_text += (
                        f"  *Пара:* {row['Учебная пара']} ({class_time.get(row['Учебная пара'])})\n"
                        f"  *Аудитория:* {row['Аудитория']}\n"
                        f"  *Дисциплина:* {row['Дисциплина']}\n\n"
                    )
        return schedule_text.strip()
    except Exception as e:
        return f"Ошибка при обработке файла: {e}"
    
# Обработчик нажатия на кнопку "На 5 дней"
@dp.message(lambda message: message.text == "На 5 дней")
async def send_5_days_schedule(message: types.Message):
    user_data = load_user_data(message.from_user.id)
    if user_data:
        group = user_data['group']

        schedule_text = get_schedule_for_next_5_days(group)

        await message.answer(schedule_text, parse_mode="Markdown", protect_content=True)
    else:
        await message.answer("Пожалуйста, выберите свою учебную группу.", reply_markup=generate_group_keyboard())

# Функция для получения расписания на сегодня
def get_schedule_for_today(group):
    try:
        file_path = generate_schedule_file(group)
        if not os.path.exists(file_path):
            return "Расписание для вашей группы не найдено."

        df = pd.read_excel(file_path)
        df['День'] = pd.to_datetime(df['День'], errors='coerce')

        today = datetime.now().date()

        today_schedule = df[df['День'] == pd.Timestamp(today)]

        if today_schedule.empty:
            return f"Расписание на сегодня ({today.strftime('%d.%m.%Y')}) отсутствует."

        today_of_week = day_names[today.strftime('%A')]
        schedule_text = f"*-----{today.strftime('%d.%m.%Y')} ({today_of_week})------*\n\n"

        for _, row in today_schedule.iterrows():
            schedule_text += (
                f"  *Пара:* {row['Учебная пара']} ({class_time.get(row['Учебная пара'])})\n"
                f"  *Аудитория:* {row['Аудитория']}\n"
                f"  *Дисциплина:* {row['Дисциплина']}\n\n"
            )
        return schedule_text.strip()
    except Exception as e:
        return f"Ошибка при обработке файла: {e}"

# Обработчик нажатия на кнопку "На сегодня"
@dp.message(lambda message: message.text == "На сегодня")
async def send_today_schedule(message: types.Message):
    user_data = load_user_data(message.from_user.id)
    if user_data:
        group = user_data['group']

        schedule_text = get_schedule_for_today(group)

        await message.answer(schedule_text, parse_mode="Markdown", protect_content=True)
    else:
        await message.answer("Пожалуйста, выберите свою учебную группу.", reply_markup=generate_group_keyboard())

# Функция для получения расписания на следующую неделю
def get_schedule_for_next_week(group):
    try:
        file_path = generate_schedule_file(group)
        if not os.path.exists(file_path):
            return "Расписание для вашей группы не найдено."

        df = pd.read_excel(file_path)
        df['День'] = pd.to_datetime(df['День'], errors='coerce')

        today = datetime.now().date()
        next_monday = today + timedelta(days=(7 - today.weekday()))
        next_sunday = next_monday + timedelta(days=6)

        next_week_schedule = df[(df['День'] >= pd.Timestamp(next_monday)) & (df['День'] <= pd.Timestamp(next_sunday))]

        if next_week_schedule.empty:
            return "Расписание на следующую неделю отсутствует."

        schedule_text = f"Расписание на следующую неделю \n(с {next_monday.strftime('%d.%m.%Y')} по {next_sunday.strftime('%d.%m.%Y')}):\n\n"

        for day in pd.date_range(next_monday, next_sunday):
            day_of_week = day_names[day.strftime('%A')]
            day_schedule = next_week_schedule[next_week_schedule['День'] == pd.Timestamp(day)]

            if not day_schedule.empty:
                schedule_text += f"*-----{day.strftime('%d.%m.%Y')} ({day_of_week})------*\n\n"
                for _, row in day_schedule.iterrows():
                    schedule_text += (
                        f"  *Пара:* {row['Учебная пара']} ({class_time.get(row['Учебная пара'])})\n"
                        f"  *Аудитория:* {row['Аудитория']}\n"
                        f"  *Дисциплина:* {row['Дисциплина']}\n\n"
                    )
        return schedule_text.strip()
    except Exception as e:
        return f"Ошибка при обработке файла: {e}"

# Обработчик нажатия на кнопку "На следующую неделю"
@dp.message(lambda message: message.text == "На следующую неделю")
async def send_next_week_schedule(message: types.Message):
    user_data = load_user_data(message.from_user.id)
    if user_data:
        group = user_data['group']

        schedule_text = get_schedule_for_next_week(group)

        await message.answer(schedule_text, parse_mode="Markdown", protect_content=True)
    else:
        await message.answer("Пожалуйста, выберите свою учебную группу.", reply_markup=generate_group_keyboard())

# Функция для удаления данных о пользователе
def delete_user_data(user_id):
    if os.path.exists(user_data_file):
        with open(user_data_file, 'r') as f:
            users_data = json.load(f)

        if str(user_id) in users_data:
            del users_data[str(user_id)]

            with open(user_data_file, 'w') as f:
                json.dump(users_data, f)
            return True
    return False        

# Обработчик кнопки "Удалить Данные"
@dp.message(lambda message: message.text == "Удалить Данные")
async def delete_data(message: types.Message):
    success = delete_user_data(message.from_user.id)
    if success:
        await message.answer("Ваши данные успешно удалены. Пожалуйста, выберите свою учебную группу.", reply_markup=generate_group_keyboard())
    else:
        await message.answer("Данные не найдены. Вы уже не привязаны к группе. Выберите группу.", reply_markup=generate_group_keyboard())

################################################   КОНЕЦ РАБОТЫ С 5 КНОПКАМИ  ##################################################

# Обработчик для новых чатов
@dp.message(F.chat.type.in_({"group", "supergroup"}))
async def handle_group_chats(message: types.Message):
    await bot.leave_chat(message.chat.id)
    print(f"Бот был добавлен в группу {message.chat.title}. Он её покинул.")

# Функция для отправки уведомлений всем пользователям
async def notify_users(message: types.Message):
    try:
        admin_id1 = 549773399
        if message.from_user.id != admin_id1:
            await message.answer("У вас нет прав для выполнения этой команды.")
            return
        notification_text = message.text.replace("/notifyusers", "").strip()
        if not notification_text:
            await message.answer("Пожалуйста, укажите текст уведомления после команды.")
            return
        if not os.path.exists(user_data_file):
            await message.answer("Список пользователей пуст. Никому отправлять сообщения.")
            return
        with open(user_data_file, 'r') as f:
            users_data = json.load(f)
        for user_id in users_data.keys():
            try:
                await bot.send_message(chat_id=user_id, text=notification_text)
            except Exception as e:
                print(f"Не удалось отправить сообщение пользователю {user_id}: {e}")
        await message.answer("Уведомления успешно отправлены.")
    except Exception as e:
        await message.answer(f"Произошла ошибка: {e}")

# Функция для получения количества пользователей в базе данных
def get_users_count():
    try:
        if os.path.exists(user_data_file):
            with open(user_data_file, 'r') as f:
                users_data = json.load(f)
            return len(users_data)
        else:
            return 0
    except Exception as e:
        return f"Произошла ошибка при подсчете пользователей: {e}"

AUTHORIZED_USERS = {"1780078217", "549773399"}

# Обработчик для загрузки нового файла расписания
@dp.message(lambda message: message.document)
async def handle_uploaded_file(message: types.Message):
    try:
        if str(message.from_user.id) not in AUTHORIZED_USERS:
            await message.answer("У вас нет прав для замены файла расписания.")
            return

        document = message.document
        if not (document.file_name.endswith('.xls') or document.file_name.endswith('.xlsx')):
            await message.answer("Пожалуйста, загрузите файл в формате .xls или .xlsx.")
            return

        file_id = document.file_id
        file_info = await bot.get_file(file_id)
        downloaded_file = await bot.download_file(file_info.file_path)

        temp_file_path = os.path.join(schedule_dir, "temp_file.xls")
        with open(temp_file_path, "wb") as f:
            f.write(downloaded_file.read())

        if os.path.exists(file_path):
            os.remove(file_path)

        df = pd.read_excel(temp_file_path)
        df.to_excel(file_path, index=False)
        os.remove(temp_file_path)

        await message.answer("Файл успешно обновлен, преобразован в .xlsx и сохранен как Sched2.xlsx.")
    except Exception as e:
        await message.answer(f"Произошла ошибка при обработке файла: {e}")

# Обработчик команды /users
@dp.message(F.text.startswith("/users"))
async def handle_users_count(message: types.Message):
    user_count = get_users_count()
    if isinstance(user_count, int):
        await message.answer(f"В базе {user_count} пользователей.")
    else:
        await message.answer(f"Ошибка: {user_count}")

# Обработчик команды /notifyusers
@dp.message(F.text.startswith("/notifyusers"))
async def handle_notify_users(message: types.Message):
    await notify_users(message)

# Запуск бота
async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
