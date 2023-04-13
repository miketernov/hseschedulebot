import telebot
from tokens import token1
from telebot import types
from icalendar import Calendar, Event
import glob
import codecs
import pendulum
import os
import difflib as difflib
from threading import Timer
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

tok = token1
RUZ_URL = "https://ruz.hse.ru/ruz/main"
bot = telebot.TeleBot(tok)
folder_path = r'C:\Users\test_\Downloads'
df = pd.DataFrame(columns=['Name', 'path_to_file'])


@bot.message_handler(commands=['start'])
def start_message(message):
    start = bot.send_message(message.chat.id,
                             "Привет, этот бот показывает расписание студентов ВШЭ. Для того, чтобы получить его, введите ФИО. Если захотите сменить профиль, введите /start и ФИО заново.")
    bot.register_next_step_handler(start, save_mail)


@bot.callback_query_handler(func=lambda call: True)
def query_handler(call):
    if call.data == "today_button":
        give_data_today(call.message)
    elif call.data == "tomorrow_button":
        give_data_tomorrow(call.message)
    elif call.data == "week_button":
        give_data_week(call.message)


def get_path():
    file_path = os.path.join(folder_path, '*')
    file = sorted(glob.iglob(file_path), key=os.path.getctime, reverse=True)
    return file[0]


def get_changes(message):
    read_changes = pd.read_excel("people.xlsx")
    for k in read_changes['Name']:
        old_path = read_changes.loc[read_changes['Name'] == k, 'path_to_file'].iloc[0]
        option = Options()
        # option.add_argument('-headless')
        driver = webdriver.Chrome(executable_path=r'C:\Chrome\chromedriver.exe', chrome_options=option)
        driver.get(RUZ_URL)
        wait = WebDriverWait(driver, 10)
        time.sleep(1)
        button = wait.until(EC.visibility_of_element_located((By.XPATH, '//button[contains(text(), "Студент")]')))
        button.click()
        time.sleep(1)
        wait = WebDriverWait(driver, 10)
        input_field = wait.until(EC.visibility_of_element_located((By.ID, 'autocomplete-student')))
        input_field.send_keys(k)
        wait = WebDriverWait(driver, 10)
        time.sleep(1)
        ul_element = wait.until(EC.visibility_of_element_located((By.ID, 'pr_id_1_list')))
        ul_element.click()
        time.sleep(2)
        wait = WebDriverWait(driver, 10)
        data_export = wait.until(
            EC.visibility_of_element_located((By.XPATH, '//button[contains(text(), "Экспорт")]')))
        data_export.click()
        time.sleep(1)
        driver.quit()
        new_path = get_path()
        file_1 = old_path
        file_2 = new_path
        filename_1 = os.path.splitext(file_1)[0]
        filename_2 = os.path.splitext(file_2)[0]
        old_path = filename_1 + ".txt"
        new_path = filename_2 + ".txt"
        os.rename(file_1, filename_1 + ".txt")
        os.rename(file_2, filename_2 + ".txt")

        file_1 = codecs.open(old_path,
                             "r", "utf_8_sig")
        file_2 = codecs.open(new_path, "r",
                             "utf_8_sig")
        file_1_line = file_1.readline()
        file_2_line = file_2.readline()

        new_file_with_changes = open(
            r'C:\Users\test_\Downloads\new_file_for' + k + '.txt', 'w+')
        while file_1_line != '' or file_2_line != '':
            file_1_line = file_1_line.rstrip()
            file_2_line = file_2_line.rstrip()
            if file_1_line != file_2_line:
                if file_2_line != '':
                    new_file_with_changes.write(file_2_line + "\n")
            file_1_line = file_1.readline()
            file_2_line = file_2.readline()

        file_1.close()
        file_2.close()
        new_file_with_changes.close()
        os.remove(old_path)
        os.remove(new_path)
        row_index = df.loc[df['Name'] == k].index
        read_changes.loc[row_index, 'path_to_file'] = r'C:\Users\test_\Downloads\new_file_for' + k + '.txt'
        file_5 = r'C:\Users\test_\Downloads\new_file_for' + k + '.txt'
        filename_5 = os.path.splitext(file_5)[0]
        os.rename(file_5, filename_5 + ".ics")
        file_6 = filename_5 + ".ics"
        g = open(file_6, 'rb')
        calendar = Calendar.from_ical(g.read())
        answer_week = "Внимание. У вас изменения:\n"
        for component in calendar.walk():
            if component.name == "VEVENT":
                description = component.get('description')
                description = description.replace("\n \n \n", "\n")
                start_lesson = component.get('dtstart')
                end = component.get('dtend')
                answer_week = answer_week + description + "\n" + "Начало: {}".format(
                    start_lesson.dt) + "\n" + "Конец: {}".format(end.dt) + "\n\n"
        bot.send_message(message.chat.id, answer_week)
    Timer(86400, get_changes).start()
        # new_df = [user_mail, get_path()]
        # if df.empty:
        #     df.loc[len(df)] = new_df
        # else:
        #     if user_mail in df['Name'].tolist():
        #         row_index = df.loc[df['Name'] == user_mail].index
        #         df.loc[row_index, 'path_to_file'] = get_path()
        #     else:
        #         df.loc[len(df)] = new_df
        # df.to_excel('people.xlsx')


def save_mail(message):
    user_mail = message.text
    if user_mail == "/start":
        start = bot.send_message(message.chat.id, "Введите ФИО заново:")
        bot.register_next_step_handler(start, save_mail)
    else:
        option = Options()
        # option.add_argument('-headless')
        driver = webdriver.Chrome(executable_path=r'C:\Chrome\chromedriver.exe', chrome_options=option)
        driver.get(RUZ_URL)
        wait = WebDriverWait(driver, 10)
        time.sleep(1)
        button = wait.until(EC.visibility_of_element_located((By.XPATH, '//button[contains(text(), "Студент")]')))
        button.click()
        time.sleep(1)
        wait = WebDriverWait(driver, 10)
        input_field = wait.until(EC.visibility_of_element_located((By.ID, 'autocomplete-student')))
        input_field.send_keys(user_mail)
        wait = WebDriverWait(driver, 10)
        time.sleep(1)
        ul_element = wait.until(EC.visibility_of_element_located((By.ID, 'pr_id_1_list')))
        if ul_element.text == "Не найдено":
            reply = bot.send_message(message.chat.id,
                                     "Либо вы ввели ваше ФИО неправильно, либо вы не являетесь студентов ВШЭ. Повторите ввод ФИО:")
            bot.register_next_step_handler(reply, save_mail)
        else:
            ul_element.click()
            time.sleep(1)
            wait = WebDriverWait(driver, 10)
            data_export = wait.until(
                EC.visibility_of_element_located((By.XPATH, '//button[contains(text(), "Экспорт")]')))
            data_export.click()
            time.sleep(1)
            driver.quit()
            new_df = [user_mail, get_path()]
            if df.empty:
                df.loc[len(df)] = new_df
            else:
                if user_mail in df['Name'].tolist():
                    row_index = df.loc[df['Name'] == user_mail].index
                    df.loc[row_index, 'path_to_file'] = get_path()
                else:
                    df.loc[len(df)] = new_df
            df.to_excel('people.xlsx')
            markup = types.InlineKeyboardMarkup(row_width=1)
            btn1 = types.InlineKeyboardButton("Расписание на сегодня", callback_data="today_button")
            btn2 = types.InlineKeyboardButton("Расписание на завтра", callback_data="tomorrow_button")
            btn3 = types.InlineKeyboardButton("Расписание на неделю", callback_data="week_button")
            markup.add(btn1, btn2, btn3)
            bot.send_message(message.chat.id, "Выберите, что вам нужно:", reply_markup=markup)


def give_data_today(message):
    g = open(get_path(), 'rb')
    calendar = Calendar.from_ical(g.read())
    answer_today = "Вот ваше расписание на сегодня:\n"
    answer_today_res = answer_today
    for component in calendar.walk():
        if component.name == "VEVENT":
            description = component.get('description')
            description = description.replace("\n \n \n", "\n")
            start_lesson = component.get('dtstart')
            start_lesson_ymd = "{}".format(start_lesson.dt)
            start_lesson_ymd = start_lesson_ymd[:10]
            today = "{}".format(pendulum.today())
            today = today[:10]
            end = component.get('dtend')
            if start_lesson_ymd == today:
                answer_today = answer_today + description + "\n" + "Начало: {}".format(
                    start_lesson.dt) + "\n" + "Конец: {}".format(end.dt) + "\n\n"
    g.close()
    if answer_today_res == answer_today:
        answer_today = "Сегодня занятий нет"
    bot.send_message(message.chat.id, answer_today)
    # driver.quit()

    # bot.register_next_step_handler(message, reply_message, answer_today, answer_tomorrow, answer_week)
    # user_answer = message.text


def give_data_tomorrow(message):
    answer_tomorrow = "Вот ваше расписание на завтра:\n"
    answer_tomorrow_res = answer_tomorrow
    g = open(get_path(), 'rb')
    calendar = Calendar.from_ical(g.read())
    for component in calendar.walk():
        if component.name == "VEVENT":
            description = component.get('description')
            description = description.replace("\n \n \n", "\n")
            start_lesson = component.get('dtstart')
            start_lesson_ymd = "{}".format(start_lesson.dt)
            start_lesson_ymd = start_lesson_ymd[:10]
            tomorrow = "{}".format(pendulum.tomorrow())
            tomorrow = tomorrow[:10]
            end = component.get('dtend')
            if start_lesson_ymd == tomorrow:
                answer_tomorrow = answer_tomorrow + description + "\n" + "Начало: {}".format(
                    start_lesson.dt) + "\n" + "Конец: {}".format(end.dt) + "\n\n"
    g.close()
    if answer_tomorrow_res == answer_tomorrow:
        answer_tomorrow = "Завтра занятий нет"
    bot.send_message(message.chat.id, answer_tomorrow)


def give_data_week(message):
    answer_week = "Вот ваше расписание на текущую неделю:\n"
    answer_week_res = answer_week
    g = open(get_path(), 'rb')
    calendar = Calendar.from_ical(g.read())
    for component in calendar.walk():
        if component.name == "VEVENT":
            description = component.get('description')
            description = description.replace("\n \n \n", "\n")
            start_lesson = component.get('dtstart')
            end = component.get('dtend')
            answer_week = answer_week + description + "\n" + "Начало: {}".format(
                start_lesson.dt) + "\n" + "Конец: {}".format(end.dt) + "\n\n"
    if answer_week_res == answer_week:
        answer_week = "На неделе занятий нет"
    bot.send_message(message.chat.id, answer_week)


if __name__ == '__main__':
    bot.infinity_polling()
