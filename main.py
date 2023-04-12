import telebot
from tokens import token1
from telebot import types  # для указание типов
from icalendar import Calendar, Event
import glob
import pendulum
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

tok = token1
RUZ_URL = "https://ruz.hse.ru/ruz/main"
bot = telebot.TeleBot(tok)
folder_path = 'C:\\Users\\79186\\Downloads'


@bot.message_handler(commands=['start'])
def start_message(message):
    start = bot.send_message(message.chat.id,
                             "Привет, этот бот показывает расписание студентов ВШЭ. Для того, чтобы получить его, введите ФИО. Если захотите сменить профиль, введите /start и ФИО заново.")
    bot.register_next_step_handler(start, save_mail)



def save_mail(message):
    user_mail = message.text
    if user_mail == "/start":
        start = bot.send_message(message.chat.id, "Введите ФИО заново:")
        bot.register_next_step_handler(start, save_mail)
    else:
        option = Options()
        # option.add_argument('-headless')
        driver = webdriver.Chrome(executable_path='C:\Chrome\chromedriver.exe', chrome_options=option)
        driver.get(RUZ_URL)

        wait = WebDriverWait(driver, 10)
        button = wait.until(EC.visibility_of_element_located((By.XPATH, '//button[contains(text(), "Студент")]')))
        button.click()
        wait = WebDriverWait(driver, 10)
        input_field = wait.until(EC.visibility_of_element_located((By.ID, 'autocomplete-student')))
        input_field.send_keys(user_mail)
        wait = WebDriverWait(driver, 10)
        ul_element = wait.until(EC.visibility_of_element_located((By.ID, 'pr_id_1_list')))
        if ul_element.text == "Не найдено":
            reply = bot.send_message(message.chat.id,
                                     "Либо вы ввели ваше ФИО неправильно, либо вы не являетесь студентов ВШЭ. Повторите ввод ФИО:")
            bot.register_next_step_handler(reply, save_mail)
        else:
            ul_element.click()
            time.sleep(1)
            wait = WebDriverWait(driver, 10)
            data_export = wait.until(EC.visibility_of_element_located((By.XPATH, '//button[contains(text(), "Экспорт")]')))
            data_export.click()
            time.sleep(1)
            file_path = os.path.join(folder_path, '*')
            file = sorted(glob.iglob(file_path), key=os.path.getctime, reverse=True)
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            btn1 = types.KeyboardButton("Расписание на сегодня")
            btn2 = types.KeyboardButton("Расписание на завтра")
            btn3 = types.KeyboardButton("Расписание на неделю")
            markup.add(btn1, btn2, btn3)
            message = bot.send_message(message.chat.id, "Выберите, что вам нужно:", reply_markup=markup)
            k = file[0]
            bot.register_next_step_handler(message, give_data, k)


def give_data(message, k):
    g = open(k, 'rb')
    calendar = Calendar.from_ical(g.read())
    answer_today = "Вот ваше расписание на сегодня:\n"
    answer_tomorrow = "Вот ваше расписание на завтра:\n"
    answer_week = "Вот ваше расписание на текущую неделю:\n"
    answer_week_res = answer_week
    answer_today_res = answer_today
    answer_tomorrow_res = answer_tomorrow
    for component in calendar.walk():
        if component.name == "VEVENT":
            description = component.get('description')
            description = description.replace("\n \n \n", "\n")
            start_lesson = component.get('dtstart')
            start_lesson_ymd = "{}".format(start_lesson.dt)
            start_lesson_ymd = start_lesson_ymd[:10]
            today = "{}".format(pendulum.today())
            today = today[:10]
            tomorrow = "{}".format(pendulum.tomorrow())
            tomorrow = tomorrow[:10]
            end = component.get('dtend')
            answer_week = answer_week + description + "\n" + "Начало: {}".format(
                start_lesson.dt) + "\n" + "Конец: {}".format(end.dt) + "\n\n"
            if start_lesson_ymd == today:
                answer_today = answer_today + description + "\n" + "Начало: {}".format(
                    start_lesson.dt) + "\n" + "Конец: {}".format(end.dt) + "\n\n"
            elif start_lesson_ymd == tomorrow:
                answer_tomorrow = answer_tomorrow + description + "\n" + "Начало: {}".format(
                    start_lesson.dt) + "\n" + "Конец: {}".format(end.dt) + "\n\n"
    g.close()
    os.remove(k)
    if answer_today_res == answer_today:
        answer_today = "Сегодня занятий нет"
    if answer_tomorrow_res == answer_tomorrow:
        answer_tomorrow = "Завтра занятий нет"
    if answer_week_res == answer_week:
        answer_week = "На неделе занятий нет"
    user_answer = message.text
    if user_answer == "Расписание на сегодня":
        bot.send_message(message.chat.id, answer_today)
    elif user_answer == "Расписание на завтра":
        bot.send_message(message.chat.id, answer_tomorrow)
    elif user_answer == "Расписание на неделю":
        bot.send_message(message.chat.id, answer_week)
    elif user_answer == "/start":
        start = bot.send_message(message.chat.id, "Введите ФИО заново:")
        bot.register_next_step_handler(start, save_mail)
    # driver.quit()


if __name__ == '__main__':
    bot.infinity_polling()
