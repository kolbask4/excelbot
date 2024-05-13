import telebot
import sqlite3
from telebot import types
from time import sleep
import excel_list
import excel_list2

bot = telebot.TeleBot('6670513034:AAEtrcpEaGcq-SEAedxb4iaCC1h9H6cHoi4')

user_state = {}
edit_msg_id = None

conn = sqlite3.connect('asd.db')
cursor = conn.cursor()

cursor.execute('''
    CREATE TABLE IF NOT EXISTS contr (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT
    )
''')

cursor.execute('''
    CREATE TABLE IF NOT EXISTS nomenkl (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        data DATE,
        tovar TEXT,
        kolvo INTEGER,
        price INTEGER,
        rashod INTEGER
    )
''')

conn.commit()
conn.close()

@bot.message_handler(commands=['start'])
def start(message):
    bot.send_message(message.chat.id, 'Добро пожаловать! Для начала выберите нужную операцию (через команды бота)')

@bot.message_handler(commands=['operation'])
def operation(message):
    global edit_msg_id
    user_id = message.chat.id
    markup = types.InlineKeyboardMarkup(row_width=2)
    markup.add(
            types.InlineKeyboardButton(text='Приход', callback_data='prihod'),
            types.InlineKeyboardButton(text='Расход', callback_data='rashod')
        )
    msg = bot.send_message(user_id, 'Выберите тип операции:', reply_markup=markup)
    edit_msg_id = msg.message_id

@bot.message_handler(commands=['document'])
def rashod(message):
    user_id = message.chat.id
    msg =  bot.send_message(user_id, 'Ваш документ в процессе создания, пожалуйста, подождите. Это займёт не более 5 секунд..')
    edit_msg_id = msg.message_id
    excel_list.create_excel_list()
    sleep(4)
    bot.edit_message_text(chat_id=user_id, message_id=edit_msg_id, text='Ваш документ готов:')
    bot.send_document(user_id, types.InputFile('report.xlsx'))

@bot.message_handler(commands=['document2'])
def rashod(message):
    user_id = message.chat.id
    msg =  bot.send_message(user_id, 'Ваш документ в процессе создания, пожалуйста, подождите. Это займёт не более 5 секунд..')
    excel_list2.create_excel_list2()
    edit_msg_id = msg.message_id
    sleep(5)
    bot.edit_message_text(chat_id=user_id, message_id=edit_msg_id, text='Ваш документ готов:')
    bot.send_document(user_id, types.InputFile('report2.xlsx'))

@bot.callback_query_handler(func=lambda call: True)
def callback_handler(call):
    global type_o
    user_id = call.message.chat.id
    if call.message:
        if call.data == 'prihod':
            type_o = 0
            bot.edit_message_text(chat_id=user_id, message_id=edit_msg_id, text='Введите дату в формате дд.мм.гггг:')
            user_state[user_id] = 'waiting_date'
            print(type_o)
        elif call.data == 'rashod':
            type_o = 1
            bot.edit_message_text(chat_id=user_id, message_id=edit_msg_id, text='Введите дату в формате дд.мм.гггг:')
            user_state[user_id] = 'waiting_date'
            print(type_o)


@bot.message_handler(func=lambda message: True)
def get_message(message):
    user_id = message.chat.id
    if user_id in user_state:
        state = user_state[user_id]
        if state == 'waiting_date':
            global date
            date_split = message.text.split('.')
            date = f'{date_split[2]}-{date_split[1]}-{date_split[0]}'
            bot.edit_message_text(chat_id=user_id, message_id=edit_msg_id, text='Введите имя поставщика:')
            bot.delete_message(user_id, message.message_id)
            user_state[user_id] = 'waiting_post'
            print(date)
        elif state == 'waiting_post':
            global post
            post = message.text
            bot.edit_message_text(chat_id=user_id, message_id=edit_msg_id, text='Введите наименовние товара, его количество и цену через пробел:')
            bot.delete_message(user_id, message.message_id)
            user_state[user_id] = 'waiting_tovar'
            print(post)
        elif state == 'waiting_tovar':
            global tovar, kolvo, price
            all_items = message.text.split(' ')
            tovar = all_items[0]
            kolvo = int(all_items[1])
            price = int(all_items[2])
            bot.edit_message_text(chat_id=user_id, message_id=edit_msg_id, text='Успешно сохранено в базу данных =)')
            bot.delete_message(user_id, message.message_id)
            del user_state[user_id]
            print(tovar, kolvo, price)

            add_prihod(date, post, tovar, kolvo, price, type_o)

def add_prihod(date, post, tovar, kolvo, price, type_o):
    conn = sqlite3.connect('asd.db')
    cursor = conn.cursor()

    cursor.execute('INSERT INTO contr (name) VALUES (?)', (post,))
    cursor.execute('INSERT INTO nomenkl (data, tovar, kolvo, price, rashod) VALUES (?, ?, ?, ?, ?)', (date, tovar, kolvo, price, type_o))

    conn.commit()
    conn.close()


bot.polling()