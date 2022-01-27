from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardButton, InlineKeyboardMarkup

olcham = KeyboardButton("O'lchamlarni olish")
malumot = KeyboardButton("Ma'lumotlarni kiritish")
button2 = KeyboardButton("Tugatish")
boshBtn = KeyboardButton("Bosh sahifa")
ostenBtn1 = KeyboardButton("Shahta o'lchamlarini olish")
ostenBtn2 = KeyboardButton("Obramleniya o'lchamlarini olish")
telefonBtn = KeyboardButton("Mening telefon raqamim", request_contact=True)
adressBtn = KeyboardButton("Joylashuvni yuborish", request_location=True)

# inline
sana = InlineKeyboardButton(text="Bugungi sana", callback_data='sana')
toxtatishBtn = InlineKeyboardButton(text="To'xtatish", callback_data='stop')
keyingi = InlineKeyboardButton(text="Davom ettirish", callback_data='next')

osten = ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True).add(olcham)
osten2 = ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True).add(malumot)
keyboard2 = ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True).add(button2)
Bosh_sahifa = ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True).add(boshBtn)
menu = ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True).add(ostenBtn1).add(ostenBtn2)
telefon = ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True).add(telefonBtn)
osten2Toxtatish = InlineKeyboardMarkup(row_width=1).add(toxtatishBtn, keyingi)
adress = ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True).add(adressBtn)
get_data = InlineKeyboardMarkup(row_width=1).add(sana)