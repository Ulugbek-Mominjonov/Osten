import logging

from datetime import datetime

from aiogram import Bot, Dispatcher, executor, types
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters import Text
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram.types import ReplyKeyboardMarkup, ReplyKeyboardRemove

import keyboard as btn
import write_file

API_TOKEN = '5024350132:AAHLxhib-RnMOBGpQWHLepX4UIIg5a9MwvU'

# Configure logging
logging.basicConfig(level=logging.INFO)

# Initialize bot and dispatcher
bot = Bot(token=API_TOKEN)
storage = MemoryStorage()
dp = Dispatcher(bot, storage=storage)

class Form(StatesGroup):
    loyiha_nomi = State()
    sana = State()
    menedjer = State()
    shahta_turi = State()
    mrlOrRl = State()
    olcham = State()
    qavat = State()
    eshiklar_soni = State()
    P = State()
    Hp = State()
    Eslatma = State()

class Form2(StatesGroup):
    loyiha_nomi = State()
    adres = State()
    injiner = State()
    montajnik = State()
    telefon = State()
    data = State()
    eshik = State()
    kenglik = State()
    qavat = State()

@dp.message_handler(commands=['start', 'info'])
async def send_welcome(message: types.Message):
    await message.answer("Assalomu alaykum O'lchamlar olish botiga hush kelibsiz!!!", reply_markup=btn.menu)


@dp.message_handler()
async def startState(message: types.Message, state: FSMContext):
    if message.text == "Shahta o'lchamlarini olish":
        await message.reply("Siz Shahta o'lchamlarini olish botni tanladingiz.\nO'lchamlarni olish tugmani bosing", reply_markup=btn.osten)
    elif message.text == "O'lchamlarni olish":
        await Form.loyiha_nomi.set()
        await message.answer("Loyiha nomini kiritng:", reply_markup=ReplyKeyboardRemove())
    elif message.text == "Obramleniya o'lchamlarini olish":
        await message.reply("Siz Obramleniya o'lchamlarini olish botni tanladingiz.\nMa'lumotni kiritish tugmani bosing", reply_markup=btn.osten2)
    elif message.text == "Ma'lumotlarni kiritish":
        await Form2.loyiha_nomi.set()
        await message.reply("Лойиҳа номи:", reply_markup=ReplyKeyboardRemove())
    elif message.text == "Bosh sahifa":
        await message.answer("Assalomu alaykum O'lchamlar olish botiga hush kelibsiz!!!", reply_markup=btn.menu)


# Osten
@dp.message_handler(state='*', commands='cancel')
@dp.message_handler(Text(equals='cancel', ignore_case=True), state='*')
async def cancel_handler(message: types.Message, state: FSMContext):
    current_state = await state.get_state()
    if current_state is None:
        return

    logging.info('Cancelling state %r', current_state)
    await state.finish()
    await message.answer('Cancelled.', reply_markup=types.ReplyKeyboardRemove())


# LOYIHA NOMI
@dp.message_handler(state=Form.loyiha_nomi)
async def process_name(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['Loyiha_nomi'] = message.text
    await Form.next()
    await message.answer("Loyiha sanasi:", reply_markup=btn.get_data)

# LOYIHA SANASI
@dp.callback_query_handler(text='sana', state=Form.sana)
async def set_date(callback: types.CallbackQuery, state: FSMContext):
    async with state.proxy() as data:
        data['Sana'] = datetime.today().strftime('%Y.%m.%d')
    await Form.next()
    await callback.message.answer("Savdo menedjeri:")

@dp.message_handler(state=Form.sana)
async def process_data(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['Sana'] = message.text
    await Form.next()
    await message.answer("Savdo menedjeri:")

# SAVDO MENEDJERI
@dp.message_handler(state=Form.menedjer)
async def process_menedjer(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['Menedjer'] = message.text
    await Form.next()

    markup = ReplyKeyboardMarkup(resize_keyboard=True, selective=True)
    markup.add("Beton", "Bloklar")
    markup.add("Po'lat Konstruksiya")

    await message.answer("Shahta turini tanlang:", reply_markup=markup)

# SHAHTA TURI
@dp.message_handler(lambda message: message.text not in ["Beton", "Bloklar", "Po'lat Konstruksiya"],
                    state=Form.shahta_turi)
async def process_type_invalid(message: types.Message):
    return await message.answer("Yaroqsiz shahta turi. Keyboarddagilardan birini tanlang:")

@dp.message_handler(state=Form.shahta_turi)
async def process_type(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['Shahta_turi'] = message.text
    await Form.next()

    markup = ReplyKeyboardMarkup(resize_keyboard=True, selective=True)
    markup.add("MR", "MRL")
    await message.answer("MR yoki MRL, birini tanlang:", reply_markup=markup)

# RL MRL
@dp.message_handler(lambda message: message.text not in ["MR", "MRL"], state=Form.mrlOrRl)
async def process_RlOrMrl_invalid(message: types.Message):
    return await message.answer("Yaroqsiz ma'lumot!!!")

@dp.message_handler(state=Form.mrlOrRl)
async def process_RlOrMrl(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['Mrl_Or_Rl'] = message.text
        data['olchamlar'] = {}
    await Form.next()
    await bot.send_photo(chat_id=message.chat.id, photo=open('rasm/rasm_1.jpg', 'rb'),
                         caption="Rasm asosida o'lchamlarni kiriting")
    await bot.send_message(message.chat.id, "A ni kiriting", reply_markup=ReplyKeyboardRemove())

# OLCHMALAR
@dp.message_handler(lambda message: not message.text.isdigit(), state=Form.olcham)
async def process_A_invalid(message: types.Message):
    return await message.answer("Raqam kiriting: Misol uchun: 20")

@dp.message_handler(state=Form.olcham)
async def process_RlOrMrl(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        if len(data['olchamlar']) == 0:
            data['olchamlar']['A'] = message.text
            await message.answer("B ni kiriting...")
        elif len(data['olchamlar']) == 1:
            data['olchamlar']['B'] = message.text
            await message.answer("C ni kiriting...")
        elif len(data['olchamlar']) == 2:
            data['olchamlar']['C'] = message.text
            await message.answer("D ni kiriting...")
        elif len(data['olchamlar']) == 3:
            data['olchamlar']['D'] = message.text
            await message.answer("E ni kiriting...")
        elif len(data['olchamlar']) == 4:
            data['olchamlar']['E'] = message.text
            data['state'] = 1
            data['qavatlar'] = {}
            await Form.next()
            await bot.send_photo(chat_id=message.chat.id, photo=open('rasm/rasm_2.jpg', 'rb'),
                             caption="Rasm asosida uning o'lchamlarini kiriting: ")
            await bot.send_message(message.chat.id, "1 qavat balandligi: ")

# QAVATLAR
@dp.message_handler(lambda message: message.text=="Tugatish", state=Form.qavat)
async def process_2_invalid(message: types.Message, state: FSMContext):
    await Form.eshiklar_soni.set()
    await message.answer("Eshiklar sonini kiriting: ", reply_markup=ReplyKeyboardRemove())

@dp.message_handler(lambda message: not message.text.isdigit(), state=Form.qavat)
async def process_E_invalid(message: types.Message):
    return await message.answer("Raqam kiriting: Misol uchun: 20")

@dp.message_handler(state=Form.qavat)
async def process_telefon(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        q = data['state']
        if q <= 15:
            if len(data['qavatlar']) == 0:
                data['qavatlar'][f'{q}'] = message.text
                data['state'] += 1
                await message.answer(f"{q + 1} qavat balandligini kiriting...", reply_markup=btn.keyboard2)
            elif len(data['qavatlar']) == 1:
                data['qavatlar'][f'{q}'] = message.text
                data['state'] += 1
                await message.answer(f"{q + 1} qavat balandligini kiriting...")
            elif len(data['qavatlar']) == 2:
                data['qavatlar'][f'{q}'] = message.text
                data['state'] += 1
                await message.answer(f"{q + 1} qavat balandligini kiriting...")
            elif len(data['qavatlar']) == 3:
                data['qavatlar'][f'{q}'] = message.text
                data['state'] += 1
                await message.answer(f"{q + 1} qavat balandligini kiriting...")
            elif len(data['qavatlar']) == 4:
                data['qavatlar'][f'{q}'] = message.text
                data['state'] += 1
                await message.answer(f"{q + 1} qavat balandligini kiriting...")
            elif len(data['qavatlar']) == 5:
                data['qavatlar'][f'{q}'] = message.text
                data['state'] += 1
                await message.answer(f"{q + 1} qavat balandligini kiriting...")
            elif len(data['qavatlar']) == 6:
                data['qavatlar'][f'{q}'] = message.text
                data['state'] += 1
                await message.answer(f"{q + 1} qavat balandligini kiriting...")
            elif len(data['qavatlar']) == 7:
                data['qavatlar'][f'{q}'] = message.text
                data['state'] += 1
                await message.answer(f"{q + 1} qavat balandligini kiriting...")
            elif len(data['qavatlar']) == 8:
                data['qavatlar'][f'{q}'] = message.text
                data['state'] += 1
                await message.answer(f"{q + 1} qavat balandligini kiriting...")
            elif len(data['qavatlar']) == 9:
                data['qavatlar'][f'{q}'] = message.text
                data['state'] += 1
                await message.answer(f"{q + 1} qavat balandligini kiriting...")
            elif len(data['qavatlar']) == 9:
                data['qavatlar'][f'{q}'] = message.text
                data['state'] += 1
                await message.answer(f"{q + 1} qavat balandligini kiriting...")
            elif len(data['qavatlar']) == 10:
                data['qavatlar'][f'{q}'] = message.text
                data['state'] += 1
                await message.answer(f"{q + 1} qavat balandligini kiriting...")
            elif len(data['qavatlar']) == 11:
                data['qavatlar'][f'{q}'] = message.text
                data['state'] += 1
                await message.answer(f"{q + 1} qavat balandligini kiriting...")
            elif len(data['qavatlar']) == 12:
                data['qavatlar'][f'{q}'] = message.text
                data['state'] += 1
                await message.answer(f"{q + 1} qavat balandligini kiriting...")
            elif len(data['qavatlar']) == 13:
                data['qavatlar'][f'{q}'] = message.text
                data['state'] += 1
                await message.answer(f"{q + 1} qavat balandligini kiriting...")
            elif len(data['qavatlar']) == 14:
                data['qavatlar'][f'{q}'] = message.text
                data['state'] += 1
                await Form.next()
                await message.answer("Eshiklar sonini kiriting: ", reply_markup=types.ReplyKeyboardRemove())

@dp.message_handler(lambda message: not message.text.isdigit(), state=Form.eshiklar_soni)
async def process_E_invalid(message: types.Message):
    return await message.answer("Iltimos Raqam kiriting: Misol uchun: 20")

@dp.message_handler(state=Form.eshiklar_soni)
async def process_eshik(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['eshiklar_soni'] = message.text
    await Form.next()
    await message.answer("Priyamka balandligini kiriting: ")

@dp.message_handler(lambda message: not message.text.isdigit(), state=Form.P)
async def process_E_invalid(message: types.Message):
    return await message.answer("Iltimos Raqam kiriting: Misol uchun: 20")

@dp.message_handler(state=Form.P)
async def process_P(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['P'] = message.text
    await Form.next()
    await message.answer("Oxirgi etaj balandligini kiriting: ")

@dp.message_handler(lambda message: not message.text.isdigit(), state=Form.Hp)
async def process_E_invalid(message: types.Message):
    return await message.answer("Iltimos Raqam kiriting: Misol uchun: 20")

@dp.message_handler(state=Form.Hp)
async def process_HP(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['Hp'] = message.text
    await Form.next()
    await message.answer("Eslatma kiriting: ")

@dp.message_handler(state=Form.Eslatma)
async def process_Eslatma(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['Eslatma'] = message.text
    keys = []
    for item in data.keys():
        keys.append(item)

    # write image
    write_file.write_img(data)

    # write excel
    write_file.write_excel_osten(keys, data)

    await state.finish()
    await message.answer("Ma'lumotlar tayyorlanyapti....")
    await bot.send_photo(chat_id=message.chat.id, photo=open('salom.png', 'rb'))
    await message.answer_document(open(f"{data['Loyiha_nomi']}.xlsx", 'rb'), reply_markup=btn.Bosh_sahifa)


# Birnima
@dp.message_handler(state=Form2.loyiha_nomi)
async def process_adress(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['Лойиҳа номи'] = message.text
    await Form2.next()
    await message.answer("Адрес:", reply_markup=btn.adress)


@dp.message_handler(state=Form2.adres)
async def process_adress(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['Адрес'] = message.text
    await Form2.next()
    await message.answer("Инженер технолог:", reply_markup=ReplyKeyboardRemove())

@dp.message_handler(state=Form2.adres, content_types=types.ContentType.LOCATION)
async def process_adress(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['Адрес'] = message.location.live_period
    await Form2.next()
    await message.answer("Инженер технолог:", reply_markup=ReplyKeyboardRemove())


@dp.message_handler(state=Form2.injiner)
async def process_injiner(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['Инженер технолог'] = message.text
    await Form2.next()
    await message.answer("Монтажник:")

@dp.message_handler(state=Form2.montajnik)
async def process_montajnik(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['Монтажник'] = message.text
    await Form2.next()
    await message.answer("Номер телефона:", reply_markup=btn.telefon)

@dp.message_handler(state=Form2.telefon, content_types=types.ContentType.CONTACT)
async def process_telefon(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['Номер телефона'] = message.contact.phone_number
    await Form2.next()
    await message.answer("Дата:", reply_markup=btn.get_data)

@dp.message_handler(state=Form2.telefon)
async def process_telefon(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['Номер телефона'] = message.text
    await Form2.next()
    await message.answer("Дата:", reply_markup=btn.get_data)

@dp.callback_query_handler(text='sana', state=Form2.data)
async def set_date(callback: types.CallbackQuery, state: FSMContext):
    async with state.proxy() as data:
        data['Дата'] = datetime.today().strftime('%Y.%m.%d')
    await Form2.next()
    await callback.message.answer("Проём двери:", reply_markup=ReplyKeyboardRemove())

@dp.message_handler(state=Form2.data)
async def process_date(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['Дата'] = message.text
    await Form2.next()
    await message.answer("Проём двери:")

@dp.message_handler(lambda message: not message.text.isdigit(), state=Form2.eshik)
async def process_E_invalid(message: types.Message):
    return await message.answer("Iltimos Raqam kiriting: Misol uchun: 20")

@dp.message_handler(state=Form2.eshik)
async def process_eshik(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['Проём двери'] = message.text
    await Form2.next()
    await message.answer("Ширина парога:")

@dp.message_handler(lambda message: not message.text.isdigit(), state=Form2.kenglik)
async def process_E_invalid(message: types.Message):
    return await message.answer("Iltimos Raqam kiriting: Misol uchun: 20")

@dp.message_handler(state=Form2.kenglik)
async def process_kenglik(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['Ширина парога'] = message.text
        data['state'] = 1
        for i in range(1, 19):
            data[f"qavat_{i}"] = {}
    await Form2.next()
    await message.answer_photo(photo=open('rasm/rasm_2_bot.png', 'rb'), caption="Rasm asosida o'lchamlarni kiriting")
    await message.answer("1 qavat ma'lumotlarini kiritsh:\nA ni kiriting...")

@dp.message_handler(lambda message: not message.text.isdigit(), state=Form2.qavat)
async def process_E_invalid(message: types.Message):
    return await message.answer("Iltimos Raqam kiriting: Misol uchun: 20")

@dp.callback_query_handler(text='stop', state=Form2.qavat)
async def set_date(callback: types.CallbackQuery, state: FSMContext):
    async with state.proxy() as data:
        # write excel
        write_file.write_excel_birnima(data)
    await state.finish()
    await callback.message.answer("Ma'lumot tayyorlanyapti...")
    await callback.message.answer_document(open(f"{data['Лойиҳа номи']}.xlsx", 'rb'), reply_markup=btn.menu)

@dp.callback_query_handler(text='next', state=Form2.qavat)
async def set_date(callback: types.CallbackQuery, state: FSMContext):
    async with state.proxy() as data:
        await callback.message.answer(f"{data['state']} qavat ma'lumotlarini kiritsh:\nA ni kiriting...")

@dp.message_handler(lambda message: not message.text.isdigit(), state=Form2.qavat)
async def process_qavat_invalid_2(message: types.Message):
    return await message.answer("Raqam kiriting: Misol uchun: 20")

# Qavat
@dp.message_handler(state=Form2.qavat)
async def process_qavat(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        q = data['state']
        if q <= 18:
            if len(data[f'qavat_{q}']) == 0:
                data[f'qavat_{q}']["A"] = message.text
                await message.answer("B ni kiriting...")
            elif len(data[f'qavat_{q}']) == 1:
                data[f'qavat_{q}']["B"] = message.text
                await message.answer("C ni kiriting...")
            elif len(data[f'qavat_{q}']) == 2:
                data[f'qavat_{q}']["C"] = message.text
                await message.answer("D ni kiriting...")
            elif len(data[f'qavat_{q}']) == 3:
                data[f'qavat_{q}']["D"] = message.text
                await message.answer("E ni kiriting...")
            elif len(data[f'qavat_{q}']) == 4:
                data[f'qavat_{q}']["E"] = message.text
                await message.answer("F ni kiriting...")
            elif len(data[f'qavat_{q}']) == 5:
                data[f'qavat_{q}']["F"] = message.text
                data['state'] += 1
                if q < 18:
                    await message.answer(f"{q}-qavat ma'lumotlari qabul qilindi", reply_markup=btn.osten2Toxtatish)
                else:
                    write_file.write_excel_birnima(data)
                    await state.finish()
                    await message.answer("Ma'lumot tayyorlanyapti...")
                    await message.answer_document(open(f"{data['Лойиҳа номи']}.xlsx", 'rb'), reply_markup=btn.menu)


if __name__ == '__main__':
    executor.start_polling(dp, skip_updates=True)
