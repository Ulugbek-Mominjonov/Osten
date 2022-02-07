from datetime import datetime

from aiogram import types
from aiogram.dispatcher import FSMContext
from aiogram.types import ReplyKeyboardMarkup, ReplyKeyboardRemove

from loader import dp, bot
import keyboard as btn
import write_file

from states import Form

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
    await callback.message.answer("O'lchovchining (F.I.Sh.)")

@dp.message_handler(state=Form.sana)
async def process_data(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['Sana'] = message.text
    await Form.next()
    await message.answer("O'lchovchining (F.I.Sh.):")

# SAVDO MENEDJERI
@dp.message_handler(state=Form.menedjer)
async def process_menedjer(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data["O'lchovchining (F.I.Sh.)"] = message.text
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
        if q <= 20:
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
                await message.answer(f"{q + 1} qavat balandligini kiriting...")
            elif len(data['qavatlar']) == 15:
                data['qavatlar'][f'{q}'] = message.text
                data['state'] += 1
                await message.answer(f"{q + 1} qavat balandligini kiriting...")
            elif len(data['qavatlar']) == 16:
                data['qavatlar'][f'{q}'] = message.text
                data['state'] += 1
                await message.answer(f"{q + 1} qavat balandligini kiriting...")
            elif len(data['qavatlar']) == 17:
                data['qavatlar'][f'{q}'] = message.text
                data['state'] += 1
                await message.answer(f"{q + 1} qavat balandligini kiriting...")
            elif len(data['qavatlar']) == 18:
                data['qavatlar'][f'{q}'] = message.text
                data['state'] += 1
                await message.answer(f"{q + 1} qavat balandligini kiriting...")
            elif len(data['qavatlar']) == 19:
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
