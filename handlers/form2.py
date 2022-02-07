from datetime import datetime

from aiogram import types
from aiogram.dispatcher import FSMContext
from aiogram.types import ReplyKeyboardRemove

from loader import dp
import keyboard as btn
import write_file

from states import Form2

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
    msg = await message.answer('.', reply_markup=ReplyKeyboardRemove())
    await msg.delete()
    await message.answer("Дата:", reply_markup=btn.get_data)

@dp.message_handler(state=Form2.telefon)
async def process_telefon(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['Номер телефона'] = message.text
    await Form2.next()
    msg = await message.answer('.', reply_markup=ReplyKeyboardRemove())
    await msg.delete()
    await message.answer("Дата:", reply_markup=btn.get_data)

@dp.callback_query_handler(text='sana', state=Form2.data)
async def set_date(callback: types.CallbackQuery, state: FSMContext):
    async with state.proxy() as data:
        data['Дата'] = datetime.today().strftime('%Y.%m.%d')
    await Form2.next()
    await callback.message.answer("Проём двери:")

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
