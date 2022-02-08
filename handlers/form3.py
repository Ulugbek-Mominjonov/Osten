
from datetime import datetime

from aiogram import types

from aiogram.dispatcher import FSMContext

from loader import dp
import keyboard as btn
import write_file

from states import Form3

arr = {}

# osten3
@dp.message_handler(state=Form3.loyiha_nomi)
async def process_loyiha(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['Loyiha nomi'] = message.text
        data['state'] = 1
        # write_file.write_excel_osten3(data)
        for i in range(1, 21):
            data[f"qavat_{i}"] = {}
    await Form3.next()
    await message.answer("1-etaj malumotlarini kiriting:\nChap A...")


@dp.message_handler(lambda message: not message.text.isdigit(), state=Form3.qavat)
async def process_qavat_invalid(message: types.Message):
    return await message.answer("Iltimos Raqam kiriting: Misol uchun: 20")


@dp.callback_query_handler(text='stop', state=Form3.qavat)
async def stop_state(callback: types.CallbackQuery, state: FSMContext):
    await Form3.next()
    await callback.message.answer("Sim o'rtasi:")


@dp.callback_query_handler(text='next', state=Form3.qavat)
async def set_date(callback: types.CallbackQuery, state: FSMContext):
    async with state.proxy() as data:
        await callback.message.answer(f"{data['state']} qavat ma'lumotlarini kiritsh:\n\nChap A...")


@dp.message_handler(state=Form3.qavat)
async def process_qavat(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        q = data['state']
        if q <= 20:
            if len(data[f'qavat_{q}']) == 0:
                data[f'qavat_{q}']["Chap A"] = int(message.text)
                await message.answer("Chap B...")
            elif len(data[f'qavat_{q}']) == 1:
                data[f'qavat_{q}']["Chap B"] = int(message.text)
                await message.answer("Chap C...")
            elif len(data[f'qavat_{q}']) == 2:
                data[f'qavat_{q}']["Chap C"] = int(message.text)
                await message.answer("O'ng D...")
            elif len(data[f'qavat_{q}']) == 3:
                data[f'qavat_{q}']["O'ng D"] = int(message.text)
                await message.answer("O'ng E...")
            elif len(data[f'qavat_{q}']) == 4:
                data[f'qavat_{q}']["O'ng E"] = int(message.text)
                await message.answer("O'ng F..")
            elif len(data[f'qavat_{q}']) == 5:
                data[f'qavat_{q}']["O'ng F"] = int(message.text)
                await message.answer("G..")
            elif len(data[f'qavat_{q}']) == 6:
                data[f'qavat_{q}']["G"] = int(message.text)
                await message.answer("I..")
            elif len(data[f'qavat_{q}']) == 7:
                data[f'qavat_{q}']["I"] = int(message.text)
                await message.answer("J..")
            elif len(data[f'qavat_{q}']) == 8:
                data[f'qavat_{q}']["J"] = int(message.text)
                await message.answer("Qavat balandligi(H)...")
            elif len(data[f'qavat_{q}']) == 9:
                data[f'qavat_{q}']["Qavat baladligi(H)"] = int(message.text)
                await message.answer("Eshik balandligi(K)...")
            elif len(data[f'qavat_{q}']) == 10:
                data[f'qavat_{q}']["Eshik baladligi(K)"] = int(message.text)
                data['state'] += 1

                if q < 21:
                    await message.answer(f"{q}-qavat ma'lumotlari qabul qilindi", reply_markup=btn.osten2Toxtatish)
                else:
                    await Form3.next()
                    await message.answer("Sim o'rtasi...")

@dp.message_handler(state=Form3.mejdu)
async def process_date(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data["Sim o'rtasi"] = message.text
    await Form3.next()
    await message.answer("Devor qalinligi:")

@dp.message_handler(state=Form3.steni)
async def process_date(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['Devor qalinligi'] = message.text
    await Form3.next()
    await message.answer("Priyamka chuqurligi:")

@dp.message_handler(state=Form3.priyamka)
async def process_date(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['Priyamka chuqurligi'] = message.text
    await Form3.next()
    await message.answer("Oxirgi qavat balandligi:")

@dp.message_handler(state=Form3.visota)
async def process_date(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['Oxirgi qavat balandligi'] = message.text
    await Form3.next()
    await message.answer("Priyamka devor turi:")

@dp.message_handler(state=Form3.stenki)
async def process_date(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['Priyamka devor turi'] = message.text
    await Form3.next()
    await message.answer("Priyamka tagida bo'sh devor bormi?", reply_markup=btn.daNet)

@dp.callback_query_handler(text='da', state=Form3.plashad_pod)
async def set_date(callback: types.CallbackQuery, state: FSMContext):
    async with state.proxy() as data:
        data["Priyamka tagida bo'sh devor bormi?"] = "Bor"
    await Form3.next()
    await callback.message.answer("Orqa tomon devor turi:", reply_markup=btn.devor_tur)

@dp.callback_query_handler(text='net', state=Form3.plashad_pod)
async def set_date(callback: types.CallbackQuery, state: FSMContext):
    async with state.proxy() as data:
        data["Priyamka tagida bo'sh devor bormi?"] = "Yo'q"
    await Form3.next()
    await callback.message.answer("Orqa tomon devor turi:", reply_markup=btn.devor_tur)

@dp.message_handler(state=Form3.plashad_pod)
async def process_date(message: types.Message, state: FSMContext):
    return await message.answer("Iltimos tugmachalardan birini tanlang!", reply_markup=btn.daNet)

@dp.callback_query_handler(text='beton', state=Form3.orqa)
async def set_date(callback: types.CallbackQuery, state: FSMContext):
    async with state.proxy() as data:
        data["orqa devor"] = "Beton"
    await Form3.next()
    await callback.message.answer("Chap yon devor turi:", reply_markup=btn.devor_tur)

@dp.callback_query_handler(text='metal', state=Form3.orqa)
async def set_date(callback: types.CallbackQuery, state: FSMContext):
    async with state.proxy() as data:
        data["orqa devor"] = "Metal konstruksiya"
    await Form3.next()
    await callback.message.answer("Chap yon devor turi:", reply_markup=btn.devor_tur)

@dp.callback_query_handler(text='gisht', state=Form3.orqa)
async def set_date(callback: types.CallbackQuery, state: FSMContext):
    async with state.proxy() as data:
        data["orqa devor"] = "G'isht"
    await Form3.next()
    await callback.message.answer("Chap yon devor turi:", reply_markup=btn.devor_tur)

@dp.message_handler(state=Form3.orqa)
async def process_date(message: types.Message, state: FSMContext):
    return await message.answer("Iltimos tugmachalardan birini tanlang!", reply_markup=btn.devor_tur)


@dp.callback_query_handler(text='beton', state=Form3.chap_yon)
async def set_date(callback: types.CallbackQuery, state: FSMContext):
    async with state.proxy() as data:
        data["chap_yon devor"] = "Beton"
    await Form3.next()
    await callback.message.answer("O'ng yon devor turi:", reply_markup=btn.devor_tur)


@dp.callback_query_handler(text='metal', state=Form3.chap_yon)
async def set_date(callback: types.CallbackQuery, state: FSMContext):
    async with state.proxy() as data:
        data["chap_yon devor"] = "Metal konstruksiya"
    await Form3.next()
    await callback.message.answer("O'ng yon devor turi:", reply_markup=btn.devor_tur)


@dp.callback_query_handler(text='gisht', state=Form3.chap_yon)
async def set_date(callback: types.CallbackQuery, state: FSMContext):
    async with state.proxy() as data:
        data["chap_yon devor"] = "G'isht"
    await Form3.next()
    await callback.message.answer("O'ng yon devor turi:", reply_markup=btn.devor_tur)


@dp.message_handler(state=Form3.chap_yon)
async def process_date(message: types.Message, state: FSMContext):
    return await message.answer("Iltimos tugmachalardan birini tanlang!", reply_markup=btn.devor_tur)

@dp.callback_query_handler(text='beton', state=Form3.ong_yon)
async def set_date(callback: types.CallbackQuery, state: FSMContext):
    async with state.proxy() as data:
        data["ong_yon devor"] = "Beton"
    await Form3.next()
    await callback.message.answer("oldi tomon chap devor turi:", reply_markup=btn.old_devor_tur)


@dp.callback_query_handler(text='metal', state=Form3.ong_yon)
async def set_date(callback: types.CallbackQuery, state: FSMContext):
    async with state.proxy() as data:
        data["ong_yon devor"] = "Metal konstruksiya"
    await Form3.next()
    await callback.message.answer("oldi tomon chap devor turi:", reply_markup=btn.old_devor_tur)


@dp.callback_query_handler(text='gisht', state=Form3.ong_yon)
async def set_date(callback: types.CallbackQuery, state: FSMContext):
    async with state.proxy() as data:
        data["ong_yon devor"] = "G'isht"
    await Form3.next()
    await callback.message.answer("oldi tomon chap devor turi:", reply_markup=btn.old_devor_tur)


@dp.message_handler(state=Form3.ong_yon)
async def process_date(message: types.Message, state: FSMContext):
    return await message.answer("Iltimos tugmachalardan birini tanlang!", reply_markup=btn.old_devor_tur)

@dp.callback_query_handler(text='beton', state=Form3.old_chap)
async def set_date(callback: types.CallbackQuery, state: FSMContext):
    async with state.proxy() as data:
        data["old_chap devor"] = "Beton"
    await Form3.next()
    await callback.message.answer("oldi tomon o'ng devor turi:", reply_markup=btn.old_devor_tur)


@dp.callback_query_handler(text='metal', state=Form3.old_chap)
async def set_date(callback: types.CallbackQuery, state: FSMContext):
    async with state.proxy() as data:
        data["old_chap devor"] = "Metal konstruksiya"
    await Form3.next()
    await callback.message.answer("oldi tomon o'ng devor turi:", reply_markup=btn.old_devor_tur)


@dp.callback_query_handler(text='gisht', state=Form3.old_chap)
async def set_date(callback: types.CallbackQuery, state: FSMContext):
    async with state.proxy() as data:
        data["old_chap devor"] = "G'isht"
    await Form3.next()
    await callback.message.answer("oldi tomon o'ng devor turi:", reply_markup=btn.old_devor_tur)

@dp.callback_query_handler(text='yoq', state=Form3.old_chap)
async def set_date(callback: types.CallbackQuery, state: FSMContext):
    async with state.proxy() as data:
        data["old_chap devor"] = ""
    await Form3.next()
    await callback.message.answer("oldi tomon o'ng devor turi:", reply_markup=btn.old_devor_tur)


@dp.message_handler(state=Form3.old_chap)
async def process_date(message: types.Message, state: FSMContext):
    return await message.answer("Iltimos tugmachalardan birini tanlang!", reply_markup=btn.old_devor_tur)


@dp.callback_query_handler(text='beton', state=Form3.old_ong)
async def set_date(callback: types.CallbackQuery, state: FSMContext):
    async with state.proxy() as data:
        data["old_ong devor"] = "Beton"
    await Form3.next()
    await callback.message.answer("Stanina ostidagi betonning balandligi:")


@dp.callback_query_handler(text='metal', state=Form3.old_ong)
async def set_date(callback: types.CallbackQuery, state: FSMContext):
    async with state.proxy() as data:
        data["old_ong devor"] = "Metal konstruksiya"
    await Form3.next()
    await callback.message.answer("Stanina ostidagi betonning balandligi:")


@dp.callback_query_handler(text='gisht', state=Form3.old_ong)
async def set_date(callback: types.CallbackQuery, state: FSMContext):
    async with state.proxy() as data:
        data["old_ong devor"] = "G'isht"
    await Form3.next()
    await callback.message.answer("Stanina ostidagi betonning balandligi:")


@dp.callback_query_handler(text='yoq', state=Form3.old_ong)
async def set_date(callback: types.CallbackQuery, state: FSMContext):
    async with state.proxy() as data:
        data["old_ong devor"] = ""
    await Form3.next()
    await callback.message.answer("Stanina ostidagi betonning balandligi:")


@dp.message_handler(state=Form3.old_ong)
async def process_date(message: types.Message, state: FSMContext):
    return await message.answer("Iltimos tugmachalardan birini tanlang!")

@dp.message_handler(state=Form3.beton)
async def process_date(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['Stanina ostidagi betonning balandligi'] = message.text
    await Form3.next()
    await message.answer("Mashina xonasi balandligi:")

@dp.message_handler(state=Form3.mashina)
async def process_date(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['Mashina xonasi balandligi'] = message.text
    await Form3.next()
    await message.answer("Eslatma:")

@dp.message_handler(state=Form3.eslatma)
async def process_date(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['Eslatma'] = message.text
    await Form3.next()
    await message.answer("Ma'lumotni kiritgan muhandis")

@dp.message_handler(state=Form3.podgatovil)
async def process_date(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['Muhandis'] = message.text
    await Form3.next()
    await message.answer("Sana:", reply_markup=btn.get_data)

@dp.callback_query_handler(text='sana', state=Form3.data)
async def set_date(callback: types.CallbackQuery, state: FSMContext):
    async with state.proxy() as data:
        data['Sana'] = datetime.today().strftime('%Y.%m.%d')
        print(data)
        write_file.write_excel_osten3(data)
    await state.finish()
    await callback.message.answer("Ma'lumot tayyorlanyapti...")
    await callback.message.answer_document(open(f"{data['Loyiha nomi']}.xlsx", 'rb'), reply_markup=btn.menu)


@dp.message_handler(state=Form3.data)
async def process_date(message: types.Message, state: FSMContext):
    async with state.proxy() as data:
        data['Sana'] = message.text
        print(data)
        write_file.write_excel_osten3(data)
    await state.finish()
    await message.answer("Ma'lumot tayyorlanyapti...")
    await message.answer_document(open(f"{data['Loyiha nomi']}.xlsx", 'rb'), reply_markup=btn.menu)