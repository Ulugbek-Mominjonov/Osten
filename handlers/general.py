import logging

from aiogram import types
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters import Text
from aiogram.types import ReplyKeyboardRemove
from loader import dp

import keyboard as btn
from states import Form, Form2, Form3

@dp.message_handler(commands=['start', 'info'])
async def send_welcome(message: types.Message):
    await message.answer("Assalomu alaykum O'lchamlar olish botiga hush kelibsiz!!!", reply_markup=btn.menu)


@dp.message_handler()
async def startState(message: types.Message, state: FSMContext):
    if message.text == "Shahta o'lchamlarini olish":
        await message.reply("Siz Shahta o'lchamlarini olish botni tanladingiz.\nO'lchamlarni olish tugmani bosing",
                            reply_markup=btn.osten)
    elif message.text == "O'lchamlarni olish":
        await Form.loyiha_nomi.set()
        await message.answer("Loyiha nomini kiritng:", reply_markup=ReplyKeyboardRemove())
    elif message.text == "Obramleniya o'lchamlarini olish":
        await message.reply(
            "Siz Obramleniya o'lchamlarini olish botni tanladingiz.\nMa'lumotni kiritish tugmani bosing",
            reply_markup=btn.osten2)
    elif message.text == "Ma'lumotlarni kiritish":
        await Form2.loyiha_nomi.set()
        await message.reply("Лойиҳа номи:", reply_markup=ReplyKeyboardRemove())
    elif message.text == "Birnima":
        await message.reply("Siz Birnima o'lchamlarini olish botni tanladingiz.\nMa'lumotni kiritish tugmani bosing",
                            reply_markup=btn.osten3)
    elif message.text == "O'lchamlarni kiritish":
        await Form3.loyiha_nomi.set()
        await message.reply("Loyiha nomi:", reply_markup=ReplyKeyboardRemove())
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
