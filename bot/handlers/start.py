from aiogram import Router
from aiogram.filters import Command, CommandStart
from aiogram.types import Message, CallbackQuery
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup

from bot.db.database import ensure_user
from bot.keyboards import work_type_keyboard, WORK_TYPE_NAMES

router = Router()


class UserState(StatesGroup):
    choosing_work_type = State()
    waiting_for_file   = State()


@router.message(CommandStart())
async def cmd_start(msg: Message, state: FSMContext):
    await ensure_user(msg.from_user.id, msg.from_user.username)
    await state.set_state(UserState.choosing_work_type)
    await msg.answer(
        "👋 Привет!\n\n"
        "Я форматирую .docx по ГОСТу — шрифт, поля, интервалы, заголовки.\n\n"
        "Выбери тип работы:",
        reply_markup=work_type_keyboard()
    )


@router.callback_query(lambda c: c.data and c.data.startswith("wt:"))
async def choose_work_type(callback: CallbackQuery, state: FSMContext):
    work_type = callback.data.split(":")[1]
    await state.update_data(work_type=work_type)
    await state.set_state(UserState.waiting_for_file)

    name = WORK_TYPE_NAMES.get(work_type, "Документ")
    await callback.message.edit_text(
        f"✅ Выбрано: <b>{name}</b>\n\n"
        "Теперь отправь .docx файл — я верну его отформатированным по ГОСТу.",
        parse_mode="HTML"
    )
    await callback.answer()


@router.message(Command("help"))
async def cmd_help(msg: Message):
    await msg.answer(
        "ℹ️ <b>Как пользоваться:</b>\n\n"
        "1. /start — выбрать тип работы\n"
        "2. Отправить .docx файл\n"
        "3. Получить готовый документ по ГОСТу\n\n"
        "💳 /buy — купить пакет документов\n"
        "📊 /quota — посмотреть остаток",
        parse_mode="HTML"
    )
