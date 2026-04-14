from aiogram import Router
from aiogram.filters import Command, CommandStart
from aiogram.types import Message, CallbackQuery
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup

from bot.db.database import ensure_user
from bot.keyboards import work_type_keyboard, example_keyboard, WORK_TYPE_NAMES

router = Router()


class UserState(StatesGroup):
    choosing_work_type   = State()
    choosing_example     = State()   # загрузить пример или стандарт
    waiting_for_example  = State()   # ждём файл-пример
    waiting_for_file     = State()   # ждём основной документ


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
    await state.set_state(UserState.choosing_example)

    name = WORK_TYPE_NAMES.get(work_type, "Документ")
    await callback.message.edit_text(
        f"✅ Выбрано: <b>{name}</b>\n\n"
        "Есть образец оформления от твоего университета?\n\n"
        "Загрузи его — я подстрою форматирование под стиль вуза.\n"
        "Или выбери стандартный ГОСТ 7.32-2017:",
        parse_mode="HTML",
        reply_markup=example_keyboard()
    )
    await callback.answer()


@router.callback_query(lambda c: c.data == "example:skip")
async def example_skip(callback: CallbackQuery, state: FSMContext):
    await state.update_data(example_config=None)
    await state.set_state(UserState.waiting_for_file)
    await callback.message.edit_text(
        "⚡ <b>Стандартный ГОСТ 7.32-2017</b>\n\n"
        "Отправь .docx файл — я верну его отформатированным.",
        parse_mode="HTML"
    )
    await callback.answer()


@router.callback_query(lambda c: c.data == "example:upload")
async def example_upload(callback: CallbackQuery, state: FSMContext):
    await state.set_state(UserState.waiting_for_example)
    await callback.message.edit_text(
        "📎 Отправь .docx файл-образец от своего университета.\n\n"
        "Я извлеку из него параметры форматирования и применю к твоему документу."
    )
    await callback.answer()


@router.message(Command("help"))
async def cmd_help(msg: Message):
    await msg.answer(
        "ℹ️ <b>Как пользоваться:</b>\n\n"
        "1. /start — выбрать тип работы\n"
        "2. Загрузить образец вуза (необязательно)\n"
        "3. Отправить .docx файл\n"
        "4. Получить готовый документ\n\n"
        "💳 /buy — купить пакет документов\n"
        "📊 /quota — посмотреть остаток",
        parse_mode="HTML"
    )
