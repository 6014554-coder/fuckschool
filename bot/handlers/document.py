from io import BytesIO

from aiogram import Router, Bot
from aiogram.types import Message, BufferedInputFile
from aiogram.fsm.context import FSMContext

from bot.db.database import can_use_free, get_paid_docs, deduct_doc, ensure_user
from bot.services.formatter import format_document
from bot.services.analyzer import analyze_example
from bot.keyboards import work_type_keyboard, buy_keyboard, example_keyboard, WORK_TYPE_NAMES
from bot.handlers.start import UserState

router = Router()

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
MAX_FILE_SIZE = 20 * 1024 * 1024  # 20 МБ


async def _download_docx(msg: Message, bot: Bot):
    """Скачивает .docx из сообщения, возвращает bytes или None с ошибкой."""
    if not msg.document or msg.document.mime_type != DOCX_MIME:
        await msg.answer("Пожалуйста, отправь файл в формате .docx")
        return None
    if msg.document.file_size > MAX_FILE_SIZE:
        await msg.answer("Файл слишком большой. Максимум — 20 МБ.")
        return None
    file = await bot.get_file(msg.document.file_id)
    buf = BytesIO()
    await bot.download_file(file.file_path, destination=buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Шаг 1: получаем файл-образец
# ---------------------------------------------------------------------------

@router.message(UserState.waiting_for_example)
async def handle_example_file(msg: Message, state: FSMContext, bot: Bot):
    docx_bytes = await _download_docx(msg, bot)
    if docx_bytes is None:
        return

    processing = await msg.answer("🔍 Анализирую стиль образца...")
    try:
        config = analyze_example(docx_bytes)
    except Exception:
        await bot.delete_message(msg.chat.id, processing.message_id)
        await msg.answer(
            "⚠️ Не удалось прочитать файл-образец.\n"
            "Убедись, что это корректный .docx файл.",
            reply_markup=example_keyboard()
        )
        raise

    await state.update_data(example_config=config)
    await state.set_state(UserState.waiting_for_file)

    font = config.get("font_name", "Times New Roman")
    size = config.get("font_size_pt", 14.0)
    m = config.get("margins", {})
    indent = config.get("body", {}).get("first_line_indent_cm", 1.25)

    await bot.delete_message(msg.chat.id, processing.message_id)
    await msg.answer(
        f"✅ <b>Стиль образца распознан:</b>\n\n"
        f"• Шрифт: {font} {size:.0f}pt\n"
        f"• Поля: {m.get('left_mm', 30):.0f}/{m.get('right_mm', 10):.0f}/"
        f"{m.get('top_mm', 20):.0f}/{m.get('bottom_mm', 20):.0f} мм (Л/П/В/Н)\n"
        f"• Отступ абзаца: {indent:.2f} см\n\n"
        "Теперь отправь .docx который нужно отформатировать.",
        parse_mode="HTML"
    )


# ---------------------------------------------------------------------------
# Шаг 2: получаем основной документ
# ---------------------------------------------------------------------------

@router.message(UserState.waiting_for_file)
async def handle_file(msg: Message, state: FSMContext, bot: Bot):
    docx_bytes = await _download_docx(msg, bot)
    if docx_bytes is None:
        return

    user_id = msg.from_user.id
    await ensure_user(user_id, msg.from_user.username)

    # Проверяем квоту
    paid = await get_paid_docs(user_id)
    free_ok = await can_use_free(user_id)

    if paid == 0 and not free_ok:
        await msg.answer(
            "❌ Лимит исчерпан.\n\n"
            "Бесплатный документ уже использован в этом месяце.\n"
            "Купи пакет, чтобы продолжить:",
            reply_markup=buy_keyboard()
        )
        return

    data = await state.get_data()
    work_type     = data.get("work_type", "course")
    example_config = data.get("example_config", None)
    work_name = WORK_TYPE_NAMES.get(work_type, "Документ")

    mode_label = "по образцу вуза" if example_config else "по ГОСТ 7.32-2017"
    processing_msg = await msg.answer(f"⏳ Форматирую {mode_label}...")

    try:
        result_bytes = format_document(docx_bytes, config=example_config)
        await deduct_doc(user_id)

        paid_after = await get_paid_docs(user_id)
        free_after = await can_use_free(user_id)
        if paid_after > 0:
            quota_text = f"Осталось платных: {paid_after}"
        elif free_after:
            quota_text = "Осталось: 1 бесплатный документ в этом месяце"
        else:
            quota_text = "Бесплатный лимит исчерпан. /buy — купить пакет"

        original_name = msg.document.file_name or "document.docx"
        name_without_ext = original_name.rsplit(".", 1)[0]
        suffix = "_вуз.docx" if example_config else "_ГОСТ.docx"
        result_name = f"{name_without_ext}{suffix}"

        await bot.delete_message(msg.chat.id, processing_msg.message_id)
        await msg.answer_document(
            document=BufferedInputFile(result_bytes, filename=result_name),
            caption=(
                f"✅ <b>{work_name}</b> отформатирована {mode_label}\n\n"
                f"📋 {quota_text}\n\n"
                "Чтобы форматировать ещё — /start"
            ),
            parse_mode="HTML"
        )
        await state.clear()

    except Exception as e:
        await bot.delete_message(msg.chat.id, processing_msg.message_id)
        await msg.answer(
            "⚠️ Не удалось обработать файл.\n\n"
            "Убедись, что файл не повреждён и попробуй снова."
        )
        raise


# ---------------------------------------------------------------------------
# Fallback
# ---------------------------------------------------------------------------

@router.message()
async def handle_unexpected(msg: Message, state: FSMContext):
    current_state = await state.get_state()
    if current_state is None:
        await msg.answer("Привет! Нажми /start чтобы начать.")
    elif current_state == UserState.choosing_work_type:
        await msg.answer("Выбери тип работы из кнопок ниже:", reply_markup=work_type_keyboard())
    elif current_state == UserState.choosing_example:
        await msg.answer("Выбери вариант форматирования:", reply_markup=example_keyboard())
    elif current_state == UserState.waiting_for_example:
        await msg.answer("Жду .docx файл-образец от вуза.")
    elif current_state == UserState.waiting_for_file:
        await msg.answer("Жду .docx файл для форматирования.")
