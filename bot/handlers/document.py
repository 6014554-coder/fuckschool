from io import BytesIO

from aiogram import Router, Bot
from aiogram.types import Message, BufferedInputFile
from aiogram.fsm.context import FSMContext

from bot.db.database import can_use_free, get_paid_docs, deduct_doc, ensure_user
from bot.services.formatter import format_document
from bot.keyboards import work_type_keyboard, buy_keyboard, WORK_TYPE_NAMES
from bot.handlers.start import UserState

router = Router()

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
MAX_FILE_SIZE = 20 * 1024 * 1024  # 20 МБ


@router.message(UserState.waiting_for_file)
async def handle_file(msg: Message, state: FSMContext, bot: Bot):
    # Проверяем что это .docx
    if not msg.document or msg.document.mime_type != DOCX_MIME:
        await msg.answer("Пожалуйста, отправь файл в формате .docx")
        return

    if msg.document.file_size > MAX_FILE_SIZE:
        await msg.answer("Файл слишком большой. Максимум — 20 МБ.")
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

    # Берём тип работы из стейта
    data = await state.get_data()
    work_type = data.get("work_type", "course")
    work_name = WORK_TYPE_NAMES.get(work_type, "Документ")

    processing_msg = await msg.answer("⏳ Форматирую документ...")

    try:
        # Скачиваем файл
        file = await bot.get_file(msg.document.file_id)
        buf = BytesIO()
        await bot.download_file(file.file_path, destination=buf)
        docx_bytes = buf.getvalue()

        # Форматируем
        result_bytes = format_document(docx_bytes)

        # Списываем документ
        await deduct_doc(user_id)

        # Остаток
        paid_after = await get_paid_docs(user_id)
        free_after = await can_use_free(user_id)
        if paid_after > 0:
            quota_text = f"Осталось платных: {paid_after}"
        elif free_after:
            quota_text = "Осталось: 1 бесплатный документ в этом месяце"
        else:
            quota_text = "Бесплатный лимит исчерпан. /buy — купить пакет"

        # Формируем имя файла
        original_name = msg.document.file_name or "document.docx"
        name_without_ext = original_name.rsplit(".", 1)[0]
        result_name = f"{name_without_ext}_ГОСТ.docx"

        await bot.delete_message(msg.chat.id, processing_msg.message_id)
        await msg.answer_document(
            document=BufferedInputFile(result_bytes, filename=result_name),
            caption=(
                f"✅ <b>{work_name}</b> отформатирована по ГОСТу\n\n"
                f"📋 {quota_text}\n\n"
                "Чтобы форматировать ещё — /start"
            ),
            parse_mode="HTML"
        )

        # Сбрасываем стейт — нужно снова выбрать тип работы
        await state.clear()

    except Exception as e:
        await bot.delete_message(msg.chat.id, processing_msg.message_id)
        await msg.answer(
            "⚠️ Не удалось обработать файл.\n\n"
            "Убедись, что файл не повреждён и попробуй снова. Если ошибка повторяется — "
            "напиши нам."
        )
        raise


@router.message()
async def handle_unexpected(msg: Message, state: FSMContext):
    """Ловит сообщения вне нужного стейта."""
    current_state = await state.get_state()
    if current_state is None:
        await msg.answer(
            "Привет! Нажми /start чтобы начать.",
        )
    elif current_state == UserState.choosing_work_type:
        await msg.answer(
            "Выбери тип работы из кнопок ниже:",
            reply_markup=work_type_keyboard()
        )
