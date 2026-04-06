import os

from aiogram import Router
from aiogram.filters import Command
from aiogram.types import (
    Message, CallbackQuery, LabeledPrice,
    PreCheckoutQuery, SuccessfulPayment
)
from aiogram.fsm.context import FSMContext

from bot.db.database import get_paid_docs, can_use_free, add_paid_docs, ensure_user
from bot.keyboards import buy_keyboard

router = Router()

PACKAGES = {
    "pack_5":  {"docs": 5,  "amount": 19900, "label": "5 документов — 199 ₽"},
    "pack_15": {"docs": 15, "amount": 44900, "label": "15 документов — 449 ₽"},
}


@router.message(Command("buy"))
async def cmd_buy(msg: Message):
    await msg.answer(
        "💳 <b>Купить пакет документов</b>\n\n"
        "После оплаты документы сразу зачислятся на твой аккаунт.\n"
        "Пакет привязан к твоему Telegram — передать нельзя.",
        parse_mode="HTML",
        reply_markup=buy_keyboard()
    )


@router.message(Command("quota"))
async def cmd_quota(msg: Message):
    user_id = msg.from_user.id
    await ensure_user(user_id, msg.from_user.username)
    paid = await get_paid_docs(user_id)
    free_ok = await can_use_free(user_id)

    lines = ["📊 <b>Твой баланс:</b>\n"]
    if paid > 0:
        lines.append(f"• Платных документов: <b>{paid}</b>")
    if free_ok:
        lines.append("• Бесплатный документ в этом месяце: <b>доступен</b>")
    else:
        lines.append("• Бесплатный документ: <b>использован</b> (сбросится в начале месяца)")

    if paid == 0 and not free_ok:
        lines.append("\n/buy — купить пакет")

    await msg.answer("\n".join(lines), parse_mode="HTML")


@router.callback_query(lambda c: c.data and c.data.startswith("buy:"))
async def process_buy(callback: CallbackQuery):
    package_key = callback.data.split(":")[1]
    pkg = PACKAGES.get(package_key)
    if not pkg:
        await callback.answer("Неизвестный пакет", show_alert=True)
        return

    provider_token = os.getenv("PAYMENTS_PROVIDER_TOKEN", "")
    if not provider_token:
        await callback.answer("Оплата временно недоступна", show_alert=True)
        return

    await callback.message.answer_invoice(
        title=f"Fuck School — {pkg['label']}",
        description=f"Пакет из {pkg['docs']} документов для ГОСТ-форматирования",
        payload=package_key,
        provider_token=provider_token,
        currency="RUB",
        prices=[LabeledPrice(label=pkg["label"], amount=pkg["amount"])],
        start_parameter="buy",
    )
    await callback.answer()


@router.pre_checkout_query()
async def pre_checkout(query: PreCheckoutQuery):
    await query.answer(ok=True)


@router.message(lambda m: m.successful_payment is not None)
async def successful_payment(msg: Message):
    payment: SuccessfulPayment = msg.successful_payment
    package_key = payment.invoice_payload
    pkg = PACKAGES.get(package_key)

    if not pkg:
        await msg.answer("Ошибка при начислении. Напиши нам, мы разберёмся.")
        return

    await add_paid_docs(
        user_id=msg.from_user.id,
        count=pkg["docs"],
        package=package_key,
        amount=payment.total_amount,
        charge_id=payment.telegram_payment_charge_id,
    )

    await msg.answer(
        f"✅ Оплата прошла!\n\n"
        f"Зачислено: <b>{pkg['docs']} документов</b>\n\n"
        f"Нажми /start чтобы форматировать.",
        parse_mode="HTML"
    )
