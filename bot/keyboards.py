from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton


def work_type_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(text="📄 Курсовая",    callback_data="wt:course"),
            InlineKeyboardButton(text="🎓 Диплом / ВКР", callback_data="wt:diploma"),
        ],
        [
            InlineKeyboardButton(text="📝 Реферат",      callback_data="wt:essay"),
            InlineKeyboardButton(text="🔬 Лабораторная", callback_data="wt:lab"),
        ],
    ])


def example_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📎 Загрузить пример",  callback_data="example:upload")],
        [InlineKeyboardButton(text="⚡ Стандартный ГОСТ",  callback_data="example:skip")],
    ])


def buy_keyboard() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="5 документов — 199 ₽",  callback_data="buy:pack_5")],
        [InlineKeyboardButton(text="15 документов — 449 ₽", callback_data="buy:pack_15")],
    ])


WORK_TYPE_NAMES = {
    "course":  "Курсовая работа",
    "diploma": "Дипломная работа / ВКР",
    "essay":   "Реферат",
    "lab":     "Лабораторная работа",
}
