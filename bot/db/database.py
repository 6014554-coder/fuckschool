import aiosqlite
import os

DB_PATH = os.path.join(os.path.dirname(__file__), "../../data.db")


async def init_db():
    async with aiosqlite.connect(DB_PATH) as db:
        await db.executescript("""
            CREATE TABLE IF NOT EXISTS users (
                user_id     INTEGER PRIMARY KEY,
                username    TEXT,
                created_at  TEXT DEFAULT (datetime('now'))
            );

            CREATE TABLE IF NOT EXISTS quotas (
                user_id         INTEGER PRIMARY KEY,
                free_used_at    TEXT,        -- дата последнего использования бесплатного
                paid_docs       INTEGER DEFAULT 0,  -- остаток платных документов
                FOREIGN KEY (user_id) REFERENCES users(user_id)
            );

            CREATE TABLE IF NOT EXISTS transactions (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id     INTEGER,
                package     TEXT,            -- 'pack_5' или 'pack_15'
                amount      INTEGER,         -- в рублях
                telegram_charge_id TEXT,
                created_at  TEXT DEFAULT (datetime('now')),
                FOREIGN KEY (user_id) REFERENCES users(user_id)
            );
        """)
        await db.commit()


async def ensure_user(user_id: int, username):
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(
            "INSERT OR IGNORE INTO users (user_id, username) VALUES (?, ?)",
            (user_id, username)
        )
        await db.execute(
            "INSERT OR IGNORE INTO quotas (user_id) VALUES (?)",
            (user_id,)
        )
        await db.commit()


async def can_use_free(user_id: int) -> bool:
    """True если бесплатный документ в этом месяце ещё не использован."""
    async with aiosqlite.connect(DB_PATH) as db:
        async with db.execute(
            "SELECT free_used_at FROM quotas WHERE user_id = ?", (user_id,)
        ) as cur:
            row = await cur.fetchone()
            if not row or row[0] is None:
                return True
            # Сравниваем год+месяц
            async with db.execute(
                "SELECT strftime('%Y-%m', free_used_at) = strftime('%Y-%m', 'now') "
                "FROM quotas WHERE user_id = ?", (user_id,)
            ) as cur2:
                same_month = await cur2.fetchone()
                return not (same_month and same_month[0])


async def get_paid_docs(user_id: int) -> int:
    async with aiosqlite.connect(DB_PATH) as db:
        async with db.execute(
            "SELECT paid_docs FROM quotas WHERE user_id = ?", (user_id,)
        ) as cur:
            row = await cur.fetchone()
            return row[0] if row else 0


async def deduct_doc(user_id: int):
    """Списывает один документ: сначала платные, потом бесплатный."""
    paid = await get_paid_docs(user_id)
    async with aiosqlite.connect(DB_PATH) as db:
        if paid > 0:
            await db.execute(
                "UPDATE quotas SET paid_docs = paid_docs - 1 WHERE user_id = ?",
                (user_id,)
            )
        else:
            await db.execute(
                "UPDATE quotas SET free_used_at = datetime('now') WHERE user_id = ?",
                (user_id,)
            )
        await db.commit()


async def add_paid_docs(user_id: int, count: int, package: str,
                        amount: int, charge_id: str):
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(
            "UPDATE quotas SET paid_docs = paid_docs + ? WHERE user_id = ?",
            (count, user_id)
        )
        await db.execute(
            "INSERT INTO transactions (user_id, package, amount, telegram_charge_id) "
            "VALUES (?, ?, ?, ?)",
            (user_id, package, amount, charge_id)
        )
        await db.commit()
