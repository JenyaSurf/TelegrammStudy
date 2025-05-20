import asyncio
import html
from pathlib import Path

from aiogram import Bot, Dispatcher, Router
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters import CommandStart
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import Message, FSInputFile

# папка для загрузки
BASE_DIR = Path(__file__).parent
TEMP_DIR = BASE_DIR / "data" / "temp"
TEMP_DIR.mkdir(parents=True, exist_ok=True)

# ваш токен
TOKEN = "7458983813:AAGTHcuK2O6uRYZm5FvfxVbtAGpwW_VAFDc"

# инициализация бота
bot    = Bot(token=TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
dp     = Dispatcher(storage=MemoryStorage())
router = Router()
dp.include_router(router)

# импорт ваших модулей
from convert.convert_doc   import convert_to_docx
from formatter.format_main import process_document

@router.message(CommandStart())
async def cmd_start(message: Message):
    await message.answer(
        "Привет! Отправь мне .docx или .doc файл отчёта — отформатирую и верну назад."
    )

@router.message()
async def handle_docs(message: Message):
    if not message.document:
        return

    name = message.document.file_name or "document"
    ext  = name.lower().rsplit(".", 1)[-1]
    if ext not in ("docx", "doc"):
        return await message.answer("❗ Пожалуйста, пришли .doc или .docx.")

    # сохраняем входной файл
    raw_path = TEMP_DIR / name
    file     = await bot.get_file(message.document.file_id)
    await bot.download_file(file.file_path, destination=raw_path)

    # если .doc — конвертим
    if ext == "doc":
        try:
            raw_path = Path(convert_to_docx(str(raw_path)))
        except Exception as e:
            msg = html.escape(str(e))
            return await message.answer(
                f"⚠️ Не удалось конвертировать .doc:\n<code>{msg}</code>",
                parse_mode="HTML"
            )

    # форматируем
    try:
        formatted_path = Path(process_document(str(raw_path)))
    except Exception as e:
        msg = html.escape(str(e))
        return await message.answer(
            f"❌ Ошибка при форматировании:\n<code>{msg}</code>",
            parse_mode="HTML"
        )

    # отправляем готовый файл
    await message.answer_document(
        document=FSInputFile(path=formatted_path, filename=formatted_path.name),
        caption="✅ Вот твой отформатированный отчёт!"
    )

async def main():
    print("🤖 Бот запущен, ждёт документов…")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
