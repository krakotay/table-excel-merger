import asyncio
import logging
import tempfile
from typing import List, Dict, Any
from aiogram import Bot, Dispatcher
from aiogram.types import Message
from aiogram.filters import Command
import os
from dotenv import load_dotenv
from aiogram import F, types
from aiogram_media_group import media_group_handler
from inn_check import check_by_inn, merge_excel
from aiogram.types import FSInputFile
from aiogram.enums.chat_action import ChatAction

load_dotenv()


BOT_TOKEN = os.getenv("TOKEN")


# Настройка логирования
logging.basicConfig(level=logging.INFO)

# Инициализация бота и диспетчера
bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()


@dp.message(F.document & F.document.file_name.endswith(".xlsx"), F.media_group_id)
@media_group_handler
async def excel_handler(messages: List[types.Message]):
    """Обрабатывает ровно два файла Excel, отправленные как медиагруппа."""
    if not messages[0].document:
        await messages[0].answer("Не удалось получить информацию о документе.")
        return
    if len(messages) != 2:
        await messages[0].answer("Пожалуйста, отправьте ровно два файла Excel одним сообщением (как группу).")
        return
    temp_files_paths = []
    media_group_id = messages[0].media_group_id
    try:
        # Скачиваем оба файла
        for msg in messages:
            bot_file_info = await bot.get_file(msg.document.file_id)
            file_path = bot_file_info.file_path
            f_name = msg.document.file_name
            if not os.path.exists("temp"):
                os.makedirs("temp")

            # Используем стандартный временный каталог системы вместо "temp"
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False, dir="temp") as temp_file:
                temp_file_path = temp_file.name
                logging.info(f"Создание временного файла: {temp_file_path} для '{f_name}'")
                await bot.download_file(file_path, destination=temp_file_path)
                logging.info(f"Файл '{f_name}' скачан во временный файл: {temp_file_path}")
                temp_files_paths.append({'path': temp_file_path, 'name': f_name})
        df = merge_excel(temp_files_paths[0]['path'], temp_files_paths[1]['path'])
        with tempfile.NamedTemporaryFile(suffix=".xlsx", prefix="ООО_", delete=False, dir="temp") as temp_file:
            temp_file_path = temp_file.name
            df.write_excel(temp_file_path)
            await messages[0].answer_document(FSInputFile(path=temp_file_path))
        await bot.send_chat_action(messages[0].chat.id, action=ChatAction.TYPING)
        wait_time_s = df.height / 1.6
        await messages[0].answer(f"Поиск ФИО, пожалуйста, подождите... Ожидаемое время: {wait_time_s} секунд")
        new_df = check_by_inn(df) 

        with tempfile.NamedTemporaryFile(suffix=".xlsx", prefix="ООО_ФИО_", delete=False, dir="temp") as temp_file:
            temp_file_path = temp_file.name
            new_df.write_excel(temp_file_path)
            await messages[0].answer_document(FSInputFile(path=temp_file_path))
        await messages[0].answer(f"Файлы '{temp_files_paths[0]['name']}' и '{temp_files_paths[1]['name']}' успешно скачаны и (условно) обработаны.")


    except Exception as e:
        logging.error(f"Ошибка при обработке группы {media_group_id}: {e}", exc_info=True)
        await messages[0].answer(f"Произошла ошибка при скачивании или обработке файлов: {e}")




@dp.message(Command("start"))
async def start(message: Message):
    await message.answer("Добро пожаловать!")




async def main():
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
