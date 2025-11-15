
import os
import asyncio
import datetime as dt

import numpy as np
import pandas as pd
import yfinance as yf

from aiogram import Bot, Dispatcher
from aiogram.filters import Command
from aiogram.types import Message, FSInputFile

from dotenv import load_dotenv

HISTORY_DAYS = 365
XAU_TICKER = "XAUUSD=X"
DXY_TICKER = "^DXY"
COL_XAU = "XAUUSD"
COL_DXY = "DXY"

def fetch_xau_dxy_data():
    end = dt.datetime.utcnow()
    start = end - dt.timedelta(days=HISTORY_DAYS)

    xau = yf.download(XAU_TICKER, start=start, end=end)
    dxy = yf.download(DXY_TICKER, start=start, end=end)

    if xau.empty or dxy.empty:
        raise RuntimeError("Не удалось загрузить данные.")

    df = pd.DataFrame({
        COL_XAU: xau["Close"],
        COL_DXY: dxy["Close"],
    }).dropna()

    df["XAU_Return"] = df[COL_XAU].pct_change()
    df["DXY_Return"] = df[COL_DXY].pct_change()

    corr = df["XAU_Return"].corr(df["DXY_Return"])
    return df, corr

def create_excel_with_chart(df, corr):
    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"xau_dxy_report_{timestamp}.xlsx"

    with pd.ExcelWriter(filename, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Data")

        workbook = writer.book
        data_ws = writer.sheets["Data"]

        summary_ws = workbook.add_worksheet("Summary")
        summary_ws.write("A1", "Correlation")
        summary_ws.write("B1", float(corr))

        num_fmt = workbook.add_format({"num_format": "0.0000"})
        summary_ws.set_column("A:A", 40)
        summary_ws.set_column("B:B", 12, num_fmt)

        chart = workbook.add_chart({"type": "line"})
        max_row = len(df)

        chart.add_series({
            "name": COL_XAU,
            "categories": ["Data", 1, 0, max_row, 0],
            "values": ["Data", 1, 1, max_row, 1],
        })
        chart.add_series({
            "name": COL_DXY,
            "categories": ["Data", 1, 0, max_row, 0],
            "values": ["Data", 1, 2, max_row, 2],
        })

        chart.set_title({"name": f"{COL_XAU} vs {COL_DXY}"})
        chart.set_x_axis({"name": "Date"})
        chart.set_y_axis({"name": "Price"})
        chart.set_legend({"position": "bottom"})

        data_ws.insert_chart("G2", chart)

    return filename

load_dotenv()
BOT_TOKEN = os.getenv("BOT_TOKEN")

async def handle_start(message: Message):
    await message.answer("Команда: /xau_dxy")

async def handle_xau_dxy(message: Message):
    await message.answer("Создаю отчёт...")

    try:
        df, corr = await asyncio.to_thread(fetch_xau_dxy_data)
        filename = await asyncio.to_thread(create_excel_with_chart, df, corr)

        doc = FSInputFile(filename)
        await message.answer_document(doc, caption=f"Correlation: {corr:.4f}")
    except Exception as e:
        await message.answer(f"Ошибка: {e}")

async def main():
    bot = Bot(token=BOT_TOKEN)
    dp = Dispatcher()

    dp.message.register(handle_start, Command("start"))
    dp.message.register(handle_xau_dxy, Command("xau_dxy"))

    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
