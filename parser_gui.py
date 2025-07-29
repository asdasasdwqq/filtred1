import os
import fitz
import cv2
import io
import numpy as np
import pandas as pd
from PIL import Image
from telegram import Update
from telegram.ext import (
    ApplicationBuilder, MessageHandler, ContextTypes, CommandHandler,
    filters
)

TOKEN = os.environ.get("TOKEN")  # ← читается из Render ENV

user_files = {}

def extract_qrs_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    detector = cv2.QRCodeDetector()
    qrs = []

    for page_index in range(len(doc)):
        page = doc.load_page(page_index)
        width, height = page.rect.width, page.rect.height
        center_point = (width / 2, height / 2)

        pix = page.get_pixmap(matrix=fitz.Matrix(6, 6))
        img_bytes = pix.tobytes("png")
        img_pil = Image.open(io.BytesIO(img_bytes))
        img_cv = cv2.cvtColor(np.array(img_pil), cv2.COLOR_RGB2BGR)

        retval, decoded_info, points, _ = detector.detectAndDecodeMulti(img_cv)
        if not retval:
            continue

        closest_qr = None
        min_dist = float('inf')

        for val, pts in zip(decoded_info, points):
            if not val:
                continue
            x_coords = pts[:, 0]
            y_coords = pts[:, 1]
            cx = float(np.mean(x_coords)) / 6
            cy = float(np.mean(y_coords)) / 6
            dist = (cx - center_point[0]) ** 2 + (cy - center_point[1]) ** 2
            if dist < min_dist:
                min_dist = dist
                closest_qr = val.strip()

        if closest_qr:
            qrs.append(closest_qr)

    return qrs

def process_files(excel_path, pdf_path):
    df = pd.read_excel(excel_path, header=4)
    qrs = extract_qrs_from_pdf(pdf_path)
    df["QR-код"] = ""
    for i in range(min(len(df), len(qrs))):
        df.at[i, "QR-код"] = qrs[i]

    drop_cols = ['№ задания', 'Фото', 'Бренд', 'Размер', 'Цвет', 'Баркод']
    df.drop(columns=[col for col in drop_cols if col in df.columns], inplace=True)

    result_path = f"result_{os.path.basename(excel_path)}"
    df.to_excel(result_path, index=False)
    return result_path

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    file = update.message.document
    file_path = f"{user_id}_{file.file_name}"
    tg_file = await file.get_file()
    await tg_file.download_to_drive(file_path)

    if user_id not in user_files:
        user_files[user_id] = {}

    if file.file_name.lower().endswith(".pdf"):
        user_files[user_id]["pdf"] = file_path
        await update.message.reply_text("PDF получен.")
    elif file.file_name.lower().endswith(".xlsx"):
        user_files[user_id]["xlsx"] = file_path
        await update.message.reply_text("Excel получен.")
    else:
        await update.message.reply_text("Поддерживаются только .pdf и .xlsx")

    if "pdf" in user_files[user_id] and "xlsx" in user_files[user_id]:
        await update.message.reply_text("Обрабатываю файлы...")
        try:
            result_path = process_files(user_files[user_id]["xlsx"], user_files[user_id]["pdf"])
            await update.message.reply_document(document=open(result_path, "rb"), filename="Результат.xlsx")
        except Exception as e:
            await update.message.reply_text(f"Ошибка обработки: {str(e)}")

        os.remove(user_files[user_id]["pdf"])
        os.remove(user_files[user_id]["xlsx"])
        os.remove(result_path)
        del user_files[user_id]

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Бот готов. Отправьте .xlsx и .pdf файлы.")

async def main():
    app = ApplicationBuilder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    await app.run_webhook(
        listen="0.0.0.0",
        port=int(os.environ.get("PORT", 8080)),
        webhook_url=f"https://{os.environ['RENDER_EXTERNAL_HOSTNAME']}/",
    )

if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
