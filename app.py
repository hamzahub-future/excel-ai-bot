import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
import openai
from fpdf import FPDF
import os
import urllib.request

# ⛔️ Вставь свой OpenAI ключ в переменную окружения!
openai.api_key = os.getenv("OPENAI_API_KEY")

st.title("📊 AI Excel Аналитик + PDF-отчёт")

uploaded_file = st.file_uploader("Загрузи Excel или CSV", type=["xlsx", "csv"])

# Подгружаем Unicode-шрифт для PDF
FONT_URL = "https://github.com/dejavu-fonts/dejavu-fonts/blob/master/ttf/DejaVuSans.ttf?raw=true"
FONT_PATH = "DejaVuSans.ttf"

if not os.path.exists(FONT_PATH):
    urllib.request.urlretrieve(FONT_URL, FONT_PATH)

if uploaded_file:
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    st.subheader("📋 Предпросмотр таблицы:")
    st.write(df.head())

    numeric_cols = df.select_dtypes(include='number').columns.tolist()

    if not numeric_cols:
        st.warning("Нет числовых колонок для анализа.")
    else:
        if st.button("📄 Сгенерировать PDF-отчёт"):
            with st.spinner("Генерация отчёта..."):
                pdf = FPDF()
                pdf.add_font("DejaVu", "", FONT_PATH, uni=True)
                pdf.add_font("DejaVu", "B", FONT_PATH, uni=True)
                pdf.set_auto_page_break(auto=True, margin=15)

                for col in numeric_cols:
                    # График
                    plt.figure(figsize=(8, 3))
                    df[col].plot(kind='line', title=col)
                    plt.tight_layout()
                    img_path = f"{col}.png"
                    plt.savefig(img_path)
                    plt.close()

                    # GPT-анализ
                    try:
                        response = openai.ChatCompletion.create(
                            model="gpt-3.5-turbo",
                            messages=[
                                {"role": "system", "content": "Ты аналитик данных. Дай краткий, понятный анализ этой числовой колонки."},
                                {"role": "user", "content": f"Колонка: {col}. Данные: {df[col].dropna().tolist()}"}
                            ],
                            max_tokens=300
                        )
                        analysis = response['choices'][0]['message']['content']
                    except Exception as e:
                        analysis = f"[Ошибка GPT: {e}]"

                    # PDF страница
                    pdf.add_page()
                    pdf.set_font("DejaVu", "B", 14)
                    pdf.cell(0, 10, f"Анализ: {col}", ln=True)
                    pdf.image(img_path, w=180)
                    pdf.set_font("DejaVu", "", 12)
                    pdf.multi_cell(0, 10, analysis)

                    os.remove(img_path)

                # Сохраняем PDF
                pdf_output_path = "report.pdf"
                pdf.output(pdf_output_path)

                with open(pdf_output_path, "rb") as f:
                    st.success("Готово! Скачай отчёт:")
                    st.download_button(
                        label="📥 Скачать PDF",
                        data=f,
                        file_name="Excel_AI_Отчёт.pdf",
                        mime="application/pdf"
                    )
