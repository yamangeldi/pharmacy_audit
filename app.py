import os
import io
import pandas as pd
import requests
import openpyxl
import streamlit as st
from roboflow import Roboflow

# Настройка страницы браузера
st.set_page_config(page_title="Аудит Выкладки", layout="centered")

st.title("💊 Автоматический аудит выкладки")
st.write("Загрузите файл матрицы (план) и файл с фотоотчетом (ссылками), чтобы нейросеть проверила наличие препаратов.")

# Окна для загрузки файлов пользователем
st.subheader("1. Загрузка данных")
matrix_file = st.file_uploader("Загрузите матрицу (matrix.xlsx)", type=["xlsx"])
report_file = st.file_uploader("Загрузите фотоотчет (Oson...xlsm или xlsx)", type=["xlsx", "xlsm"])

# Поле для ввода API ключа (чтобы не хранить его в коде)
api_key_input = st.text_input("Введите ваш API-ключ Roboflow", type="password")

if st.button("🚀 Запустить проверку", type="primary"):
    if not matrix_file or not report_file or not api_key_input:
        st.error("Пожалуйста, загрузите оба файла и введите API-ключ!")
    else:
        try:
            # 1. Авторизация
            st.info("Подключение к нейросети...")
            rf = Roboflow(api_key=api_key_input)
            project = rf.workspace().project("uz_ir_pharmacy")
            model = project.version(1).model # Поменяй на 2, когда обучишь новую версию
            
            # 2. Чтение Матрицы
            st.info("Чтение планов матрицы...")
            df_matrix = pd.read_excel(matrix_file)
            df_matrix = df_matrix.dropna(subset=['roboflow_name'])

            matrix_plans = {}
            all_unique_products = set()
            for index, row in df_matrix.iterrows():
                shelf_name = str(row['Column Name']).strip()
                product = str(row['roboflow_name']).strip()
                target = int(row['Target packs'])
                
                if shelf_name not in matrix_plans:
                    matrix_plans[shelf_name] = {}
                matrix_plans[shelf_name][product] = target
                all_unique_products.add(product)

            all_unique_products = sorted(list(all_unique_products))

            # 3. Чтение Отчета
            st.info("Открытие фотоотчета...")
            wb = openpyxl.load_workbook(report_file, data_only=True)
            ws = wb.active
            
            # Для теста в вебе берем первые 10 строк. Потом можно будет убрать лимит
            max_rows = min(ws.max_row + 1, 12) 
            
            master_results = []
            
            # Элементы интерфейса для прогресса
            progress_bar = st.progress(0)
            status_text = st.empty()

            # 4. Основной конвейер
            total_items = max_rows - 2
            for i, row_idx in enumerate(range(2, max_rows)):
                status_text.text(f"Анализ строки {row_idx} (аптека {ws.cell(row=row_idx, column=3).value})...")
                progress_bar.progress((i + 1) / total_items)
                
                visit_id = ws.cell(row=row_idx, column=1).value
                visit_date = ws.cell(row=row_idx, column=2).value
                pharmacy_id = ws.cell(row=row_idx, column=3).value
                shelf_name = str(ws.cell(row=row_idx, column=7).value).strip()
                
                cell_url = ws.cell(row=row_idx, column=9)
                image_url = cell_url.hyperlink.target if hasattr(cell_url, 'hyperlink') and cell_url.hyperlink else str(cell_url.value)

                if not image_url or not str(image_url).startswith("http"):
                    continue

                row_data = {
                    "ID Визита": visit_id,
                    "Дата": visit_date,
                    "ID Аптеки": pharmacy_id,
                    "Название полки": shelf_name,
                    "Ссылка на фото": image_url
                }

                temp_filename = f"temp_visit_{row_idx}.jpg"
                try:
                    response = requests.get(image_url, timeout=15)
                    if response.status_code == 200:
                        with open(temp_filename, "wb") as f:
                            f.write(response.content)
                            
                    prediction = model.predict(temp_filename, confidence=40, overlap=30).json()
                    found_items = prediction["predictions"]

                    shelf_fact = {}
                    for item in found_items:
                        cls = item["class"]
                        shelf_fact[cls] = shelf_fact.get(cls, 0) + 1

                    current_plan = matrix_plans.get(shelf_name, {})
                    sum_of_percentages = 0
                    planned_items_count = 0

                    for product in all_unique_products:
                        col_name = f"% {product}"
                        target = current_plan.get(product, 0)

                        if target > 0:
                            fact = shelf_fact.get(product, 0)
                            pct = min((fact / target) * 100, 100)
                            row_data[col_name] = round(pct, 1)
                            sum_of_percentages += pct
                            planned_items_count += 1
                        else:
                            row_data[col_name] = "Not in Plan"

                    if planned_items_count > 0:
                        row_data["ИТОГО ПО ПОЛКЕ (%)"] = round(sum_of_percentages / planned_items_count, 1)
                    else:
                        row_data["ИТОГО ПО ПОЛКЕ (%)"] = "Нет плана"

                    master_results.append(row_data)

                except Exception as e:
                    pass
                finally:
                    if os.path.exists(temp_filename):
                        os.remove(temp_filename)

            # 5. Вывод результата в интерфейс
            if master_results:
                base_cols = ["ID Визита", "Дата", "ID Аптеки", "Название полки", "Ссылка на фото", "ИТОГО ПО ПОЛКЕ (%)"]
                product_cols = [f"% {p}" for p in all_unique_products]
                df_final = pd.DataFrame(master_results)[base_cols + product_cols]
                
                st.success("✅ Анализ завершен!")
                st.dataframe(df_final) # Показываем таблицу прямо на сайте
                
                # Создаем кнопку для скачивания файла
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_final.to_excel(writer, index=False)
                
                st.download_button(
                    label="📥 Скачать итоговый Excel",
                    data=buffer.getvalue(),
                    file_name="Финальный_Аудит_Построчный.xlsx",
                    mime="application/vnd.ms-excel"
                )
            else:
                st.warning("Данные не найдены.")
                
        except Exception as e:
            st.error(f"Произошла ошибка: {e}")