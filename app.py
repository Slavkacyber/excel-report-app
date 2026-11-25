import streamlit as st
import pandas as pd
import io

# Настройка страницы
st.set_page_config(page_title="DataFlow | Автоматизация отчетов", page_icon="📊", layout="centered")

# Стилизация (немного CSS для красоты)
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stButton>button { 
        width: 100%; 
        background-color: #10b981; 
        color: white; 
        font-weight: bold; 
        height: 50px;
        border-radius: 8px;
    }
    .stButton>button:hover { background-color: #059669; }
    .title { color: #1e3a8a; text-align: center; font-family: 'Manrope', sans-serif; }
    </style>
""", unsafe_allow_html=True)

# Заголовки
st.markdown("<h1 class='title'>📊 DataFlow</h1>", unsafe_allow_html=True)
st.markdown("<h3 style='text-align: center; color: #6b7280;'>Загрузите сырую выгрузку — получите готовый отчет за 5 секунд</h3>", unsafe_allow_html=True)
st.write("---")

# Блок загрузки файла
uploaded_file = st.file_uploader("Выберите файл Excel или CSV", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        # Читаем файл
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, sep=None, engine='python')
        else:
            df = pd.read_excel(uploaded_file)
        
        st.success(f"Файл загружен! Найдено строк: {len(df)}")
        
        # Настройки отчета
        with st.expander("⚙️ Настройки отчета", expanded=True):
            col1, col2 = st.columns(2)
            with col1:
                group_col = st.selectbox("Группировать по:", df.columns)
            with col2:
                # Ищем числовые колонки для агрегации
                numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
                if numeric_cols:
                    agg_col = st.selectbox("Суммировать:", numeric_cols)
                else:
                    st.warning("Числовых колонок не найдено.")
                    agg_col = None
        
        # Генерация отчета по кнопке
        if st.button("🚀 Сгенерировать отчет"):
            if agg_col:
                with st.spinner("Магия... Считаем отчет..."):
                    # Логика сводки
                    summary = df.groupby(group_col).agg(
                        Сумма=(agg_col, 'sum'),
                        Количество=(agg_col, 'count'),
                        Среднее=(agg_col, 'mean')
                    ).reset_index()
                    
                    # Сортировка по убыванию суммы
                    summary = summary.sort_values(by='Сумма', ascending=False)
                    
                    # Сохранение в Excel в оперативной памяти
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        summary.to_excel(writer, sheet_name='Сводка', index=False)
                        
                        # Красивое форматирование Excel
                        workbook  = writer.book
                        worksheet = writer.sheets['Сводка']
                        
                        # Формат заголовков
                        header_format = workbook.add_format({
                            'bold': True,
                            'bg_color': '#1e3a8a',
                            'font_color': 'white',
                            'border': 1
                        })
                        
                        # Применяем формат к шапке
                        for col_num, value in enumerate(summary.columns.values):
                            worksheet.write(0, col_num, value, header_format)
                            worksheet.set_column(col_num, col_num, 20)
                            
                        # Формат цифр
                        money_fmt = workbook.add_format({'num_format': '#,##0.00'})
                        worksheet.set_column(1, 3, 15, money_fmt)
                        
                    processed_data = output.getvalue()
                    
                    st.success("Отчет готов!")
                    
                    # Кнопка скачивания
                    st.download_button(
                        label="⬇️ Скачать готовый Excel-отчет",
                        data=processed_data,
                        file_name=f"report_{group_col}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.error("Для генерации отчета нужна хотя бы одна числовая колонка.")
                
    except Exception as e:
        st.error(f"Ошибка обработки файла: {e}")

st.write("---")
st.markdown("<p style='text-align: center; color: #9ca3af;'>© 2024 DataFlow. Автоматизация отчетов</p>", unsafe_allow_html=True)