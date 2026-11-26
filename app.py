import streamlit as st
import pandas as pd
import numpy as np
import io
import base64
from datetime import datetime

# 1. НАСТРОЙКА СТРАНИЦЫ И ДИЗАЙН
st.set_page_config(page_title="Сводка | Комбайн для отчетов", page_icon="↗", layout="wide")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=Manrope:wght@600;700;800&display=swap');
    
    :root {
        --ink: #1c2923;
        --muted: #68766f;
        --green: #278a5c;
        --green-dark: #176b45;
        --paper: #f7f9f7;
        --line: #e4eae5;
        --card-bg: #ffffff;
    }
    
    /* Основной фон */
    .stApp {
        background-color: var(--paper);
        font-family: 'Inter', sans-serif;
        color: var(--ink);
    }
    
    /* Заголовки */
    h1, h2, h3 { font-family: 'Manrope', sans-serif !important; font-weight: 800 !important; letter-spacing: -0.5px; }
    
    /* Стилизация вкладок (Tabs) */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: transparent;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        border-radius: 10px 10px 0 0;
        padding: 0 24px;
        font-family: 'Manrope', sans-serif;
        font-weight: 600;
        color: var(--muted);
    }
    .stTabs [aria-selected="true"] {
        background-color: var(--card-bg);
        color: var(--green);
        border-bottom: 3px solid var(--green);
    }

    /* Карточки */
    .metric-card {
        background: var(--card-bg);
        padding: 24px;
        border-radius: 16px;
        border: 1px solid var(--line);
        box-shadow: 0 4px 12px rgba(0,0,0,0.03);
        text-align: center;
    }
    .metric-value {
        font-size: 32px;
        font-family: 'Manrope', sans-serif;
        font-weight: 800;
        color: var(--green);
        margin: 10px 0;
    }
    .metric-label {
        font-size: 14px;
        color: var(--muted);
        text-transform: uppercase;
        letter-spacing: 1px;
        font-weight: 600;
    }

    /* Кнопки */
    .stButton>button {
        background-color: var(--green) !important;
        color: white !important;
        border-radius: 10px !important;
        padding: 10px 24px !important;
        font-weight: 600 !important;
        width: 100%;
        transition: all 0.2s ease;
    }
    .stButton>button:hover {
        background-color: var(--green-dark) !important;
        transform: translateY(-2px);
    }

    /* Кнопка скачивания */
    .stDownloadButton>button {
        background-color: white !important;
        color: var(--green) !important;
        border: 2px solid var(--green) !important;
        border-radius: 10px !important;
        font-weight: 700 !important;
        padding: 12px 24px !important;
    }
    .stDownloadButton>button:hover {
        background-color: #f0f9f4 !important;
    }

    /* Скрываем стандартные элементы Streamlit */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* Стилизация инпутов */
    .stSelectbox>div>div, .stTextInput>div>div {
        border-radius: 8px;
        border-color: var(--line);
    }
</style>
""", unsafe_allow_html=True)

# 2. ФУНКЦИИ ОБРАБОТКИ И ЭКСПОРТА
def load_file(file):
    try:
        if file.name.endswith('.csv'):
            return pd.read_csv(file, sep=None, engine='python')
        elif file.name.endswith('.xls'):
            return pd.read_excel(file, engine='xlrd')
        else:
            return pd.read_excel(file)
    except Exception as e:
        st.error(f"Ошибка чтения файла: {e}")
        return None

def export_to_excel(df, sheet_name, style=False):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        if style:
            workbook  = writer.book
            worksheet = writer.sheets[sheet_name]
            header_format = workbook.add_format({'bold': True, 'bg_color': '#1c2923', 'font_color': 'white', 'border': 1})
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
                worksheet.set_column(col_num, col_num, 20)
            money_fmt = workbook.add_format({'num_format': '#,##0.00'})
            worksheet.set_column(1, len(df.columns), 15, money_fmt)
    return output.getvalue()

# 3. ИНТЕРФЕЙС
st.markdown("<h1 style='text-align: center; margin-bottom: 5px;'>Ваши Excel-отчёты готовы <span style='color:#278a5c'>за 5 секунд</span></h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #68766f; font-size: 18px; margin-bottom: 40px;'>Полноценный комбайн для аналитики вашего бизнеса.</p>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Загрузите ваш файл (Excel или CSV)", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    df = load_file(uploaded_file)
    
    if df is not None:
        # Создаем 4 вкладки
        tab1, tab2, tab3, tab4 = st.tabs(["📊 Сводка и Итоги", "📈 Анализ трендов", "🔍 Поиск аномалий", "🏆 ABC-анализ"])

        # --- ВКЛАДКА 1: СВОДКА ---
        with tab1:
            st.markdown("### Группировка и базовые метрики")
            col1, col2 = st.columns(2)
            
            # Ищем числовые колонки, чтобы не давать пользователю выбрать текст
            numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
            
            with col1:
                group_col = st.selectbox("По какому столбцу группировать?", df.columns)
            with col2:
                if numeric_cols:
                    agg_col = st.selectbox("Какой столбец суммировать?", numeric_cols)
                else:
                    st.warning("В файле нет числовых колонок!")
                    agg_col = None

            if agg_col:
                if st.button("Сгенерировать сводку"):
                    summary = df.groupby(group_col).agg(
                        Сумма=(agg_col, 'sum'),
                        Кол_во=(agg_col, 'count'),
                        Среднее=(agg_col, 'mean')
                    ).reset_index().sort_values(by='Сумма', ascending=False)
                    
                    # Вывод красивых метрик
                    m1, m2, m3 = st.columns(3)
                    m1.markdown(f"<div class='metric-card'><div class='metric-label'>Общая сумма</div><div class='metric-value'>{summary['Сумма'].sum():,.2f}</div></div>", unsafe_allow_html=True)
                    m2.markdown(f"<div class='metric-card'><div class='metric-label'>Всего записей</div><div class='metric-value'>{len(df)}</div></div>", unsafe_allow_html=True)
                    m3.markdown(f"<div class='metric-card'><div class='metric-label'>Лидер</div><div class='metric-value'>{summary.iloc[0][group_col]}</div></div>", unsafe_allow_html=True)
                    
                    st.write("---")
                    st.dataframe(summary, use_container_width=True)
                    st.download_button("⬇️ Скачать сводку в Excel", export_to_excel(summary, "Сводка", True), "svodka.xlsx")

        # --- ВКЛАДКА 2: АНАЛИЗ ТРЕНДОВ ---
        with tab2:
            st.markdown("### Динамика показателей")
            st.info("Выберите столбец с датой и числовой столбец, чтобы увидеть график изменений во времени.")
            
            date_col = st.selectbox("Столбец с датой:", df.columns, key="date_sel")
            trend_col = st.selectbox("Столбец для анализа:", numeric_cols, key="trend_sel")
            
            if date_col and trend_col:
                if st.button("Построить график"):
                    try:
                        # Пытаемся преобразовать в дату
                        df[date_col] = pd.to_datetime(df[date_col], errors='coerce', dayfirst=True)
                        trend_data = df.dropna(subset=[date_col]).groupby(pd.Grouper(key=date_col, freq='ME'))[trend_col].sum().reset_index()
                        
                        st.line_chart(trend_data.set_index(date_col)[trend_col])
                        st.dataframe(trend_data, use_container_width=True)
                    except Exception as e:
                        st.error(f"Не удалось построить график: {e}. Проверьте формат дат.")

        # --- ВКЛАДКА 3: ПОИСК АНОМАЛИЙ (ВЫБРОСОВ) ---
        with tab3:
            st.markdown("### Поиск выбросов (Аномалий)")
            st.info("Инструмент находит значения, которые сильно отклоняются от среднего (выше или ниже нормы на 1.5 IQR). Это может быть ошибкой ввода или мошенничеством.")
            
            anom_col = st.selectbox("Выберите числовой столбец для поиска аномалий:", numeric_cols, key="anom_sel")
            
            if anom_col:
                if st.button("Найти аномалии"):
                    # Математика поиска выбросов
                    Q1 = df[anom_col].quantile(0.25)
                    Q3 = df[anom_col].quantile(0.75)
                    IQR = Q3 - Q1
                    lower_bound = Q1 - 1.5 * IQR
                    upper_bound = Q3 + 1.5 * IQR
                    
                    anomalies = df[(df[anom_col] < lower_bound) | (df[anom_col] > upper_bound)]
                    
                    if len(anomalies) > 0:
                        st.warning(f"Найдено {len(anomalies)} аномальных записей (ниже {lower_bound:,.2f} или выше {upper_bound:,.2f}).")
                        st.dataframe(anomalies, use_container_width=True)
                        st.download_button("⬇️ Скачать аномалии в Excel", export_to_excel(anomalies, "Аномалии", True), "anomalii.xlsx")
                    else:
                        st.success("Аномалий не найдено. Все данные в пределах нормы.")

        # --- ВКЛАДКА 4: ABC-АНАЛИЗ ---
        with tab4:
            st.markdown("### ABC-анализ")
            st.info("Правило Парето: 20% клиентов/товаров приносят 80% выручки. Инструмент разобьет ваши данные на 3 группы (A, B, C).")
            
            abc_col = st.selectbox("По какому столбцу группировать?", df.columns, key="abc_group")
            abc_val = st.selectbox("По какому столбцу считать?", numeric_cols, key="abc_val")
            
            if abc_col and abc_val:
                if st.button("Сделать ABC-анализ"):
                    abc_df = df.groupby(abc_col)[abc_val].sum().reset_index()
                    abc_df = abc_df.sort_values(by=abc_val, ascending=False)
                    
                    # Считаем кумулятивную долю
                    abc_df['Доля %'] = abc_df[abc_val] / abc_df[abc_val].sum() * 100
                    abc_df['Накопительная доля %'] = abc_df['Доля %'].cumsum()
                    
                    # Присваиваем категории
                    abc_df['Категория'] = np.where(abc_df['Накопительная доля %'] <= 80, 'A (Важные)', 
                                                   np.where(abc_df['Накопительная доля %'] <= 95, 'B (Средние)', 'C (Слабые)'))
                    
                    # Красим категории в цвета
                    st.dataframe(abc_df.style.apply(lambda x: ['background: #dcfce7' if v == 'A (Важные)' else 'background: #fef9c3' if v == 'B (Средние)' else 'background: #fee2e2' for v in x], subset=['Категория']), use_container_width=True)
                    
                    st.download_button("⬇️ Скачать ABC-анализ в Excel", export_to_excel(abc_df, "ABC", True), "abc_analiz.xlsx")

    # Подвал
    st.markdown("<br><br><p style='text-align: center; color: #9ca3af; font-size: 12px;'>© 2024 Сводка. Отчёты без рутины.</p>", unsafe_allow_html=True)
