import pandas as pd
import sqlite3
from openpyxl.styles import Font, PatternFill
import streamlit as st

@st.cache_data
def report():
    df = pd.read_excel('dataset/online_retail_II.xlsx') 
    
    df_clean = df.dropna()
    df_unique = df.drop_duplicates()
    df['InvoiceDate'] = pd.to_datetime(df['InvoiceDate'], errors='coerce')
    df['Month'] = df['InvoiceDate'].dt.strftime('%Y-%m')
    df['Price'] = df['Quantity'] * df['Price']
    
    revenueMonth = (
        df.groupby('Month')['Price']
        .sum()
        .reset_index(name='Revenue')
        )
    #print("Выручка по месяцам:", "\n", revenueMonth)
    
    topProducts = (
        df.groupby(['StockCode', 'Description'])['Price']
        .sum()
        .sort_values(ascending=False)
        .head(10)
        .reset_index(name='Revenue')
    )
    #print("\n", "Топ-10 продуктов:", "\n", topProducts)
    
    topByQuantity = (
        df.groupby(['StockCode', 'Description'])['Quantity']
        .sum()
        .sort_values(ascending=False)
        .head(10)
        .reset_index(name='Quantity')
    )
    topProducts['StockCode'] = topProducts['StockCode'].astype(str)
    topByQuantity['StockCode'] = topByQuantity['StockCode'].astype(str)
    #print("\n", "Топ-10 продуктов по количеству:", "\n", topByQuantity)
    
    averageMonthCheck = (
        df.groupby('Month')['Price']
        .mean()
        .reset_index(name='AverageCheck')
    )
    #print("\n", "Средний чек по месяцам:", "\n", averageMonthCheck)
    return revenueMonth, topProducts, topByQuantity, averageMonthCheck

def database(revenueMonth, topProducts, topByQuantity, averageMonthCheck):
    tables = {
        'revenue_month': revenueMonth,
        'top_products': topProducts,
        'top_by_quantity': topByQuantity,
        'average_month_check': averageMonthCheck,
    }
    conn = sqlite3.connect('database.db')
    try:
        for table_name, frame in tables.items():
            frame = frame.copy()
            if 'Month' in frame.columns:
                frame['Month'] = frame['Month'].astype(str)
            frame.to_sql(table_name, conn, if_exists='replace', index=False)
    finally:
        conn.close()
    
def makeExcelReport(revenueMonth, topProducts, topByQuantity, averageMonthCheck):
    with pd.ExcelWriter('report.xlsx', engine='openpyxl') as writer:
        revenueMonth.to_excel(writer, sheet_name='Выручка по месяцам', index=False)
        topProducts.to_excel(writer, sheet_name='Топ продаж', index=False)
        topByQuantity.to_excel(writer, sheet_name='Топ по количеству', index=False)
        averageMonthCheck.to_excel(writer, sheet_name='Средний чек по месяцам', index=False)

def dashboard(revenueMonth, topProducts, topByQuantity, averageMonthCheck):
    if revenueMonth.empty:
        st.warning('Данных нет — сначала запусти пайплайн загрузки.')
        st.stop()
        
    months = revenueMonth['Month'].astype(str).unique()
    selected = st.sidebar.multiselect('Месяцы', months, default=list(months))
    filteredRevenue = revenueMonth[revenueMonth['Month'].astype(str).isin(selected)]
    filteredAverage = averageMonthCheck[
        averageMonthCheck['Month'].astype(str).isin(selected)
    ]
    
    col1, col2, col3 = st.columns(3)
    col1.metric('Выручка всего', f"{filteredRevenue['Revenue'].sum():,.0f}")
    col2.metric('Средний чек', f"{filteredAverage['AverageCheck'].mean():,.0f}")
    col3.metric('Месяцев выбрано', len(selected))
        
    st.subheader("Выручка по месяцам")
    st.dataframe(filteredRevenue)
    
    st.subheader("Топ-10 продуктов по выручке")
    st.dataframe(topProducts)
    
    st.subheader("Топ-10 продуктов по количеству")
    st.dataframe(topByQuantity)
    
    st.subheader("Средний чек по месяцам")
    st.dataframe(filteredAverage)

def main():
    st.title("Отчет по продажам")
    with st.spinner('Загружаю и рассчитываю отчет...'):
        revenueMonth, topProducts, topByQuantity, averageMonthCheck = report()

    dashboard(revenueMonth, topProducts, topByQuantity, averageMonthCheck)

    if not st.session_state.get('exports_generated', False):
        database(revenueMonth, topProducts, topByQuantity, averageMonthCheck)
        makeExcelReport(revenueMonth, topProducts, topByQuantity, averageMonthCheck)
        st.session_state['exports_generated'] = True


if __name__ == "__main__":
    main()
