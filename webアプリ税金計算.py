import streamlit as st
import pandas as pd
from io import BytesIO
import openpyxl  # Excelファイル操作用ライブラリのインポート


# 計算処理の関数
def perform_calculations(df):
    df['取引日時'] = pd.to_datetime(df['取引日時'])
    df = df.sort_values(by='取引日時')
    df = df.drop(columns=['取引ID', '注文ID', 'タイプ', 'M/T'], errors='ignore')

    columns = ['取引日時'] + [col for col in df.columns if col != '取引日時']
    df = df[columns]

    # 初期化
    total_quantity_bought_list = []
    total_quantity_sold_list = []
    average_acquisition_price_list = []
    total_purchase_value_list = []
    daily_sales_value_list = []
    total_sales_value_list = []
    daily_profit_loss_list = []
    cumulative_profit_list = []
    cumulative_loss_list = []
    daily_tax_due_list = []
    cumulative_tax_due_list = []
    daily_quantity_bought_list = []

    total_quantity_bought = 0
    total_quantity_sold = 0
    total_purchase_value = 0
    total_sales_value = 0
    average_acquisition_price = 0
    cumulative_profit = 0
    cumulative_loss = 0
    cumulative_tax_due = 0

    # 行ごとの処理
    for index, row in df.iterrows():
        if row['売/買'] == '買':
            daily_quantity_bought = row['数量']
            total_quantity_bought += daily_quantity_bought
            total_purchase_value += row['数量'] * row['価格']
            average_acquisition_price = total_purchase_value / total_quantity_bought
            daily_sales_value = 0
            daily_profit_loss = 0
            daily_tax_due = 0
        else:
            daily_quantity_bought = 0
            total_quantity_sold += row['数量']
            daily_sales_value = row['数量'] * row['価格']
            total_sales_value += daily_sales_value
            sales_proceeds = daily_sales_value
            acquisition_cost = average_acquisition_price * row['数量']
            daily_profit_loss = sales_proceeds - acquisition_cost
            if daily_profit_loss > 0:
                cumulative_profit += daily_profit_loss
            else:
                cumulative_loss += daily_profit_loss
            tax_rate = 0.20
            daily_tax_due = max(0, daily_profit_loss) * tax_rate
            cumulative_tax_due += daily_tax_due

        total_quantity_bought_list.append(total_quantity_bought)
        total_quantity_sold_list.append(total_quantity_sold)
        average_acquisition_price_list.append(average_acquisition_price)
        total_purchase_value_list.append(total_purchase_value)
        daily_sales_value_list.append(daily_sales_value)
        total_sales_value_list.append(total_sales_value)
        daily_profit_loss_list.append(daily_profit_loss)
        cumulative_profit_list.append(cumulative_profit)
        cumulative_loss_list.append(cumulative_loss)
        daily_tax_due_list.append(daily_tax_due)
        cumulative_tax_due_list.append(cumulative_tax_due)
        daily_quantity_bought_list.append(daily_quantity_bought)

    # データフレームに計算結果を追加
    df['買いの合計枚数'] = daily_quantity_bought_list
    df['合計購入枚数'] = total_quantity_bought_list
    df['保有総量'] = df['合計購入枚数'] - total_quantity_sold_list
    df['買いの合計金額'] = total_purchase_value_list
    df['平均取得単価'] = average_acquisition_price_list
    df['その日の売り金額'] = daily_sales_value_list
    df['売りの合計枚数'] = total_quantity_sold_list
    df['売りの合計金額'] = total_sales_value_list
    df['利益/損失'] = daily_profit_loss_list
    df['合計利益'] = cumulative_profit_list
    df['合計損失'] = cumulative_loss_list
    df['税額'] = daily_tax_due_list
    df['合計税額'] = cumulative_tax_due_list

    return df


# Streamlitアプリの構成
st.title("取引データ計算アプリ")

# ファイルアップロード
uploaded_file = st.file_uploader("取引データをアップロードしてください (CSVまたはExcel形式)", type=["csv", "xlsx"])

if uploaded_file is not None:
    # ファイル形式に応じた読み込み
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        elif uploaded_file.name.endswith(".xlsx"):
            df = pd.read_excel(uploaded_file, engine='openpyxl')  # 明示的にopenpyxlを指定
        else:
            st.error("対応しているファイル形式は CSV または Excel (.xlsx) です。")
            st.stop()

        st.success("ファイルの読み込みに成功しました！")

        # 計算処理
        st.write("計算を開始します...")
        result_df = perform_calculations(df)
        st.write("計算が完了しました。結果を以下に表示します:")
        st.dataframe(result_df)

        # 結果を保存してダウンロード可能に
        output_format = st.radio("出力形式を選択してください:", ["Excel", "CSV"])
        if output_format == "Excel":
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False, sheet_name='結果')
            output.seek(0)
            st.download_button(
                label="計算結果をダウンロード (Excel)",
                data=output,
                file_name="calculated_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        elif output_format == "CSV":
            csv = result_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="計算結果をダウンロード (CSV)",
                data=csv,
                file_name="calculated_results.csv",
                mime="text/csv"
            )
    except Exception as e:
        st.error(f"エラーが発生しました: {e}")
else:
    st.info("ファイルをアップロードしてください。")
