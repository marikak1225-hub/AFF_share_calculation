import pandas as pd
import streamlit as st
import plotly.express as px
import io
from datetime import date

st.title("メディア別シェア率分析ツール（期間指定＋目標割り振り）")

# ファイルアップロード
uploaded_file = st.file_uploader("Excelファイルをアップロードしてください", type=["xlsx"])

if uploaded_file:
    # Excel読み込み
    df = pd.read_excel(uploaded_file, sheet_name="【楽天カード】", engine="openpyxl", header=None)

    # 日付列検出（3行目）
    date_row = df.iloc[2]
    date_cols = []
    valid_dates = []
    for i, val in enumerate(date_row):
        try:
            parsed_date = pd.to_datetime(val, errors="raise")
            if not pd.isna(parsed_date):
                date_cols.append(i)
                valid_dates.append(parsed_date)
        except:
            continue

    if not valid_dates:
        st.error("有効な日付が見つかりません。Excelのフォーマットを確認してください。")
    else:
        # 期間選択
        start_date = st.date_input("開始日", value=valid_dates[0].date())
        end_date = st.date_input("終了日", value=valid_dates[-1].date())

        # 選択範囲の列インデックス
        selected_cols = [i for i, d in zip(date_cols, valid_dates) if start_date <= d.date() <= end_date]

        # メディア別集計（D列、合計除外）
        result = []
        for idx in range(len(df)):
            media_name = df.iloc[idx, 3]
            if pd.notna(media_name) and "合計" not in str(media_name):
                forecast_sum = 0
                actual_sum = 0

                # 下の行でForecastとActualを探す
                for j in range(idx + 1, min(idx + 10, len(df))):
                    label = str(df.iloc[j, 19]).strip()
                    if label == "Forecast":
                        forecast_sum = df.iloc[j, selected_cols].sum()
                    elif label.startswith("Actual"):
                        actual_sum = df.iloc[j, selected_cols].sum()

                if forecast_sum > 0 or actual_sum > 0:
                    result.append({"Media": media_name, "Forecast": forecast_sum, "Actual": actual_sum})

        result_df = pd.DataFrame(result)

        if not result_df.empty:
            # シェア率計算
            total_forecast = result_df["Forecast"].sum()
            total_actual = result_df["Actual"].sum()
            result_df["Forecast Share %"] = (result_df["Forecast"] / total_forecast * 100).round(2)
            result_df["Actual Share %"] = (result_df["Actual"] / total_actual * 100).round(2)

            # テーブル表示（シェア率）
            st.subheader("メディア別詳細シェア率")
            st.dataframe(result_df)

            # 円グラフ表示
            st.plotly_chart(px.pie(result_df, names="Media", values="Forecast", title="Forecastシェア率"))
            st.plotly_chart(px.pie(result_df, names="Media", values="Actual", title="Actualシェア率"))

            # Excelエクスポート（シェア率）
            output_share = io.BytesIO()
            with pd.ExcelWriter(output_share, engine="openpyxl") as writer:
                result_df.to_excel(writer, index=False, sheet_name="Media_Share")

            st.download_button(
                label="📥 シェア率テーブルをExcelでダウンロード",
                data=output_share.getvalue(),
                file_name="media_share.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # 目標件数入力
            user_target = st.number_input("全体目標件数を入力してください", min_value=0, value=1000)
            result_df["Allocated Target"] = (user_target * result_df["Actual Share %"] / 100).round(0)

            # 割り振り結果テーブル（Media, Actual, Actual Share %, Allocated Targetのみ）
            allocation_df = result_df[["Media", "Actual", "Actual Share %", "Allocated Target"]]

            st.subheader("目標件数割り振り結果")
            st.dataframe(allocation_df)

            # Excelエクスポート（割り振り結果）
            output_alloc = io.BytesIO()
            with pd.ExcelWriter(output_alloc, engine="openpyxl") as writer:
                allocation_df.to_excel(writer, index=False, sheet_name="Target_Allocation")

            st.download_button(
                label="📥 割り振り結果をExcelでダウンロード",
                data=output_alloc.getvalue(),
                file_name="media_target_allocation.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("対象データがありません。Excelを確認してください。")
