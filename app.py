import pandas as pd
import plotly.express as px
import streamlit as st

from logic import analyze, load_jisseki, load_kwzweb, to_excel

st.set_page_config(
    page_title="施工通知 期日確認",
    layout="wide",
)

st.title("施工通知 期日確認アプリ")
st.caption("主工種フラグ=1 の作業実績に対応する施工通知が、期日までに発行されているか確認します。")

# ── 期日ルール設定 ────────────────────────────────
with st.expander("期日ルール設定", expanded=False):
    c1, c2 = st.columns(2)
    months_before = c1.number_input("施行月の何ヶ月前か", min_value=1, max_value=6, value=1, step=1)
    deadline_day = c2.number_input("何日まで", min_value=1, max_value=31, value=15, step=1)
    st.info(f"現在のルール：施行月の **{months_before} ヶ月前の {deadline_day} 日** までに施工通知を発行")

st.divider()

# ── ファイルアップロード ──────────────────────────
st.subheader("ファイルアップロード")
col1, col2 = st.columns(2)
kwzweb_file = col1.file_uploader(
    "① kwzweb.csv（施工通知マスタ）",
    type=["csv"],
)
jisseki_file = col2.file_uploader(
    "② 作業実績データ.xlsx",
    type=["xlsx"],
)

if not (kwzweb_file and jisseki_file):
    st.info("2 つのファイルをアップロードすると分析が始まります。")
    st.stop()

# ── 分析 ─────────────────────────────────────────
with st.spinner("データを読み込んで分析中..."):
    try:
        kwzweb_df = load_kwzweb(kwzweb_file)
        jisseki_df = load_jisseki(jisseki_file)
        result_df, skipped = analyze(kwzweb_df, jisseki_df, int(deadline_day), int(months_before))
    except Exception as e:
        st.error(f"エラーが発生しました: {e}")
        st.stop()

st.divider()

# ── 集計メトリクス ───────────────────────────────
st.subheader("集計結果")

total = len(result_df)
ok = int((result_df["判定"] == "OK").sum())
ng = total - ok
ok_rate = ok / total * 100 if total > 0 else 0.0

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("読込件数（主工種=1）", f"{total + skipped:,} 件")
m2.metric("集計対象", f"{total:,} 件")
m3.metric("OK", f"{ok:,} 件")
m4.metric("NG", f"{ng:,} 件")
m5.metric("OK 率", f"{ok_rate:.1f} %")

if skipped > 0:
    st.caption(f"※ kwzweb に対応する施工通知が存在しない {skipped} 件は集計対象外")

# ── グラフ ────────────────────────────────────────
c1, c2 = st.columns(2)

with c1:
    fig_pie = px.pie(
        values=[ok, ng],
        names=["OK", "NG"],
        color=["OK", "NG"],
        color_discrete_map={"OK": "#00CC96", "NG": "#EF553B"},
        title="OK / NG 割合",
        hole=0.45,
    )
    fig_pie.update_traces(textinfo="percent+label")
    st.plotly_chart(fig_pie, use_container_width=True)

with c2:
    area_summary = (
        result_df.groupby(["箇所名", "判定"])
        .size()
        .reset_index(name="件数")
    )
    fig_bar = px.bar(
        area_summary,
        x="箇所名",
        y="件数",
        color="判定",
        color_discrete_map={"OK": "#00CC96", "NG": "#EF553B"},
        title="箇所別 OK / NG 件数",
        barmode="stack",
        text_auto=True,
    )
    fig_bar.update_xaxes(tickangle=30)
    st.plotly_chart(fig_bar, use_container_width=True)

# 施行月別グラフ
if "施行日" in result_df.columns:
    monthly = result_df.copy()
    monthly["施行年月"] = pd.to_datetime(monthly["施行日"]).dt.to_period("M").astype(str)
    monthly_summary = (
        monthly.groupby(["施行年月", "判定"])
        .size()
        .reset_index(name="件数")
    )
    fig_monthly = px.bar(
        monthly_summary,
        x="施行年月",
        y="件数",
        color="判定",
        color_discrete_map={"OK": "#00CC96", "NG": "#EF553B"},
        title="施行月別 OK / NG 件数",
        barmode="stack",
        text_auto=True,
    )
    st.plotly_chart(fig_monthly, use_container_width=True)

st.divider()

# ── 詳細一覧 ──────────────────────────────────────
st.subheader("詳細一覧")

show_ng_only = st.checkbox("NG のみ表示", value=False)
display_df = result_df[result_df["判定"] != "OK"].copy() if show_ng_only else result_df.copy()

for col in ["施行日", "期日"]:
    display_df[col] = pd.to_datetime(display_df[col]).dt.strftime("%Y/%m/%d")
display_df["発行日"] = pd.to_datetime(display_df["発行日"]).dt.strftime("%Y/%m/%d %H:%M")

st.dataframe(
    display_df,
    use_container_width=True,
    hide_index=True,
    column_config={
        "判定": st.column_config.TextColumn("判定", width="small"),
        "施行日": st.column_config.TextColumn("施行日", width="small"),
        "期日": st.column_config.TextColumn("期日", width="small"),
        "発行日": st.column_config.TextColumn("発行日"),
    },
)

st.caption(f"表示件数: {len(display_df):,} 件")

st.divider()

# ── Excel ダウンロード ────────────────────────────
excel_bytes = to_excel(result_df)
st.download_button(
    label="Excel ダウンロード（集計・全件・NG一覧）",
    data=excel_bytes,
    file_name="施工通知_期日確認結果.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
