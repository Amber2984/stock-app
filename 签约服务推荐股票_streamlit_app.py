import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="签约服务推荐股票统计工具", layout="wide")
st.title("📊 签约服务推荐股票交易统计工具")

uploaded_file = st.file_uploader("请上传 Excel 文件（如：副本2025.6.6-7.8号股票明细汇总.xlsx）", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)

        # 数据准备
        df["成交金额"] = pd.to_numeric(df["成交金额"], errors="coerce")
        df["手续费"] = pd.to_numeric(df["手续费"], errors="coerce")
        df["是否签约客户"] = df["是否签约"].notna() & (df["是否签约"] != "#N/A")
        df["是否双融账户"] = df["双融账户"].notna()
        df = df[df["买卖方向"] == "证券买入"].copy()

        # 团队分类映射
        dept_map = {"财富中心": "投顾团队", "营销中心": "营销团队"}
        df["团队名称"] = df["部门"].map(dept_map).fillna("独立客户")

        # 创建中间列避免 boolean index 与 groupby 不匹配错误
        df["成交金额_签约"] = df.apply(lambda row: row["成交金额"] if row["是否签约客户"] else 0, axis=1)
        df["手续费_签约"] = df.apply(lambda row: row["手续费"] if row["是否签约客户"] else 0, axis=1)
        df["成交金额_双融"] = df.apply(lambda row: row["成交金额"] if row["是否双融账户"] else 0, axis=1)
        df["手续费_双融"] = df.apply(lambda row: row["手续费"] if row["是否双融账户"] else 0, axis=1)

        # 聚合
        summary = df.groupby(["交收日期", "团队名称"]).agg(
            买入客户数=("客户代码", pd.Series.nunique),
            总成交金额_万=("成交金额", lambda x: round(x.sum() / 10000, 2)),
            总佣金收入_元=("手续费", lambda x: round(x.sum(), 2)),
            其中签约客户数=("是否签约客户", "sum"),
            其中签约成交金额_万=("成交金额_签约", lambda x: round(x.sum() / 10000, 2)),
            签约佣金收入_元=("手续费_签约", lambda x: round(x.sum(), 2)),
            签约客户佣金占比=("手续费", lambda x: round(x[df["是否签约客户"]].sum() / x.sum(), 2) if x.sum() > 0 else 0),
            双融账户买入户数=("是否双融账户", "sum"),
            双融账户买入金额_万=("成交金额_双融", lambda x: round(x.sum() / 10000, 2)),
            双融账户佣金收入_元=("手续费_双融", lambda x: round(x.sum(), 2))
        ).reset_index()

        # 保证每个日期包含所有团队
        team_order = ["投顾团队", "营销团队", "独立客户"]
        all_dates = summary["交收日期"].unique()
        full_index = pd.MultiIndex.from_product([all_dates, team_order], names=["交收日期", "团队名称"])
        summary = summary.set_index(["交收日期", "团队名称"]).reindex(full_index, fill_value=0).reset_index()

        # 格式化列名
        summary.columns = [
            "日期", "团队名称", "买入客户数", "总成交金额（万）", "总佣金收入（元）", "其中签约客户数",
            "其中签约成交金额（万）", "签约佣金收入（元）", "签约客户佣金占比",
            "双融账户买入户数", "双融账户买入金额（万）", "双融账户佣金收入（元）"
        ]

        st.success("✅ 数据处理成功！以下是结果预览：")
        st.dataframe(summary, use_container_width=True)

        # 下载按钮
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            summary.to_excel(writer, index=False, sheet_name="统计结果")
        st.download_button(
            label="📥 点击下载统计结果 Excel 文件",
            data=output.getvalue(),
            file_name="签约服务推荐股票交易统计结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ 处理数据时发生错误：{e}")
