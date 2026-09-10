# -*- coding: utf-8 -*-
"""资金分配工具（目标余额法）页面"""

import re
import pandas as pd
from io import BytesIO
from common import parse_amount  # 复用共享函数
import streamlit as st

# ========== 银行名称映射（简称 → 标准名称） ==========
BANK_NAME_MAP = {
    "中信": "中信",
    "上实": "上实",
    "浦发": "浦发（五羊支行）",
    "招行": "招行0001",
    "交行": "交通银行（流花支行）",
    "工行": "工商银行（大德路支行）",
    "中行": "中国银行",
    "南商": "南洋商业银行",
    "农行": "农业银行",
    "光大": "光大银行",
    "汇丰": "汇丰银行",
    "民生": "民生银行",
    "兴业": "兴业银行",
    "广发": "广发银行",
    "招行0019": "招行0019",
    "花旗": "花旗银行",
}


def to_excel_download(df_in: pd.DataFrame, df_out: pd.DataFrame) -> BytesIO:
    """把转入/转出清单打包成一个带总计行的 Excel"""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        in_total = df_in["实际调拨金额（万元）"].astype(float).sum() if not df_in.empty else 0
        in_out = df_in.copy()
        if not in_out.empty:
            in_out = pd.concat([
                in_out,
                pd.DataFrame([{"银行": "总计", "变动（万元）": in_total,
                               "实际调拨金额（万元）": in_total, "备注（自定义）": ""}])
            ], ignore_index=True)
        in_out.to_excel(writer, sheet_name="需转入", index=False)

        out_total = df_out["实际调拨金额（万元）"].astype(float).sum() if not df_out.empty else 0
        out_out = df_out.copy()
        if not out_out.empty:
            out_out = pd.concat([
                out_out,
                pd.DataFrame([{"银行": "总计", "变动（万元）": out_total,
                               "实际调拨金额（万元）": out_total, "备注（自定义）": ""}])
            ], ignore_index=True)
        out_out.to_excel(writer, sheet_name="需转出", index=False)
    buffer.seek(0)
    return buffer


def with_total_row(df: pd.DataFrame) -> pd.DataFrame:
    """在表格末尾追加一行『总计』"""
    if df.empty:
        return df
    total = df["实际调拨金额（万元）"].astype(float).sum()
    return pd.concat([
        df,
        pd.DataFrame([{"银行": "总计", "变动（万元）": total,
                       "实际调拨金额（万元）": total, "备注（自定义）": ""}])
    ], ignore_index=True)


def parse_plan(text):
    """把粘贴文本解析成 [(银行简称, 金额或None), ...]"""
    parsed = []
    for line in text.strip().split('\n'):
        if "分配" in line and "不分配" not in line and re.search(r'\d+\s*万', line):
            if not re.search(r'[\u4e00-\u9fa5]{2,}\d+\s*万', line):
                continue
        line = re.sub(r'^\s*\d+\s*[、.）)]\s*', '', line)
        if not line.strip():
            continue
        m = re.match(
            r'([\u4e00-\u9fa5A-Za-z]+)\s*[（(][^）)]*[)）]\s*([\d.]+)\s*万',
            line,
        )
        if m:
            parsed.append((m.group(1), float(m.group(2))))
            continue
        m = re.match(r'([\u4e00-\u9fa5A-Za-z]+)\s*([\d.]+)\s*万', line)
        if m:
            parsed.append((m.group(1), float(m.group(2))))
            continue
        m2 = re.match(r'([\u4e00-\u9fa5A-Za-z]+)\s*不分配', line)
        if m2:
            parsed.append((m2.group(1), None))
    return parsed


def main():
    st.title("💰 资金分配工具（目标余额法）")
    st.markdown("上传资金表，设定目标余额，未填写的银行默认不变。")

    uploaded_file = st.file_uploader("📂 上传资金表（.xlsx）", type=["xlsx"])
    if uploaded_file is None:
        st.info("👆 请先上传资金表")
        return

    excel_file = pd.ExcelFile(uploaded_file)
    sheet_names = excel_file.sheet_names
    date_sheets = [s for s in sheet_names if re.match(r'^\d+$', s)]
    if not date_sheets:
        st.error("未找到数字命名的工作表（如 1,2,3...）")
        return

    selected_date = st.selectbox("📆 选择日期", sorted(date_sheets, key=int), index=len(date_sheets) - 1)

    df = pd.read_excel(uploaded_file, sheet_name=selected_date, header=None)
    header_row = None
    for i in range(10):
        if df.iloc[i, 0] == "银行名称":
            header_row = i
            break
    if header_row is None:
        st.error("未找到“银行名称”行")
        return

    data = df.iloc[header_row + 1:].copy()
    data.columns = df.iloc[header_row]
    data = data.dropna(subset=[data.columns[0]], how='all')

    balance_col = "当日台帐余额（万元）"
    if balance_col not in data.columns:
        for col in data.columns:
            if "余额" in col and "万元" in col:
                balance_col = col
                break
        else:
            st.error("未找到余额列")
            return

    bank_col = data.columns[0]
    df_balance = data[[bank_col, balance_col]].copy()
    df_balance = df_balance[df_balance[bank_col].notna()]
    df_balance = df_balance[~df_balance[bank_col].str.contains("合计", na=False)]
    df_balance[balance_col] = pd.to_numeric(df_balance[balance_col], errors='coerce')
    df_balance = df_balance.dropna(subset=[balance_col])
    df_balance.columns = ["银行", "当前余额（万元）"]

    st.subheader(f"📊 当前余额（{selected_date}号）")
    st.dataframe(df_balance, use_container_width=True)

    # ========== 目标余额输入 ==========
    st.subheader("🎯 设定目标余额")
    st.markdown("粘贴计划，**支持带序号、括号备注、“不分配”**，自动抓取银行和金额。示例：")
    st.code("1、中信2100万\n4、招行100万(银行要求降低存款基数)\n13、广发不分配", language="text")
    plan_text = st.text_area("把计划文本整段粘贴到这里", height=180)

    plan_df = df_balance.copy()
    plan_df["目标余额（万元）"] = plan_df["当前余额（万元）"]

    if plan_text.strip():
        for raw_name, val in parse_plan(plan_text):
            mapped_name = BANK_NAME_MAP.get(raw_name, raw_name)
            for idx, row in plan_df.iterrows():
                bank_full = row["银行"]
                if bank_full == mapped_name or raw_name in bank_full or bank_full in raw_name:
                    if val is not None:
                        plan_df.at[idx, "目标余额（万元）"] = val
                    break

    edited_plan = st.data_editor(plan_df, use_container_width=True, num_rows="fixed", key="plan_editor")

    # ========== 计算 ==========
    if "result" not in st.session_state:
        st.session_state.result = None

    if st.button("🧮 计算资金流向"):
        result = edited_plan.copy()
        result["目标余额（万元）"] = pd.to_numeric(result["目标余额（万元）"], errors='coerce').fillna(result["当前余额（万元）"])
        result["变动（万元）"] = result["目标余额（万元）"] - result["当前余额（万元）"]
        result["资金流向"] = result["变动（万元）"].apply(
            lambda x: "🔴 需转出" if x < 0 else ("🟢 需转入" if x > 0 else "无变动")
        )
        st.session_state.result = result

    if st.session_state.result is not None:
        result = st.session_state.result

        total_cur = result["当前余额（万元）"].sum()
        total_tgt = result["目标余额（万元）"].sum()

        st.subheader("📈 汇总")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("当前总余额", f"{total_cur:.2f} 万元")
        col2.metric("目标总余额", f"{total_tgt:.2f} 万元")
        col3.metric("总转入", f"{result[result['变动（万元）']>0]['变动（万元）'].sum():.2f} 万元")
        col4.metric("总转出", f"{abs(result[result['变动（万元）']<0]['变动（万元）'].sum()):.2f} 万元")

        # 变动明细表 + 总计行
        st.subheader("📋 各银行变动明细")
        detail = result[["银行", "当前余额（万元）", "目标余额（万元）", "变动（万元）", "资金流向"]].copy()
        detail_total = pd.DataFrame([{
            "银行": "总计",
            "当前余额（万元）": detail["当前余额（万元）"].astype(float).sum(),
            "目标余额（万元）": detail["目标余额（万元）"].astype(float).sum(),
            "变动（万元）": detail["变动（万元）"].astype(float).sum(),
            "资金流向": "",
        }])
        st.dataframe(pd.concat([detail, detail_total], ignore_index=True), use_container_width=True, hide_index=True)

        # ========== 资金调拨指令 ==========
        st.subheader("🔁 资金调拨指令（可编辑实际金额和备注）")
        inflow = result[result["变动（万元）"] > 0].copy()
        outflow = result[result["变动（万元）"] < 0].copy()

        if not inflow.empty:
            inflow["实际调拨金额（万元）"] = inflow["变动（万元）"]
            inflow["备注（自定义）"] = ""
        else:
            inflow = pd.DataFrame(columns=["银行", "变动（万元）", "实际调拨金额（万元）", "备注（自定义）"])
        if not outflow.empty:
            outflow["实际调拨金额（万元）"] = outflow["变动（万元）"]
            outflow["备注（自定义）"] = ""
        else:
            outflow = pd.DataFrame(columns=["银行", "变动（万元）", "实际调拨金额（万元）", "备注（自定义）"])

        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### 🟢 需转入")
            edited_in = st.data_editor(
                inflow[["银行", "变动（万元）", "实际调拨金额（万元）", "备注（自定义）"]],
                use_container_width=True, num_rows="fixed", key="inflow_editor")
            st.dataframe(with_total_row(edited_in), use_container_width=True, hide_index=True)
        with c2:
            st.markdown("#### 🔴 需转出")
            edited_out = st.data_editor(
                outflow[["银行", "变动（万元）", "实际调拨金额（万元）", "备注（自定义）"]],
                use_container_width=True, num_rows="fixed", key="outflow_editor")
            st.dataframe(with_total_row(edited_out.assign(**{
                "实际调拨金额（万元）": edited_out["实际调拨金额（万元）"].astype(float).abs()
            })), use_container_width=True, hide_index=True)

        # 守恒校验
        in_total = edited_in["实际调拨金额（万元）"].astype(float).sum()
        out_total = edited_out["实际调拨金额（万元）"].astype(float).abs().sum()
        diff = in_total - out_total
        if abs(diff) < 1e-6:
            st.success(f"✅ 转入({in_total:.2f}) = 转出({out_total:.2f})，金额平衡")
        else:
            st.warning(f"⚠️ 转入 {in_total:.2f} 万 ≠ 转出 {out_total:.2f} 万，差额 {abs(diff):.2f} 万（可编辑实际金额后核对）")

        # 导出 Excel
        st.subheader("📤 最终调拨清单")
        final_in = edited_in[edited_in["实际调拨金额（万元）"].astype(float) > 0]
        final_out = edited_out[edited_out["实际调拨金额（万元）"].astype(float) < 0]
        dl = to_excel_download(final_in, final_out)
        st.download_button(
            label="⬇️ 下载调拨清单（Excel）",
            data=dl,
            file_name=f"资金调拨清单_{selected_date}号.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.caption("Excel 含两个工作表：『需转入』『需转出』，每个表末尾均带总计行。")


if __name__ == "__main__":
    main()
