# -*- coding: utf-8 -*-
"""智能对账工具页面"""

import streamlit as st
import pandas as pd
from common import (  # 复用共享函数
    parse_amount, normalize_name, find_col,
    extract_invoice_numbers, parse_date,
)


def load_bank_data(uploaded_bank):
    """读取直联支付单（银行流水），支持多sheet，自动找表头"""
    xls_bank = pd.ExcelFile(uploaded_bank)
    sheet_names = xls_bank.sheet_names

    target_sheets = []
    if sheet_names:
        first_sheet = sheet_names[0]
        date_pattern = first_sheet.split('电子')[0].strip() if '电子' in first_sheet else first_sheet
        for sn in sheet_names:
            if date_pattern in sn:
                target_sheets.append(sn)
    if not target_sheets:
        target_sheets = sheet_names

    bank_raw_list = []
    for sn in target_sheets:
        df = pd.read_excel(xls_bank, sheet_name=sn, dtype=str, header=None)
        header_idx = None
        for i, row in df.iterrows():
            row_str = ' '.join([str(x) for x in row.tolist()])
            if '单据编号' in row_str or '收款人' in row_str or '付款金额' in row_str:
                header_idx = i
                break
        if header_idx is not None:
            df.columns = df.iloc[header_idx].tolist()
            df = df.iloc[header_idx + 1:].reset_index(drop=True)
        else:
            df.columns = df.iloc[0].tolist()
            df = df.iloc[1:].reset_index(drop=True)
        df = df.dropna(how='all')
        df = df[~df.apply(lambda r: r.astype(str).str.contains('打印|共 （').any(), axis=1)]
        bank_raw_list.append(df)

    return pd.concat(bank_raw_list, ignore_index=True)


def load_payment_data(uploaded_account):
    """读取付款表（财务账），自动检测表头位置"""
    df = pd.read_excel(uploaded_account, dtype=str, header=None)
    header_idx = None
    for i, row in df.iterrows():
        row_str = ' '.join([str(x) for x in row.tolist()])
        if '供应商' in row_str or '付款金额' in row_str or '含税金额' in row_str or '发票号' in row_str:
            header_idx = i
            break
    if header_idx is not None:
        df.columns = df.iloc[header_idx].tolist()
        df = df.iloc[header_idx + 1:].reset_index(drop=True)
    else:
        df.columns = df.iloc[0].tolist()
        df = df.iloc[1:].reset_index(drop=True)
    return df


def match_records(bank_proc, account_proc, amount_col_b, amount_col_a, name_col_b, name_col_a, invoice_col_a=None):
    """三层匹配：金额+名称精确 → 模糊 → 发票号"""
    account_proc = account_proc.copy()
    account_proc["已匹配"] = False
    match_results = []

    if invoice_col_a and invoice_col_a in account_proc.columns:
        account_proc["发票号列表"] = account_proc[invoice_col_a].apply(extract_invoice_numbers)

    for _, b_row in bank_proc.iterrows():
        b_amt = b_row["金额_num"]
        b_name = b_row["名称_norm"]
        match_type = "⚠️ 未匹配"
        match_detail = ""

        if pd.notna(b_amt) and b_name:
            for a_idx, a_row in account_proc[~account_proc["已匹配"]].iterrows():
                a_amt = a_row["金额_num"]
                a_name = a_row["名称_norm"]
                if pd.notna(a_amt) and abs(b_amt - a_amt) < 0.01:
                    if b_name == a_name:
                        account_proc.at[a_idx, "已匹配"] = True
                        match_type = "✅ 匹配成功（金额+名称精确）"
                        match_detail = f"付款表第{a_idx+1}行"
                        break
                    elif b_name in a_name or a_name in b_name:
                        account_proc.at[a_idx, "已匹配"] = True
                        match_type = "✅ 匹配成功（金额+名称模糊）"
                        match_detail = f"付款表第{a_idx+1}行"
                        break

        res = b_row.to_dict()
        res["匹配状态"] = match_type
        res["匹配说明"] = match_detail
        match_results.append(res)

    return match_results, account_proc


def main():
    st.title("🤖 智能对账工具")
    st.markdown("""
    ### 使用说明
    1. **左侧**上传 **直联支付单**（银行流水/实际付款记录）
    2. **右侧**上传 **付款表**（器械付款明细/财务账）
    3. 系统自动按 **金额 + 供应商/收款人名称** 进行匹配
    4. 绿色=匹配成功，红色=未匹配（需人工核查）
    """)

    col1, col2 = st.columns(2)
    with col1:
        uploaded_bank = st.file_uploader("📂 银行流水（直联支付单）", type=["xlsx", "xls"], key="bank")
    with col2:
        uploaded_account = st.file_uploader("📂 付款表（器械付款明细）", type=["xlsx", "xls"], key="acc")

    if not (uploaded_bank and uploaded_account):
        st.info("👆 请上传两个文件后开始对账")
        return

    try:
        with st.spinner("正在读取文件..."):
            bank_raw = load_bank_data(uploaded_bank)
            account_raw = load_payment_data(uploaded_account)

        bank_cols = bank_raw.columns.tolist()
        acc_cols = account_raw.columns.tolist()

        amount_col_bank = find_col(bank_cols, ["付款金额", "金额"])
        name_col_bank = find_col(bank_cols, ["收款人名称", "收款人", "对手", "户名", "供应商名称"])
        date_col_bank = find_col(bank_cols, ["单据日期", "日期"])
        doc_col_bank = find_col(bank_cols, ["单据编号"])

        amount_col_acc = find_col(acc_cols, ["付款金额", "含税金额", "金额"])
        name_col_acc = find_col(acc_cols, ["供应商名称", "供应商", "收款人名称", "收款人", "户名"])
        invoice_col_acc = find_col(acc_cols, ["发票号", "发票"])

        st.subheader("🔧 列映射确认")
        c1, c2 = st.columns(2)
        with c1:
            st.write("**银行流水**")
            all_b = [""] + bank_cols
            amount_col_bank_sel = st.selectbox("金额列", all_b, index=all_b.index(amount_col_bank) if amount_col_bank in all_b else 0, key="ab")
            name_col_bank_sel = st.selectbox("名称列(收款人)", all_b, index=all_b.index(name_col_bank) if name_col_bank in all_b else 0, key="nb")
            date_col_bank_sel = st.selectbox("日期列(可选)", all_b, index=all_b.index(date_col_bank) if date_col_bank in all_b else 0, key="db")
        with c2:
            st.write("**付款表**")
            all_a = [""] + acc_cols
            amount_col_acc_sel = st.selectbox("金额列", all_a, index=all_a.index(amount_col_acc) if amount_col_acc in all_a else 0, key="aa")
            name_col_acc_sel = st.selectbox("名称列(供应商)", all_a, index=all_a.index(name_col_acc) if name_col_acc in all_a else 0, key="na")
            invoice_col_acc_sel = st.selectbox("发票号列(可选)", all_a, index=all_a.index(invoice_col_acc) if invoice_col_acc in all_a else 0, key="ia")

        if not all([amount_col_bank_sel, amount_col_acc_sel, name_col_bank_sel, name_col_acc_sel]):
            st.warning("⚠️ 请确认4个必填列（金额+名称）都已正确映射！")
            st.write("银行流水列：", bank_cols)
            st.write("付款表列：", acc_cols)
            return

        bank_proc = bank_raw.copy()
        account_proc = account_raw.copy()
        bank_proc["金额_num"] = bank_proc[amount_col_bank_sel].apply(parse_amount)
        account_proc["金额_num"] = account_proc[amount_col_acc_sel].apply(parse_amount)
        bank_proc["名称_norm"] = bank_proc[name_col_bank_sel].apply(normalize_name)
        account_proc["名称_norm"] = account_proc[name_col_acc_sel].apply(normalize_name)
        bank_proc = bank_proc[bank_proc["金额_num"].notna()].reset_index(drop=True)
        account_proc = account_proc[account_proc["金额_num"].notna()].reset_index(drop=True)
        summary_mask_b = bank_proc[amount_col_bank_sel].astype(str).str.contains('付款|合计|总计', na=False)
        bank_proc = bank_proc[~summary_mask_b].reset_index(drop=True)
        summary_mask_a = account_proc[amount_col_acc_sel].astype(str).str.contains('付款|合计|总计', na=False)
        account_proc = account_proc[~summary_mask_a].reset_index(drop=True)

        st.info(f"📊 银行流水: **{len(bank_proc)}** 条 | 付款表: **{len(account_proc)}** 条")
        if len(bank_proc) == 0 or len(account_proc) == 0:
            st.error("❌ 有效数据行为0，请检查列映射是否正确！")
            return

        with st.spinner("正在匹配..."):
            match_results, account_proc = match_records(
                bank_proc, account_proc,
                amount_col_bank_sel, amount_col_acc_sel,
                name_col_bank_sel, name_col_acc_sel,
                invoice_col_acc_sel if invoice_col_acc_sel else None,
            )

        result_df = pd.DataFrame(match_results)
        total_bank = len(bank_proc)
        matched_count = sum(1 for x in match_results if "✅" in x["匹配状态"])
        match_rate = matched_count / total_bank if total_bank > 0 else 0
        unmatched_in_account = account_proc[~account_proc["已匹配"]]

        st.subheader("📊 对账结果")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("匹配成功率", f"{match_rate:.1%}")
        m2.metric("✅ 已匹配", matched_count)
        m3.metric("⚠️ 银行流水未匹配", total_bank - matched_count)
        m4.metric("⚠️ 付款表未匹配", len(unmatched_in_account))

        bank_total = bank_proc["金额_num"].sum()
        acc_total = account_proc["金额_num"].sum()
        diff = bank_total - acc_total
        st.write(f"💰 银行流水总金额: **{bank_total:,.2f}** | 付款表总金额: **{acc_total:,.2f}** | 差异: **{diff:,.2f}**")

        if matched_count == 0:
            st.warning("⚠️ **未匹配到任何记录**，请检查：两个文件是否同一批次、供应商名称是否一致、金额是否有精度差异。")

        tab1, tab2, tab3 = st.tabs(["✅ 匹配成功", "⚠️ 银行流水未匹配", "⚠️ 付款表未匹配"])
        with tab1:
            matched_df = result_df[result_df["匹配状态"].str.contains("✅")]
            st.write(f"共 {len(matched_df)} 条")
            if len(matched_df) > 0:
                show_cols = [c for c in matched_df.columns if c not in ["名称_norm", "金额_num"]]
                st.dataframe(matched_df[show_cols], use_container_width=True)
        with tab2:
            unmatched_bank = result_df[~result_df["匹配状态"].str.contains("✅")]
            st.write(f"共 {len(unmatched_bank)} 条银行流水未在付款表中找到匹配")
            if len(unmatched_bank) > 0:
                show_cols = [c for c in unmatched_bank.columns if c not in ["名称_norm", "金额_num"]]
                st.dataframe(unmatched_bank[show_cols], use_container_width=True)
        with tab3:
            st.write(f"共 {len(unmatched_in_account)} 条付款表记录未在银行流水中找到匹配")
            if len(unmatched_in_account) > 0:
                show_cols = [c for c in unmatched_in_account.columns if c not in ["已匹配", "名称_norm", "金额_num", "发票号列表"]]
                st.dataframe(unmatched_in_account[show_cols], use_container_width=True)

        # 导出
        st.subheader("📥 导出结果")
        out = pd.ExcelWriter("对账结果.xlsx", engine='openpyxl')
        result_df.drop(columns=["名称_norm", "金额_num"], errors='ignore').to_excel(out, sheet_name="银行流水匹配结果", index=False)
        unmatched_in_account.drop(columns=["已匹配", "名称_norm", "金额_num", "发票号列表"], errors='ignore').to_excel(out, sheet_name="付款表未匹配", index=False)
        out.close()
        with open("对账结果.xlsx", "rb") as f:
            st.download_button(
                label="📥 下载对账结果 Excel",
                data=f.read(),
                file_name="对账结果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as e:
        st.error(f"❌ 处理出错: {e}")
        import traceback
        st.text(traceback.format_exc())


if __name__ == "__main__":
    main()
