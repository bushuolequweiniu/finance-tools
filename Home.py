# -*- coding: utf-8 -*-
"""首页：介绍两个功能，提供导航入口"""

import streamlit as st

st.set_page_config(page_title="财务办公工具箱", layout="wide")

st.title("📊 财务办公工具箱")
st.markdown(
    """
    欢迎使用！本工具箱集成了日常财务对账与资金调度的常用功能，
    **左侧导航栏**选择对应功能即可使用。
    """
)

st.divider()

c1, c2 = st.columns(2)

with c1:
    st.subheader("🤖 智能对账工具")
    st.markdown(
        """
        - 上传 **直联支付单**（银行流水）
        - 上传 **付款表**（器械付款明细 / 财务账）
        - 自动按「金额 + 供应商名称」匹配
        - 支持发票号二次匹配，结果可导出 Excel
        """
    )
    "st.page_link("pages/1_Reconciliation.py", label="👉 去对账", icon="💼")"

with c2:
    st.subheader("💰 资金分配工具")
    st.markdown(
        """
        - 上传资金表，选择日期
        - 粘贴分配计划（支持序号 / 括号备注 / 「不分配」）
        - 自动计算各银行转入 / 转出
        - 调拨清单可导出 Excel
        """
    )
    st.page_link("pages/2_Allocation.py", label="👉 去分配资金 →", icon="💰")

st.divider()
st.caption("💡 提示：两个功能共享同一套数据，可先后配合使用。")
