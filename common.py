# -*- coding: utf-8 -*-
"""共享工具函数（两个页面共用）"""

import re
import pandas as pd


def parse_amount(val):
    """把各种格式的金额字符串转成 float"""
    if pd.isna(val):
        return None
    s = str(val).strip()
    s = re.sub(r'[,\s，　]', '', s)
    s = (s.replace('０', '0').replace('１', '1').replace('２', '2').replace('３', '3')
            .replace('４', '4').replace('５', '5').replace('６', '6').replace('７', '7')
            .replace('８', '8').replace('９', '9'))
    try:
        return float(s)
    except Exception:
        return None


def parse_date(val):
    if pd.isna(val):
        return None
    try:
        return pd.to_datetime(val).date()
    except Exception:
        try:
            return pd.to_datetime(float(val), unit='D', origin='1899-12-30').date()
        except Exception:
            return None


def normalize_name(val):
    """标准化供应商 / 收款人名称，便于匹配"""
    if pd.isna(val):
        return ""
    s = re.sub(r'\s+', '', str(val).strip())
    s = s.replace('（', '(').replace('）', ')')
    if '：' in s:
        s = s.split('：')[-1]
    if ':' in s:
        s = s.split(':')[-1]
    return s


def find_col(cols, keywords):
    """优先精确匹配，再模糊包含（排除编码类列）"""
    for c in cols:
        for kw in keywords:
            if str(c).strip() == kw:
                return c
    exclude = ["编码", "代码", "code", "ID", "id"]
    candidates = []
    for c in cols:
        for kw in keywords:
            if kw in str(c):
                if any(ex in str(c) for ex in exclude):
                    continue
                candidates.append(c)
                break
    if candidates:
        return candidates[0]
    for c in cols:
        for kw in keywords:
            if kw in str(c):
                return c
    return None


def extract_invoice_numbers(val):
    """从发票号字段提取所有发票号码列表"""
    if pd.isna(val):
        return []
    s = str(val).strip().replace("'", "")
    parts = re.split(r'[、，,;；\s\n]+', s)
    return [p.strip() for p in parts if re.match(r'^\d{10,}$', p.strip())]
