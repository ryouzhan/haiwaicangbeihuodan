#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""海外仓备货发货单智能处理工具 (Web版) - 格式100%严格对齐标准发货单导出模板"""

from collections import defaultdict
from datetime import datetime
import io
import json
import os
import re
import sys
from typing import Any, Dict, List, Optional, Tuple
from cryptography.fernet import Fernet
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

# ==================== 1. 页面全局配置与极简样式 ====================
st.set_page_config(
    page_title="发货单处理工具",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
<style>
    .stApp {
        background-color: #F8FAFC !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
        color: #1E293B !important;
    }
    #MainMenu, footer, header {visibility: hidden;}

    .header-box { padding: 1.5rem 0 0.8rem 0; margin-bottom: 0.6rem; }
    .header-badge {
        display: inline-block; padding: 3px 10px; font-size: 0.72rem; font-weight: 600;
        text-transform: uppercase; background: #E2E8F0; color: #475569;
        border-radius: 9999px; margin-bottom: 0.3rem;
    }
    .header-title { font-size: 1.85rem; font-weight: 700; color: #0F172A; margin: 0; }
    .header-subtitle { font-size: 0.88rem; color: #64748B; margin-top: 0.2rem; }

    div[data-testid="stPopover"] > button {
        border-radius: 20px !important;
        padding: 4px 14px !important;
        font-size: 0.82rem !important;
        background: #FFFFFF !important;
        border: 1px solid #CBD5E1 !important;
        color: #334155 !important;
        height: auto !important;
        display: inline-flex !important;
        align-items: center !important;
        box-shadow: 0 1px 2px rgba(0,0,0,0.04) !important;
        transition: all 0.2s ease;
    }
    div[data-testid="stPopover"] > button:hover {
        border-color: #2563EB !important;
        color: #1E40AF !important;
        background: #F8FAFC !important;
    }

    [data-testid="stFileUploader"] section {
        background-color: #FFFFFF !important;
        border: 1.5px dashed #CBD5E1 !important;
        border-radius: 12px !important;
        padding: 1.5rem 1rem !important;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.02) !important;
    }
    [data-testid="stFileUploader"] section:hover {
        border-color: #2563EB !important;
        background-color: #F8FAFC !important;
    }

    .metric-container {
        display: grid;
        grid-template-columns: repeat(6, 1fr);
        gap: 0.8rem;
        margin: 1.2rem 0 1.5rem 0;
    }
    .metric-card {
        background: #FFFFFF;
        padding: 1rem 0.8rem;
        border-radius: 10px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 1px 2px rgba(0,0,0,0.03);
    }
    .metric-title { font-size: 0.75rem; font-weight: 500; color: #64748B; margin-bottom: 0.3rem; }
    .metric-num { font-size: 1.45rem; font-weight: 700; color: #0F172A; line-height: 1.1; }
    .metric-unit { font-size: 0.75rem; font-weight: 500; color: #94A3B8; margin-left: 0.2rem; }
</style>
""",
    unsafe_allow_html=True,
)


# ==================== 2. 商品库在线解密与映射 ====================
MAPPING_FILE = "sku_mapping.json"


def load_sku_mapping() -> Dict[str, str]:
  if os.path.exists(MAPPING_FILE):
    try:
      with open(MAPPING_FILE, "r", encoding="utf-8") as f:
        return json.load(f)
    except Exception:
      pass
  return {}


def save_sku_mapping(mapping: Dict[str, str]) -> None:
  try:
    with open(MAPPING_FILE, "w", encoding="utf-8") as f:
      json.dump(mapping, f, ensure_ascii=False, indent=2)
  except Exception:
    pass


def get_secret_key() -> Optional[bytes]:
  try:
    if "COMMODITIES_KEY" in st.secrets:
      return st.secrets["COMMODITIES_KEY"].encode()
  except Exception:
    pass
  if os.path.exists("secret.key"):
    try:
      with open("secret.key", "rb") as f:
        return f.read().strip()
    except Exception:
      pass
  return None


def parse_raw_table_bytes(raw_bytes: bytes) -> Optional[pd.DataFrame]:
  try:
    if raw_bytes.startswith(b"PK\x03\x04") or raw_bytes.startswith(
        b"\xd0\xcf\x11\xe0"
    ):
      return pd.read_excel(io.BytesIO(raw_bytes))
    else:
      try:
        return pd.read_csv(io.BytesIO(raw_bytes), encoding="utf-8-sig")
      except Exception:
        return pd.read_csv(io.BytesIO(raw_bytes), encoding="gbk")
  except Exception:
    return None


def load_active_commodities() -> Tuple[Optional[pd.DataFrame], str]:
  key = get_secret_key()
  if os.path.exists("commodities.dat"):
    if not key:
      return None, "未配置解密密钥 (请在 Secrets 填入 COMMODITIES_KEY)"
    try:
      with open("commodities.dat", "rb") as f:
        cipher_data = f.read()
      cipher = Fernet(key)
      decrypted_bytes = cipher.decrypt(cipher_data)
      df = parse_raw_table_bytes(decrypted_bytes)
      if df is not None:
        df.columns = [str(c).strip() for c in df.columns]
        return df, "商品库 (已加密安全同步)"
    except Exception as e:
      return None, f"解密失败: {e}"

  valid_exts = (".xlsx", ".xls", ".csv")
  candidates = []
  for fname in os.listdir("."):
    if fname.startswith("~$"):
      continue
    if any(fname.lower().endswith(ext) for ext in valid_exts) and "commodit" in fname.lower():
      full_path = os.path.join(".", fname)
      candidates.append((full_path, fname, os.path.getmtime(full_path)))

  if candidates:
    candidates.sort(key=lambda x: x, reverse=True)
    latest_path, latest_name, _ = candidates[0]
    try:
      with open(latest_path, "rb") as f:
        df = parse_raw_table_bytes(f.read())
      if df is not None:
        df.columns = [str(c).strip() for c in df.columns]
        return df, latest_name
    except Exception:
      pass

  return None, "未检测到商品库"


# ==================== 3. 复合表格解析与单号精准提取 ====================
def parse_raw_order_file(uploaded_file):
  filename = str(uploaded_file.name).lower()
  df_raw = (
      pd.read_csv(uploaded_file, header=None)
      if filename.endswith(".csv")
      else pd.read_excel(uploaded_file, header=None)
  )

  cut_idx = len(df_raw)
  split_keywords = ["备货单号", "辅料SKU", "关联备货单号", "三方仓入库单号"]
  for idx in range(1, len(df_raw)):
    first_cell = str(df_raw.iloc[idx, 0]).strip()
    if any(k in first_cell for k in split_keywords):
      cut_idx = idx
      break

  goods_headers = [str(c).strip() for c in df_raw.iloc[0].values]
  goods_df = df_raw.iloc[1:cut_idx].copy()
  goods_df.columns = goods_headers
  goods_df = goods_df.loc[:, ~goods_df.columns.str.startswith("Unnamed")]
  goods_df.dropna(how="all", inplace=True)

  # 从【物流信息】中定位【关联备货单号】
  related_order_code = ""
  for row_idx in range(cut_idx, len(df_raw)):
    row_vals = [str(v).strip() for v in df_raw.iloc[row_idx].values]
    if "关联备货单号" in row_vals:
      col_idx = row_vals.index("关联备货单号")
      if row_idx + 1 < len(df_raw):
        val = str(df_raw.iloc[row_idx + 1, col_idx]).strip()
        if val and val not in ("nan", "None"):
          related_order_code = val
          break

  if not related_order_code:
    for row_idx in range(cut_idx, len(df_raw)):
      for cell in df_raw.iloc[row_idx].dropna():
        s = str(cell).strip()
        m = re.search(r"\b(OWS[A-Za-z0-9\-]+|FBA[A-Za-z0-9\-]+)\b", s)
        if m:
          related_order_code = m.group(1)
          break
      if related_order_code:
        break

  return goods_df, (related_order_code or "")


def get_warehouse_code(addr):
  """提取物流中心编码，附带美西/美东分区，如 'GYR3(美西)' 或 'AWD仓'"""
  addr = str(addr).strip()
  if not addr or addr in ("nan", "None"):
    return "AWD仓"
  u_addr = addr.upper()
  if "AWD" in u_addr:
    return "AWD仓"

  east = ["ABE8", "AVP1", "DCA6", "TEB9", "PHL7", "BDL3", "RDU2", "TEB6"]
  central = [
      "DFW6",
      "FOE1",
      "MDW2",
      "IND7",
      "MEM1",
      "RFD2",
      "IND9",
      "AKR1",
      "ITX3",
  ]
  west = [
      "IUTE",
      "ONT8",
      "LAS1",
      "LAX9",
      "GYR2",
      "GYR3",
      "ABQ2",
      "SCK4",
      "SMF3",
      "SBD1",
      "PSP3",
      "TCY1",
  ]

  for w in west:
    if w in u_addr:
      return f"{w}(美西)"
  for c in central:
    if c in u_addr:
      return f"{c}(美中)"
  for e in east:
    if e in u_addr:
      return f"{e}(美东)"
  return addr


def extract_pcs_from_text(text):
  match = re.search(r"(\d+)\s*(?:pc|pcs|PC|PCS|只|件|套)", str(text), re.I)
  return int(match.group(1)) if match else 1


# ==================== 4. 核心计算与严格标准格式构建 ====================
def process_shipment_data(
    goods_df, order_code, commodities_df, sku_mapping=None
):
  goods_df["SKU"] = goods_df["SKU"].astype(str).str.strip()

  comm_dict = {}
  if commodities_df is not None and not commodities_df.empty:
    sku_col = next(
        (c for c in commodities_df.columns if str(c).strip().upper() == "SKU"),
        None,
    )
    if sku_col:
      for _, r in commodities_df.iterrows():
        s = str(r[sku_col]).strip()
        comm_dict[s] = r.to_dict()

  merged_rows = []
  missing_skus = []

  for _, row in goods_df.iterrows():
    raw_sku = str(row["SKU"]).strip()
    lookup_sku = (
        sku_mapping.get(raw_sku, raw_sku) if sku_mapping else raw_sku
    )

    matched = comm_dict.get(lookup_sku)
    if not matched and lookup_sku.upper().startswith("ZF-"):
      matched = comm_dict.get(lookup_sku[3:])
    if not matched and not lookup_sku.upper().startswith("ZF-"):
      matched = comm_dict.get(f"ZF-{lookup_sku}")
    if not matched:
      matched = comm_dict.get(raw_sku, {})

    if not matched:
      missing_skus.append(raw_sku)

    item = {**row.to_dict()}

    # 品名、图片与供应商
    item["_品名"] = item.get("品名") or matched.get(
        "品名", matched.get("中文品名", "")
    )
    item["_图片"] = item.get("商品图片") or item.get("图片") or matched.get("图片", "")
    item["_供应商"] = matched.get(
        "供应商名称", matched.get("供应商", matched.get("商品品牌", ""))
    )

    # 单箱数量（箱规）
    local_carton = row.get("单箱数量") or row.get("单箱数量(pcs)") or 0
    cloud_carton = matched.get("单箱数量(pcs)", matched.get("单箱数量", 0))
    final_carton = pd.to_numeric(
        local_carton if local_carton and local_carton != 0 else cloud_carton,
        errors="coerce",
    )
    item["_单箱数量"] = (
        int(final_carton)
        if not np.isnan(final_carton) and final_carton > 0
        else 0
    )

    # 单品 PCS
    pcs_val = pd.to_numeric(
        matched.get("单品PCS", matched.get("PCS", 0)), errors="coerce"
    )
    item["_单品PCS"] = (
        int(pcs_val)
        if not np.isnan(pcs_val) and pcs_val > 0
        else extract_pcs_from_text(item["_品名"])
    )

    # 采购单价
    price_val = 0.0
    for p_col in [
        "采购单价(CNY)",
        "采购成本(￥)",
        "指定采购单价(CNY)",
        "采购单价",
        "单价",
    ]:
      if p_col in matched:
        val = pd.to_numeric(matched[p_col], errors="coerce")
        if not np.isnan(val) and val > 0:
          price_val = val
          break
    item["_单价"] = price_val

    # 规格尺寸与重量
    def get_dim(keys, fallback=0.0):
      for k in keys:
        if k in matched:
          v = pd.to_numeric(matched[k], errors="coerce")
          if not np.isnan(v) and v > 0:
            return float(v)
      return float(fallback)

    item["_重量"] = get_dim(
        ["单箱重量(kg)", "重量(kg)", "毛重(kg)", "单箱重量"],
        row.get("单箱重量(kg)", 0),
    )
    item["_长"] = get_dim(
        ["箱规长(cm)", "外箱长(cm)", "长(cm)", "长"], row.get("箱规长(cm)", 0)
    )
    item["_宽"] = get_dim(
        ["箱规宽(cm)", "外箱宽(cm)", "宽(cm)", "宽"], row.get("箱规宽(cm)", 0)
    )
    item["_高"] = get_dim(
        ["箱规高(cm)", "外箱高(cm)", "高(cm)", "高"], row.get("箱规高(cm)", 0)
    )

    merged_rows.append(item)

  df_m = pd.DataFrame(merged_rows)

  # ---------- 1. 构建严格匹配的【详细数据】(23列) ----------
  detail_cols = [
      "SKU",
      "品名",
      "商品图片",
      "发货量",
      "箱数",
      "单箱数量",
      "物流中心编码",
      "供应商",
      "货件编号",
      "ReferenceId",
      "箱号",
      "总箱数编号",
      "外箱重量(kg)",
      "外箱总重量(kg)",
      "外箱长(cm)",
      "外箱宽(cm)",
      "外箱高(cm)",
      "外箱体积(m³)",
      "外箱总体积(m³)",
      "外箱总体积重(kg)",
      "创建时间",
      "发货时间",
      "物流商",
  ]

  df_detail = pd.DataFrame(columns=detail_cols)
  df_detail["SKU"] = df_m["SKU"]
  df_detail["品名"] = df_m["_品名"]
  df_detail["商品图片"] = df_m["_图片"]

  # 单箱数量
  df_detail["单箱数量"] = df_m["_单箱数量"]

  # 箱数与发货量
  if "箱数" in df_m.columns and "备货量" in df_m.columns:
    df_detail["箱数"] = (
        pd.to_numeric(df_m["箱数"], errors="coerce").fillna(0).astype(int)
    )
    df_detail["发货量"] = (
        pd.to_numeric(df_m["备货量"], errors="coerce").fillna(0).astype(int)
    )
    df_detail["发货量"] = np.where(
        df_detail["发货量"] > 0,
        df_detail["发货量"],
        df_detail["箱数"] * df_detail["单箱数量"],
    )
  else:
    raw_qty = (
        pd.to_numeric(
            df_m.get("发货量", df_m.get("申报量", df_m.get("备货量", 0))),
            errors="coerce",
        )
        .fillna(0)
        .astype(int)
    )
    safe_c = np.where(df_detail["单箱数量"] > 0, df_detail["单箱数量"], 1)
    df_detail["箱数"] = np.where(
        df_detail["单箱数量"] > 0, np.ceil(raw_qty / safe_c).astype(int), 0
    )
    df_detail["发货量"] = raw_qty

  # 物流中心编码与货件编号
  addr_raw = ""
  for col in ["物流中心编码", "收货仓库", "配送地址"]:
    if col in df_m.columns and not df_m[col].dropna().empty:
      addr_raw = str(df_m[col].dropna().iloc[0]).strip()
      break
  df_detail["物流中心编码"] = get_warehouse_code(addr_raw)

  df_detail["供应商"] = df_m["_供应商"]
  df_detail["货件编号"] = order_code
  df_detail["ReferenceId"] = df_m.get("ReferenceId", "")

  # 自动生成连续箱号 (如 1-9, 10-18)
  box_nums = []
  current_box_start = 1
  for b_count in df_detail["箱数"]:
    if b_count <= 0:
      box_nums.append("")
    elif b_count == 1:
      box_nums.append(str(current_box_start))
      current_box_start += 1
    else:
      end_box = current_box_start + b_count - 1
      box_nums.append(f"{current_box_start}-{end_box}")
      current_box_start = end_box + 1
  df_detail["箱号"] = box_nums

  total_box_count = int(df_detail["箱数"].sum())
  df_detail["总箱数编号"] = total_box_count if total_box_count > 0 else ""

  # 重量与尺寸
  df_detail["外箱重量(kg)"] = np.where(
      df_m["_重量"] > 0, df_m["_重量"].round(2), ""
  )
  total_w = (df_m["_重量"] * df_detail["箱数"]).round(2)
  df_detail["外箱总重量(kg)"] = np.where(total_w > 0, total_w, "")

  df_detail["外箱长(cm)"] = np.where(
      df_m["_长"] > 0, df_m["_长"].round(1), ""
  )
  df_detail["外箱宽(cm)"] = np.where(
      df_m["_宽"] > 0, df_m["_宽"].round(1), ""
  )
  df_detail["外箱高(cm)"] = np.where(
      df_m["_高"] > 0, df_m["_高"].round(1), ""
  )

  # 体积与体积重
  vol = (df_m["_长"] * df_m["_宽"] * df_m["_高"]) / 1000000
  df_detail["外箱体积(m³)"] = np.where(vol > 0, vol.round(4), "")
  tot_vol = (vol * df_detail["箱数"]).round(4)
  df_detail["外箱总体积(m³)"] = np.where(tot_vol > 0, tot_vol, "")

  tot_vwt = (tot_vol * 167).round(2)
  df_detail["外箱总体积重(kg)"] = np.where(tot_vwt > 0, tot_vwt, "")

  # 时间与物流商
  created_t = df_m.get("创建时间", "")
  if isinstance(created_t, pd.Series) and not created_t.dropna().empty:
    df_detail["创建时间"] = created_t.iloc[0]
  else:
    df_detail["创建时间"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

  df_detail["发货时间"] = df_m.get("发货时间", "")
  df_detail["物流商"] = df_m.get("物流商", "")

  # 空值彻底格式化为空字符串（留空）
  df_detail.fillna("", inplace=True)
  df_detail.replace({"nan": "", "None": "", np.nan: ""}, inplace=True)

  # ---------- 2. 构建严格匹配的【汇总结果】(9列) ----------
  # 按照模板定义：['物流中心编码', '货件编号', 'ReferenceId', '品名', '供应商', '总箱数', '外箱总重量', '外箱总体积', '外箱总体积重(kg)']
  num_boxes = int(df_detail["箱数"].sum())
  sum_weight = (
      pd.to_numeric(df_detail["外箱总重量(kg)"], errors="coerce")
      .fillna(0)
      .sum()
  )
  sum_volume = (
      pd.to_numeric(df_detail["外箱总体积(m³)"], errors="coerce")
      .fillna(0)
      .sum()
  )
  sum_vol_weight = (
      pd.to_numeric(df_detail["外箱总体积重(kg)"], errors="coerce")
      .fillna(0)
      .sum()
  )

  summary_row = {
      "物流中心编码": df_detail["物流中心编码"].iloc[0]
      if not df_detail.empty
      else "AWD仓",
      "货件编号": order_code,
      "ReferenceId": df_detail["ReferenceId"].iloc[0]
      if not df_detail.empty
      else "",
      "品名": df_detail["品名"].iloc[0] if not df_detail.empty else "",
      "供应商": df_detail["供应商"].iloc[0] if not df_detail.empty else "",
      "总箱数": num_boxes,
      "外箱总重量": round(sum_weight, 2) if sum_weight > 0 else "",
      "外箱总体积": round(sum_volume, 4) if sum_volume > 0 else "",
      "外箱总体积重(kg)": round(sum_vol_weight, 2)
      if sum_vol_weight > 0
      else "",
  }
  df_summary = pd.DataFrame([summary_row])
  df_summary.fillna("", inplace=True)
  df_summary.replace({"nan": "", "None": "", np.nan: ""}, inplace=True)

  # 附加计算字段供网页看板使用
  total_pcs_sum = int((df_detail["发货量"] * df_m["_单品PCS"]).sum())
  total_val_sum = (df_detail["发货量"] * df_m["_单价"]).round(2).sum()

  kpi_metrics = {
      "total_box": num_boxes,
      "total_sets": int(df_detail["发货量"].sum()),
      "total_pcs": total_pcs_sum,
      "total_weight": f"{sum_weight:,.2f}" if sum_weight > 0 else "0",
      "total_volume": f"{sum_volume:,.3f}" if sum_volume > 0 else "0",
      "total_val": f"￥{total_val_sum:,.2f}",
  }

  return df_detail, df_summary, kpi_metrics, list(set(missing_skus))


# ==================== 5. 专业 Excel 导出美化 (原汁原味) ====================
def export_and_beautify(df_detail, df_summary):
  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df_detail.to_excel(writer, sheet_name="详细数据", index=False)
    df_summary.to_excel(writer, sheet_name="汇总结果", index=False)

  wb = load_workbook(output)
  header_fill = PatternFill(
      start_color="1E293B", end_color="1E293B", fill_type="solid"
  )
  header_font = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
  stripe_fill = PatternFill(
      start_color="F8FAFC", end_color="F8FAFC", fill_type="solid"
  )
  thin_border = Border(
      left=Side(style="thin", color="CBD5E1"),
      right=Side(style="thin", color="CBD5E1"),
      top=Side(style="thin", color="CBD5E1"),
      bottom=Side(style="thin", color="CBD5E1"),
  )
  center_align = Alignment(
      horizontal="center", vertical="center", wrap_text=True
  )

  for sheetname in ["详细数据", "汇总结果"]:
    ws = wb[sheetname]
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row)):
      for cell in row:
        cell.border = thin_border
        cell.alignment = center_align
        if i == 0:
          cell.fill = header_fill
          cell.font = header_font
        elif i % 2 == 0:
          cell.fill = stripe_fill

    for col in ws.columns:
      max_len = 0
      col_letter = col[0].column_letter
      for cell in col:
        v = str(cell.value) if cell.value is not None else ""
        try:
          length = len(v.encode("gbk"))
          if length > max_len:
            max_len = length
        except:
          pass
      ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 48)

  final_stream = io.BytesIO()
  wb.save(final_stream)
  return final_stream.getvalue()


# ==================== 6. 主程序与界面交互 ====================
def main():
  st.markdown(
      """
    <div class="header-box">
        <div class="header-badge">Shipment Generator V9.2</div>
        <h1 class="header-title">发货单智能生成工具</h1>
        <p class="header-subtitle">输出格式100%对齐标准模板 · 物流关联单号联动 · 智能商品库规格匹配</p>
    </div>
    """,
      unsafe_allow_html=True,
  )

  # 1. 加载特殊映射
  if "sku_mapping" not in st.session_state:
    st.session_state["sku_mapping"] = load_sku_mapping()
  current_mapping = st.session_state["sku_mapping"]

  # 2. 自动定位商品库
  active_df, table_label = load_active_commodities()

  custom_uploaded = st.session_state.get("custom_commodities", None)
  if custom_uploaded is not None:
    table_pill_label = f"🟢 自定义: {custom_uploaded.name} ▾"
  elif active_df is not None:
    table_pill_label = f"🟢 {table_label} ▾"
  else:
    table_pill_label = f"🔴 {table_label} ▾"

  map_count = len(current_mapping)
  mapping_pill_label = (
      f"⚡ 特殊映射 ({map_count}条) ▾" if map_count > 0 else "⚡ 特殊映射 ▾"
  )

  col_p1, col_p2, _ = st.columns([1.5, 1.2, 1.3])

  with col_p1:
    with st.popover(table_pill_label):
      st.caption("临时更换商品库（仅本次生效）：")
      st.file_uploader(
          "上传替代商品列表",
          type=["xlsx", "xls", "csv"],
          label_visibility="collapsed",
          key="custom_commodities",
      )
      if custom_uploaded is not None and st.button(
          "恢复默认商品库", use_container_width=True
      ):
        del st.session_state["custom_commodities"]
        st.rerun()

  with col_p2:
    with st.popover(mapping_pill_label):
      st.caption("双击编辑，支持从 Excel 复制两列直接粘贴：")
      rows = [
          {"面单SKU": k, "商品库SKU": v} for k, v in current_mapping.items()
      ]
      if not rows:
        rows = [{"面单SKU": "", "商品库SKU": ""}]
      df_mapping = pd.DataFrame(rows)

      edited_df = st.data_editor(
          df_mapping,
          num_rows="dynamic",
          use_container_width=True,
          hide_index=True,
          height=200,
          column_config={
              "面单SKU": st.column_config.TextColumn(
                  "备货单 SKU", required=True
              ),
              "商品库SKU": st.column_config.TextColumn(
                  "商品库 SKU", required=True
              ),
          },
          key="sku_mapping_editor",
      )

      c_btn1, c_btn2 = st.columns(2)
      if c_btn1.button("保存规则", type="primary", use_container_width=True):
        new_map = {}
        for _, r in edited_df.iterrows():
          src = str(r.get("面单SKU", "")).strip()
          tgt = str(r.get("商品库SKU", "")).strip()
          if src and tgt and src not in ("nan", "None"):
            new_map[src] = tgt
        st.session_state["sku_mapping"] = new_map
        save_sku_mapping(new_map)
        st.rerun()

      if c_btn2.button("清空全部", use_container_width=True):
        st.session_state["sku_mapping"] = {}
        save_sku_mapping({})
        st.rerun()

  if custom_uploaded is not None:
    active_df = parse_raw_table_bytes(custom_uploaded.getvalue())

  st.write("")

  # 3. 主文件上传
  uploaded_file = st.file_uploader(
      "请上传备货单/发货单 Excel 或 CSV 文件",
      type=["xlsx", "xls", "csv"],
      help="自动识别商品行并从【物流信息】中定位【关联备货单号】",
  )

  if uploaded_file is not None:
    try:
      with st.spinner("正在解析物流信息并匹配商品库数据..."):
        goods_df, related_order_code = parse_raw_order_file(uploaded_file)
        df_detail, df_summary, kpi, missing_skus = process_shipment_data(
            goods_df, related_order_code, active_df, sku_mapping=current_mapping
        )

      # 顶部 KPI 看板
      st.markdown(
          f"""
            <div class="metric-container">
                <div class="metric-card">
                    <div class="metric-title">货件编号 (关联备货单)</div>
                    <div class="metric-num" style="font-size:1.05rem; word-break:break-all;">{related_order_code or '—'}</div>
                </div>
                <div class="metric-card">
                    <div class="metric-title">总装箱量</div>
                    <div class="metric-num">{kpi['total_box']:,}<span class="metric-unit">箱</span></div>
                </div>
                <div class="metric-card">
                    <div class="metric-title">总备货量</div>
                    <div class="metric-num">{kpi['total_sets']:,}<span class="metric-unit">套</span></div>
                </div>
                <div class="metric-card">
                    <div class="metric-title">总 PCS (单件实物)</div>
                    <div class="metric-num">{kpi['total_pcs']:,}<span class="metric-unit">件</span></div>
                </div>
                <div class="metric-card">
                    <div class="metric-title">实重 / 总体积</div>
                    <div class="metric-num" style="font-size:1.1rem;">{kpi['total_weight']}<span class="metric-unit">kg</span> / {kpi['total_volume']}<span class="metric-unit">m³</span></div>
                </div>
                <div class="metric-card">
                    <div class="metric-title">货件总货值</div>
                    <div class="metric-num">{kpi['total_val']}</div>
                </div>
            </div>
            """,
          unsafe_allow_html=True,
      )

      if missing_skus:
        st.warning(
            f"⚠️ 提示：共有 {len(missing_skus)} 个 SKU 未在商品库中匹配到规格信息："
            f" {', '.join(missing_skus[:8])}{'...' if len(missing_skus) > 8 else ''}"
        )

      # 导出按钮
      ts = datetime.now().strftime("%Y%m%d_%H%M%S")
      filename_tag = f"_{related_order_code}" if related_order_code else ""
      out_filename = f"发货单处理结果{filename_tag}_{ts}.xlsx"
      excel_bytes = export_and_beautify(df_detail, df_summary)

      st.download_button(
          label="⬇️ 导出标准发货单与汇总表 (.xlsx)",
          data=excel_bytes,
          file_name=out_filename,
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          type="primary",
          use_container_width=True,
      )

      # 严格对应模板的选项卡展示
      tab1, tab2 = st.tabs(["📝 详细数据", "📊 汇总结果"])
      with tab1:
        st.dataframe(df_detail, use_container_width=True)
      with tab2:
        st.dataframe(df_summary, use_container_width=True)

    except Exception as e:
      st.error(f"❌ 数据处理失败：{str(e)}")


if __name__ == "__main__":
  main()
