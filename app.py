import streamlit as st
import pandas as pd
import os
import json
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from datetime import datetime

# ---------- 页面配置 ----------
st.set_page_config(page_title="合并并校验同料不同价", layout="wide")
st.title("📊 合并并校验同料不同价（单价冲突版）")

# ---------- 初始化 session_state ----------
if "mapping_columns" not in st.session_state:
    st.session_state.mapping_columns = [
        {"export_name": "客户料号", "col_a": "", "col_b": ""},
        {"export_name": "HQ料号", "col_a": "", "col_b": ""},
        {"export_name": "Category", "col_a": "", "col_b": ""},
        {"export_name": "Usag", "col_a": "", "col_b": ""},
        {"export_name": "单价", "col_a": "", "col_b": ""},
        {"export_name": "Extend", "col_a": "", "col_b": ""},
    ]
if "df_a" not in st.session_state:
    st.session_state.df_a = None
    st.session_state.sheet_a = None
    st.session_state.file_a_name = None
if "df_b" not in st.session_state:
    st.session_state.df_b = None
    st.session_state.sheet_b = None
    st.session_state.file_b_name = None
if "output_dir" not in st.session_state:
    st.session_state.output_dir = os.getcwd()

CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "mapping_config.json")

# ---------- 辅助函数 ----------
def select_sheet(file_path, key):
    xl = pd.ExcelFile(file_path)
    sheets = xl.sheet_names
    if len(sheets) == 1:
        return sheets[0]
    else:
        return st.selectbox(f"选择 {key} 的工作表", sheets, key=f"sheet_{key}")

def load_config():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            config = json.load(f)
            if "mapping_columns" in config:
                st.session_state.mapping_columns = config["mapping_columns"]
            if "output_dir" in config and os.path.exists(config["output_dir"]):
                st.session_state.output_dir = config["output_dir"]
        st.success("已加载历史配置")

def save_config():
    config = {
        "mapping_columns": st.session_state.mapping_columns,
        "output_dir": st.session_state.output_dir
    }
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    st.success("配置已保存")

def standardize(df, table):
    out = pd.DataFrame(index=df.index)
    for cfg in st.session_state.mapping_columns:
        col = cfg["export_name"]
        src = cfg["col_a"] if table == "a" else cfg["col_b"]
        if src and src in df.columns:
            out[col] = df[src]
        else:
            out[col] = None
    if table == "a":
        fname = os.path.splitext(os.path.basename(st.session_state.file_a_name))[0]
        sheet = st.session_state.sheet_a
        out["数据来源"] = f"{fname}({sheet})"
    else:
        fname = os.path.splitext(os.path.basename(st.session_state.file_b_name))[0]
        sheet = st.session_state.sheet_b
        out["数据来源"] = f"{fname}({sheet})"
    return out

def process():
    if st.session_state.df_a is None or st.session_state.df_b is None:
        st.error("请先上传表A和表B")
        return
    if not st.session_state.mapping_columns:
        st.error("请先配置字段映射")
        return

    has_key = any(c["export_name"] == "客户料号" and (c["col_a"] or c["col_b"]) for c in st.session_state.mapping_columns)
    has_price = any(c["export_name"] == "单价" and (c["col_a"] or c["col_b"]) for c in st.session_state.mapping_columns)
    if not has_key:
        st.error("客户料号必须映射有效列")
        return
    if not has_price:
        st.error("单价必须映射有效列")
        return

    with st.spinner("正在处理数据..."):
        df_a_std = standardize(st.session_state.df_a, "a")
        df_b_std = standardize(st.session_state.df_b, "b")
        df_merged = pd.concat([df_a_std, df_b_std], ignore_index=True)

        df_merged["单价"] = pd.to_numeric(df_merged["单价"], errors="coerce")

        a_price = df_a_std.dropna(subset=["单价"]).groupby("客户料号")["单价"].first().to_dict()
        b_price = df_b_std.dropna(subset=["单价"]).groupby("客户料号")["单价"].first().to_dict()

        conflict_pns = set()
        for pn in a_price:
            if pn in b_price and abs(a_price[pn] - b_price[pn]) > 0.0001:
                conflict_pns.add(pn)

        df_merged["是否单价冲突"] = df_merged["客户料号"].isin(conflict_pns)

        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_filename = f"合并校验结果_{now}_V1.0.xlsx"
        out_path = os.path.join(st.session_state.output_dir, out_filename)

        df_merged.to_excel(out_path, index=False, engine="openpyxl")

        wb = load_workbook(out_path)
        ws = wb.active
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        headers = [cell.value for cell in ws[1]]
        conflict_col_idx = headers.index("是否单价冲突") + 1
        for row in range(2, ws.max_row + 1):
            if ws.cell(row=row, column=conflict_col_idx).value is True:
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=row, column=col).fill = yellow_fill
        wb.save(out_path)

        with open(out_path, "rb") as f:
            st.download_button(
                label="📥 下载处理结果",
                data=f,
                file_name=out_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        st.success(f"处理完成！文件已保存至：{out_path}")
        st.info("黄色行表示同一客户料号在A、B表中单价不同")

# ---------- 侧边栏：文件上传与配置 ----------
with st.sidebar:
    st.header("1️⃣ 上传文件")
    uploaded_a = st.file_uploader("表A (Excel)", type=["xlsx", "xls"], key="a")
    uploaded_b = st.file_uploader("表B (Excel)", type=["xlsx", "xls"], key="b")

    if uploaded_a:
        st.session_state.file_a_name = uploaded_a.name
        sheets_a = pd.ExcelFile(uploaded_a).sheet_names
        if len(sheets_a) == 1:
            st.session_state.sheet_a = sheets_a[0]
        else:
            st.session_state.sheet_a = st.selectbox("表A 工作表", sheets_a, key="sheet_a_sel")
        st.session_state.df_a = pd.read_excel(uploaded_a, sheet_name=st.session_state.sheet_a)
        st.success(f"表A加载成功，{len(st.session_state.df_a)}行")

    if uploaded_b:
        st.session_state.file_b_name = uploaded_b.name
        sheets_b = pd.ExcelFile(uploaded_b).sheet_names
        if len(sheets_b) == 1:
            st.session_state.sheet_b = sheets_b[0]
        else:
            st.session_state.sheet_b = st.selectbox("表B 工作表", sheets_b, key="sheet_b_sel")
        st.session_state.df_b = pd.read_excel(uploaded_b, sheet_name=st.session_state.sheet_b)
        st.success(f"表B加载成功，{len(st.session_state.df_b)}行")

    st.header("2️⃣ 字段映射配置")
    st.caption("导出列名：最终输出的列名称。表A列/表B列：选择实际列名。")

    # 动态生成可编辑表格
    if st.session_state.df_a is not None and st.session_state.df_b is not None:
        cols_a = [""] + list(st.session_state.df_a.columns)
        cols_b = [""] + list(st.session_state.df_b.columns)
        mapping_data = []
        for m in st.session_state.mapping_columns:
            mapping_data.append({
                "导出列名": m["export_name"],
                "表A列": m["col_a"],
                "表B列": m["col_b"]
            })
        edited = st.data_editor(
            mapping_data,
            column_config={
                "导出列名": st.column_config.TextColumn("导出列名", required=True),
                "表A列": st.column_config.SelectboxColumn("表A列", options=cols_a),
                "表B列": st.column_config.SelectboxColumn("表B列", options=cols_b),
            },
            use_container_width=True,
            num_rows="dynamic",
            key="mapping_editor"
        )
        # 更新 mapping_columns
        new_mapping = []
        for row in edited:
            if row["导出列名"].strip():
                new_mapping.append({
                    "export_name": row["导出列名"].strip(),
                    "col_a": row["表A列"] if row["表A列"] != "" else None,
                    "col_b": row["表B列"] if row["表B列"] != "" else None
                })
        st.session_state.mapping_columns = new_mapping

    col1, col2 = st.columns(2)
    with col1:
        if st.button("💾 保存配置"):
            save_config()
    with col2:
        if st.button("📂 加载配置"):
            load_config()

    st.header("3️⃣ 输出目录")
    output_dir = st.text_input("输出文件夹路径", st.session_state.output_dir)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
        st.session_state.output_dir = output_dir

# ---------- 主区域：处理按钮 ----------
if st.button("🚀 开始合并并校验单价", type="primary"):
    process()