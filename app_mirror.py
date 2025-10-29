# =====================================
# Streamlit Web App: 模拟Project：人事用合同记录表自动审核（四输出表版 + 漏填检查 + 驻店客户版）
# =====================================

import streamlit as st
import pandas as pd
import time
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
from io import BytesIO

# =====================================
# 🏁 应用标题与说明
# =====================================
st.title("📊 模拟实际运用环境Project：人事用合同记录表自动审核系统（四Sheet + 漏填检查 + 驻店客户版）")

# =====================================
# 📂 上传文件区：要求上传 5 个 xlsx 文件
# =====================================
uploaded_files = st.file_uploader(
    "请上传以下文件：记录表、放款明细、字段、二次明细、重卡数据",
    type="xlsx",
    accept_multiple_files=True
)

if not uploaded_files or len(uploaded_files) < 5:
    st.warning("⚠️ 请上传所有 5 个文件后继续")
    st.stop()
else:
    st.success("✅ 文件上传完成")

# =====================================
# 🧰 工具函数区（文件定位、列名模糊匹配、日期/数值处理）
# =====================================

# 按关键字查找文件（文件名包含关键字）
def find_file(files_list, keyword):
    for f in files_list:
        if keyword in f.name:
            return f
    raise FileNotFoundError(f"❌ 未找到包含关键词「{keyword}」的文件")

# 统一列名格式（去空格、转小写）
def normalize_colname(c): return str(c).strip().lower()

# 按关键字匹配列名（支持 exact 精确匹配与模糊匹配）
def find_col(df, keyword, exact=False):
    key = keyword.strip().lower()
    for col in df.columns:
        cname = normalize_colname(col)
        if (exact and cname == key) or (not exact and key in cname):
            return col
    return None

# 查找sheet（sheet名包含关键字即可）
def find_sheet(xls, keyword):
    for s in xls.sheet_names:
        if keyword in s:
            return s
    raise ValueError(f"❌ 未找到包含关键词「{keyword}」的sheet")

# 统一数值解析（去逗号、转float、处理百分号）
def normalize_num(val):
    if pd.isna(val): return None
    s = str(val).replace(",", "").strip()
    if s in ["", "-", "nan"]: return None
    try:
        if "%" in s: return float(s.replace("%", "")) / 100
        return float(s)
    except ValueError:
        return s

# 日期匹配（年/月/日完全一致）
def same_date_ymd(a, b):
    try:
        da = pd.to_datetime(a, errors='coerce')
        db = pd.to_datetime(b, errors='coerce')
        if pd.isna(da) or pd.isna(db): return False
        return (da.year, da.month, da.day) == (db.year, db.month, db.day)
    except Exception:
        return False

# =====================================
# 🔍 核心对比函数：逐行比对并标红错误单元格
# =====================================
def compare_fields_and_mark(row_idx, row, main_df, main_kw, ref_df, ref_kw,
                            ref_contract_col, ws, red_fill, exact=False, skip_counter=None):
    errors = 0
    main_col = find_col(main_df, main_kw, exact=exact)
    ref_col = find_col(ref_df, ref_kw, exact=exact)
    if not main_col or not ref_col or not ref_contract_col:
        return 0

    # 获取当前合同号
    contract_no = str(row.get(contract_col_main)).strip()
    if pd.isna(contract_no) or contract_no in ["", "nan"]:
        return 0

    # 在参考表中查找同合同号数据
    ref_rows = ref_df[ref_df[ref_contract_col].astype(str).str.strip() == contract_no]
    if ref_rows.empty:
        return 0

    ref_val = ref_rows.iloc[0][ref_col]
    main_val = row.get(main_col)

    # 如果字段表中城市经理为空 → 跳过检查并计数
    if main_kw == "城市经理":
        if pd.isna(ref_val) or str(ref_val).strip() in ["", "-", "nan", "none", "null"]:
            if skip_counter is not None: skip_counter[0] += 1
            return 0

    # 同为NaN则认为一致
    if pd.isna(main_val) and pd.isna(ref_val):
        return 0

    # 日期字段专用对比逻辑
    if any(k in main_kw for k in ["日期", "时间"]) or any(k in ref_kw for k in ["日期", "时间"]):
        if not same_date_ymd(main_val, ref_val):
            errors = 1
    else:
        # 数值/文本比对
        main_num = normalize_num(main_val)
        ref_num = normalize_num(ref_val)
        if isinstance(main_num, (int, float)) and isinstance(ref_num, (int, float)):
            diff = abs(main_num - ref_num)
            # 保证金比例容差 ±0.005
            if main_kw == "保证金比例" and ref_kw == "保证金比例_2":
                if diff > 0.005: errors = 1
            else:
                if diff > 1e-6: errors = 1
        else:
            if str(main_num).strip().lower().replace(".0", "") != str(ref_num).strip().lower().replace(".0", ""):
                errors = 1

    # 若存在错误则标红
    if errors:
        ws.cell(row_idx + 3, list(main_df.columns).index(main_col) + 1).fill = red_fill
    return errors

# =====================================
# 🧮 单sheet检查函数
# =====================================
def check_one_sheet(sheet_keyword):
    start_time = time.time()
    xls_main = pd.ExcelFile(main_file)

    # 查找目标sheet
    try:
        target_sheet = find_sheet(xls_main, sheet_keyword)
    except ValueError:
        st.warning(f"⚠️ 未找到包含「{sheet_keyword}」的sheet，跳过。")
        return 0, None, 0, set()

    # 读取目标sheet（第二行为表头）
    main_df = pd.read_excel(xls_main, sheet_name=target_sheet, header=1)

    # 创建临时输出文件并保留空行
    output_path = f"记录表_{sheet_keyword}_审核标注版.xlsx"
    empty_row = pd.DataFrame([[""] * len(main_df.columns)], columns=main_df.columns)
    pd.concat([empty_row, main_df], ignore_index=True).to_excel(output_path, index=False)

    # 打开Excel用于写入标注
    wb = load_workbook(output_path)
    ws = wb.active
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # 查找合同号列
    global contract_col_main
    contract_col_main = find_col(main_df, "合同")
    if not contract_col_main:
        st.error(f"❌ 在「{sheet_keyword}」中未找到合同列。")
        return 0, None, 0, set()

    total_errors = 0
    skip_city_manager = [0]
    contracts_seen = set()

    # 添加Streamlit进度条
    progress = st.progress(0)
    status = st.empty()

    # === 遍历每一行进行字段比对 ===
    for idx, row in main_df.iterrows():
        if pd.isna(row.get(contract_col_main)):
            continue
        contracts_seen.add(str(row.get(contract_col_main)).strip())

        # 放款明细对照
        for main_kw, ref_kw in mapping_fk.items():
            total_errors += compare_fields_and_mark(idx, row, main_df, main_kw, fk_df, ref_kw, contract_col_fk, ws, red_fill)
        # 字段表对照（含城市经理跳过计数）
        for main_kw, ref_kw in mapping_zd.items():
            exact = (main_kw == "城市经理")
            total_errors += compare_fields_and_mark(idx, row, main_df, main_kw, zd_df, ref_kw, contract_col_zd, ws, red_fill, exact=exact, skip_counter=skip_city_manager)
        # 二次明细对照
        for main_kw, ref_kw in mapping_ec.items():
            total_errors += compare_fields_and_mark(idx, row, main_df, main_kw, ec_df, ref_kw, contract_col_ec, ws, red_fill)
        # 重卡数据对照
        for main_kw, ref_kw in mapping_zk.items():
            total_errors += compare_fields_and_mark(idx, row, main_df, main_kw, zk_df, ref_kw, contract_col_zk, ws, red_fill)

        progress.progress((idx + 1) / len(main_df))
        if (idx + 1) % 10 == 0:
            status.text(f"检查「{sheet_keyword}」... {idx+1}/{len(main_df)}")

    # 若有错误则对应合同号黄标
    cidx = list(main_df.columns).index(contract_col_main) + 1
    for r in range(len(main_df)):
        if any(ws.cell(r + 3, c).fill == red_fill for c in range(1, len(main_df.columns) + 1)):
            ws.cell(r + 3, cidx).fill = yellow_fill

    # 导出检查结果
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    st.download_button(
        label=f"📥 下载 {sheet_keyword} 审核标注版",
        data=output,
        file_name=f"记录表_{sheet_keyword}_审核标注版.xlsx"
    )

    elapsed = time.time() - start_time
    st.success(f"✅ {sheet_keyword} 检查完成，共 {total_errors} 处错误，用时 {elapsed:.2f} 秒。")
    return total_errors, elapsed, skip_city_manager[0], contracts_seen

# =====================================
# 📖 文件读取：按关键字识别五份文件
# =====================================
main_file = find_file(uploaded_files, "记录表")
fk_file = find_file(uploaded_files, "放款明细")
zd_file = find_file(uploaded_files, "字段")
ec_file = find_file(uploaded_files, "二次明细")
zk_file = find_file(uploaded_files, "重卡数据")

# 各文件sheet读取（模糊匹配sheet名）
fk_df = pd.read_excel(pd.ExcelFile(fk_file), sheet_name=find_sheet(pd.ExcelFile(fk_file), "本司"))
zd_df = pd.read_excel(pd.ExcelFile(zd_file), sheet_name=find_sheet(pd.ExcelFile(zd_file), "重卡"))
ec_df = pd.read_excel(ec_file)
zk_df = pd.read_excel(zk_file)

# 合同列定位
contract_col_fk = find_col(fk_df, "合同")
contract_col_zd = find_col(zd_df, "合同")
contract_col_ec = find_col(ec_df, "合同")
contract_col_zk = find_col(zk_df, "合同")

# 对照字段映射表
mapping_fk = {"授信方": "授信", "租赁本金": "本金", "租赁期限月": "租赁期限月", "客户经理": "客户经理", "起租收益率": "收益率", "主车台数": "主车台数", "挂车台数": "挂车台数"}
mapping_zd = {"保证金比例": "保证金比例_2", "项目提报人": "提报", "起租时间": "起租日_商", "租赁期限月": "总期数_商_资产", "所属省区": "区域", "城市经理": "城市经理"}
mapping_ec = {"二次时间": "出本流程时间"}
mapping_zk = {"授信方": "授信方"}

# =====================================
# 🧾 多sheet循环 + 驻店客户表
# =====================================
sheet_keywords = ["二次", "部分担保", "随州", "驻店客户"]
total_all = elapsed_all = skip_total = 0
contracts_seen_all_sheets = set()

# 循环处理四张sheet
for kw in sheet_keywords:
    count, used, skipped, seen = check_one_sheet(kw)
    total_all += count
    elapsed_all += used or 0
    skip_total += skipped
    contracts_seen_all_sheets.update(seen)

st.success(f"🎯 全部审核完成，共 {total_all} 处错误，总耗时 {elapsed_all:.2f} 秒。")

# =====================================
# 🕵️ 漏填检查：跳过“是否车管家=是”与“提成类型=联合租赁/驻店”
# =====================================
field_contracts = zd_df[contract_col_zd].dropna().astype(str).str.strip()
col_car_manager = find_col(zd_df, "是否车管家", exact=True)
col_bonus_type = find_col(zd_df, "提成类型", exact=True)

missing_contracts_mask = (~field_contracts.isin(contracts_seen_all_sheets))

# 跳过“车管家=是”
if col_car_manager:
    missing_contracts_mask &= ~(zd_df[col_car_manager].astype(str).str.strip().str.lower() == "是")
# 跳过“联合租赁/驻店”
if col_bonus_type:
    missing_contracts_mask &= ~(
        zd_df[col_bonus_type].astype(str).str.strip().isin(["联合租赁", "驻店"])
    )

# 标记漏填
zd_df_missing = zd_df.copy()
zd_df_missing["漏填检查"] = ""
zd_df_missing.loc[missing_contracts_mask, "漏填检查"] = "❗ 漏填"
漏填合同数 = zd_df_missing["漏填检查"].eq("❗ 漏填").sum()
st.warning(f"⚠️ 共发现 {漏填合同数} 个合同在记录表中未出现（已排除车管家、联合租赁、驻店）")

# =====================================
# 📤 导出字段表（含漏填标注 + 仅漏填版）
# =====================================
yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

# 全字段表（含漏填标注）
wb = Workbook()
ws = wb.active
for c_idx, c in enumerate(zd_df_missing.columns, 1): ws.cell(1, c_idx, c)
for r_idx, row in enumerate(zd_df_missing.itertuples(index=False), 2):
    for c_idx, v in enumerate(row, 1):
        ws.cell(r_idx, c_idx, v)
        if zd_df_missing.columns[c_idx-1] == "漏填检查" and v == "❗ 漏填":
            ws.cell(r_idx, c_idx).fill = yellow_fill

output_all = BytesIO()
wb.save(output_all)
output_all.seek(0)
st.download_button("📥 下载字段表漏填标注版", output_all, "字段表_漏填标注版.xlsx")

# 仅漏填合同
zd_df_only_missing = zd_df_missing[zd_df_missing["漏填检查"] == "❗ 漏填"].copy()
if not zd_df_only_missing.empty:
    wb2 = Workbook()
    ws2 = wb2.active
    for c_idx, c in enumerate(zd_df_only_missing.columns, 1): ws2.cell(1, c_idx, c)
    for r_idx, row in enumerate(zd_df_only_missing.itertuples(index=False), 2):
        for c_idx, v in enumerate(row, 1):
            ws2.cell(r_idx, c_idx, v)
            if zd_df_only_missing.columns[c_idx-1] == "漏填检查" and v == "❗ 漏填":
                ws2.cell(r_idx, c_idx).fill = yellow_fill
    out2 = BytesIO()
    wb2.save(out2)
    out2.seek(0)
    st.download_button("📥 下载仅漏填字段表", out2, "字段表_仅漏填.xlsx")

# =====================================
# ✅ 结束提示
# =====================================
st.success("✅ 所有检查、标注与导出完成！")
