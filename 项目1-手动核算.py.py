import os
import json
import pandas as pd
import numpy as np
from pathlib import Path
import re
from datetime import datetime, date

# ==================== 敏感信息配置区（外发时替换为占位符）====================
# 路径（本地真实值，外发替换为占位符）
PATHS = {
    'input_folder':     r'F:\...\BI待处理数据-项目1-手动核算-py',
    'field_map_file':   r'F:\...\field_map.json',
    'match_file':       r'F:\...\活动名称对应一二级分类.xlsx',
    'check_file2':      r'F:\...\收入成本校验表--python校验.xlsx',
    'second_arch_file': r'F:\...\分公司目录-新二级部门标签.xlsx',
}

# 业务规则（本地真实值，外发替换为占位符）
RULES = {
    'arch_replace_map': {
        '【旧分公司A】': '【新分公司A】',
        '【旧分公司B】': '【新分公司B】',
        '【旧分公司C】': '【新分公司C】',
        '【旧总部】':   '【新总部】',
    },
    'special_manager_id':  0,
    'special_second_arch': '【特殊分公司名称】',
    'edu_first_category':  '【教育一级分类】',
    'edu_third_arch':      '【教育三级部门】',
    'exempt_category_for_agent':  '【豁免分类编码】',
    'special_aggregate_category': '【特殊汇总分类编码】',
    'invoice_exempt_category':    '【开票豁免分类编码】',
    'reminder_categories': ['【提醒分类1】', '【提醒分类2】'],
    'allowed_arch_exceptions': ['【例外部门1】', '【例外部门2】'],
    'no_arch_label': '无归属部门',
}

# 脱敏后的字段名常量（代码全程使用这些，不含任何原始字段名）
COL_MERCHANT_NO   = '商户编号'
COL_MERCHANT_NAME = '商户名'
COL_EMP_ID        = '员工编号'
COL_EMP_NAME      = '员工姓名'
COL_AGENT_ID      = '渠道编号'
COL_AGENT_NAME    = '渠道名称'
COL_ARCH1         = '组织层级1'
COL_ARCH2         = '组织层级2'
COL_ARCH3         = '组织层级3'
COL_ARCH4         = '组织层级4'
COL_OPS_ID        = '运营经理编号'
COL_OPS_NAME      = '运营经理姓名'
COL_INDUSTRY      = '行业分类'
COL_SALES_ID      = '销售员编号'
COL_SALES_NAME    = '销售员姓名'
COL_TXN_COUNT     = '有效交易笔数'
COL_TXN_AMT       = '有效交易金额'
COL_REWARD        = '补贴金额'
COL_REWARD_BASE   = '补贴/基数'
COL_RATIO         = '分润比例'
COL_SETTLE        = '支出金额'
COL_PROFIT        = '净收益'
COL_CHANNEL       = '支付渠道'
COL_BIZ_CAT1      = '业务大类'
COL_BIZ_CAT2      = '业务小类'
COL_BILL_PERIOD   = '账期'
COL_BI_PERIOD     = '报表期'
COL_TABLE_NAME    = '数据表'
COL_INVOICE       = '发票抬头'
COL_PROJECT       = '项目名'
COL_PROJECT_TYPE  = '项目类型'
COL_PROJECT_CAT   = '项目类别'
COL_SETTLE_METHOD = '结算方式'
COL_AGENT_SETTLE  = '渠道结算标记'
COL_EMP_SETTLE    = '员工结算标记'
COL_EMP_D0        = '员工即时分润'
COL_AGENT_D0      = '渠道即时分润'

# 脱敏后的标准表头（代码全程使用）
EXPECTED_HEADERS = [
    COL_MERCHANT_NO, COL_MERCHANT_NAME,
    COL_EMP_ID, COL_EMP_NAME,
    COL_AGENT_ID, COL_AGENT_NAME,
    COL_ARCH1, COL_ARCH2, COL_ARCH3, COL_ARCH4,
    COL_OPS_ID, COL_OPS_NAME, COL_INDUSTRY,
    COL_SALES_ID, COL_SALES_NAME,
    COL_TXN_COUNT, COL_TXN_AMT,
    COL_REWARD, COL_REWARD_BASE, COL_RATIO,
    COL_SETTLE, COL_PROFIT,
    COL_CHANNEL, COL_BIZ_CAT1, COL_BIZ_CAT2,
    COL_BILL_PERIOD, COL_BI_PERIOD, COL_TABLE_NAME,
    COL_INVOICE, COL_PROJECT,
    COL_PROJECT_TYPE, COL_PROJECT_CAT, COL_SETTLE_METHOD,
    COL_AGENT_SETTLE, COL_EMP_SETTLE,
    COL_EMP_D0, COL_AGENT_D0,
]
# ==========================================================================

# 从外部加载原始→脱敏映射（代码里看不到原始字段名）
with open(PATHS['field_map_file'], 'r', encoding='utf-8') as f:
    FIELD_MAP = json.load(f)

folder_path         = PATHS['input_folder']
MATCH_FILE          = PATHS['match_file']
CHECK_FILE2         = PATHS['check_file2']
SECOND_ARCH_FILE    = PATHS['second_arch_file']

MULTIPLY_COLS        = [COL_TXN_AMT, COL_REWARD, COL_REWARD_BASE, COL_SETTLE, COL_PROFIT]
COMPANY_LEVEL_COLS   = [COL_ARCH1, COL_ARCH2, COL_ARCH3, COL_ARCH4]
SPECIAL_MANAGER_ID   = RULES['special_manager_id']
SPECIAL_SECOND_ARCH  = RULES['special_second_arch']
REMINDER_CATEGORIES  = RULES['reminder_categories']
ARCH_REPLACE_MAP     = RULES['arch_replace_map']
NO_ARCH_LABEL        = RULES['no_arch_label']
ALLOWED_ARCH_EXCEPTIONS = RULES['allowed_arch_exceptions']
EXEMPT_CATEGORY_FOR_AGENT  = RULES['exempt_category_for_agent']
SPECIAL_AGGREGATE_CATEGORY = RULES['special_aggregate_category']
INVOICE_EXEMPT_CATEGORY    = RULES['invoice_exempt_category']
EDU_FIRST_CATEGORY         = RULES['edu_first_category']
EDU_THIRD_ARCH             = RULES['edu_third_arch']

def get_last_month_str():
    today = date.today()
    first_of_month = today.replace(day=1)
    last_month = first_of_month - pd.DateOffset(months=1)
    return last_month.strftime("%Y-%m")

LAST_MONTH = get_last_month_str()
LAST_MONTH_DATE = datetime.strptime(LAST_MONTH, "%Y-%m").replace(day=1)
print(f"报表期基准（上个月）：{LAST_MONTH}")
print("项目4 将自动将 报表期 纠正为此值。\n")

# ==================== 列名重命名（原始 → 脱敏） ====================
def rename_columns(df):
    """按外部映射将原始字段名替换为脱敏字段名"""
    return df.rename(columns=FIELD_MAP)

# ==================== 加载有效组织层级2标签 ====================
def load_valid_second_arch_labels():
    file_path = SECOND_ARCH_FILE
    if not os.path.exists(file_path):
        print(f"警告：未找到组织层级目录文件：{file_path}，跳过校验")
        return set()
    try:
        df = pd.read_excel(file_path, dtype=str)
        df.columns = [str(c).strip() for c in df.columns]
        target_col = None
        for col in df.columns:
            if '标签' in col or '部门' in col:
                target_col = col
                break
        if target_col is None:
            target_col = df.columns[0]
        labels = set(df[target_col].dropna().astype(str).str.strip())
        print(f"已加载组织层级目录，共 {len(labels)} 个有效标签")
        return labels
    except Exception as e:
        print(f"加载组织层级目录失败：{e}，跳过校验")
        return set()

VALID_SECOND_ARCH_LABELS = load_valid_second_arch_labels()

def validate_second_arch(df, valid_labels):
    errors = []
    if COL_ARCH2 not in df.columns or COL_ARCH1 not in df.columns:
        return df, pd.DataFrame()
    df[COL_ARCH2] = df[COL_ARCH2].replace(ARCH_REPLACE_MAP)
    for idx, row in df.iterrows():
        second = row.get(COL_ARCH2)
        first = row.get(COL_ARCH1)
        reasons = []
        if pd.notna(second) and str(second).strip().endswith('分公司'):
            if pd.isna(first) or str(first).strip() != '分公司':
                reasons.append(f"二级部门为'{second}'，但一级部门不是'分公司'（实际：{first}）")
        if pd.notna(second) and str(second).strip() not in valid_labels:
            reasons.append(f"二级部门'{second}'不在组织层级目录中")
        if reasons:
            errors.append({
                '行索引': idx,
                '组织层级1': first,
                '组织层级2': second,
                '错误原因': '；'.join(reasons)
            })
    errors_df = pd.DataFrame(errors)
    if not errors_df.empty:
        print(f"   [警告] 发现 {len(errors_df)} 行组织层级校验错误。")
    return df, errors_df

# ==================== 商户编号清洗函数 ====================
def clean_merchant_no(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s == "":
        return ""
    digits = re.sub(r'\D', '', s)
    if not digits:
        return ""
    return digits[:15]

# ==================== 核心函数 ====================
def match_header_case_insensitive(df, expected_headers):
    actual_lower = [col.lower() for col in df.columns]
    expected_lower = [h.lower() for h in expected_headers]
    if actual_lower == expected_lower:
        df.columns = expected_headers
        return True
    return False

def convert_id_columns(df):
    id_cols = [col for col in df.columns if "编号" in col]
    for col in id_cols:
        if df[col].isnull().all():
            continue
        numeric_series = pd.to_numeric(df[col], errors='coerce')
        invalid_mask = df[col].notna() & numeric_series.isna()
        if invalid_mask.any():
            invalid_rows = df.index[invalid_mask].tolist()
            display = invalid_rows[:5]
            rows_str = f"{display} ... (共 {len(invalid_rows)} 行)" if len(invalid_rows) > 5 else str(invalid_rows)
            raise ValueError(f"列 '{col}' 中存在非数值内容，行索引：{rows_str}")
        df[col] = numeric_series
    return df

def format_date_columns(df):
    """账期转为日期，报表期强制纠正为上月"""
    if COL_BILL_PERIOD in df.columns:
        df[COL_BILL_PERIOD] = pd.to_datetime(df[COL_BILL_PERIOD], errors='coerce')
    if COL_BI_PERIOD in df.columns:
        df[COL_BI_PERIOD] = LAST_MONTH_DATE
        print(f"   [提示] 已将所有行的 报表期 统一纠正为 {LAST_MONTH}")
    return df

def check_uniqueness(df, file_name):
    errors = []
    if COL_EMP_ID in df.columns and COL_EMP_NAME in df.columns:
        non_null = df[[COL_EMP_ID, COL_EMP_NAME]].dropna()
        if not non_null.empty:
            grouped = non_null.groupby(COL_EMP_ID)[COL_EMP_NAME].nunique()
            dup_ids = grouped[grouped > 1].index.tolist()
            if dup_ids:
                conflict_rows = non_null[non_null[COL_EMP_ID].isin(dup_ids)].index.tolist()
                display = conflict_rows[:5]
                rows_str = f"{display} ... (共 {len(conflict_rows)} 行)" if len(conflict_rows) > 5 else str(conflict_rows)
                errors.append(f"员工编号对应多个员工姓名，冲突行索引：{rows_str}")
    if COL_AGENT_ID in df.columns and COL_AGENT_NAME in df.columns:
        non_null = df[[COL_AGENT_ID, COL_AGENT_NAME]].dropna()
        if not non_null.empty:
            grouped = non_null.groupby(COL_AGENT_ID)[COL_AGENT_NAME].nunique()
            dup_ids = grouped[grouped > 1].index.tolist()
            if dup_ids:
                conflict_rows = non_null[non_null[COL_AGENT_ID].isin(dup_ids)].index.tolist()
                display = conflict_rows[:5]
                rows_str = f"{display} ... (共 {len(conflict_rows)} 行)" if len(conflict_rows) > 5 else str(conflict_rows)
                errors.append(f"渠道编号对应多个渠道名称，冲突行索引：{rows_str}")
    if COL_AGENT_ID in df.columns and COL_EMP_ID in df.columns and COL_BIZ_CAT2 in df.columns:
        non_exempt = df[df[COL_BIZ_CAT2] != EXEMPT_CATEGORY_FOR_AGENT]
        non_null = non_exempt[[COL_AGENT_ID, COL_EMP_ID]].dropna()
        if not non_null.empty:
            grouped = non_null.groupby(COL_AGENT_ID)[COL_EMP_ID].nunique()
            dup_ids = grouped[grouped > 1].index.tolist()
            if dup_ids:
                conflict_rows = non_null[non_null[COL_AGENT_ID].isin(dup_ids)].index.tolist()
                display = conflict_rows[:5]
                rows_str = f"{display} ... (共 {len(conflict_rows)} 行)" if len(conflict_rows) > 5 else str(conflict_rows)
                errors.append(f"渠道编号对应多个员工编号（豁免分类除外），冲突行索引：{rows_str}")
    if errors:
        raise ValueError("\n".join(errors))

def check_admin_arch_when_have_agent(df, file_name):
    key_cols = [COL_EMP_ID, COL_EMP_NAME, COL_AGENT_ID, COL_AGENT_NAME]
    def is_valid_series(series):
        return series.notna() & (series != 0) & (series != "0") & (series != "") & (series != pd.NA)
    has_agent = pd.Series(False, index=df.index)
    for col in key_cols:
        if col in df.columns:
            has_agent = has_agent | is_valid_series(df[col])
    empty_arch_mask = pd.Series(False, index=df.index)
    for col in COMPANY_LEVEL_COLS:
        if col in df.columns:
            is_empty = df[col].isna() | (df[col].astype(str).str.strip() == '')
            empty_arch_mask = empty_arch_mask | is_empty
    invalid_mask = has_agent & empty_arch_mask
    if invalid_mask.any():
        invalid_rows = df.index[invalid_mask].tolist()
        display = invalid_rows[:5]
        rows_str = f"{display} ... (共 {len(invalid_rows)} 行)" if len(invalid_rows) > 5 else str(invalid_rows)
        raise ValueError(f"存在员工/渠道有效信息，但组织层级列存在空值，不符合要求。行索引：{rows_str}")

def aggregate_special_category(df):
    mask = df[COL_BIZ_CAT2] == SPECIAL_AGGREGATE_CATEGORY
    if not mask.any():
        return df
    special_df = df[mask].copy()
    group_cols = [
        COL_EMP_ID, COL_EMP_NAME, COL_AGENT_ID, COL_AGENT_NAME,
        COL_ARCH1, COL_ARCH2, COL_ARCH3, COL_ARCH4,
        COL_OPS_ID, COL_OPS_NAME, COL_INDUSTRY, COL_SALES_ID, COL_SALES_NAME
    ]
    agg_cols = [COL_TXN_COUNT, COL_TXN_AMT, COL_REWARD, COL_REWARD_BASE, COL_SETTLE, COL_PROFIT]
    for col in agg_cols:
        if col in special_df.columns:
            special_df[col] = pd.to_numeric(special_df[col], errors='coerce')
    grouped = special_df.groupby(group_cols, as_index=False)[agg_cols].sum()
    aggregated = pd.DataFrame(columns=df.columns)
    for _, row in grouped.iterrows():
        new_row = {col: pd.NA for col in df.columns}
        for col in group_cols:
            if col in row:
                new_row[col] = row[col]
        for col in agg_cols:
            if col in row:
                new_row[col] = row[col]
        first_row = special_df.iloc[0]
        for col in [COL_CHANNEL, COL_BIZ_CAT1, COL_BIZ_CAT2, COL_BILL_PERIOD, COL_BI_PERIOD, COL_TABLE_NAME,
                    COL_INVOICE, COL_PROJECT, COL_PROJECT_TYPE, COL_PROJECT_CAT, COL_SETTLE_METHOD,
                    COL_AGENT_SETTLE, COL_EMP_SETTLE, COL_EMP_D0, COL_AGENT_D0]:
            if col in first_row:
                new_row[col] = first_row[col]
        new_row[COL_MERCHANT_NO] = pd.NA
        new_row[COL_MERCHANT_NAME] = pd.NA
        new_row[COL_RATIO] = pd.NA
        aggregated = pd.concat([aggregated, pd.DataFrame([new_row])], ignore_index=True)
    df = df[~mask].copy()
    df = pd.concat([df, aggregated], ignore_index=True)
    print(f"   [特殊汇总] 分类 '{SPECIAL_AGGREGATE_CATEGORY}' 已按指定维度汇总，原始 {mask.sum()} 行汇总为 {len(aggregated)} 行")
    return df

def check_and_clean_data(df, file_name):
    if not match_header_case_insensitive(df, EXPECTED_HEADERS):
        if len(df.columns) != len(EXPECTED_HEADERS):
            if len(df.columns) > len(EXPECTED_HEADERS):
                df = df.iloc[:, :len(EXPECTED_HEADERS)]
            else:
                for i in range(len(EXPECTED_HEADERS) - len(df.columns)):
                    df[f"_临时列_{i}"] = pd.NA
        df.columns = EXPECTED_HEADERS
        print(f"   [提示] 文件 {file_name} 表头列数已自动调整并替换为标准表头")

    df = df.map(lambda x: pd.NA if isinstance(x, str) and x == "(空白)" else x)

    if COL_MERCHANT_NO in df.columns:
        df[COL_MERCHANT_NO] = df[COL_MERCHANT_NO].apply(clean_merchant_no)
        non_empty = df[COL_MERCHANT_NO].str.len().notna() & (df[COL_MERCHANT_NO] != "")
        length_issues = df.loc[non_empty, COL_MERCHANT_NO].str.len()
        if (length_issues != 15).any():
            print(f"   [提示] 有 {(length_issues != 15).sum()} 行商户编号清洗后长度不是15位，请核实源数据")

    for col in [COL_BIZ_CAT2, COL_TABLE_NAME]:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: x.replace("数励", "激励") if isinstance(x, str) else x)

    if COL_TABLE_NAME in df.columns and COL_BIZ_CAT2 in df.columns:
        mask_empty = df[COL_TABLE_NAME].isna()
        if mask_empty.any():
            df.loc[mask_empty, COL_TABLE_NAME] = df.loc[mask_empty, COL_BIZ_CAT2]
            print(f"   [提示] 已为 {mask_empty.sum()} 行填充数据表（=业务小类）")

    df = convert_id_columns(df)

    if COL_EMP_ID in df.columns and COL_ARCH2 in df.columns:
        mask = df[COL_EMP_ID] == SPECIAL_MANAGER_ID
        df.loc[mask, COL_ARCH2] = SPECIAL_SECOND_ARCH

    def is_empty_or_zero(val):
        if pd.isna(val):
            return True
        if isinstance(val, (int, float)) and val == 0:
            return True
        if isinstance(val, str) and val.strip() in ("0", ""):
            return True
        return False

    arch_cols = [col for col in COMPANY_LEVEL_COLS if col in df.columns]
    for idx, row in df.iterrows():
        col_vals = [row.get(col) for col in arch_cols]
        empty_flags = [is_empty_or_zero(v) for v in col_vals]
        if any(empty_flags) and not all(empty_flags):
            raise ValueError(
                f"组织层级四列必须同时有值或同时为空，但第 {idx} 行存在部分为空的情况。"
                f" 各列值：{dict(zip(arch_cols, col_vals))}"
            )

    for idx, row in df.iterrows():
        all_empty = all(is_empty_or_zero(row[col]) for col in arch_cols)
        if all_empty:
            for col in arch_cols:
                df.at[idx, col] = NO_ARCH_LABEL

    if COL_ARCH2 in df.columns:
        for old_name, new_name in ARCH_REPLACE_MAP.items():
            mask = df[COL_ARCH2] == old_name
            if mask.any():
                df.loc[mask, COL_ARCH2] = new_name
                print(f"   [提示] 已将 {mask.sum()} 行 '{old_name}' 替换为 '{new_name}'")

    df, errors_df = validate_second_arch(df, VALID_SECOND_ARCH_LABELS)

    if COL_ARCH1 in df.columns and COL_ARCH2 in df.columns:
        mask_level1_fens = df[COL_ARCH1] == "分公司"
        if mask_level1_fens.any():
            invalid_mask = mask_level1_fens & (~df[COL_ARCH2].astype(str).str.contains("分公司", na=False)) & (~df[COL_ARCH2].isin(ALLOWED_ARCH_EXCEPTIONS))
            if invalid_mask.any():
                invalid_rows = df.index[invalid_mask].tolist()
                display = invalid_rows[:5]
                rows_str = f"{display} ... (共 {len(invalid_rows)} 行)" if len(invalid_rows) > 5 else str(invalid_rows)
                raise ValueError(f"当组织层级1为'分公司'时，组织层级2必须包含'分公司'字样，或者是 '{ALLOWED_ARCH_EXCEPTIONS}' 之一。不符合的行索引：{rows_str}")

    check_admin_arch_when_have_agent(df, file_name)
    check_uniqueness(df, file_name)

    if COL_BIZ_CAT1 in df.columns and COL_ARCH3 in df.columns:
        mask_edu = df[COL_BIZ_CAT1] == EDU_FIRST_CATEGORY
        if mask_edu.any():
            invalid_mask = mask_edu & (df[COL_ARCH3].astype(str).str.strip() != EDU_THIRD_ARCH)
            if invalid_mask.any():
                invalid_rows = df.index[invalid_mask].tolist()
                display = invalid_rows[:5]
                rows_str = f"{display} ... (共 {len(invalid_rows)} 行)" if len(invalid_rows) > 5 else str(invalid_rows)
                raise ValueError(f"业务大类为'{EDU_FIRST_CATEGORY}'时，组织层级3必须为'{EDU_THIRD_ARCH}'。不符合的行索引：{rows_str}")

    if COL_INVOICE in df.columns and COL_EMP_NAME in df.columns and COL_AGENT_NAME in df.columns:
        def is_valid_invoice(row):
            invoice = row.get(COL_INVOICE)
            if pd.isna(invoice) or str(invoice).strip() == "":
                return True
            if row.get(COL_BIZ_CAT2) == INVOICE_EXEMPT_CATEGORY:
                return True
            manager = row.get(COL_EMP_NAME)
            agent = row.get(COL_AGENT_NAME)
            invoice_str = str(invoice).strip()
            manager_str = str(manager).strip() if not pd.isna(manager) else ""
            agent_str = str(agent).strip() if not pd.isna(agent) else ""
            return invoice_str == manager_str or invoice_str == agent_str
        invalid_mask = ~df.apply(is_valid_invoice, axis=1)
        if invalid_mask.any():
            invalid_rows = df.index[invalid_mask].tolist()
            display = invalid_rows[:5]
            rows_str = f"{display} ... (共 {len(invalid_rows)} 行)" if len(invalid_rows) > 5 else str(invalid_rows)
            print(f"   [警告] 发票抬头非空但既不等于员工姓名也不等于渠道名称（且非豁免分类），行索引：{rows_str}")

    if COL_RATIO in df.columns:
        def fix_ratio(val):
            if isinstance(val, str) and "/" in val:
                return 0.0
            return val
        df[COL_RATIO] = df[COL_RATIO].apply(fix_ratio)
        df[COL_RATIO] = pd.to_numeric(df[COL_RATIO], errors='coerce')

    for col in MULTIPLY_COLS:
        if col not in df.columns:
            raise ValueError(f"调分字段 '{col}' 不存在于表头中！")
        numeric_col = pd.to_numeric(df[col], errors='coerce')
        invalid_mask = df[col].notna() & numeric_col.isna()
        if invalid_mask.any():
            invalid_rows = df.index[invalid_mask].tolist()
            display = invalid_rows[:5]
            rows_str = f"{display} ... (共 {len(invalid_rows)} 行)" if len(invalid_rows) > 5 else str(invalid_rows)
            sample_vals = df.loc[invalid_rows[0], col] if invalid_rows else ""
            raise ValueError(f"字段 '{col}' 包含非数值内容，无法调分，行索引：{rows_str}，例如：{sample_vals}")
        df[col] = (numeric_col.round(2) * 100).round(0).astype('Int64')

    if COL_PROFIT in df.columns and COL_REWARD in df.columns and COL_SETTLE in df.columns:
        profit_null = df[COL_PROFIT].isna()
        if profit_null.any():
            df.loc[profit_null, COL_PROFIT] = df[COL_REWARD] - df[COL_SETTLE]
            print(f"   [提示] 已为 {profit_null.sum()} 行填充净收益（补贴金额 - 支出金额）")
        mask_non_null = df[COL_REWARD].notna() & df[COL_SETTLE].notna()
        if mask_non_null.any():
            correct_profit = df[COL_REWARD] - df[COL_SETTLE]
            diff_mask = (df[COL_PROFIT] != correct_profit) & mask_non_null
            if diff_mask.any():
                df.loc[diff_mask, COL_PROFIT] = correct_profit
                print(f"   [修正] 已强制修正 {diff_mask.sum()} 行净收益为 补贴金额 - 支出金额")

    df = aggregate_special_category(df)
    df = format_date_columns(df)
    return df, errors_df

def get_target_sheet(xl):
    sheet_names = xl.sheet_names
    if not sheet_names:
        raise ValueError("工作簿中没有工作表")
    for name in sheet_names:
        if name.lower() == "bi":
            return name
    return sheet_names[-1]

def save_excel_with_date_format(df, output_path, sheet_name):
    df_copy = df.copy()
    date_cols = [COL_BILL_PERIOD, COL_BI_PERIOD]
    for col in date_cols:
        if col in df_copy.columns:
            if pd.api.types.is_datetime64_any_dtype(df_copy[col]):
                df_copy[col] = df_copy[col].dt.strftime("%Y-%m")
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_copy.to_excel(writer, sheet_name=sheet_name, index=False)

def process_file(input_path, output_dir):
    file_name = os.path.basename(input_path)
    print(f"正在处理：{file_name}")
    try:
        xl = pd.ExcelFile(input_path)
        target_sheet = get_target_sheet(xl)
        df = pd.read_excel(input_path, sheet_name=target_sheet, dtype=object, keep_default_na=False)
        df = rename_columns(df)   # 关键：读入后立即脱敏字段名
        df = df.replace(r'^\s*$', pd.NA, regex=True)
    except Exception as e:
        raise Exception(f"读取文件失败：{e}")
    try:
        df_clean, errors_df = check_and_clean_data(df.copy(), file_name)
    except Exception as e:
        raise Exception(f"校验或清洗失败：{e}")

    if COL_BIZ_CAT2 in df_clean.columns:
        found = df_clean[df_clean[COL_BIZ_CAT2].isin(REMINDER_CATEGORIES)][COL_BIZ_CAT2].unique()
        if len(found) > 0:
            print(f"   [提醒] 文件 {file_name} 包含以下需要调整四级标签和增值划分的分类：")
            for cat in found:
                print(f"          - {cat}")

    output_path = os.path.join(output_dir, f"{Path(file_name).stem}_processed.xlsx")
    save_excel_with_date_format(df_clean, output_path, target_sheet)
    print(f"处理成功，已保存至：{output_path}")
    return df_clean, errors_df

def generate_summary(output_dir, all_errors):
    processed_files = list(Path(output_dir).glob("*_processed.xlsx"))
    if not processed_files:
        print("未找到任何处理后的文件，无法生成汇总清单。")
        return
    all_data = []
    invalid_bi_records = []
    for file_path in processed_files:
        try:
            df = pd.read_excel(file_path, dtype=object)
            required_cols = [COL_BIZ_CAT2, COL_BI_PERIOD, COL_BILL_PERIOD, COL_REWARD, COL_REWARD_BASE, COL_SETTLE, COL_PROFIT, COL_CHANNEL, COL_BIZ_CAT1]
            if not all(col in df.columns for col in required_cols):
                print(f"跳过文件 {file_path.name}：缺少必要的列")
                continue
            bi_series = df[COL_BI_PERIOD]
            if bi_series.dtype == 'object':
                bi_str = []
                for v in bi_series:
                    if pd.isna(v):
                        bi_str.append("")
                    elif isinstance(v, datetime):
                        bi_str.append(v.strftime("%Y-%m"))
                    else:
                        try:
                            dt = pd.to_datetime(v)
                            if pd.notna(dt):
                                bi_str.append(dt.strftime("%Y-%m"))
                            else:
                                bi_str.append(str(v).strip())
                        except:
                            bi_str.append(str(v).strip())
                df["报表期_str"] = bi_str
            else:
                df["报表期_str"] = bi_series.dt.strftime("%Y-%m")
            df["报表期校验"] = df["报表期_str"].apply(lambda x: "合规" if x == LAST_MONTH else f"不合规（应为{LAST_MONTH}）")
            invalid = df[df["报表期校验"] != "合规"].copy()
            if not invalid.empty:
                invalid["来源文件"] = file_path.name
                invalid_bi_records.append(invalid)
            sub_df = df[required_cols].copy()
            for col in [COL_REWARD, COL_REWARD_BASE, COL_SETTLE, COL_PROFIT]:
                sub_df[col] = pd.to_numeric(sub_df[col], errors='coerce') / 100.0
            sub_df = sub_df[sub_df[COL_BIZ_CAT2].notna()]
            if not sub_df.empty:
                sub_df["报表期_str"] = df["报表期_str"]
                sub_df["报表期校验"] = df["报表期校验"]
                all_data.append(sub_df)
        except Exception as e:
            print(f"读取文件 {file_path.name} 失败：{e}")

    if invalid_bi_records:
        invalid_all = pd.concat(invalid_bi_records, ignore_index=True)
        output_invalid_path = os.path.join(output_dir, "报表期异常明细.xlsx")
        try:
            if os.path.exists(output_invalid_path):
                os.remove(output_invalid_path)
            with pd.ExcelWriter(output_invalid_path, engine='openpyxl') as writer:
                invalid_all.to_excel(writer, sheet_name="异常记录", index=False)
            print(f"\n⚠️ 报表期异常记录已导出至：{output_invalid_path}，请核查。")
        except PermissionError:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            new_path = output_invalid_path.replace(".xlsx", f"_{timestamp}.xlsx")
            with pd.ExcelWriter(new_path, engine='openpyxl') as writer:
                invalid_all.to_excel(writer, sheet_name="异常记录", index=False)
            print(f"\n原文件被占用，已另存至：{new_path}")
    else:
        print("\n所有报表期均已纠正，无异常明细。")

    if not all_data:
        print("没有有效数据可汇总。")
        return

    combined = pd.concat(all_data, ignore_index=True)
    group_cols = [COL_BIZ_CAT2, "报表期_str", COL_BILL_PERIOD]
    summary = combined.groupby(group_cols).agg(
        **{
            COL_CHANNEL: (COL_CHANNEL, "first"),
            COL_BIZ_CAT1: (COL_BIZ_CAT1, "first"),
            COL_REWARD: (COL_REWARD, "sum"),
            "补贴金额成本基数": (COL_REWARD_BASE, "sum"),
            COL_SETTLE: (COL_SETTLE, "sum"),
            COL_PROFIT: (COL_PROFIT, "sum"),
            "报表期状态": ("报表期校验", lambda x: "⚠️存在异常" if (x != "合规").any() else "全部合规"),
        }
    ).reset_index()
    summary.rename(columns={"报表期_str": COL_BI_PERIOD}, inplace=True)
    for col in [COL_REWARD, "补贴金额成本基数", COL_SETTLE, COL_PROFIT]:
        summary[col] = summary[col].round(2)
    summary["提醒信息"] = summary[COL_BIZ_CAT2].apply(
        lambda x: "需调整四级标签和增值划分" if x in REMINDER_CATEGORIES else ""
    )

    try:
        match_df = pd.read_excel(MATCH_FILE, dtype=str)
        match_df.columns = [col.strip() for col in match_df.columns]
        # 匹配表里字段名也需脱敏，约定它使用同样字段名
        match_df = rename_columns(match_df)
        match_df = match_df[[COL_BIZ_CAT2, COL_CHANNEL, COL_BIZ_CAT1]].drop_duplicates(subset=[COL_BIZ_CAT2])
        match_df.rename(columns={COL_CHANNEL: '匹配渠道', COL_BIZ_CAT1: '匹配业务大类'}, inplace=True)
        summary = summary.merge(match_df, on=COL_BIZ_CAT2, how='left')
        summary['渠道一致'] = np.where((summary[COL_CHANNEL] == summary['匹配渠道']) | summary['匹配渠道'].isna(), '一致', '不一致')
        summary['业务大类一致'] = np.where((summary[COL_BIZ_CAT1] == summary['匹配业务大类']) | summary['匹配业务大类'].isna(), '一致', '不一致')
    except Exception as e:
        print(f"加载匹配表失败：{e}，跳过匹配校验")
        for c in ['匹配渠道', '匹配业务大类', '渠道一致', '业务大类一致']:
            summary[c] = ''

    try:
        check_df = pd.read_excel(CHECK_FILE2, sheet_name='Sheet1', dtype=str)
        check_df.columns = [col.strip() for col in check_df.columns]
        # 线下校验表字段名也需要外部映射，此处使用约定的脱敏名
        check_cols = [COL_BIZ_CAT2, '补贴金额（收入）', '支出金额（成本）']
        avail = [c for c in check_cols if c in check_df.columns]
        if len(avail) == 3:
            check_df = check_df[check_cols].drop_duplicates(subset=[COL_BIZ_CAT2])
            for col in ['补贴金额（收入）', '支出金额（成本）']:
                check_df[col] = pd.to_numeric(check_df[col], errors='coerce')
            check_df.rename(columns={'补贴金额（收入）': '线下补贴金额', '支出金额（成本）': '线下支出金额'}, inplace=True)
            summary = summary.merge(check_df, on=COL_BIZ_CAT2, how='left')
            summary['补贴差异'] = summary[COL_REWARD].fillna(0) - summary['线下补贴金额'].fillna(0)
            summary['支出差异'] = summary[COL_SETTLE].fillna(0) - summary['线下支出金额'].fillna(0)
            summary.loc[summary[COL_REWARD].isna() & summary['线下补贴金额'].isna(), '补贴差异'] = np.nan
            summary.loc[summary[COL_SETTLE].isna() & summary['线下支出金额'].isna(), '支出差异'] = np.nan
        else:
            raise ValueError("线下校验表列名不匹配")
    except Exception as e:
        print(f"加载收入成本校验表失败：{e}")
        for c in ['线下补贴金额', '线下支出金额', '补贴差异', '支出差异']:
            summary[c] = ''

    final_cols = [
        COL_BIZ_CAT2, COL_BI_PERIOD, COL_BILL_PERIOD, COL_CHANNEL, COL_BIZ_CAT1,
        '匹配渠道', '渠道一致', '匹配业务大类', '业务大类一致',
        COL_REWARD, COL_REWARD_BASE, COL_SETTLE, COL_PROFIT,
        '线下补贴金额', '补贴差异', '线下支出金额', '支出差异',
        '报表期状态', '提醒信息'
    ]
    summary = summary[[col for col in final_cols if col in summary.columns]]

    output_summary_path = os.path.join(output_dir, "汇总统计清单.xlsx")
    try:
        if os.path.exists(output_summary_path):
            os.remove(output_summary_path)
        with pd.ExcelWriter(output_summary_path, engine='openpyxl') as writer:
            summary.to_excel(writer, sheet_name="按业务小类_期间汇总", index=False)
            if all_errors:
                error_df = pd.concat(all_errors, ignore_index=True)
                error_df.to_excel(writer, sheet_name="组织层级校验错误", index=False)
        print(f"\n汇总统计清单已保存至：{output_summary_path}")
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        new_path = output_summary_path.replace(".xlsx", f"_{timestamp}.xlsx")
        with pd.ExcelWriter(new_path, engine='openpyxl') as writer:
            summary.to_excel(writer, sheet_name="按业务小类_期间汇总", index=False)
            if all_errors:
                error_df = pd.concat(all_errors, ignore_index=True)
                error_df.to_excel(writer, sheet_name="组织层级校验错误", index=False)
        print(f"\n汇总统计清单已另存至：{new_path}")

    if all_errors:
        print(f"\n⚠️ 组织层级校验错误 {len(pd.concat(all_errors, ignore_index=True))} 条，详见汇总文件。")

def merge_all_processed(output_dir):
    processed_files = list(Path(output_dir).glob("*_processed.xlsx"))
    if not processed_files:
        print("未找到任何处理后的文件，无法合并。")
        return
    all_dfs = []
    for file_path in processed_files:
        try:
            df = pd.read_excel(file_path, dtype=object)
            df["来源文件"] = file_path.stem
            all_dfs.append(df)
        except Exception as e:
            print(f"读取文件 {file_path.name} 失败：{e}")
    if not all_dfs:
        print("没有可合并的数据。")
        return
    merged = pd.concat(all_dfs, ignore_index=True, sort=False)
    for col in EXPECTED_HEADERS:
        if col not in merged.columns:
            merged[col] = pd.NA
    merged = merged[EXPECTED_HEADERS + ["来源文件"]]
    output_merged_path = os.path.join(output_dir, "合并数据总表.xlsx")
    try:
        if os.path.exists(output_merged_path):
            os.remove(output_merged_path)
        with pd.ExcelWriter(output_merged_path, engine='openpyxl') as writer:
            merged.to_excel(writer, sheet_name="合并数据", index=False)
        print(f"\n合并后的总表已保存至：{output_merged_path}，共 {len(merged)} 行")
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        new_path = output_merged_path.replace(".xlsx", f"_{timestamp}.xlsx")
        with pd.ExcelWriter(new_path, engine='openpyxl') as writer:
            merged.to_excel(writer, sheet_name="合并数据", index=False)

def batch_process(input_folder, output_folder=None):
    if output_folder is None:
        output_folder = os.path.join(input_folder, "processed_output")
    os.makedirs(output_folder, exist_ok=True)
    exts = (".xlsx", ".xls")
    files = [f for f in os.listdir(input_folder) if f.lower().endswith(exts)]
    if not files:
        print("未找到Excel文件。")
        return
    success_count = 0
    fail_list = []
    all_errors = []
    for file in files:
        input_path = os.path.join(input_folder, file)
        try:
            _, errors_df = process_file(input_path, output_folder)
            if not errors_df.empty:
                errors_df['来源文件'] = file
                all_errors.append(errors_df)
            success_count += 1
        except Exception as e:
            fail_list.append((file, str(e)))
            print(f"   [错误] {file} 处理失败：{e}")
    print(f"\n处理完成。成功：{success_count}，失败：{len(fail_list)}")
    if fail_list:
        print("失败文件列表：")
        for name, err in fail_list:
            print(f"  - {name}: {err}")
    if success_count > 0:
        generate_summary(output_folder, all_errors)
        merge_all_processed(output_folder)

if __name__ == "__main__":
    print(f"开始处理，目标文件夹：{folder_path}")
    batch_process(folder_path)
    input("全部处理完成，按回车键退出...")