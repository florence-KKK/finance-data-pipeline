import os
import sys
import json
import pandas as pd
import numpy as np
from pathlib import Path
from datetime import date, datetime
import warnings
warnings.filterwarnings("ignore")

# ==================== 全局错误捕获 ====================
def global_exception_handler(exc_type, exc_value, exc_traceback):
    print("\n" + "="*60)
    print("发生未捕获的异常：")
    print(f"类型：{exc_type}")
    print(f"值：{exc_value}")
    import traceback
    traceback.print_tb(exc_traceback)
    print("="*60)
    input("按回车键退出...")
    sys.exit(1)

sys.excepthook = global_exception_handler

# ==================== 敏感信息配置区（外发时替换为占位符）====================
PATHS = {
    'input_folder':     r'F:\...\BI待处理数据-项目2-系统导出-py',
    'field_map_file':   r'F:\...\field_map.json',
    'match_file':       r'F:\...\活动名称对应一二级分类.xlsx',
    'second_arch_file': r'F:\...\分公司目录-新二级部门标签.xlsx',
    'output_dir':       r'F:\...\processed_output_split',
}

RULES = {
    'arch_replace_map': {
        '【旧分公司A】': '【新分公司A】',
        '【旧分公司B】': '【新分公司B】',
        '【旧分公司C】': '【新分公司C】',
        '【旧总部】':   '【新总部】',
    },
    'special_manager_id':  0,
    'special_second_arch': '【特殊分公司名称】',
    'reminder_categories': ['【提醒分类1】', '【提醒分类2】'],
    'allowed_arch_exceptions': ['【例外部门1】', '【例外部门2】'],
    'no_arch_label': '无归属部门',
    'match_sheet_name': '名称对应活动类型',
}

# 脱敏后的字段名常量（代码全程使用）
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

# 脱敏后的标准表头
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

# 源数据列（来自系统导出的原始数据，脱敏后仅保留列名）
SOURCE_COLS = [
    COL_MERCHANT_NO, COL_MERCHANT_NAME,
    COL_EMP_ID, COL_EMP_NAME,
    COL_AGENT_ID, COL_AGENT_NAME,
    COL_ARCH1, COL_ARCH2, COL_ARCH3, COL_ARCH4,
    COL_OPS_ID, COL_OPS_NAME, COL_INDUSTRY,
    COL_SALES_ID, COL_SALES_NAME,
    COL_TXN_COUNT, COL_TXN_AMT,
    COL_REWARD, COL_REWARD_BASE, COL_RATIO,
    COL_SETTLE, COL_PROFIT,
    '活动名称', COL_BILL_PERIOD,
]

AMOUNT_COLS = [COL_TXN_AMT, COL_REWARD, COL_REWARD_BASE, COL_SETTLE, COL_PROFIT]
COMPANY_LEVEL_COLS = [COL_ARCH1, COL_ARCH2, COL_ARCH3, COL_ARCH4]
# ==========================================================================

with open(PATHS['field_map_file'], 'r', encoding='utf-8') as f:
    FIELD_MAP = json.load(f)

INPUT_FOLDER     = PATHS['input_folder']
MATCH_FILE       = PATHS['match_file']
MATCH_SHEET      = RULES['match_sheet_name']
OUTPUT_DIR       = PATHS['output_dir']
SECOND_ARCH_FILE = PATHS['second_arch_file']

SPECIAL_MANAGER_ID  = RULES['special_manager_id']
SPECIAL_SECOND_ARCH = RULES['special_second_arch']
REMINDER_CATEGORIES = RULES['reminder_categories']
ARCH_REPLACE_MAP    = RULES['arch_replace_map']
NO_ARCH_LABEL       = RULES['no_arch_label']
ALLOWED_ARCH_EXCEPTIONS = RULES['allowed_arch_exceptions']

def get_last_month_str():
    today = date.today()
    first_of_month = today.replace(day=1)
    last_month = first_of_month - pd.DateOffset(months=1)
    return last_month.strftime("%Y-%m")

LAST_MONTH = get_last_month_str()
print(f"当前系统日期：{date.today()}")
print(f"报表期基准（上个月）：{LAST_MONTH}")
print("项目2 报表期将自动纠正为此值。\n")

# ==================== 列名重命名（原始 → 脱敏） ====================
def rename_columns(df):
    return df.rename(columns=FIELD_MAP)

# ==================== 加载组织层级目录 ====================
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

# ==================== 辅助函数 ====================
def find_input_file(folder):
    if not os.path.exists(folder):
        print(f"错误：文件夹不存在 {folder}")
        return None
    exts = (".xlsx", ".xls")
    files = [f for f in os.listdir(folder) if f.lower().endswith(exts) and not f.startswith('~$')]
    if not files:
        return None
    if len(files) == 1:
        return os.path.join(folder, files[0])
    else:
        print("发现多个 Excel 文件，请选择要处理的文件：")
        for i, f in enumerate(files, 1):
            print(f"  {i}. {f}")
        while True:
            try:
                choice = input("请输入序号（1-{}）：".format(len(files)))
                idx = int(choice) - 1
                if 0 <= idx < len(files):
                    return os.path.join(folder, files[idx])
                else:
                    print("序号无效，请重新输入。")
            except ValueError:
                print("请输入数字序号。")

def load_match_table():
    try:
        df_match = pd.read_excel(MATCH_FILE, sheet_name=MATCH_SHEET, dtype=str)
        # 匹配表字段名也需要脱敏处理
        df_match = rename_columns(df_match)
        required = ['活动名称', COL_CHANNEL, COL_BIZ_CAT1, COL_BIZ_CAT2, COL_TABLE_NAME]
        for col in required:
            if col not in df_match.columns:
                raise ValueError(f"匹配文件缺少列：{col}")
        df_match = df_match.drop_duplicates(subset=["活动名称"])
        return df_match.set_index("活动名称")
    except Exception as e:
        raise Exception(f"加载匹配文件失败：{e}")

def convert_to_standard(df_source, match_dict):
    df_source.columns = [col.strip() for col in df_source.columns]
    df_std = pd.DataFrame(columns=EXPECTED_HEADERS)
    for col in SOURCE_COLS:
        if col in df_source.columns:
            df_std[col] = df_source[col]
        else:
            df_std[col] = pd.NA
    df_std[COL_BI_PERIOD] = LAST_MONTH
    if "活动名称" not in df_std.columns:
        raise ValueError("转换后的数据中缺少“活动名称”列")
    all_activities = df_std["活动名称"].dropna().unique()
    missing = [act for act in all_activities if act not in match_dict.index]
    if missing:
        print("\n错误：以下活动名称在匹配表中找不到对应记录，请补充匹配文件后重新运行：")
        for act in missing:
            print(f"  - {act}")
        raise ValueError(f"缺少 {len(missing)} 个活动名称的匹配信息")
    def get_match(activity_name, field):
        if pd.isna(activity_name):
            return pd.NA
        row = match_dict.loc[activity_name]
        return row.get(field, pd.NA)
    df_std[COL_CHANNEL]     = df_std["活动名称"].apply(lambda x: get_match(x, COL_CHANNEL))
    df_std[COL_BIZ_CAT1]    = df_std["活动名称"].apply(lambda x: get_match(x, COL_BIZ_CAT1))
    df_std[COL_BIZ_CAT2]    = df_std["活动名称"].apply(lambda x: get_match(x, COL_BIZ_CAT2))
    df_std[COL_TABLE_NAME]  = df_std["活动名称"].apply(lambda x: get_match(x, COL_TABLE_NAME))

    empty_channel_mask = df_std[COL_CHANNEL].isna() | (df_std[COL_CHANNEL].astype(str).str.strip() == "")
    empty_firstcat_mask = df_std[COL_BIZ_CAT1].isna() | (df_std[COL_BIZ_CAT1].astype(str).str.strip() == "")
    if empty_channel_mask.any():
        empty_activities = df_std.loc[empty_channel_mask, "活动名称"].dropna().unique()
        print(f"\n⚠️ 警告：以下活动名称匹配到的【支付渠道】为空，请检查匹配表：")
        for act in empty_activities:
            print(f"  - {act}")
    if empty_firstcat_mask.any():
        empty_activities = df_std.loc[empty_firstcat_mask, "活动名称"].dropna().unique()
        print(f"\n⚠️ 警告：以下活动名称匹配到的【业务大类】为空，请检查匹配表：")
        for act in empty_activities:
            print(f"  - {act}")

    for col in [COL_INVOICE, COL_PROJECT, COL_PROJECT_TYPE, COL_PROJECT_CAT,
                COL_SETTLE_METHOD, COL_AGENT_SETTLE, COL_EMP_SETTLE,
                COL_EMP_D0, COL_AGENT_D0]:
        df_std[col] = pd.NA
    df_std.drop(columns=["活动名称"], inplace=True)
    return df_std

def check_and_clean_data(df, file_name):
    if len(df.columns) != len(EXPECTED_HEADERS):
        raise ValueError(f"列数不一致！实际列数：{len(df.columns)}，期望列数：{len(EXPECTED_HEADERS)}")
    df.columns = EXPECTED_HEADERS

    df = df.map(lambda x: pd.NA if isinstance(x, str) and x == "(空白)" else x)

    for col in [COL_BIZ_CAT2, COL_TABLE_NAME]:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: x.replace("数励", "激励") if isinstance(x, str) else x)

    id_cols = [col for col in df.columns if "编号" in col]
    for col in id_cols:
        if df[col].isnull().all():
            continue
        numeric_series = pd.to_numeric(df[col], errors='coerce')
        invalid_mask = df[col].notna() & numeric_series.isna()
        if invalid_mask.any():
            invalid_rows = df.index[invalid_mask].tolist()
            raise ValueError(f"列 '{col}' 中存在非数值内容，行索引：{invalid_rows}")
        df[col] = numeric_series

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

    for idx, row in df.iterrows():
        all_empty = all(is_empty_or_zero(row[col]) for col in COMPANY_LEVEL_COLS)
        if all_empty:
            for col in COMPANY_LEVEL_COLS:
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
                raise ValueError(f"当组织层级1为'分公司'时，组织层级2必须包含'分公司'字样，或者是 '{ALLOWED_ARCH_EXCEPTIONS}' 之一。不符合的行索引：{invalid_rows}")

    key_cols = [COL_EMP_ID, COL_EMP_NAME, COL_AGENT_ID, COL_AGENT_NAME]
    def is_valid_series(series):
        return series.notna() & (series != 0) & (series != "0") & (series != "") & (series != pd.NA)
    has_agent = pd.Series(False, index=df.index)
    for col in key_cols:
        if col in df.columns:
            has_agent = has_agent | is_valid_series(df[col])
    is_all_wuguis = df[COMPANY_LEVEL_COLS].eq(NO_ARCH_LABEL).all(axis=1)
    invalid_mask = has_agent & is_all_wuguis
    if invalid_mask.any():
        invalid_rows = df.index[invalid_mask].tolist()
        raise ValueError(f"存在员工/渠道有效信息，但组织层级全为'{NO_ARCH_LABEL}'，不符合要求。行索引：{invalid_rows}")

    errors = []
    if COL_EMP_ID in df.columns and COL_EMP_NAME in df.columns:
        non_null = df[[COL_EMP_ID, COL_EMP_NAME]].dropna()
        if not non_null.empty:
            grouped = non_null.groupby(COL_EMP_ID)[COL_EMP_NAME].nunique()
            dup_ids = grouped[grouped > 1].index.tolist()
            if dup_ids:
                conflict_rows = non_null[non_null[COL_EMP_ID].isin(dup_ids)].index.tolist()
                errors.append(f"员工编号对应多个员工姓名，冲突行索引：{conflict_rows}")
    if COL_AGENT_ID in df.columns and COL_AGENT_NAME in df.columns:
        non_null = df[[COL_AGENT_ID, COL_AGENT_NAME]].dropna()
        if not non_null.empty:
            grouped = non_null.groupby(COL_AGENT_ID)[COL_AGENT_NAME].nunique()
            dup_ids = grouped[grouped > 1].index.tolist()
            if dup_ids:
                conflict_rows = non_null[non_null[COL_AGENT_ID].isin(dup_ids)].index.tolist()
                errors.append(f"渠道编号对应多个渠道名称，冲突行索引：{conflict_rows}")
    if COL_AGENT_ID in df.columns and COL_EMP_ID in df.columns:
        non_null = df[[COL_AGENT_ID, COL_EMP_ID]].dropna()
        if not non_null.empty:
            grouped = non_null.groupby(COL_AGENT_ID)[COL_EMP_ID].nunique()
            dup_ids = grouped[grouped > 1].index.tolist()
            if dup_ids:
                conflict_rows = non_null[non_null[COL_AGENT_ID].isin(dup_ids)].index.tolist()
                errors.append(f"渠道编号对应多个员工编号，冲突行索引：{conflict_rows}")
    if errors:
        raise ValueError("\n".join(errors))

    if COL_RATIO in df.columns:
        def fix_ratio(val):
            if isinstance(val, str) and "/" in val:
                return 0.0
            return val
        df[COL_RATIO] = df[COL_RATIO].apply(fix_ratio)
        df[COL_RATIO] = pd.to_numeric(df[COL_RATIO], errors='coerce')

    for col in AMOUNT_COLS:
        if col not in df.columns:
            raise ValueError(f"金额字段 '{col}' 不存在")
        numeric_col = pd.to_numeric(df[col], errors='coerce')
        invalid_mask = df[col].notna() & numeric_col.isna()
        if invalid_mask.any():
            invalid_rows = df.index[invalid_mask].tolist()
            sample_vals = df.loc[invalid_rows[0], col] if invalid_rows else ""
            raise ValueError(f"字段 '{col}' 包含非数值内容，行索引：{invalid_rows}，例如：{sample_vals}")
        not_integer = (numeric_col.notna()) & (numeric_col != numeric_col.round(0))
        if not_integer.any():
            warn_rows = df.index[not_integer].tolist()
            print(f"   [警告] 字段 '{col}' 中存在非整数值（将四舍五入取整），受影响行数：{len(warn_rows)}")
            numeric_col = numeric_col.round(0)
        df[col] = numeric_col.astype('Int64')

    if COL_PROFIT in df.columns and COL_REWARD in df.columns and COL_SETTLE in df.columns:
        profit_null = df[COL_PROFIT].isna()
        if profit_null.any():
            df.loc[profit_null, COL_PROFIT] = df[COL_REWARD] - df[COL_SETTLE]
            print(f"   [提示] 已为 {profit_null.sum()} 行填充净收益（补贴金额 - 支出金额）")
        calc_profit = df[COL_REWARD] - df[COL_SETTLE]
        mask_non_null = df[COL_PROFIT].notna() & df[COL_REWARD].notna() & df[COL_SETTLE].notna()
        diff = (df[COL_PROFIT] - calc_profit).abs()
        invalid_rows = diff[mask_non_null & (diff != 0)].index.tolist()
        if invalid_rows:
            sample = df.loc[invalid_rows[0], [COL_REWARD, COL_SETTLE, COL_PROFIT]].to_dict()
            raise ValueError(f"核验失败：净收益 ≠ 补贴金额 - 支出金额，受影响行索引：{invalid_rows}，示例行：{sample}")

    if COL_BILL_PERIOD in df.columns:
        df[COL_BILL_PERIOD] = pd.to_datetime(df[COL_BILL_PERIOD], errors='coerce')
    if COL_BI_PERIOD in df.columns:
        df[COL_BI_PERIOD] = datetime.strptime(LAST_MONTH, "%Y-%m").replace(day=1)
        print(f"   [提示] 已将 报表期 统一纠正为 {LAST_MONTH}")

    return df, errors_df

def save_excel_with_date_format(df, output_path, sheet_name):
    df_copy = df.copy()
    date_cols = [COL_BILL_PERIOD, COL_BI_PERIOD]
    for col in date_cols:
        if col in df_copy.columns:
            try:
                if pd.api.types.is_datetime64_any_dtype(df_copy[col]):
                    df_copy[col] = df_copy[col].dt.strftime("%Y-%m")
                else:
                    df_copy[col] = pd.to_datetime(df_copy[col], errors='coerce').dt.strftime("%Y-%m")
            except Exception as e:
                print(f"   [警告] 日期列 '{col}' 格式化失败：{e}")
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df_copy.to_excel(writer, sheet_name=sheet_name, index=False)
    except PermissionError:
        raise PermissionError(f"文件被占用，无法写入：{output_path}")
    except Exception as e:
        raise Exception(f"保存Excel文件失败：{e}")

def split_by_secondary_category(df_clean, output_dir, errors_df):
    print("\n开始按业务小类拆分...")
    if COL_BIZ_CAT2 not in df_clean.columns:
        raise ValueError("数据中缺少业务小类列，无法按分类拆分")

    if "活动名称" in df_clean.columns:
        empty_cat = df_clean[COL_BIZ_CAT2].isna()
        if empty_cat.any():
            df_clean.loc[empty_cat, COL_BIZ_CAT2] = df_clean.loc[empty_cat, "活动名称"]
            print(f"   [提示] 已将 {empty_cat.sum()} 行空业务小类替换为活动名称")

    non_empty = df_clean[COL_BIZ_CAT2].dropna()
    if len(non_empty) == 0:
        df_clean[COL_BIZ_CAT2] = "未分类"
        print("   [警告] 业务小类全为空，已全部设为'未分类'")
        non_empty = df_clean[COL_BIZ_CAT2]

    categories = non_empty.unique()
    os.makedirs(output_dir, exist_ok=True)
    output_cols = EXPECTED_HEADERS
    failed_categories = []

    for cat in categories:
        df_cat = df_clean[df_clean[COL_BIZ_CAT2] == cat].copy()
        for col in output_cols:
            if col not in df_cat.columns:
                df_cat[col] = pd.NA
        df_cat = df_cat[output_cols]

        bill_period = "未知期间"
        if COL_BILL_PERIOD in df_cat.columns and not df_cat[COL_BILL_PERIOD].isna().all():
            first_valid = df_cat[COL_BILL_PERIOD].dropna().iloc[0]
            try:
                if isinstance(first_valid, datetime):
                    bill_period = first_valid.strftime("%Y%m")
                else:
                    dt = pd.to_datetime(first_valid)
                    bill_period = dt.strftime("%Y%m")
            except:
                bill_period = "未知期间"

        safe_cat = str(cat).replace("/", "_").replace("\\", "_").replace(":", "_").replace("*", "_").replace("?", "_").replace("\"", "_").replace("<", "_").replace(">", "_").replace("|", "_")
        output_file = os.path.join(output_dir, f"{safe_cat}-{bill_period}.xlsx")
        try:
            save_excel_with_date_format(df_cat, output_file, "数据")
            print(f"  已拆分：{cat} (账期 {bill_period}) -> {output_file} ({len(df_cat)} 行)")
        except Exception as e:
            error_msg = f"拆分失败 {cat}：{e}"
            print(f"  [错误] {error_msg}")
            failed_categories.append((cat, str(e)))

    if failed_categories:
        print("\n以下分类拆分失败：")
        for cat, err in failed_categories:
            print(f"  - {cat}: {err}")

    if not errors_df.empty:
        err_path = os.path.join(output_dir, "组织层级校验错误.xlsx")
        try:
            errors_df.to_excel(err_path, index=False)
            print(f"  已输出组织层级校验错误至：{err_path}（共 {len(errors_df)} 行）")
        except Exception as e:
            print(f"  写入错误文件失败：{e}")

def generate_summary(output_dir, errors_df):
    processed_files = list(Path(output_dir).glob("*.xlsx"))
    processed_files = [f for f in processed_files if "组织层级校验错误" not in f.name]
    if not processed_files:
        print("未找到任何拆分后的文件，无法生成汇总清单。")
        return
    all_data = []
    invalid_bi_records = []
    for file_path in processed_files:
        try:
            df = pd.read_excel(file_path, dtype=object)
            required_cols = [COL_BIZ_CAT2, COL_BI_PERIOD, COL_BILL_PERIOD, COL_REWARD, COL_REWARD_BASE, COL_SETTLE, COL_PROFIT]
            if not all(col in df.columns for col in required_cols):
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
            print(f"\n⚠️ 报表期异常记录已导出至：{output_invalid_path}")
        except PermissionError:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            new_path = output_invalid_path.replace(".xlsx", f"_{timestamp}.xlsx")
            with pd.ExcelWriter(new_path, engine='openpyxl') as writer:
                invalid_all.to_excel(writer, sheet_name="异常记录", index=False)

    if not all_data:
        print("没有有效数据可汇总。")
        return

    combined = pd.concat(all_data, ignore_index=True)
    group_cols = [COL_BIZ_CAT2, "报表期_str", COL_BILL_PERIOD]
    summary = combined.groupby(group_cols).agg(
        **{
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

    summary = summary[[COL_BIZ_CAT2, COL_BI_PERIOD, COL_BILL_PERIOD, COL_REWARD, "补贴金额成本基数", COL_SETTLE, COL_PROFIT, "报表期状态"]]
    summary = summary.rename(columns={"补贴金额成本基数": COL_REWARD_BASE})
    summary["提醒信息"] = summary[COL_BIZ_CAT2].apply(
        lambda x: "需调整四级标签和增值划分" if x in REMINDER_CATEGORIES else ""
    )

    output_summary_path = os.path.join(output_dir, "汇总统计清单.xlsx")
    try:
        if os.path.exists(output_summary_path):
            os.remove(output_summary_path)
        with pd.ExcelWriter(output_summary_path, engine='openpyxl') as writer:
            summary.to_excel(writer, sheet_name="按业务小类_期间汇总", index=False)
            if not errors_df.empty:
                errors_df.to_excel(writer, sheet_name="组织层级校验错误", index=False)
        print(f"\n汇总统计清单已保存至：{output_summary_path}")
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        new_path = output_summary_path.replace(".xlsx", f"_{timestamp}.xlsx")
        with pd.ExcelWriter(new_path, engine='openpyxl') as writer:
            summary.to_excel(writer, sheet_name="按业务小类_期间汇总", index=False)
            if not errors_df.empty:
                errors_df.to_excel(writer, sheet_name="组织层级校验错误", index=False)
        print(f"\n汇总统计清单已另存至：{new_path}")

    if not errors_df.empty:
        print(f"\n⚠️ 组织层级校验错误 {len(errors_df)} 条，详见汇总文件。")

def merge_all_processed(output_dir):
    processed_files = list(Path(output_dir).glob("*.xlsx"))
    processed_files = [f for f in processed_files if "组织层级校验错误" not in f.name]
    if not processed_files:
        print("未找到任何拆分后的文件，无法合并。")
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

def main():
    print("项目2-系统导出 按业务小类拆分处理开始...")
    print(f"输入文件夹：{INPUT_FOLDER}")
    print(f"输出目录：{OUTPUT_DIR}")

    if not os.path.exists(INPUT_FOLDER):
        print(f"错误：文件夹不存在 {INPUT_FOLDER}")
        input("按回车键退出...")
        return

    input_file = find_input_file(INPUT_FOLDER)
    if input_file is None:
        print("错误：未找到任何 Excel 文件。")
        input("按回车键退出...")
        return
    print(f"原始文件：{input_file}")

    try:
        df_raw = pd.read_excel(input_file, dtype=object, keep_default_na=False)
        df_raw = rename_columns(df_raw)   # 关键：读入后立即脱敏
        df_raw = df_raw.map(lambda x: x.strip() if isinstance(x, str) else x)
        df_raw = df_raw.replace(r'^\s*$', pd.NA, regex=True)
        if df_raw.empty or len(df_raw) == 0:
            print("错误：原始数据为空。")
            input("按回车键退出...")
            return
        print(f"成功读取原始数据，共 {len(df_raw)} 行")
    except Exception as e:
        print(f"读取原始文件失败：{e}")
        input("按回车键退出...")
        return

    try:
        match_dict = load_match_table()
        print("成功加载活动名称匹配表")
    except Exception as e:
        print(f"加载匹配表失败：{e}")
        input("按回车键退出...")
        return

    try:
        df_std = convert_to_standard(df_raw, match_dict)
        print("格式转换完成")
    except Exception as e:
        print(f"格式转换失败：{e}")
        input("按回车键退出...")
        return

    try:
        df_clean, errors_df = check_and_clean_data(df_std.copy(), "系统导出")
        print("数据校验清洗通过")
    except Exception as e:
        print(f"校验清洗失败：{e}")
        input("按回车键退出...")
        return

    split_by_secondary_category(df_clean, OUTPUT_DIR, errors_df)
    generate_summary(OUTPUT_DIR, errors_df)
    merge_all_processed(OUTPUT_DIR)

    print("\n项目2-系统导出 全部处理完成！")
    input("按回车键退出...")

if __name__ == "__main__":
    main()