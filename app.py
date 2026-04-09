import streamlit as st
import pandas as pd
import datetime
from io import BytesIO

# ---------- 配置 ----------
st.set_page_config(page_title="EHS 协作平台", layout="wide")

DATA_FILE = "ehs_data.xlsx"

# 系统内部使用的标准字段（环境因素）
ENV_STD_FIELDS = [
    "生命周期", "活动/产品/服务", "环境因素", "状态", "环境影响",
    "控制手段", "a", "b", "c", "d", "e", "评价", "x", "y", "评价.1", "SEA判定"
]
# 系统内部使用的标准字段（危险源）
HAZ_STD_FIELDS = [
    "编号", "部门/工序", "活动/过程", "危险源", "状态", "涉及设备",
    "频度(hr/d)", "暴露人员（全天）", "可能事件/事故", "L", "E", "N", "C", "D",
    "现行控制方式", "重要性判定", "建议控制措施"
]

# ---------- 初始化数据文件 ----------
def init_data_file():
    env_df = pd.DataFrame(columns=ENV_STD_FIELDS + ["部门", "最后修改时间"])
    haz_df = pd.DataFrame(columns=HAZ_STD_FIELDS + ["最后修改时间"])
    with pd.ExcelWriter(DATA_FILE, engine="openpyxl") as writer:
        env_df.to_excel(writer, sheet_name="环境因素", index=False)
        haz_df.to_excel(writer, sheet_name="危险源", index=False)
    return env_df, haz_df

def load_data():
    try:
        env = pd.read_excel(DATA_FILE, sheet_name="环境因素")
        haz = pd.read_excel(DATA_FILE, sheet_name="危险源")
    except:
        env, haz = init_data_file()
    if "最后修改时间" not in env.columns:
        env["最后修改时间"] = ""
    if "最后修改时间" not in haz.columns:
        haz["最后修改时间"] = ""
    return env, haz

def save_data(env_df, haz_df):
    with pd.ExcelWriter(DATA_FILE, engine="openpyxl") as writer:
        env_df.to_excel(writer, sheet_name="环境因素", index=False)
        haz_df.to_excel(writer, sheet_name="危险源", index=False)

def update_modified_rows(original_subset, edited_subset, full_df, department_col="部门"):
    if original_subset.empty and edited_subset.empty:
        return full_df, 0
    if not edited_subset.empty and department_col in edited_subset.columns:
        dept = edited_subset.iloc[0][department_col] if len(edited_subset) > 0 else None
    elif not original_subset.empty and department_col in original_subset.columns:
        dept = original_subset.iloc[0][department_col]
    else:
        dept = None
    if dept is None:
        return full_df, 0
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    full_df = full_df[full_df[department_col] != dept]
    if not edited_subset.empty:
        for idx in edited_subset.index:
            if pd.isna(edited_subset.at[idx, "最后修改时间"]) or edited_subset.at[idx, "最后修改时间"] == "":
                edited_subset.at[idx, "最后修改时间"] = now_str
        full_df = pd.concat([full_df, edited_subset], ignore_index=True)
    modified_count = len(edited_subset) + len(original_subset) - 2 * len(set(original_subset.index) & set(edited_subset.index))
    return full_df, modified_count

# ---------- 增强版上传函数：自动处理多行表头 ----------
def upload_environment_auto(uploaded_file):
    """专门处理环境因素的三行表头Excel"""
    try:
        df_raw = pd.read_excel(uploaded_file, header=None)
        header_row = df_raw.iloc[2]
        col_names = []
        last_val = None
        for val in header_row:
            if pd.notna(val):
                last_val = str(val).strip()
            col_names.append(last_val)
        new_cols = []
        count_map = {}
        for name in col_names:
            if name is None:
                name = "Unnamed"
            if name in count_map:
                count_map[name] += 1
                new_cols.append(f"{name}.{count_map[name]}")
            else:
                count_map[name] = 0
                new_cols.append(name)
        data_rows = df_raw.iloc[3:].copy()
        data_rows.columns = new_cols
        mapping = {}
        for std in ENV_STD_FIELDS:
            if std in data_rows.columns:
                mapping[std] = std
            else:
                for col in data_rows.columns:
                    if str(col).strip().lower() == std.lower():
                        mapping[std] = col
                        break
        result_df = pd.DataFrame()
        for std in ENV_STD_FIELDS:
            if std in mapping:
                result_df[std] = data_rows[mapping[std]]
            else:
                result_df[std] = ""
        if "部门" not in result_df.columns:
            result_df["部门"] = ""
        result_df["最后修改时间"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        result_df = result_df.dropna(how='all')
        return result_df
    except Exception as e:
        st.error(f"自动解析失败：{e}")
        return None

def upload_hazard_auto(uploaded_file):
    """处理危险源表（两行表头）"""
    try:
        df = pd.read_excel(uploaded_file, header=1)
        existing_cols = [c for c in HAZ_STD_FIELDS if c in df.columns]
        result_df = df[existing_cols].copy()
        for col in HAZ_STD_FIELDS:
            if col not in result_df.columns:
                result_df[col] = ""
        result_df["最后修改时间"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        return result_df
    except Exception as e:
        st.error(f"自动解析失败：{e}")
        return None

# ---------- 侧边栏部门选项 ----------
def get_departments(env_df, haz_df):
    depts = set()
    if "部门" in env_df.columns:
        depts.update(env_df["部门"].dropna().unique())
    if "部门/工序" in haz_df.columns:
        depts.update(haz_df["部门/工序"].dropna().unique())
    default_depts = ["生产部", "设备部", "安环部", "质量部", "仓储部", "行政部"]
    depts.update(default_depts)
    return sorted([d for d in depts if str(d) != ""])

# ---------- 页面布局 ----------
st.title("EHS 协作平台")

with st.sidebar:
    page = st.radio("选择表格", ["环境因素识别表", "危险源识别表"])
    st.markdown("---")
    env_df, haz_df = load_data()
    all_depts = get_departments(env_df, haz_df)
    selected_dept = st.selectbox("选择部门", ["全部"] + all_depts)
    st.caption("「全部」为只读查看，选具体部门可编辑")

if page == "环境因素识别表":
    df_full = env_df.copy()
    table_name = "环境因素"
    dept_col = "部门"
else:
    df_full = haz_df.copy()
    table_name = "危险源"
    dept_col = "部门/工序"

st.header(page)

if selected_dept == "全部":
    display_df = df_full.copy()
    editable = False
    st.info("只读模式（全部部门）")
else:
    if dept_col in df_full.columns:
        display_df = df_full[df_full[dept_col] == selected_dept].copy()
    else:
        display_df = pd.DataFrame(columns=df_full.columns)
    editable = True
    st.success(f"编辑模式：{selected_dept}")

if editable:
    edited_df = st.data_editor(display_df, num_rows="dynamic", use_container_width=True, key=f"{page}_editor")
    if st.button("保存本部门修改", type="primary"):
        original_subset = display_df.copy()
        new_full_df, modified_count = update_modified_rows(original_subset, edited_df, df_full, dept_col)
        if page == "环境因素识别表":
            save_data(new_full_df, haz_df)
        else:
            save_data(env_df, new_full_df)
        st.success(f"保存成功！影响 {modified_count} 行。")
        st.rerun()
else:
    st.dataframe(display_df, use_container_width=True)

st.markdown("---")
col1, col2 = st.columns(2)
with col1:
    output = BytesIO()
    if page == "环境因素识别表":
        df_to_download = df_full
        sheet_name = "环境因素"
    else:
        df_to_download = df_full
        sheet_name = "危险源"
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_to_download.to_excel(writer, sheet_name=sheet_name, index=False)
    st.download_button(f"下载完整{page}", data=output.getvalue(),
                       file_name=f"{page}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with col2:
    uploaded_file = st.file_uploader(f"上传并替换{page}（自动解析表头）", type=["xlsx", "xls"], key=f"upload_{page}")
    if uploaded_file:
        if page == "环境因素识别表":
            new_df = upload_environment_auto(uploaded_file)
        else:
            new_df = upload_hazard_auto(uploaded_file)
        if new_df is not None:
            if page == "环境因素识别表":
                save_data(new_df, haz_df)
            else:
                save_data(env_df, new_df)
            st.success(f"{page} 已成功替换！")
            st.rerun()

st.sidebar.markdown("---")
st.sidebar.caption("数据存储在云端Excel，每次保存会更新最后修改时间")
