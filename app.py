import streamlit as st
import pandas as pd
import datetime
from io import BytesIO

# ---------- 配置 ----------
st.set_page_config(page_title="EHS 协作平台", layout="wide")

DATA_FILE = "ehs_data.xlsx"

# 环境因素模板列（根据你的模板）
ENV_COLS = [
    "生命周期", "活动/产品/服务", "环境因素", "状态", "环境影响",
    "控制手段", "a", "b", "c", "d", "e", "评价", "x", "y", "评价.1", "SEA判定"
]
# 危险源模板列
HAZ_COLS = [
    "编号", "部门/工序", "活动/过程", "危险源", "状态", "涉及设备",
    "频度(hr/d)", "暴露人员（全天）", "可能事件/事故", "L", "E", "N", "C", "D",
    "现行控制方式", "重要性判定", "建议控制措施"
]


# ---------- 初始化数据文件 ----------
def init_data_file():
    env_df = pd.DataFrame(columns=ENV_COLS + ["部门", "最后修改时间"])
    haz_df = pd.DataFrame(columns=HAZ_COLS + ["最后修改时间"])
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
    """
    将编辑后的子集（某个部门的数据）合并回完整DataFrame。
    参数：
        original_subset: 编辑前该部门的数据（DataFrame）
        edited_subset: 编辑后该部门的数据（DataFrame）
        full_df: 完整的DataFrame（包含所有部门）
        department_col: 部门列名
    返回：
        新的完整DataFrame，以及被修改的行数（新增+修改+删除）
    """
    if original_subset.empty and edited_subset.empty:
        return full_df, 0
    # 获取当前部门
    if not edited_subset.empty and department_col in edited_subset.columns:
        dept = edited_subset.iloc[0][department_col] if len(edited_subset) > 0 else None
    elif not original_subset.empty and department_col in original_subset.columns:
        dept = original_subset.iloc[0][department_col]
    else:
        dept = None
    if dept is None:
        return full_df, 0

    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # 从完整数据中删除该部门的所有行
    full_df = full_df[full_df[department_col] != dept]
    # 为编辑后的数据添加最后修改时间（修改的行已经标记过，新增行也标记）
    if not edited_subset.empty:
        # 对于编辑后的子集，如果某行原本没有最后修改时间（新增行），则设置当前时间
        for idx in edited_subset.index:
            if pd.isna(edited_subset.at[idx, "最后修改时间"]) or edited_subset.at[idx, "最后修改时间"] == "":
                edited_subset.at[idx, "最后修改时间"] = now_str
        # 合并回完整数据
        full_df = pd.concat([full_df, edited_subset], ignore_index=True)
    # 计算修改数量：原本该部门的行数 + 新增行数 - 删除行数（简单计数）
    modified_count = len(edited_subset) + len(original_subset) - 2 * len(
        set(original_subset.index) & set(edited_subset.index))  # 粗略
    return full_df, modified_count


# ---------- 上传功能（提取模板列）----------
def upload_excel(uploaded_file, sheet_type):
    try:
        df = pd.read_excel(uploaded_file)
        if sheet_type == "环境因素":
            matched_cols = []
            for col in ENV_COLS:
                found = None
                for c in df.columns:
                    if str(c).strip() == col:
                        found = c
                        break
                if found:
                    matched_cols.append(found)
                else:
                    matched_cols.append(col)
            df = df[matched_cols] if matched_cols else pd.DataFrame(columns=ENV_COLS)
            df.columns = ENV_COLS
            if "部门" not in df.columns:
                df["部门"] = ""
            return df
        else:
            matched_cols = []
            for col in HAZ_COLS:
                found = None
                for c in df.columns:
                    if str(c).strip() == col:
                        found = c
                        break
                if found:
                    matched_cols.append(found)
                else:
                    matched_cols.append(col)
            df = df[matched_cols] if matched_cols else pd.DataFrame(columns=HAZ_COLS)
            df.columns = HAZ_COLS
            return df
    except Exception as e:
        st.error(f"读取文件失败：{e}")
        return None


# ---------- 侧边栏部门选项 ----------
def get_departments(env_df, haz_df):
    depts = set()
    if "部门" in env_df.columns:
        depts.update(env_df["部门"].dropna().unique())
    if "部门/工序" in haz_df.columns:
        depts.update(haz_df["部门/工序"].dropna().unique())
    # 也可以从模板中预设一些部门
    default_depts = ["生产部", "设备部", "安环部", "质量部", "仓储部", "行政部"]
    depts.update(default_depts)
    return sorted([d for d in depts if str(d) != ""])


# ---------- 页面布局 ----------
st.title("EHS 协作平台")

# 左侧边栏
with st.sidebar:
    page = st.radio("选择表格", ["环境因素识别表", "危险源识别表"])
    st.markdown("---")
    # 部门筛选
    env_df, haz_df = load_data()
    all_depts = get_departments(env_df, haz_df)
    selected_dept = st.selectbox("选择部门", ["全部"] + all_depts)
    st.caption("• 选择「全部」：只读查看所有部门数据\n• 选择具体部门：可编辑该部门数据")

# 根据选择的表格加载对应数据
if page == "环境因素识别表":
    df_full = env_df.copy()
    table_name = "环境因素"
    dept_col = "部门"
else:
    df_full = haz_df.copy()
    table_name = "危险源"
    dept_col = "部门/工序"

st.header(page)

# 根据部门筛选数据
if selected_dept == "全部":
    display_df = df_full.copy()
    editable = False
    st.info("当前为只读模式（全部部门），如需编辑请选择具体部门。")
else:
    # 筛选该部门数据
    if dept_col in df_full.columns:
        display_df = df_full[df_full[dept_col] == selected_dept].copy()
    else:
        display_df = pd.DataFrame(columns=df_full.columns)
    editable = True
    st.success(f"当前编辑模式：{selected_dept}，仅显示并允许修改本部门数据。")

# 显示表格（根据是否可编辑选择组件）
if editable:
    edited_df = st.data_editor(display_df, num_rows="dynamic", use_container_width=True, key=f"{page}_editor")
    # 保存按钮
    if st.button("💾 保存本部门修改", type="primary"):
        # 获取原始该部门的数据（在修改前）
        original_subset = display_df.copy()
        # 合并回总表
        new_full_df, modified_count = update_modified_rows(original_subset, edited_df, df_full, dept_col)
        # 更新全局变量并保存
        if page == "环境因素识别表":
            save_data(new_full_df, haz_df)
        else:
            save_data(env_df, new_full_df)
        st.success(f"保存成功！{selected_dept} 的数据已更新，共影响 {modified_count} 行。")
        st.rerun()
else:
    st.dataframe(display_df, use_container_width=True)

# 下载和上传（管理员功能）
st.markdown("---")
col1, col2 = st.columns(2)
with col1:
    # 下载当前表格（整个表，不是筛选后的）
    output = BytesIO()
    if page == "环境因素识别表":
        df_to_download = df_full
        sheet_name = "环境因素"
    else:
        df_to_download = df_full
        sheet_name = "危险源"
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_to_download.to_excel(writer, sheet_name=sheet_name, index=False)
    st.download_button(f"📥 下载完整{page}", data=output.getvalue(),
                       file_name=f"{page}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with col2:
    # 上传替换整个表格（仅提取模板列）
    uploaded_file = st.file_uploader(f"上传并替换{page}（仅提取模板列）", type=["xlsx", "xls"], key=f"upload_{page}")
    if uploaded_file:
        new_df = upload_excel(uploaded_file, table_name)
        if new_df is not None:
            # 添加最后修改时间列
            new_df["最后修改时间"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            if page == "环境因素识别表":
                save_data(new_df, haz_df)
            else:
                save_data(env_df, new_df)
            st.success(f"{page} 已替换！")
            st.rerun()

st.sidebar.markdown("---")
st.sidebar.caption("数据存储在云端Excel中，每次保存会更新修改行的「最后修改时间」列。")