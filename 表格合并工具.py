import streamlit as st
import pandas as pd
from io import BytesIO

# ======================
# 清洗函数（完全不变）
# ======================
def clean_data(df):
    df = df.copy()
    if "姓名" in df.columns:
        df["姓名"] = df["姓名"].astype(str).str.strip().replace("nan", "")
    for col in df.columns:
        if any(x in col for x in ["学号", "考号", "准考证", "学籍", "编号"]):
            df[col] = df[col].astype(str).str.strip()
            df[col] = pd.to_numeric(df[col], errors="coerce")
            df[col] = df[col].apply(lambda x: str(int(x)) if pd.notna(x) else "")
    return df

# ======================
# 紧凑布局 + 主表支持选择Sheet
# ======================
def main():
    # 紧凑页面配置
    st.set_page_config(page_title="成绩合并工具", layout="wide", initial_sidebar_state="collapsed")
    st.markdown("<style>.block-container {padding-top: 1rem; padding-bottom: 0.5rem;}</style>", unsafe_allow_html=True)
    st.markdown("<style>div[data-testid='stExpander'] {padding: 0.5rem;}</style>", unsafe_allow_html=True)
    st.markdown("<style>.stDivider {margin: 0.5rem 0;}</style>", unsafe_allow_html=True)

    st.title("📊 成绩自动合并工具（紧凑交互版）")
    st.markdown("主表/附表均支持选择工作表，可视化配置匹配规则，一键合并导出")
    st.divider()

    # 会话状态
    if "merge_configs" not in st.session_state:
        st.session_state.merge_configs = []

    # --------------------------
    # 1. 主表（支持选择Sheet，紧凑排版）
    # --------------------------
    st.subheader("📂 1. 上传并选择主表")
    col_file, col_sheet = st.columns([3, 1])
    with col_file:
        main_file = st.file_uploader("主表Excel", type=["xlsx", "xls"], label_visibility="collapsed")
    main_df = None
    main_cols = []
    if main_file:
        try:
            excel_main = pd.ExcelFile(main_file)
            sheets_main = excel_main.sheet_names
            with col_sheet:
                st.markdown("")
                main_sheet = st.selectbox("工作表", sheets_main, label_visibility="collapsed")
            main_df = pd.read_excel(main_file, sheet_name=main_sheet)
            main_df = clean_data(main_df)
            main_cols = list(main_df.columns)
            st.success(f"主表就绪：{len(main_df)}行 | 字段：{', '.join(main_cols[:6])}{'...' if len(main_cols)>6 else ''}")
        except Exception as e:
            st.error(f"主表读取失败：{str(e)}")

    st.divider()

    # --------------------------
    # 2. 附表配置区（紧凑）
    # --------------------------
    st.subheader("⚙️ 2. 添加/配置合并附表")
    col_info, col_add = st.columns([4, 1])
    with col_info:
        st.caption("可添加多个附表，分别配置匹配键、成绩映射")
    with col_add:
        if st.button("➕ 添加附表", use_container_width=True):
            st.session_state.merge_configs.append({"file": None, "sheet": "", "match_map": {}, "cols_map": {}})
            st.rerun()

    # 逐个渲染附表配置（紧凑布局）
    for idx, cfg in enumerate(st.session_state.merge_configs):
        with st.expander(f"附表{idx+1}", expanded=False):
            col_del, col_file, col_sheet_sub = st.columns([1,3,1])
            with col_del:
                if st.button("🗑️", key=f"del_{idx}"):
                    st.session_state.merge_configs.pop(idx)
                    st.rerun()
            with col_file:
                sub_file = st.file_uploader("上传附表", type=["xlsx", "xls"], key=f"f_{idx}", label_visibility="collapsed")
            if sub_file:
                try:
                    excel_sub = pd.ExcelFile(sub_file)
                    sheets_sub = excel_sub.sheet_names
                    with col_sheet_sub:
                        st.markdown("")
                        sheet_sel = st.selectbox("工作表", sheets_sub, key=f"sh_{idx}", label_visibility="collapsed")
                    cfg["sheet"] = sheet_sel
                    cfg["file"] = sub_file
                    sub_df = pd.read_excel(sub_file, sheet_name=sheet_sel)
                    sub_cols = list(sub_df.columns)


                    col_1, col_2 = st.columns([1,1])
                    with col_1:

                        # 匹配字段
                        st.markdown("**匹配规则（主表→附表）**")
                        match_cnt = st.number_input("匹配数", min_value=1, max_value=5, value=2, key=f"mc_{idx}")
                        match_map = {}
                        for m in range(match_cnt):
                            c1, c2 = st.columns(2)
                            mk = c1.selectbox(f"主表{m+1}", main_cols if main_cols else ["先传主表"], key=f"mk_{idx}_{m}")
                            sk = c2.selectbox(f"附表{m+1}", sub_cols, key=f"sk_{idx}_{m}")
                            match_map[mk] = sk
                        cfg["match_map"] = match_map

                    with col_2:
                        # 成绩映射
                        st.markdown("**成绩映射（附表→新列名）**")
                        col_cnt = st.number_input("映射数", min_value=1, max_value=10, value=1, key=f"cc_{idx}")
                        cols_map = {}
                        for c in range(col_cnt):
                            c1, c2 = st.columns(2)
                            sc = c1.selectbox(f"附表字段{c+1}", sub_cols, key=f"sc_{idx}_{c}")
                            nc = c2.text_input(f"新列名{c+1}", value=sc, key=f"nc_{idx}_{c}")
                            cols_map[sc] = nc
                        cfg["cols_map"] = cols_map

                except Exception as e:
                    st.error(f"附表{idx+1}异常：{str(e)}")

    st.divider()

    # --------------------------
    # 3. 执行与结果区
    # --------------------------
    run_btn = st.button("🚀 一键合并导出", type="primary", use_container_width=True)
    st.divider()
    st.subheader("✅ 合并日志（不会覆盖）")
    log_area = st.empty()
    table_box = st.empty()
    down_box = st.empty()

    if run_btn:
        if main_df is None:
            st.error("请先上传并选择主表工作表！")
            return
        if len(st.session_state.merge_configs) == 0:
            st.warning("请至少添加一个附表！")
            return

        try:
            df = main_df.copy()
            log_lines = []
            log_lines.append("ℹ️ 开始合并处理...")
            log_area.write("\n".join(log_lines))

            for idx, cfg in enumerate(st.session_state.merge_configs):
                # 🔥 修复：跳过空附表、未上传文件的附表
                if cfg["file"] is None:
                    continue
                if not cfg["sheet"] or not cfg["match_map"] or not cfg["cols_map"]:
                    continue
                with st.spinner(f"合并附表{idx+1}"):
                    df_sub = pd.read_excel(cfg["file"], sheet_name=cfg["sheet"])
                    df_sub = clean_data(df_sub)

                    match_map = cfg["match_map"]
                    cols_map = cfg["cols_map"]

                    # 调换顺序：第一步仅提取附表需要用到的原始列
                    source_match_cols = list(match_map.values())
                    source_score_cols = list(cols_map.keys())
                    df_sub_temp = df_sub[source_match_cols + source_score_cols].copy()

                    # 第二步再执行匹配键重命名
                    df_sub_temp = df_sub_temp.rename(columns={v: k for k, v in match_map.items()})
                    match_keys = list(match_map.keys())

                    # 成绩字段映射改名
                    df_sub_temp.rename(columns=cols_map, inplace=True)

                    col_positions = {}
                    for col in cols_map.values():
                        if col in df.columns:
                            col_positions[col] = df.columns.get_loc(col)
                            df.drop(columns=[col], inplace=True)

                    # 左连接（现在绝对不会有 _x _y）
                    merged = pd.merge(df, df_sub_temp, on=match_keys, how="left")
                    # 【增强】把新列放回原来的位置
                    for col, pos in col_positions.items():
                        cols = list(merged.columns)
                        cols.remove(col)
                        cols.insert(pos, col)
                        merged = merged[cols]
                    df = merged

                    def make_key(row):
                        return tuple(str(row[k]).strip() if pd.notna(row[k]) else "" for k in match_keys)
                    exist_keys = set(df.apply(make_key, axis=1))
                    new_students = df_sub_temp[~df_sub_temp.apply(make_key, axis=1).isin(exist_keys)]
                    if not new_students.empty:
                        df = pd.concat([df, new_students], ignore_index=True)

                    log_lines.append(f"✅ 附表{idx+1}完成 | 新增：{len(new_students)}人")
                    log_area.write("\n".join(log_lines))

            log_lines.append(f"🎉 全部合并完成！总数据：{len(df)}行 * {len(df.columns)}列")
            log_lines.append(f"🎉 空值数量:{df.isnull().sum().sum()}， 重复行数: {df.duplicated().sum() if len(df) > 0 else 0}")
            log_area.write("\n".join(log_lines))

            table_box.dataframe(df, height=350, use_container_width=True)

            buffer = BytesIO()
            df.to_excel(buffer, index=False, engine="openpyxl")
            buffer.seek(0)
            down_box.download_button("📥 下载最终成绩_merge稳定版.xlsx", buffer,
                                     file_name="最终成绩_merge稳定版1.xlsx",
                                     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                     use_container_width=True)
        except Exception as e:
            st.error(f"合并异常：{str(e)}")

if __name__ == "__main__":
    main()
