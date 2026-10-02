"""
智能Excel操作助手 - MVP 前端
基于 Streamlit 实现，提供文件上传、数据预览、自然语言问答、对话历史等功能。
"""

import io
import re
import time
import pandas as pd
import streamlit as st
import duckdb

# ============================================================
# 页面配置
# ============================================================
st.set_page_config(
    page_title="智能Excel助手",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# 自定义样式
st.markdown(
    """
    <style>
    .main .block-container { padding-top: 2rem; padding-bottom: 2rem; }
    .stChatMessage { padding: 1rem; border-radius: 0.75rem; margin-bottom: 0.5rem; }
    .formula-box {
        background-color: #f0f7ff;
        border: 1px solid #b8d4ff;
        border-radius: 0.5rem;
        padding: 0.75rem 1rem;
        font-family: 'Courier New', monospace;
        font-size: 0.95rem;
        color: #1a56db;
        word-break: break-all;
    }
    .upload-hint { color: #6b7280; font-size: 0.875rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# 常量配置
# ============================================================
MAX_FILE_SIZE_MB = 20
ALLOWED_EXTENSIONS = ["xlsx", "xls"]
PREVIEW_ROWS = 10
TABLE_NAME = "excel_data"

# ============================================================
# Session State 初始化
# ============================================================
if "messages" not in st.session_state:
    st.session_state.messages = []
if "df" not in st.session_state:
    st.session_state.df = None
if "file_name" not in st.session_state:
    st.session_state.file_name = None
if "con" not in st.session_state:
    st.session_state.con = None


# ============================================================
# Mock 后端逻辑（后续替换为真实 LLM 调用）
# ============================================================
def mock_llm_generate(question: str, columns: list) -> dict:
    """
    模拟大模型生成 SQL 和 Excel 公式。
    返回: {"type": "query"|"formula"|"error", "sql": str, "formula": str, "explanation": str}
    """
    time.sleep(1.0)  # 模拟模型生成耗时

    question_lower = question.lower()

    # ---- 公式类问题 ----
    formula_patterns = [
        ("求和", "SUM"),
        ("平均", "AVERAGE"),
        ("最大值", "MAX"),
        ("最大", "MAX"),
        ("最小值", "MIN"),
        ("最小", "MIN"),
        ("计数", "COUNT"),
        ("个数", "COUNT"),
        ("vlookup", "VLOOKUP"),
        ("查找", "VLOOKUP"),
    ]
    for keyword, func_name in formula_patterns:
        if keyword in question_lower:
            col_letter = "B" if columns else "A"
            if columns:
                # 尝试匹配问题中提到的列名
                for col in columns:
                    if col in question:
                        col_idx = columns.index(col) + 1
                        col_letter = chr(64 + col_idx) if col_idx <= 26 else "A" + chr(64 + col_idx - 26)
                        break
            formula = f"={func_name}({col_letter}2:{col_letter}1000)"
            return {
                "type": "formula",
                "formula": formula,
                "explanation": f"该公式使用 {func_name} 函数计算 {col_letter} 列的对应结果，"
                               f"数据范围为第 2 行到第 1000 行（可根据实际数据量调整）。",
            }

    # ---- 查询类问题 ----
    numeric_cols = [c for c in columns if c and any(k in c.lower() for k in ["销售额", "销量", "金额", "价格", "数量", "收入", "利润", "cost", "sales", "amount", "price", "quantity"])]
    text_cols = [c for c in columns if c not in numeric_cols]

    if "前" in question and ("高" in question_lower or "大" in question_lower):
        sort_col = numeric_cols[0] if numeric_cols else columns[0]
        group_col = text_cols[0] if text_cols else columns[0]
        match = re.search(r"前\s*(\d+)", question)
        n = int(match.group(1)) if match else 5
        sql = (
            f'SELECT "{group_col}", SUM("{sort_col}") AS "{sort_col}_合计" '
            f'FROM {TABLE_NAME} '
            f'GROUP BY "{group_col}" '
            f'ORDER BY "{sort_col}_合计" DESC '
            f'LIMIT {n}'
        )
        return {"type": "query", "sql": sql, "explanation": f"按 {group_col} 分组，计算 {sort_col} 的总和并取前 {n} 名。"}

    if "后" in question and ("低" in question_lower or "小" in question_lower):
        sort_col = numeric_cols[0] if numeric_cols else columns[0]
        group_col = text_cols[0] if text_cols else columns[0]
        match = re.search(r"后\s*(\d+)", question)
        n = int(match.group(1)) if match else 5
        sql = (
            f'SELECT "{group_col}", SUM("{sort_col}") AS "{sort_col}_合计" '
            f'FROM {TABLE_NAME} '
            f'GROUP BY "{group_col}" '
            f'ORDER BY "{sort_col}_合计" ASC '
            f'LIMIT {n}'
        )
        return {"type": "query", "sql": sql, "explanation": f"按 {group_col} 分组，计算 {sort_col} 的总和并取后 {n} 名。"}

    if "总" in question_lower or "合计" in question or "总计" in question:
        if numeric_cols:
            agg_expr = ", ".join([f'SUM("{c}") AS "{c}_总计"' for c in numeric_cols])
            sql = f"SELECT {agg_expr} FROM {TABLE_NAME}"
            return {"type": "query", "sql": sql, "explanation": "计算所有数值列的总计。"}

    if "平均" in question_lower:
        if numeric_cols:
            agg_expr = ", ".join([f'AVG("{c}") AS "{c}_平均值"' for c in numeric_cols])
            sql = f"SELECT {agg_expr} FROM {TABLE_NAME}"
            return {"type": "query", "sql": sql, "explanation": "计算所有数值列的平均值。"}

    if "多少" in question or "几个" in question or "数量" in question_lower:
        if columns:
            group_col = text_cols[0] if text_cols else columns[0]
            sql = f'SELECT "{group_col}", COUNT(*) AS "数量" FROM {TABLE_NAME} GROUP BY "{group_col}" ORDER BY "数量" DESC LIMIT 10'
            return {"type": "query", "sql": sql, "explanation": f"按 {group_col} 分组统计数量。"}

    # 默认：返回全表前 10 行
    sql = f"SELECT * FROM {TABLE_NAME} LIMIT 10"
    return {"type": "query", "sql": sql, "explanation": "默认展示数据前 10 行，您可以尝试更具体的问题，如「销售额最高的5个产品」。"}


def execute_query(sql: str, con) -> pd.DataFrame:
    """执行 DuckDB 查询并返回结果 DataFrame。"""
    time.sleep(0.5)  # 模拟查询耗时
    result = con.execute(sql).fetchdf()
    return result


# ============================================================
# 侧边栏：文件上传
# ============================================================
with st.sidebar:
    st.header("📁 文件上传")
    st.markdown(
        f'<p class="upload-hint">支持 .xlsx / .xls 格式，最大 {MAX_FILE_SIZE_MB}MB</p>',
        unsafe_allow_html=True,
    )

    uploaded_file = st.file_uploader(
        "选择 Excel 文件",
        type=ALLOWED_EXTENSIONS,
        accept_multiple_files=False,
        help=f"上传您的 Excel 文件，系统将自动解析并支持自然语言查询。单文件不超过 {MAX_FILE_SIZE_MB}MB。",
    )

    if uploaded_file is not None:
        # 检查文件大小
        file_size_mb = uploaded_file.size / (1024 * 1024)
        if file_size_mb > MAX_FILE_SIZE_MB:
            st.error(f"❌ 文件过大（{file_size_mb:.1f}MB），请上传不超过 {MAX_FILE_SIZE_MB}MB 的文件。")
        else:
            try:
                with st.spinner("正在解析文件..."):
                    # 读取 Excel
                    df = pd.read_excel(uploaded_file)
                    st.session_state.df = df
                    st.session_state.file_name = uploaded_file.name

                    # 注册到 DuckDB
                    con = duckdb.connect(database=":memory:", read_only=False)
                    con.register(TABLE_NAME, df)
                    st.session_state.con = con

                    # 重置对话
                    st.session_state.messages = []

                st.success(f"✅ 文件上传成功！共 {len(df)} 行，{len(df.columns)} 列。")
            except Exception as e:
                st.error(f"❌ 文件解析失败：{str(e)}")
                st.session_state.df = None
                st.session_state.con = None

    st.divider()

    # 数据预览
    if st.session_state.df is not None:
        st.subheader("📋 数据预览")
        st.caption(f"文件：{st.session_state.file_name}")
        st.dataframe(
            st.session_state.df.head(PREVIEW_ROWS),
            use_container_width=True,
            hide_index=True,
            height=min(300, 50 + len(st.session_state.df.head(PREVIEW_ROWS)) * 35),
        )
        st.caption(f"显示前 {min(PREVIEW_ROWS, len(st.session_state.df))} 行，共 {len(st.session_state.df)} 行")

        # 列信息
        with st.expander("查看列信息", expanded=False):
            col_info = pd.DataFrame({
                "列名": st.session_state.df.columns,
                "类型": st.session_state.df.dtypes.astype(str).values,
                "非空值数": st.session_state.df.notna().sum().values,
            })
            st.dataframe(col_info, use_container_width=True, hide_index=True)

        if st.button("🔄 重新开始", type="secondary", use_container_width=True):
            st.session_state.messages = []
            st.session_state.df = None
            st.session_state.file_name = None
            st.session_state.con = None
            st.rerun()
    else:
        st.info("👆 请先上传 Excel 文件以开始使用。")
        st.markdown("### 💡 示例问题")
        st.markdown("- 销售额最高的5个产品")
        st.markdown("- 各地区销量合计")
        st.markdown("- 计算A列平均值的公式")
        st.markdown("- 总销售额是多少")


# ============================================================
# 主区域：对话界面
# ============================================================
st.title("📊 智能 Excel 助手")
st.caption("用自然语言查询 Excel 数据，无需记忆函数和 SQL")

# 欢迎语 / 未上传提示
if st.session_state.df is None:
    st.info("👈 请从左侧上传您的 Excel 文件，上传成功后即可开始提问。")
else:
    st.success(f"当前文件：**{st.session_state.file_name}**（{len(st.session_state.df)} 行 × {len(st.session_state.df.columns)} 列）")

st.divider()

# 对话历史展示
chat_container = st.container()

with chat_container:
    if not st.session_state.messages:
        if st.session_state.df is not None:
            st.chat_message("assistant").markdown(
                f"您好！我已成功读取 **{st.session_state.file_name}**。\n\n"
                f"您可以问我任何关于这份数据的问题，例如：\n"
                f"- 「销售额最高的5个产品」\n"
                f"- 「各地区销量合计」\n"
                f"- 「求平均值的公式是什么」\n\n"
                f"请在下方输入框中输入您的问题 👇"
            )
    else:
        for msg in st.session_state.messages:
            with st.chat_message(msg["role"]):
                # 用户消息
                if msg["role"] == "user":
                    st.markdown(msg["content"])
                # 助手消息
                else:
                    result = msg["content"]
                    if result.get("type") == "error":
                        st.error(f"⚠️ {result['message']}")
                        if result.get("detail"):
                            with st.expander("查看详细信息"):
                                st.code(result["detail"], language="text")
                    else:
                        if result.get("explanation"):
                            st.markdown(result["explanation"])

                        if result.get("type") == "query":
                            # 展示生成的 SQL
                            with st.expander("查看生成的 SQL", expanded=False):
                                st.code(result["sql"], language="sql")

                            # 展示查询结果表格
                            if "result_df" in result and result["result_df"] is not None:
                                result_df = result["result_df"]
                                st.markdown("**查询结果：**")
                                st.dataframe(
                                    result_df,
                                    use_container_width=True,
                                    hide_index=True,
                                    height=min(400, 50 + len(result_df) * 35),
                                )
                                st.caption(f"共 {len(result_df)} 行结果")

                        elif result.get("type") == "formula":
                            # 展示 Excel 公式
                            st.markdown("**生成的 Excel 公式：**")
                            formula = result["formula"]
                            st.markdown(
                                f'<div class="formula-box">{formula}</div>',
                                unsafe_allow_html=True,
                            )
                            # 复制按钮（使用 st.code + clipboard_copy 功能）
                            st.code(formula, language="excel")
                            st.caption("👆 点击代码块右上角的复制按钮即可复制公式")


# 输入框
if st.session_state.df is not None:
    if prompt := st.chat_input("输入您的问题，例如「销售额最高的5个产品」"):
        # 添加用户消息
        st.session_state.messages.append({"role": "user", "content": prompt})

        # 显示用户消息
        with st.chat_message("user"):
            st.markdown(prompt)

        # 助手回复
        with st.chat_message("assistant"):
            try:
                # 状态 1：模型生成中
                with st.status("🤖 模型正在生成...", expanded=False) as status:
                    columns = list(st.session_state.df.columns)
                    result = mock_llm_generate(prompt, columns)
                    status.update(label="✅ 模型生成完成", state="complete", expanded=False)

                if result["type"] == "query":
                    # 状态 2：查询执行中
                    with st.status("⚡ 查询执行中...", expanded=False) as status:
                        result_df = execute_query(result["sql"], st.session_state.con)
                        result["result_df"] = result_df
                        status.update(label="✅ 查询完成", state="complete", expanded=False)

                    # 展示结果
                    if result.get("explanation"):
                        st.markdown(result["explanation"])

                    with st.expander("查看生成的 SQL", expanded=False):
                        st.code(result["sql"], language="sql")

                    st.markdown("**查询结果：**")
                    st.dataframe(
                        result_df,
                        use_container_width=True,
                        hide_index=True,
                        height=min(400, 50 + len(result_df) * 35),
                    )
                    st.caption(f"共 {len(result_df)} 行结果")

                elif result["type"] == "formula":
                    if result.get("explanation"):
                        st.markdown(result["explanation"])

                    st.markdown("**生成的 Excel 公式：**")
                    formula = result["formula"]
                    st.markdown(
                        f'<div class="formula-box">{formula}</div>',
                        unsafe_allow_html=True,
                    )
                    st.code(formula, language="excel")
                    st.caption("👆 点击代码块右上角的复制按钮即可复制公式")

                # 保存到对话历史
                st.session_state.messages.append({"role": "assistant", "content": result})

            except Exception as e:
                error_msg = {
                    "type": "error",
                    "message": "抱歉，处理您的问题时出现了异常，请尝试换一种问法。",
                    "detail": str(e),
                }
                st.error(f"⚠️ {error_msg['message']}")
                with st.expander("查看详细信息"):
                    st.code(error_msg["detail"], language="text")
                st.session_state.messages.append({"role": "assistant", "content": error_msg})

        # 滚动到底部（通过 rerun 实现）
        # st.rerun()  # 可选：如需强制刷新可开启
else:
    st.chat_input("请先上传 Excel 文件后再提问", disabled=True)


# ============================================================
# 页脚
# ============================================================
st.divider()
st.caption("💡 提示：您的所有数据都在本地处理，不会上传到任何服务器。")
