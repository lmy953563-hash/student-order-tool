import streamlit as st
import pandas as pd
import io
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="学生订购数据", layout="wide")
st.title("📊 学生订购数据智能分析看板")

# 导出样式函数
def apply_excel_style(ws):
    # 列宽设置：前两列 3.5cm (约 13-14 字符宽度)，第三列 7cm (约 26-27 字符宽度)
    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 26.5 
    
    for row in ws.iter_rows():
        ws.row_dimensions[row[0].row].height = 20
        for cell in row:
            cell.font = Font(name='宋体', size=12)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            cell.fill = PatternFill(start_color='E6F0FF' if cell.row == 1 else 'FDF5E6', fill_type='solid')

uploaded_file = st.sidebar.file_uploader("请上传 Excel 文件", type=["xlsx"])

if uploaded_file is not None:
    try:
        # 数据处理
        df = pd.read_excel(uploaded_file, sheet_name="学生订购信息汇总表")
        col_mapping = {c: str(c).strip().replace('\ufeff', '') for c in df.columns}
        df = df.rename(columns=col_mapping).rename(columns={c: '姓名' for c in df.columns if '姓名' in c})
        df = df.rename(columns={c: '班级' for c in df.columns if '班级' in c})
        df = df.rename(columns={c: '学科' for c in df.columns if '学科' in c})
        
        df_clean = df.dropna(subset=['姓名', '班级', '学科'])
        df_clean['学科'] = df_clean['学科'].astype(str)
        
        # 1. 核心看板
        st.subheader("📈 核心统计数据")
        df_exploded = df_clean.copy()
        df_exploded['学科'] = df_exploded['学科'].str.split('/')
        df_exploded = df_exploded.explode('学科')
        df_exploded['学科'] = df_exploded['学科'].str.strip()
        
        subj_totals = df_exploded.groupby('学科').size()
        cols = st.columns(2 + len(subj_totals))
        cols[0].metric("订购学生总人数", f"{len(df_clean['姓名'].unique())} 人")
        cols[1].metric("学科总订购数量", f"{len(df_exploded)} 科次")
        for i, (subj, count) in enumerate(subj_totals.items()):
            cols[2 + i].metric(f"{subj} 总数", f"{count} 科")
        
        st.markdown("---")

        # 2. 订购信息
        st.subheader("🔍 订购信息")
        all_classes = sorted(df_clean['班级'].unique())
        selected_class = st.selectbox("🎯 按班级筛选显示：", all_classes)
        
        display_df = df_clean[df_clean['班级'] == selected_class].groupby(['班级', '姓名'])['学科'].first().reset_index()
        st.dataframe(display_df, use_container_width=True, hide_index=True)

        # 3. 整合导出
        col1, col2 = st.columns([1, 4])
        if col1.button("导出选中班级"):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                display_df.to_excel(writer, index=False, sheet_name="导出")
                apply_excel_style(writer.sheets["导出"])
            st.download_button("下载 Excel", output.getvalue(), f"{selected_class}_导出.xlsx")
            
        if col2.button("一键导出所有班级数据"):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for cls in all_classes:
                    cls_data = df_clean[df_clean['班级'] == cls].groupby(['班级', '姓名'])['学科'].first().reset_index()
                    cls_data.to_excel(writer, index=False, sheet_name=str(cls))
                    apply_excel_style(writer.sheets[str(cls)])
            st.download_button("下载全部班级汇总", output.getvalue(), "全部班级导出.xlsx")

        st.markdown("---")
        st.subheader("📊 各班级学科统计表")
        pivot_df = df_exploded.groupby(['班级', '学科']).size().unstack(fill_value=0)
        pivot_df['总计'] = pivot_df.sum(axis=1)
        st.dataframe(pivot_df, use_container_width=True)

    except Exception as e:
        st.error(f"❌ 读取错误: {e}")
