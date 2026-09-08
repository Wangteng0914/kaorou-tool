"""
南京烤肉烧烤店铺数据处理工具
本地运行：streamlit run kaorou_tool.py
"""
import streamlit as st
import pandas as pd
import re
import io
import requests
from pypinyin import lazy_pinyin
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from urllib.parse import quote

st.set_page_config(page_title="烤肉店铺数据处理工具", page_icon="🍖", layout="wide")

# ============ 样式定义 ============
HEADER_FONT_WHITE = Font(bold=True, color="FFFFFF", size=11)
HEADER_FONT_DARK = Font(bold=True, color="333333", size=11)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
THIN_BORDER = Border(
    left=Side(style="thin", color="DDDDDD"),
    right=Side(style="thin", color="DDDDDD"),
    top=Side(style="thin", color="DDDDDD"),
    bottom=Side(style="thin", color="DDDDDD"),
)
ALT_FILL = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")

COLUMNS = [
    ("所属区域", 100),
    ("店铺名称", 300),
    ("联系电话", 150),
    ("大众点评", 150),
    ("营业状态", 200),
    ("联系地址", 400),
]

SHEET_STYLES = {
    "烤肉店": ("FF6B6B", HEADER_FONT_WHITE),
    "烧烤店": ("4ECDC4", HEADER_FONT_WHITE),
    "其他店": ("FFE66D", HEADER_FONT_DARK),
}


# ============ 工具函数 ============
def extract_region(filename):
    """从文件名提取区县名"""
    m = re.search(r"南京市(.+?)[区内]", filename)
    if m:
        return m.group(1) + "区"
    return "未知"


def is_mobile(phone):
    """判断是否11位手机号"""
    phone = str(phone).strip()
    return bool(re.match(r"^1\d{10}$", phone))


def classify_store(name):
    """分类：烤肉店/烧烤店/其他店"""
    if "烤肉" in name:
        return "烤肉店"
    elif "烧烤" in name:
        return "烧烤店"
    else:
        return "其他店"


def sort_key(row):
    """排序：数字开头优先，然后拼音A-Z，同名品牌连续"""
    name = str(row["店铺名称"])
    # 提取品牌名（去掉括号分店后缀，用于同名品牌连续）
    brand = re.sub(r"[（(].*?[）)]", "", name).strip()
    brand_pinyin = "".join(lazy_pinyin(brand))
    # 数字开头标记
    is_number_start = 1 if name and name[0].isdigit() else 0
    # 英文开头标记
    is_english_start = 1 if name and name[0].encode("utf-8").isalpha() else 0
    full_pinyin = "".join(lazy_pinyin(name))
    return (0 if is_number_start else 1 if is_english_start else 2, brand_pinyin, full_pinyin)


def dianping_url(name):
    """大众点评搜索链接（南京 city_id=5）"""
    return f"https://www.dianping.com/search/keyword/5/0_{quote(name)}"


def build_excel(data_dict, status_map=None):
    """生成带样式的Excel"""
    wb = Workbook()
    wb.remove(wb.active)

    for sheet_name, df in data_dict.items():
        ws = wb.create_sheet(sheet_name)
        header_color, header_font = SHEET_STYLES[sheet_name]
        count = len(df)

        # Row1: 分类总数
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLUMNS))
        c = ws.cell(row=1, column=1, value=f"{sheet_name}共{count}家")
        c.font = Font(bold=True, size=12, color="333333")
        c.alignment = CENTER
        c.fill = PatternFill(start_color="F0F0F0", end_color="F0F0F0", fill_type="solid")

        # Row2: 表头
        hfill = PatternFill(start_color=header_color, end_color=header_color, fill_type="solid")
        for ci, (col_name, width) in enumerate(COLUMNS, 1):
            c = ws.cell(row=2, column=ci, value=col_name)
            c.font = header_font
            c.alignment = CENTER
            c.fill = hfill
            c.border = THIN_BORDER
            ws.column_dimensions[get_column_letter(ci)].width = width / 7

        # Row3+: 数据
        for ri, (_, row) in enumerate(df.iterrows(), 3):
            is_alt = (ri - 3) % 2 == 1
            for ci, (col_name, _) in enumerate(COLUMNS, 1):
                if col_name == "大众点评":
                    url = dianping_url(row["店铺名称"])
                    c = ws.cell(row=ri, column=ci)
                    c.value = f'=HYPERLINK("{url}","点击查看")'
                    c.font = Font(color="0563C1", underline="single")
                elif col_name == "营业状态":
                    val = status_map.get(row["店铺名称"], "") if status_map else ""
                    c = ws.cell(row=ri, column=ci, value=val)
                elif col_name == "联系电话":
                    c = ws.cell(row=ri, column=ci, value=str(row["联系电话"]))
                    c.number_format = "@"
                else:
                    c = ws.cell(row=ri, column=ci, value=row[col_name])

                c.alignment = LEFT if col_name in ["店铺名称", "联系地址"] else CENTER
                c.border = THIN_BORDER
                if is_alt:
                    c.fill = ALT_FILL

        ws.freeze_panes = "A3"
        ws.row_dimensions[1].height = 28
        ws.row_dimensions[2].height = 24

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def verify_amap_batch(df, api_key, progress_bar=None, status_text=None):
    """高德POI核验（仅烤肉店）"""
    status_map = {}
    total = len(df)
    for idx, (_, row) in enumerate(df.iterrows()):
        name = row["店铺名称"]
        address = row["联系地址"]
        try:
            # text search
            url = "https://restapi.amap.com/v3/place/text"
            params = {
                "key": api_key,
                "keywords": name,
                "city": "南京",
                "citylimit": "true",
                "extensions": "all",
                "offset": 10,
                "page": 1,
            }
            resp = requests.get(url, params=params, timeout=10)
            data = resp.json()

            if data.get("status") != "1" or not data.get("pois"):
                status_map[name] = "❌已闭店"
                continue

            # 精准匹配：店名完全一致或地址匹配
            matched = None
            for poi in data["pois"]:
                poi_name = poi.get("name", "")
                poi_addr = poi.get("address", "")
                if poi_name == name or name in poi_name or poi_addr in address:
                    matched = poi
                    break
            if not matched:
                matched = data["pois"][0]

            poi_name = matched.get("name", "")
            # 检查暂停营业标记
            if any(kw in poi_name for kw in ["暂停营业", "停业", "装修中", "已关闭"]):
                status_map[name] = "⚠️暂停营业"
                continue

            # 检查营业时间和评分
            biz_ext = matched.get("biz_ext", {})
            opentime = matched.get("opentime2", "") or matched.get("open_time", "") or biz_ext.get("opentime", "")
            rating = float(matched.get("rating", 0) or biz_ext.get("rating", 0) or 0)

            if opentime and rating >= 3.0:
                status_map[name] = "✅正常营业"
            elif not opentime and rating <= 2.5:
                status_map[name] = "⚠️暂停营业"
            else:
                status_map[name] = "✅正常营业"

        except Exception as e:
            status_map[name] = f"❓核验失败({str(e)[:20]})"

        if progress_bar:
            progress_bar.progress((idx + 1) / total)
        if status_text:
            status_text.text(f"核验中 {idx+1}/{total}：{name} → {status_map.get(name, '')}")

    return status_map


# ============ 主界面 ============
st.title("🍖 南京烤肉烧烤店铺数据处理工具")
st.caption("上传区县Excel → 自动清洗/分类/排序 → 高德POI核验 → 导出飞书表格")

# ---- 侧边栏 ----
with st.sidebar:
    st.header("⚙️ 设置")
    uploaded_files = st.file_uploader(
        "上传区县Excel文件（可多选）",
        type=["xls", "xlsx"],
        accept_multiple_files=True,
    )
    amap_key = st.text_input("高德Web-Service API Key（选填，用于营业状态核验）", type="password")
    verify_scope = st.radio("核验范围", ["仅烤肉店", "全部店铺", "不核验"], index=0)
    st.divider()
    st.caption("数据仅在本地处理，不上传任何服务器（高德API调用除外）")

# ---- 主流程 ----
if not uploaded_files:
    st.info("👈 请在左侧上传区县Excel文件开始处理")
    st.stop()

# Step1: 合并清洗
with st.status("📥 步骤1：合并清洗数据", expanded=True) as status:
    all_rows = []
    for f in uploaded_files:
        try:
            df = pd.read_excel(f, dtype=str)
            region = extract_region(f.name)
            df["所属区域"] = region
            all_rows.append(df)
            st.write(f"  ✅ {f.name} → {region}，{len(df)}行")
        except Exception as e:
            st.error(f"  ❌ {f.name} 读取失败：{e}")

    if not all_rows:
        st.error("没有有效数据")
        st.stop()

    merged = pd.concat(all_rows, ignore_index=True)
    st.write(f"合并后：{len(merged)}行")

    # 过滤11位手机号
    merged["联系电话"] = merged["联系电话"].astype(str).str.strip()
    mobile_mask = merged["联系电话"].apply(is_mobile)
    before = len(merged)
    merged = merged[mobile_mask].copy()
    st.write(f"过滤非手机号：{before} → {len(merged)}行（剔除{before - len(merged)}行）")

    # 去重
    before = len(merged)
    merged = merged.drop_duplicates(subset=["店铺名称", "联系电话"]).copy()
    st.write(f"去重（店名+电话）：{before} → {len(merged)}行")

    # 保留字段
    merged = merged[["所属区域", "店铺名称", "联系电话", "联系地址"]].copy()
    st.write(f"**清洗完成：共{len(merged)}家有效店铺**")
    status.update(label=f"✅ 步骤1完成：{len(merged)}家", state="complete")

# Step2: 分类
with st.status("📂 步骤2：店铺分类", expanded=True) as status:
    merged["分类"] = merged["店铺名称"].apply(classify_store)
    cat_counts = merged["分类"].value_counts()
    col1, col2, col3 = st.columns(3)
    col1.metric("烤肉店", cat_counts.get("烤肉店", 0))
    col2.metric("烧烤店", cat_counts.get("烧烤店", 0))
    col3.metric("其他店", cat_counts.get("其他店", 0))
    status.update(label="✅ 步骤2完成", state="complete")

# Step3: 排序
with st.status("🔤 步骤3：拼音排序（数字开头→拼音A-Z，同名品牌连续）", expanded=True) as status:
    sorted_dfs = {}
    for cat in ["烤肉店", "烧烤店", "其他店"]:
        df_cat = merged[merged["分类"] == cat].copy()
        df_cat = df_cat.sort_values(by="店铺名称", key=lambda s: s.map(
            lambda x: sort_key({"店铺名称": x})
        )).reset_index(drop=True)
        sorted_dfs[cat] = df_cat
        st.write(f"  {cat}：{len(df_cat)}家（排序后）")
    status.update(label="✅ 步骤3完成", state="complete")

# Step4: 高德核验（可选）
status_map = None
if verify_scope != "不核验" and amap_key:
    with st.status("🗺️ 步骤4：高德POI营业状态核验", expanded=True) as status:
        verify_df = sorted_dfs["烤肉店"] if verify_scope == "仅烤肉店" else merged
        st.write(f"核验范围：{verify_scope}，共{len(verify_df)}家")
        progress_bar = st.progress(0)
        status_text = st.empty()
        status_map = verify_amap_batch(verify_df, amap_key, progress_bar, status_text)

        # 统计
        status_counts = pd.Series(list(status_map.values())).value_counts()
        st.write("**核验结果统计：**")
        for s, c in status_counts.items():
            st.write(f"  {s}：{c}家")

        # 异常店铺列表
        abnormal = {k: v for k, v in status_map.items() if v.startswith("⚠️") or v.startswith("❌")}
        if abnormal:
            st.warning(f"发现{len(abnormal)}家异常店铺：")
            for name, s in abnormal.items():
                st.write(f"  {s} | {name}")
        status.update(label=f"✅ 步骤4完成：{len(status_map)}家已核验", state="complete")
elif verify_scope != "不核验" and not amap_key:
    st.warning("⚠️ 未输入高德API Key，跳过营业状态核验。如需核验请在左侧输入Key。")

# Step5: 生成Excel并下载
with st.status("📊 步骤5：生成飞书表格", expanded=True) as status:
    excel_buf = build_excel(sorted_dfs, status_map)
    status.update(label="✅ Excel生成完成", state="complete")

st.success("🎉 全部处理完成！")

# 下载按钮
st.download_button(
    label="📥 下载Excel文件（导入飞书表格使用）",
    data=excel_buf,
    file_name="南京烤肉烧烤店铺汇总.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

# 数据预览
st.divider()
st.subheader("📋 数据预览")
tab1, tab2, tab3 = st.tabs([f"烤肉店({len(sorted_dfs['烤肉店'])})", f"烧烤店({len(sorted_dfs['烧烤店'])})", f"其他店({len(sorted_dfs['其他店'])})"])
with tab1:
    display_df = sorted_dfs["烤肉店"].copy()
    if status_map:
        display_df["营业状态"] = display_df["店铺名称"].map(status_map).fillna("")
    st.dataframe(display_df, use_container_width=True, height=400)
with tab2:
    display_df = sorted_dfs["烧烤店"].copy()
    if status_map:
        display_df["营业状态"] = display_df["店铺名称"].map(status_map).fillna("")
    st.dataframe(display_df, use_container_width=True, height=400)
with tab3:
    display_df = sorted_dfs["其他店"].copy()
    if status_map:
        display_df["营业状态"] = display_df["店铺名称"].map(status_map).fillna("")
    st.dataframe(display_df, use_container_width=True, height=400)

# 使用说明
with st.expander("📖 使用说明"):
    st.markdown("""
    **快速上手：**
    1. 左侧上传所有区县Excel文件（支持.xls/.xlsx，可多选）
    2. 输入高德Web-Service API Key（可选，用于营业状态核验）
    3. 选择核验范围（仅烤肉店/全部/不核验）
    4. 工具自动完成：合并→过滤手机号→去重→分类→拼音排序→高德核验
    5. 点击"下载Excel文件"，然后在飞书中导入该Excel即可

    **导入飞书：**
    - 飞书表格 → 导入 → 选择下载的.xlsx文件
    - 导入后样式（表头配色、冻结、超链接、手机号文本格式）自动保留

    **高德API Key获取：**
    - 访问 https://console.amap.com/ → 应用管理 → 创建应用 → 添加Key（服务平台选"Web服务"）
    - 免费额度：每日30万次调用，足够处理数千家店铺
    """)
