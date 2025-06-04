# scraper.py
# -*- coding: utf-8 -*-
import os
import io
import re
import urllib.parse
from datetime import datetime

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

st.set_page_config(page_title="臺北馬拉松賽事爬蟲", layout="wide")


# 1. 取得檔案所在目錄，再拼出 MSJH.TTC 的路徑
FONT_PATH = os.path.join(os.path.dirname(__file__), "MSJH.TTC")

# 2. 嘗試註冊字型──若找不到檔案或發生其他錯誤，就用 warning 提示
try:
    pdfmetrics.registerFont(TTFont("MSJH-Light", FONT_PATH))
    pdfmetrics.registerFont(TTFont("MSJH-Bold", FONT_PATH))
    pdfmetrics.registerFontFamily(
        "MSJH",
        normal="MSJH-Light", bold="MSJH-Bold",
        italic="MSJH-Light", boldItalic="MSJH-Bold"
    )
except Exception as e:
    st.warning(
        f"⚠️ 註冊「微軟正黑體」字型時發生錯誤：{e}\n"
        "請確認 msJh.ttc 是否存在於系統可讀路徑。"
    )

# 建立 ParagraphStyle，供 PDF 內所有中文字格使用
chinese_style = ParagraphStyle(
    name="ChineseStyle",
    fontName="MSJH-Regular",
    fontSize=10,    # 後續依欄位數量動態調整
    leading=12,
    wordWrap="CJK",
    alignment=1
)

# ===========================
# 1. 模擬 ASP.NET PostBack 的函式
# ===========================
def get_filtered_soup(year: str, region: str, rtype: str):
    base_url = "http://www.taipeimarathon.org.tw/contest.aspx"
    session = requests.Session()

    # 1) 初次 GET
    try:
        r0 = session.get(base_url, timeout=10)
    except Exception as e:
        raise Exception(f"無法連到 {base_url}：{e}")
    if r0.status_code != 200:
        raise Exception(f"第一次 GET 時，HTTP 狀態碼 {r0.status_code}，非 200。")
    soup = BeautifulSoup(r0.text, "html.parser")

    # 2) 抓三個 <select> 的 name 屬性
    sel_year = soup.find("select", id="Year")
    sel_region = soup.find("select", id="DropDownList1")
    sel_type = soup.find("select", id="type")

    if sel_year is None:
        raise Exception("找不到 <select id='Year'>，可能網站架構改版。")
    if sel_region is None:
        raise Exception("找不到 <select id='DropDownList1'> (行政區)。")
    if sel_type is None:
        raise Exception("找不到 <select id='type'> (賽事類型)。")

    year_name = sel_year.get("name")
    region_name = sel_region.get("name")
    type_name = sel_type.get("name")
    if not year_name or not region_name or not type_name:
        raise Exception("無法擷取三個 <select> 的 name 屬性，可能 HTML 改版。")

    # 3) 拿 <option selected> 的預設值
    default_year_opt = sel_year.find("option", selected=True)
    default_region_opt = sel_region.find("option", selected=True)
    default_type_opt = sel_type.find("option", selected=True)
    if default_year_opt is None or default_region_opt is None or default_type_opt is None:
        raise Exception("某些 <select> 找不到被選中的 <option>。")
    default_year_val = default_year_opt.get("value", "")
    default_region_val = default_region_opt.get("value", "")
    default_type_val = default_type_opt.get("value", "")

    # Helper: extract hidden inputs
    def extract_hidden_inputs(soup_inner):
        form = soup_inner.find("form", id="aspnetForm")
        if form is None:
            form = soup_inner.find("form")
            if form is None:
                raise Exception("頁面上找不到任何 <form>，無法擷取隱藏欄位 (VIEWSTATE)。")
        data = {}
        for inp in form.find_all("input", {"type": "hidden"}):
            if inp.has_attr("name"):
                data[inp["name"]] = inp.get("value", "")
        return data

    # Helper: PostBack
    def postback(curr_soup, event_target_name: str, selected_values: dict):
        payload = extract_hidden_inputs(curr_soup)
        for k, v in selected_values.items():
            payload[k] = v
        payload["__EVENTTARGET"] = event_target_name
        payload["__EVENTARGUMENT"] = ""
        try:
            r = session.post(base_url, data=payload, timeout=10)
        except Exception as e:
            raise Exception(f"PostBack 時網路錯誤：{e}")
        if r.status_code != 200:
            raise Exception(f"PostBack (EVENTTARGET={event_target_name}) 時，HTTP {r.status_code}。")
        return BeautifulSoup(r.text, "html.parser")

    current_soup = soup

    # 4) 處理年份
    if year not in ["now", "all"]:
        values = {
            year_name: year,
            region_name: default_region_val,
            type_name: default_type_val
        }
        current_soup = postback(current_soup, year_name, values)

    # 5) 處理行政區
    selected_region_val = region if region != "all" else default_region_val
    if selected_region_val not in ["all", default_region_val]:
        values = {
            year_name: year if year not in ["now", "all"] else default_year_val,
            region_name: selected_region_val,
            type_name: default_type_val
        }
        current_soup = postback(current_soup, region_name, values)

    # 6) 處理賽事類型
    selected_type_val = rtype if rtype != "all" else default_type_val
    if selected_type_val not in ["all", default_type_val]:
        values = {
            year_name: year if year not in ["now", "all"] else default_year_val,
            region_name: selected_region_val,
            type_name: selected_type_val
        }
        current_soup = postback(current_soup, type_name, values)

    return current_soup


# ===========================
# 2. 解析 HTML table → DataFrame (含賽事名稱超連結)
# ===========================
def parse_table_to_df(soup):
    table = soup.find("table", id="ctl00_ContentPlaceHolder1_GridView1")
    if table is None:
        for tbl in soup.find_all("table"):
            header_cells = [th.get_text(strip=True) for th in tbl.find_all("th")]
            if "賽事名稱" in header_cells and "日期" in header_cells:
                table = tbl
                break
    if table is None:
        return pd.DataFrame()

    # 取 header row
    header_row = table.find("tr")
    raw_headers = [cell.get_text(strip=True) for cell in header_row.find_all(["th", "td"])]
    valid_indices = [i for i, h in enumerate(raw_headers) if h.strip() != ""]
    headers = [raw_headers[i] for i in valid_indices]

    # 找「賽事名稱」位置
    if "賽事名稱" not in headers:
        return pd.DataFrame()
    event_idx = headers.index("賽事名稱")

    data = []
    event_links = []

    for row in table.find_all("tr")[1:]:
        cells = row.find_all("td")
        if not cells:
            continue
        texts = [cell.get_text(separator=" ", strip=True) for cell in cells]
        if all(txt == "" for txt in texts):
            continue

        # 1) 文字部分
        filtered_texts = [texts[i] for i in valid_indices]
        data.append(filtered_texts)

        # 2) 賽事名稱超連結
        td_event = row.find_all("td")[valid_indices[event_idx]]
        a_tag = td_event.find("a", href=True)
        if a_tag:
            href = a_tag["href"]
            if href.startswith("http"):
                full_url = href
            else:
                full_url = urllib.parse.urljoin("http://www.taipeimarathon.org.tw/contest.aspx", href)
        else:
            full_url = ""
        event_links.append(full_url)

    df = pd.DataFrame(data, columns=headers)
    df["賽事連結"] = event_links

    # 解析「日期」
    def parse_event_date(event_name: str, date_str: str):
        year_val = datetime.now().year
        m_year = re.match(r"^\s*(\d{4})", event_name)
        if m_year:
            year_val = int(m_year.group(1))
        m2 = re.match(r"(\d{1,2})/(\d{1,2})", date_str)
        if not m2:
            return None
        month = int(m2.group(1))
        day = int(m2.group(2))
        try:
            return datetime(year=year_val, month=month, day=day)
        except ValueError:
            return None

    parsed_dates = []
    for _, r in df.iterrows():
        evn = r.get("賽事名稱", "")
        dt_txt = r.get("日期", "")
        parsed_dates.append(parse_event_date(evn, dt_txt))
    df["parsed_date"] = parsed_dates

    # 產生地圖連結
    def make_maps_link(addr: str) -> str:
        base = "https://www.google.com/maps/search/?api=1&query="
        return base + urllib.parse.quote(addr)

    def make_maps_embed(addr: str) -> str:
        base = "https://maps.google.com/maps?q="
        return base + urllib.parse.quote(addr) + "&output=embed"

    if "地點" in df.columns:
        df["地點連結"] = df["地點"].apply(make_maps_link)
        df["地點嵌入URL"] = df["地點"].apply(make_maps_embed)
    else:
        df["地點連結"] = ""
        df["地點嵌入URL"] = ""

    return df


# ===========================
# 3. Streamlit 介面
# ===========================
st.title("🏃‍♂️ 臺北馬拉松賽事查詢")

st.markdown(
    """
- **說明**：此 App 會自動向 [臺北馬拉松賽事列表](http://www.taipeimarathon.org.tw/contest.aspx) 進行爬蟲，  
  同時抓取「賽事名稱」原始連結，並能在前端和 PDF 生成可點擊連結。  
- 結果顯示在主畫面，並可立即選擇地點預覽地圖，或下載 PDF。  
"""
)

with st.sidebar:
    st.header("🔍 篩選條件")

    year_options = ["all", "now"] + [str(y) for y in range(2005, 2026)]
    year_display = {
        "all": "歷史賽事 (全部年份)",
        "now": "目前賽事",
        **{str(y): str(y) for y in range(2005, 2026)}
    }
    year_sel = st.selectbox("年份", options=year_options, format_func=lambda x: year_display[x])

    region_options = ["all", "北", "中", "南", "東", "其他"]
    region_display = {
        "all": "全部",
        "北": "北部",
        "中": "中部",
        "南": "南部",
        "東": "東部",
        "其他": "其他"
    }
    region_sel = st.selectbox("行政區", options=region_options, format_func=lambda x: region_display[x])

    type_options = ["public", "all", "1", "2", "3", "4", "5", "6", "7", "8"]
    type_display = {
        "public": "全國賽會",
        "all": "全部",
        "1": "超級馬拉松",
        "2": "馬拉松",
        "3": "半程馬拉松",
        "4": "10k~半馬",
        "5": "10k以下",
        "6": "休閒活動",
        "7": "鐵人賽",
        "8": "接力賽"
    }
    type_sel = st.selectbox("賽事類型", options=type_options, format_func=lambda x: type_display[x])

    st.markdown("---")
    keyword = st.text_input("關鍵字 (可空白)", placeholder="例如「臺南」")

    if st.button("開始查詢"):
        with st.spinner("資料擷取中，請稍候…"):
            try:
                final_soup = get_filtered_soup(year_sel, region_sel, type_sel)
                if final_soup is None:
                    st.error("❌ 取得篩選後的頁面回傳 None。請稍後再試。")
                else:
                    df = parse_table_to_df(final_soup)
                    if df.empty:
                        st.warning("⚠️ 查無任何賽事。請檢查篩選條件或改成「目前賽事」。")
                    else:
                        if keyword.strip() != "":
                            mask = df.apply(lambda r: r.astype(str).str.contains(keyword).any(), axis=1)
                            df = df[mask]
                        st.session_state["df"] = df
            except Exception as e:
                st.error(f"❌ 取得或解析時發生錯誤：{e}")

# 主畫面：顯示查詢結果
if "df" in st.session_state and not st.session_state["df"].empty:
    df_all = st.session_state["df"]

    # 1. 先排序，不要動原始 session_state["df"]
    df_sorted = df_all.sort_values(by=["parsed_date"], ascending=True, na_position="last")

    # 2. 建立「地點 → 地點嵌入URL」的映射，供地圖預覽使用
    location_to_embed = dict(zip(df_sorted["地點"], df_sorted["地點嵌入URL"]))

    # 3. 接著將 display_df 設為排序後的 DataFrame，但把 parsed_date 與 地點嵌入URL 兩欄一起 drop
    display_df = df_sorted.drop(columns=["parsed_date", "地點嵌入URL"])

    st.subheader("📋 查詢結果")
    st.dataframe(display_df, use_container_width=True)

    st.markdown("---")
    # 4. 地點預覽：因為我們已經從 df_sorted 建了 location_to_embed，這裡改從映射取 url
    st.subheader("🗺️ 地點預覽")
    addr_list = display_df["地點"].unique().tolist()
    sel_addr = st.selectbox("選擇地點", options=addr_list, key="addr_selector")
    
    # 從映射表取出對應的地圖嵌入 URL
    embed_url = location_to_embed.get(sel_addr, "")
    if embed_url:
        st.components.v1.iframe(embed_url, width=700, height=450)
    else:
        st.warning("❌ 無法取得該地點的嵌入 URL。")

    st.markdown("---")
    st.subheader("📥 下載 PDF")

    # 5. 以下處理 PDF 生成：data_for_pdf 只需考慮 display_df，已經不含「地點嵌入URL」
    headers = list(display_df.columns)
    data_for_pdf = []

    # 5.a 處理表頭：純文字 + 灰底
    header_row = []
    for col in headers:
        header_para = Paragraph(
            col,  # 純文字
            ParagraphStyle(
                name="HeaderStyle",
                fontName="MSJH-Regular",       # 使用微軟正黑體家族
                fontSize=10,
                leading=12,
                alignment=1,           # 置中
                textColor=colors.black,
                backColor=colors.HexColor("#D3D3D3")
            )
        )
        header_row.append(header_para)
    data_for_pdf.append(header_row)

    # 5.b 處理表身：針對「賽事連結」與「地點連結」做超連結，其它純文字
    for row in display_df.itertuples(index=False):
        row_cells = []
        for i, value in enumerate(row):
            col_name = headers[i]

            if col_name == "賽事連結":
                url = value  # 這裡的 value 就是原始報名頁面 URL
                if url.strip():
                    cell_para = Paragraph(
                        f'<link href="{url}">連結</link>',
                        chinese_style
                    )
                else:
                    cell_para = Paragraph("", chinese_style)

            elif col_name == "地點連結":
                url = value  # 這裡的 value 就是 Google Map 的 URL
                if url.strip():
                    cell_para = Paragraph(
                        f'<link href="{url}">地圖連結</link>',
                        chinese_style
                    )
                else:
                    cell_para = Paragraph("", chinese_style)

            else:
                # 其餘欄位都用純文字顯示
                cell_para = Paragraph(str(value), chinese_style)

            row_cells.append(cell_para)
        data_for_pdf.append(row_cells)

    # 6. 建立 PDF：A4 橫向
    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=landscape(A4),
        leftMargin=20, rightMargin=20, topMargin=20, bottomMargin=20
    )

    # 7. 動態計算字型與欄寬
    num_cols = len(headers)
    page_width = landscape(A4)[0] - doc.leftMargin - doc.rightMargin
    col_width = page_width / num_cols

    if col_width < 50:
        target_font_size = 6
    elif col_width < 80:
        target_font_size = 8
    else:
        target_font_size = 10

    chinese_style.fontSize = target_font_size
    chinese_style.leading = target_font_size + 2

    # 8. 建立 Table 並設定欄寬
    table = Table(data_for_pdf, colWidths=[col_width] * num_cols)

    # 9. 設定 TableStyle：網格線、置中、表頭 padding
    table_style = TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
        ("TOPPADDING", (0, 0), (-1, 0), 6),
    ])
    table.setStyle(table_style)

    # 10. Build PDF
    doc.build([table])

    # 11. 取得 PDF bytes
    pdf_buffer.seek(0)
    pdf_bytes = pdf_buffer.read()
    pdf_buffer.close()

    # 12. 組檔名：「年份_行政區_賽事類型[_關鍵字].pdf」
    fn_year = year_sel if year_sel not in ["now", "all"] else "allYears"
    fn_region = region_sel
    fn_type = type_sel
    fn_keyword = keyword.strip().replace(" ", "_") if keyword.strip() != "" else ""
    filename_parts = [fn_year, fn_region, fn_type]
    if fn_keyword:
        filename_parts.append(fn_keyword)
    pdf_filename = "_".join(filename_parts) + ".pdf"

    # 13. 提供下載按鈕
    st.download_button(
        label="⬇️ 下載查詢結果 (PDF)",
        data=pdf_bytes,
        file_name=pdf_filename,
        mime="application/pdf"
    )