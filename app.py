# -*- coding: utf-8 -*-
"""
CIC LOGIS EXPORT CONSOLE (모바일 웹/PC 겸용 수출팀 일일 정산 및 회계 자동 검산 시스템)
"""

import os
import re
import csv
import io
import json
import datetime
import urllib.request
import urllib.error
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# PDF 생성 라이브러리 연동
try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib import colors
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False

st.set_page_config(page_title="CIC LOGIS 정산 콘솔", page_icon="📊", layout="wide")

THEME_COLOR = "#0F172A"

COMPANY_INFO = {
    "NAME_EN": "CIC LOGIS CO.,LTD.",
    "NAME_KR": "(주)씨아이씨로지스틱",
    "ADDR_EN": "D4-block, 95, Jayumuyeok-ro, Jung-gu, Incheon, Korea",
    "ADDR_KR": "인천광역시 중구 자유무역로 95 D4블록 301호",
    "TEL": "+82-32-744-4221~6",
    "FAX": "+82-70-4015-5833",
    "EMAIL": "ciclogis@ciclogis.co.kr"
}

def get_pdf_korean_font():
    font_paths = [
        "NanumGothic.ttf",
        "./NanumGothic.ttf",
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
        "C:/Windows/Fonts/malgun.ttf"
    ]
    for p in font_paths:
        if os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont("KoreanFont", p))
                return "KoreanFont"
            except Exception:
                continue
    return "Helvetica"

def draw_cic_letterhead_background(canvas, doc):
    canvas.saveState()
    w, h = doc.pagesize

    foot_x = 20
    foot_y = 16
    canvas.setFont("Helvetica-Bold", 7.5)
    canvas.setFillColor(colors.HexColor("#2563EB"))
    canvas.drawString(foot_x, foot_y + 18, COMPANY_INFO["NAME_EN"])

    kor_font = get_pdf_korean_font()
    canvas.setFont(kor_font, 6.5)
    canvas.setFillColor(colors.HexColor("#F43F5E"))
    canvas.drawString(foot_x, foot_y + 10, COMPANY_INFO["ADDR_EN"] + " / " + COMPANY_INFO["ADDR_KR"])
    canvas.drawString(foot_x, foot_y + 2, "Tel: " + COMPANY_INFO["TEL"] + " | Fax: " + COMPANY_INFO["FAX"] + " | Email: " + COMPANY_INFO["EMAIL"])

    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(colors.HexColor("#94A3B8"))
    canvas.drawRightString(w - 20, foot_y + 2, "Page " + str(canvas.getPageNumber()) + " | CIC LOGIS EXPORT CONSOLE")

    canvas.restoreState()

st.markdown(
    '<div style="background-color:' + THEME_COLOR + '; padding:15px; border-radius:8px; color:white; margin-bottom:20px;">'
    '<h3 style="margin:0;">📊 CIC LOGIS EXPORT CONSOLE (모바일 웹/PC 겸용 정산 시스템)</h3>'
    '<p style="margin:5px 0 0 0; font-size:13px; color:#94A3B8;">일일 정산 원장 및 회계 자동 검산 시스템</p>'
    '</div>',
    unsafe_allow_html=True
)

if "raw_records" not in st.session_state:
    st.session_state.raw_records = []

with st.sidebar:
    st.header("⚙️ 데이터 불러오기")
    fetch_mode = st.radio("방식 선택", ["구글 시트 연동", "엑셀 파일 업로드"])
    
    sheet_url = ""
    uploaded_file = None

    if fetch_mode == "구글 시트 연동":
        sheet_url = st.text_input("구글 스프레드시트 주소 (URL)")
        fetch_btn = st.button("🌐 구글시트 가져오기", use_container_width=True)
    else:
        uploaded_file = st.file_uploader("엑셀 파일 선택 (.xlsx)", type=["xlsx", "xls", "xlsm"])
        fetch_btn = st.button("📁 엑셀 파일 열기", use_container_width=True)

# 데이터 로드 처리
if fetch_btn:
    rows_data = []
    if fetch_mode == "구글 시트 연동" and sheet_url:
        match_id = re.search(r'/spreadsheets/d/([a-zA-Z0-9-_]+)', sheet_url)
        if match_id:
            sheet_id = match_id.group(1)
            match_gid = re.search(r'[#&?]gid=([0-9]+)', sheet_url)
            gid = match_gid.group(1) if match_gid else "0"
            csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
            try:
                req = urllib.request.Request(csv_url, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req, timeout=12) as response:
                    raw_bytes = response.read()
                try:
                    csv_text = raw_bytes.decode('utf-8-sig')
                except UnicodeDecodeError:
                    csv_text = raw_bytes.decode('cp949', errors='ignore')
                reader = csv.reader(io.StringIO(csv_text))
                rows_data = list(reader)
                st.success("구글 시트 연동 성공!")
            except Exception as e:
                st.error(f"구글 시트 로드 실패: {e}")
        else:
            st.error("올바른 구글 시트 주소가 아닙니다.")
    elif fetch_mode == "엑셀 파일 업로드" and uploaded_file:
        try:
            wb = openpyxl.load_workbook(uploaded_file, data_only=True)
            ws = wb.active
            for r in range(1, ws.max_row + 1):
                rows_data.append([ws.cell(r, c).value for c in range(1, ws.max_column + 1)])
            st.success("엑셀 파일 로드 성공!")
        except Exception as e:
            st.error(f"엑셀 파일 로드 실패: {e}")

    if rows_data:
        records = []
        header_idx = -1
        for idx, row in enumerate(rows_data[:10]):
            row_str = " ".join([str(c) for c in row if c is not None])
            if "입고일" in row_str or "거래처" in row_str or "M.B/L" in row_str:
                header_idx = idx
                break

        start_row = header_idx + 1 if header_idx != -1 else 0
        curr_year = datetime.datetime.now().year

        for row in rows_data[start_row:]:
            if not any(row):
                continue
            vals = [c if c is not None else "" for c in row]
            while len(vals) < 18:
                vals.append("")

            def parse_to_flt(val):
                if val is None or val == "" or val == "-" or str(val).strip() == "-":
                    return 0.0
                try:
                    clean_s = str(val).replace(",", "").replace("₩", "").replace("원", "").replace("$", "").strip()
                    return float(clean_s)
                except (ValueError, TypeError):
                    return 0.0

            date_val = str(vals[0]).strip()
            client_val = str(vals[1]).strip()
            cat_val = str(vals[2]).strip()
            mbl_val = str(vals[3]).strip()
            qty_val = parse_to_flt(vals[4])
            gwt_val = parse_to_flt(vals[5])
            cwt_val = parse_to_flt(vals[6])
            supply_val = parse_to_flt(vals[7])
            etc_val = parse_to_flt(vals[8])
            vat_val = parse_to_flt(vals[9])
            air_val = parse_to_flt(vals[10])
            cargo_val = parse_to_flt(vals[11])
            sales_val = parse_to_flt(vals[12])
            profit_val = parse_to_flt(vals[13])
            remarks_val = str(vals[14]).strip()
            net_price_val = parse_to_flt(vals[15] if len(vals) > 15 else 0)
            sell_price_val = parse_to_flt(vals[16] if len(vals) > 16 else 0)

            records.append({
                "date": date_val if date_val else datetime.datetime.now().strftime("%Y-%m-%d"),
                "client": client_val,
                "category": cat_val,
                "mbl": mbl_val,
                "qty": qty_val,
                "gwt": gwt_val,
                "cwt": cwt_val,
                "supply": supply_val,
                "etc": etc_val,
                "vat": vat_val,
                "air": air_val,
                "cargo": cargo_val,
                "sales": sales_val,
                "profit": profit_val,
                "remarks": remarks_val,
                "net_price": net_price_val,
                "sell_price": sell_price_val
            })
        st.session_state.raw_records = records

records = st.session_state.raw_records
if records:
    st.subheader(f"📑 일일 거래 원장 데이터 (총 {len(records)}건)")
    
    # KPI 요약 카드
    tot_sales = sum(r.get("sales", 0.0) for r in records)
    tot_profit = sum(r.get("profit", 0.0) for r in records)
    tot_chg_wt = sum(max(r.get("gwt", 0.0), r.get("cwt", 0.0)) for r in records)
    margin_rate = (tot_profit / tot_sales * 100.0) if tot_sales > 0 else 0.0

    kpi1, kpi2, kpi3 = st.columns(3)
    kpi1.metric("총 매출액 (Revenue)", f"₩ {tot_sales:,.0f}")
    kpi2.metric("영업이익 (PROFIT)", f"₩ {tot_profit:,.0f}", f"마진율 {margin_rate:.1f}%")
    kpi3.metric("총 물동량 (Chg.WT)", f"{tot_chg_wt:,.1f} kg")

    st.divider()

    table_data = []
    for r in records:
        table_data.append({
            "입고일": r["date"],
            "거래처": r["client"],
            "구분": r["category"],
            "M.B/L": r["mbl"],
            "수량": f"{r['qty']:,.0f}",
            "G.WT": f"{r['gwt']:,.1f}",
            "C.WT": f"{r['cwt']:,.1f}",
            "공급가액": f"{r['supply']:,.0f}",
            "기타비용": f"{r['etc']:,.0f}",
            "VAT": f"{r['vat']:,.0f}",
            "항공운임": f"{r['air']:,.0f}",
            "화물운임": f"{r['cargo']:,.0f}",
            "매출": f"{r['sales']:,.0f}",
            "PROFIT": f"{r['profit']:,.0f}",
            "기타내용": r["remarks"]
        })
    
    st.dataframe(table_data, use_container_width=True)
else:
    st.info("👈 왼쪽 사이드바에서 구글 스프레드시트 주소를 입력하거나 엑셀 파일을 업로드해 주세요.")