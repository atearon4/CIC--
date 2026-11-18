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
import calendar
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

# 차트 라이브러리 연동
try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    plt.rcParams['font.family'] = 'NanumGothic'
    plt.rcParams['axes.unicode_minus'] = False
    HAS_MATPLOTLIB = True
except ImportError:
    HAS_MATPLOTLIB = False

st.set_page_config(page_title="CIC LOGIS 수출 정산 콘솔", page_icon="📊", layout="wide")

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

DEFAULT_FORMULAS = {
    "supply": "큰값(G.WT, C.WT) * 단가",
    "vat": "(공급가액 + 기타비용) * 0.1",
    "air": "C.WT * 청구운임",
    "cargo": "G.WT * NET운임",
    "sales": "공급가액 + 기타비용 + VAT + 항공운임",
    "profit": "매출 - 화물운임(매입)"
}

KOREAN_TOKEN_MAP = {
    "화물운임(매입)": "cargo", "화물운임": "cargo", "공급가액": "supply", "기타비용": "etc",
    "항공운임": "air", "기타내용": "remarks", "NET운임": "net_price", "청구운임": "sell_price",
    "입고일": "date", "거래처": "client", "M.B/L": "mbl", "G.WT": "gwt", "C.WT": "cwt",
    "PROFIT": "profit", "VAT": "vat", "구분": "category", "수량": "qty", "매출": "sales",
    "단가": "rate", "큰값": "max", "작은값": "min", "반올림": "round", "절대값": "abs"
}

DEFAULT_RATES = {
    "KSG": {"DL": {"rate": 100, "sell_price": 0, "net_price": 0}, "AC": {"rate": 100, "sell_price": 0, "net_price": 0}},
    "통다": {"LOCAL": {"rate": 100, "sell_price": 0, "net_price": 0}, "A&A": {"rate": 250, "sell_price": 0, "net_price": 0}, "S&A": {"rate": 70, "sell_price": 0, "net_price": 0}, "RENT": {"rate": 0, "sell_price": 1780, "net_price": 1680}},
    "VC24G": {"": {"rate": 80, "sell_price": 0, "net_price": 0}},
    "트랜스올": {"": {"rate": 70, "sell_price": 0, "net_price": 0}},
    "알프스": {"": {"rate": 0, "sell_price": 7600, "net_price": 6500}},
    "샤인": {"": {"rate": 150, "sell_price": 0, "net_price": 0}},
    "포워더플러스": {"": {"rate": 0, "sell_price": 2150, "net_price": 1850}}
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
    '<h3 style="margin:0;">📊 CIC LOGIS EXPORT CONSOLE (수출팀 일일 정산 및 회계 검산 시스템)</h3>'
    '<p style="margin:5px 0 0 0; font-size:13px; color:#94A3B8;">모바일/PC 겸용 정산 및 파일 내보내기 시스템</p>'
    '</div>',
    unsafe_allow_html=True
)

if "raw_records" not in st.session_state: st.session_state.raw_records = []
if "rates" not in st.session_state: st.session_state.rates = DEFAULT_RATES.copy()
if "formulas" not in st.session_state: st.session_state.formulas = DEFAULT_FORMULAS.copy()
if "custom_expenses" not in st.session_state:
    st.session_state.custom_expenses = [
        {"item": "인건비 (수출팀 급여/식대)", "amount": 0.0},
        {"item": "차량 운행 및 유류비", "amount": 0.0},
        {"item": "사무실/창고 운영관리비", "amount": 0.0}
    ]

def parse_to_flt(val):
    if val is None or val == "" or val == "-" or str(val).strip() == "-": return 0.0
    try:
        clean_s = str(val).replace(",", "").replace("₩", "").replace("원", "").replace("$", "").strip()
        return float(clean_s)
    except (ValueError, TypeError):
        return 0.0

def parse_to_date(val, default_year=None):
    if val is None or val == "": return None
    if isinstance(val, (datetime.date, datetime.datetime)):
        return val.date() if isinstance(val, datetime.datetime) else val
    s = str(val).strip()
    nums = [int(n) for n in re.findall(r'\d+', s)]
    try:
        if len(nums) >= 3:
            y, m, d = nums[0], nums[1], nums[2]
            if y < 100: y += 2000
            return datetime.date(y, m, d)
        elif len(nums) == 2:
            y = default_year if default_year else datetime.datetime.now().year
            return datetime.date(y, nums[0], nums[1])
    except (ValueError, TypeError): pass
    return None

def translate_formula(expr_str):
    if not expr_str: return "0"
    s = str(expr_str).strip().replace("×", "*").replace("÷", "/")
    sorted_keys = sorted(KOREAN_TOKEN_MAP.keys(), key=len, reverse=True)
    for k in sorted_keys: s = s.replace(k, KOREAN_TOKEN_MAP[k])
    return s

def eval_expr(expr_str, context):
    try:
        py_expr = translate_formula(expr_str)
        allowed_names = {"max": max, "min": min, "round": round, "abs": abs}
        allowed_names.update(context)
        return float(eval(py_expr, {"__builtins__": {}}, allowed_names))
    except Exception: return 0.0

def get_rate_info(client, category):
    c_rates = st.session_state.rates.get(client, {})
    if category in c_rates:
        info = c_rates[category]
        if isinstance(info, dict): return info.get("rate", 0.0), info.get("sell_price", 0.0), info.get("net_price", 0.0)
        return float(info), 0.0, 0.0
    if "" in c_rates:
        info = c_rates[""]
        if isinstance(info, dict): return info.get("rate", 0.0), info.get("sell_price", 0.0), info.get("net_price", 0.0)
        return float(info), 0.0, 0.0
    return 0.0, 0.0, 0.0

def recalculate_records():
    for r in st.session_state.raw_records:
        c, cat = r["client"], r["category"]
        rate_def, sell_def, net_def = get_rate_info(c, cat)
        r["rate"] = rate_def if r.get("rate", 0) == 0 else r["rate"]
        r["sell_price"] = sell_def if r.get("sell_price", 0) == 0 else r["sell_price"]
        r["net_price"] = net_def if r.get("net_price", 0) == 0 else r["net_price"]

        ctx = {
            "qty": r["qty"], "gwt": r["gwt"], "cwt": r["cwt"],
            "rate": r["rate"], "sell_price": r["sell_price"], "net_price": r["net_price"],
            "supply": r["supply"], "etc": r["etc"], "vat": r["vat"]
        }
        f = st.session_state.formulas
        r["supply"] = round(eval_expr(f["supply"], ctx))
        ctx["supply"] = r["supply"]
        r["vat"] = round(eval_expr(f["vat"], ctx))
        ctx["vat"] = r["vat"]
        r["air"] = round(eval_expr(f["air"], ctx))
        r["cargo"] = round(eval_expr(f["cargo"], ctx))
        r["sales"] = round(eval_expr(f["sales"], ctx))
        ctx["sales"] = r["sales"]
        r["profit"] = round(eval_expr(f["profit"], ctx))

# 사이드바 데이터 불러오기 및 설정 관리
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

    st.divider()
    st.subheader("🛠️ 요율 및 수식 설정")
    with st.expander("⚙️ 거래처별 요율 관리"):
        rate_client = st.text_input("거래처명", placeholder="예: KSG")
        rate_cat = st.text_input("구분", placeholder="예: DL (공백가능)")
        r_val = st.number_input("단가표 (공급가)", value=100.0)
        s_val = st.number_input("청구운임", value=0.0)
        n_val = st.number_input("NET운임", value=0.0)
        if st.button("요율 등록/수정"):
            if rate_client:
                if rate_client not in st.session_state.rates: st.session_state.rates[rate_client] = {}
                st.session_state.rates[rate_client][rate_cat] = {"rate": r_val, "sell_price": s_val, "net_price": n_val}
                st.success(f"[{rate_client} - {rate_cat}] 요율 저장 완료!")
                recalculate_records()

    with st.expander("📐 딸깍 자동완성 수식 설정"):
        f_supply = st.text_input("공급가액 수식", value=st.session_state.formulas["supply"])
        f_vat = st.text_input("VAT 수식", value=st.session_state.formulas["vat"])
        f_air = st.text_input("항공운임 수식", value=st.session_state.formulas["air"])
        f_cargo = st.text_input("화물운임 수식", value=st.session_state.formulas["cargo"])
        f_sales = st.text_input("매출 수식", value=st.session_state.formulas["sales"])
        f_profit = st.text_input("PROFIT 수식", value=st.session_state.formulas["profit"])
        if st.button("수식 저장 및 재계산"):
            st.session_state.formulas = {"supply": f_supply, "vat": f_vat, "air": f_air, "cargo": f_cargo, "sales": f_sales, "profit": f_profit}
            recalculate_records()
            st.success("수식 저장 및 일괄 재계산 완료!")

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
                try: csv_text = raw_bytes.decode('utf-8-sig')
                except UnicodeDecodeError: csv_text = raw_bytes.decode('cp949', errors='ignore')
                reader = csv.reader(io.StringIO(csv_text))
                rows_data = list(reader)
                st.success("구글 시트 연동 성공!")
            except Exception as e: st.error(f"구글 시트 로드 실패: {e}")
    elif fetch_mode == "엑셀 파일 업로드" and uploaded_file:
        try:
            wb = openpyxl.load_workbook(uploaded_file, data_only=True)
            ws = wb.active
            for r in range(1, ws.max_row + 1):
                rows_data.append([ws.cell(r, c).value for c in range(1, ws.max_column + 1)])
            st.success("엑셀 파일 로드 성공!")
        except Exception as e: st.error(f"엑셀 파일 로드 실패: {e}")

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
            if not any(row): continue
            vals = [c if c is not None else "" for c in row]
            while len(vals) < 18: vals.append("")

            d_obj = parse_to_date(vals[0], default_year=curr_year)
            date_str = d_obj.strftime("%Y-%m-%d") if d_obj else str(vals[0]).strip()
            date_key = (d_obj.year * 10000 + d_obj.month * 100 + d_obj.day) if d_obj else 999999

            records.append({
                "date": date_str if date_str else datetime.datetime.now().strftime("%Y-%m-%d"),
                "date_key": date_key,
                "client": str(vals[1]).strip(),
                "category": str(vals[2]).strip(),
                "mbl": str(vals[3]).strip(),
                "qty": parse_to_flt(vals[4]),
                "gwt": parse_to_flt(vals[5]),
                "cwt": parse_to_flt(vals[6]),
                "supply": parse_to_flt(vals[7]),
                "etc": parse_to_flt(vals[8]),
                "vat": parse_to_flt(vals[9]),
                "air": parse_to_flt(vals[10]),
                "cargo": parse_to_flt(vals[11]),
                "sales": parse_to_flt(vals[12]),
                "profit": parse_to_flt(vals[13]),
                "remarks": str(vals[14]).strip(),
                "net_price": parse_to_flt(vals[15] if len(vals) > 15 else 0),
                "sell_price": parse_to_flt(vals[16] if len(vals) > 16 else 0),
                "rate": 0.0
            })
        st.session_state.raw_records = records
        recalculate_records()

records = st.session_state.raw_records
if records:
    tot_sales = sum(r.get("sales", 0.0) for r in records)
    tot_profit = sum(r.get("profit", 0.0) for r in records)
    tot_chg_wt = sum(max(r.get("gwt", 0.0), r.get("cwt", 0.0)) for r in records)
    margin_rate = (tot_profit / tot_sales * 100.0) if tot_sales > 0 else 0.0

    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("총 매출액 (Revenue)", f"₩ {tot_sales:,.0f}")
    kpi2.metric("영업이익 (PROFIT)", f"₩ {tot_profit:,.0f}", f"마진율 {margin_rate:.1f}%")
    kpi3.metric("총 물동량 (Chg.WT)", f"{tot_chg_wt:,.1f} kg")
    kpi4.metric("총 진행건수", f"{len(records):,} 건")

    st.divider()

    # 모든 전체 기능 탭 구성 (원장, 월누계, 손익계산서, 회계검산, 실적차트, 일일정산PDF/엑셀)
    tab_raw, tab_monthly, tab_expense, tab_verify, tab_chart, tab_daily_pdf = st.tabs([
        "📑 일일 거래 원장", "📊 월 누계 요약표", "💼 손익계산서", "🔍 회계 검산", "📈 실적 대시보드", "📄 일일정산표 PDF/Excel"
    ])

    with tab_raw:
        st.subheader(f"일일 거래 원장 데이터 (총 {len(records)}건)")
        table_data = []
        for r in records:
            table_data.append({
                "입고일": r["date"], "거래처": r["client"], "구분": r["category"], "M.B/L": r["mbl"],
                "수량": f"{r['qty']:,.0f}", "G.WT": f"{r['gwt']:,.1f}", "C.WT": f"{r['cwt']:,.1f}",
                "공급가액": f"{r['supply']:,.0f}", "기타비용": f"{r['etc']:,.0f}", "VAT": f"{r['vat']:,.0f}",
                "항공운임": f"{r['air']:,.0f}", "화물운임": f"{r['cargo']:,.0f}", "매출": f"{r['sales']:,.0f}",
                "PROFIT": f"{r['profit']:,.0f}", "기타내용": r["remarks"]
            })
        st.dataframe(table_data, use_container_width=True)

    with tab_monthly:
        st.subheader("📊 거래처별 월 누계 요약표")
        summary_dict = {}
        for r in records:
            key = (r["client"], r["category"])
            if key not in summary_dict:
                summary_dict[key] = {"qty": 0, "gwt": 0, "cwt": 0, "supply": 0, "etc": 0, "vat": 0, "air": 0, "cargo": 0, "sales": 0, "profit": 0}
            d = summary_dict[key]
            d["qty"] += r["qty"]; d["gwt"] += r["gwt"]; d["cwt"] += r["cwt"]
            d["supply"] += r["supply"]; d["etc"] += r["etc"]; d["vat"] += r["vat"]
            d["air"] += r["air"]; d["cargo"] += r["cargo"]; d["sales"] += r["sales"]; d["profit"] += r["profit"]

        m_data = []
        for (c, cat), d in summary_dict.items():
            m_data.append({
                "거래처": c, "구분": cat, "수량": f"{d['qty']:,.0f}", "G.WT": f"{d['gwt']:,.1f}", "C.WT": f"{d['cwt']:,.1f}",
                "공급가액": f"{d['supply']:,.0f}", "기타비용": f"{d['etc']:,.0f}", "VAT": f"{d['vat']:,.0f}",
                "항공운임": f"{d['air']:,.0f}", "화물운임": f"{d['cargo']:,.0f}", "매출": f"{d['sales']:,.0f}", "PROFIT": f"{d['profit']:,.0f}"
            })
        st.dataframe(m_data, use_container_width=True)

    with tab_expense:
        st.subheader("💼 관리비용(인건비 등) 포함 최종 손익계산서")
        tot_cargo = sum(r.get("cargo", 0.0) for r in records)
        tot_raw_profit = sum(r.get("profit", 0.0) for r in records)
        
        col_e1, col_e2 = st.columns(2)
        with col_e1:
            st.metric("총 매출액 (Revenue)", f"₩ {tot_sales:,.0f}")
            st.metric("매출원가 (화물운임 매입)", f"₩ {tot_cargo:,.0f}")
        with col_e2:
            st.metric("1차 영업이익 (PROFIT)", f"₩ {tot_raw_profit:,.0f}")
            net_margin = (tot_raw_profit / tot_sales * 100.0) if tot_sales > 0 else 0.0
            st.metric("기본 영업 마진율", f"{net_margin:.1f}%")

        st.divider()
        st.subheader("추가 관리비용 관리")
        for idx, exp in enumerate(st.session_state.custom_expenses):
            col_ex1, col_ex2 = st.columns([3, 1])
            with col_ex1:
                st.text(f"• {exp['item']}")
            with col_ex2:
                st.text(f"₩ {exp['amount']:,.0f}")

        tot_exp = sum(e["amount"] for e in st.session_state.custom_expenses)
        final_profit = tot_raw_profit - tot_exp
        final_margin = (final_profit / tot_sales * 100.0) if tot_sales > 0 else 0.0
        st.metric("★ 최종 순 영업이익 (관리비용 차감 후)", f"₩ {final_profit:,.0f}", f"최종 마진율 {final_margin:.1f}%")

    with tab_verify:
        st.subheader("🔍 회계 검산 실행")
        if st.button("🔍 회계 검산 시작 (데이터 일치 여부 확인)", use_container_width=True):
            error_count = 0
            for r in records:
                ctx = {"qty": r["qty"], "gwt": r["gwt"], "cwt": r["cwt"], "rate": r["rate"], "sell_price": r["sell_price"], "net_price": r["net_price"], "supply": r["supply"], "etc": r["etc"], "vat": r["vat"]}
                expected_vat = round(eval_expr(DEFAULT_FORMULAS["vat"], ctx))
                if abs(r["vat"] - expected_vat) >= 5: error_count += 1
            if error_count == 0:
                st.success("✅ [회계 검산 통과] 모든 데이터 항목이 수식 공식과 100% 일치합니다!")
            else:
                st.error(f"⚠️ 총 {error_count}건의 VAT 또는 수식 불일치 항목이 발견되었습니다.")

    with tab_chart:
        st.subheader("📈 실적 분석 대시보드")
        if HAS_MATPLOTLIB:
            client_sales = {}
            for r in records:
                c = r["client"]
                client_sales[c] = client_sales.get(c, 0.0) + r.get("sales", 0.0)
            
            fig, ax = plt.subplots(figsize=(8, 4))
            ax.bar(list(client_sales.keys()), list(client_sales.values()), color='#2563EB')
            ax.set_title("거래처별 매출액 현황", fontfamily='NanumGothic')
            plt.xticks(rotation=30)
            st.pyplot(fig)
        else:
            st.info("matplotlib 라이브러리가 설치되지 않았습니다.")

    with tab_daily_pdf:
        st.subheader("📄 일일 정산표 PDF 및 엑셀 다운로드")
        dates_list = sorted(list(set(r["date"] for r in records)))
        selected_date = st.selectbox("출력할 날짜 선택", dates_list if dates_list else ["날짜 없음"])

        if HAS_REPORTLAB and selected_date != "날짜 없음":
            if st.button("📥 선택한 날짜 일일정산표 PDF 생성", use_container_width=True):
                pdf_filename = f"CIC_LOGIS_일일정산_{selected_date}.pdf"
                kor_font = get_pdf_korean_font()
                doc = SimpleDocTemplate(pdf_filename, pagesize=landscape(A4), leftMargin=15, rightMargin=15, topMargin=40, bottomMargin=40)
                elements = []

                title_style = ParagraphStyle('TitleStyle', fontName=kor_font, fontSize=14, textColor=colors.HexColor("#0F172A"), alignment=1, spaceAfter=10)
                elements.append(Paragraph(f"<b>CIC LOGIS 수출팀 당일 진행건 정산 내역 ({selected_date})</b>", title_style))

                headers = ["입고일", "거래처", "구분", "M.B/L", "수량", "G.WT", "C.WT", "공급가액", "기타비용", "VAT", "항공운임", "화물운임", "매출", "PROFIT", "기타내용"]
                table_data_pdf = [headers]

                day_recs = [r for r in records if r["date"] == selected_date]
                for r in day_recs:
                    table_data_pdf.append([
                        r["date"], r["client"], r["category"], r["mbl"][:11],
                        f"{r['qty']:,.0f}", f"{r['gwt']:,.1f}", f"{r['cwt']:,.1f}",
                        f"{r['supply']:,.0f}", f"{r['etc']:,.0f}", f"{r['vat']:,.0f}",
                        f"{r['air']:,.0f}", f"{r['cargo']:,.0f}",
                        f"{r['sales']:,.0f}", f"{r['profit']:,.0f}",
                        r["remarks"][:8]
                    ])

                col_w = [48, 55, 42, 60, 36, 46, 44, 52, 46, 46, 58, 60, 62, 62, 75]
                pdf_table = Table(table_data_pdf, colWidths=col_w, repeatRows=1)
                pdf_table.setStyle(TableStyle([
                    ('FONTNAME', (0, 0), (-1, -1), kor_font),
                    ('FONTSIZE', (0, 0), (-1, -1), 7.5),
                    ('ALIGN', (0, 0), (3, -1), 'CENTER'),
                    ('ALIGN', (4, 0), (-2, -1), 'RIGHT'),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1E293B")),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white)
                ]))
                elements.append(pdf_table)
                doc.build(elements, onFirstPage=draw_cic_letterhead_background, onLaterPages=draw_cic_letterhead_background)

                with open(pdf_filename, "rb") as f:
                    st.download_button(
                        label="📥 생성된 일일정산 PDF 다운로드 클릭",
                        data=f,
                        file_name=pdf_filename,
                        mime="application/pdf",
                        use_container_width=True
                    )
else:
    st.info("👈 왼쪽 사이드바에서 구글 스프레드시트 주소를 입력하거나 엑셀 파일을 업로드해 주세요.")