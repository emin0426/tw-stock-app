# -*- coding: utf-8 -*-
"""
台股進階三條件選股 & 主力籌碼成本健檢 — 手機網頁版 App v2.4 (15:30 統一更新 + 雙層共享快取版)
核心升級：
1. 統一資料更新基準點為【每日台灣時間 15:30】：
   - 15:30 前：強制排除今日未定型資料，嚴格採用「前一個完整交易日」收盤價與三大法人籌碼。
   - 15:30 後：採用「今日」完整收盤價與證交所／櫃買中心最新三大法人買賣超。
   - 介面頂部自動偵測並明確標註【目前採用之資料日期 (YYYY-MM-DD)】。
2. 雙層極速快取架構（防多人流量爆量與 Yahoo 限流）：
   - 第一層（主動預抓快取）：自動讀取每日 15:30 由 GitHub Actions 主動預先算好的 `daily_cache/<市場>.json`，預設參數下 0.1 秒秒開！
   - 第二層（共享記憶體快取）：若為自訂參數或尚未預抓，首位訪客掃描完後自動共享快取給所有後續訪客，直到下一個 15:30 自動換新。
3. 內建【投資免責聲明警語】與【選配式密碼鎖（預設關閉，需鎖定時至 Streamlit Secrets 設定 APP_PASSWORD 即可）】。
"""

import io
import os
import re
import ssl
import json
import time
from datetime import datetime, timezone, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed
import urllib.request

import numpy as np
import pandas as pd
import yfinance as yf
import streamlit as st
import altair as alt
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==========================================
# 頁面設定與手機版 CSS 美化
# ==========================================
st.set_page_config(
    page_title="台股選股與主力籌碼健檢",
    page_icon="📈",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    .block-container { padding-top: 1.2rem; padding-bottom: 2.5rem; max-width: 820px; }
    .stButton > button { width: 100%; border-radius: 10px; font-weight: 600; padding: 0.55rem 1rem; }
    .card-box {
        border-radius: 12px;
        padding: 14px 16px;
        margin-bottom: 14px;
        border: 1px solid #d0d7de;
        background-color: #ffffff;
        box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    }
    .card-star { border-left: 6px solid #198754; background-color: #f4fbf7; }
    .card-watch { border-left: 6px solid #0d6efd; background-color: #f4f8ff; }
    .card-warn { border-left: 6px solid #dc3545; background-color: #fff5f5; }
    .card-title { font-size: 1.08rem; font-weight: 700; margin-bottom: 6px; color: #1f2328; }
    .card-badge { font-size: 0.9rem; font-weight: 600; margin-bottom: 10px; }
    .sec-head { font-weight: 700; color: #1F4E78; margin-top: 8px; margin-bottom: 3px; font-size: 0.95rem; }
    .sec-body { font-size: 0.9rem; line-height: 1.55; color: #24292f; white-space: pre-line; }
    .disclaimer-box {
        margin-top: 24px;
        padding: 12px 14px;
        border-radius: 8px;
        background-color: #f6f8fa;
        border: 1px solid #d8dee4;
        font-size: 0.82rem;
        color: #57606a;
        line-height: 1.5;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ==========================================
# 選配密碼鎖（預設關閉，若在 Streamlit Secrets 設定 APP_PASSWORD 則自動啟用）
# ==========================================
def check_optional_password():
    try:
        secret_pwd = st.secrets.get("APP_PASSWORD", "")
    except Exception:
        secret_pwd = ""
    if not secret_pwd:
        return True
    if st.session_state.get("auth_ok"):
        return True
    st.title("🔒 台股選股與主力籌碼健檢")
    pwd_in = st.text_input("請輸入存取密碼：", type="password")
    if st.button("登入", type="primary"):
        if pwd_in == str(secret_pwd):
            st.session_state["auth_ok"] = True
            st.rerun()
        else:
            st.error("密碼錯誤，請重新輸入。")
    return False


if not check_optional_password():
    st.stop()


MODE_LABELS = {
    "top": "熱門權值股",
    "twse": "上市股",
    "tpex": "上櫃股",
    "emerging": "興櫃股",
}

DEFAULT_TOP_STOCKS = {
    "2330.TW": ("台積電", "上市股"), "2317.TW": ("鴻海", "上市股"), "2454.TW": ("聯發科", "上市股"),
    "2308.TW": ("台達電", "上市股"), "2382.TW": ("廣達", "上市股"), "2881.TW": ("富邦金", "上市股"),
    "2882.TW": ("國泰金", "上市股"), "2891.TW": ("中信金", "上市股"), "2303.TW": ("聯電", "上市股"),
    "3711.TW": ("日月光投控", "上市股"), "2412.TW": ("中華電", "上市股"), "2886.TW": ("兆豐金", "上市股"),
    "3231.TW": ("緯創", "上市股"), "2884.TW": ("玉山金", "上市股"), "2357.TW": ("華碩", "上市股"),
    "6669.TW": ("緯穎", "上市股"), "2345.TW": ("智邦", "上市股"), "3037.TW": ("欣興", "上市股"),
    "2379.TW": ("瑞昱", "上市股"), "3008.TW": ("大立光", "上市股"), "2603.TW": ("長榮", "上市股"),
    "2609.TW": ("陽明", "上市股"), "2615.TW": ("萬海", "上市股"), "2002.TW": ("中鋼", "上市股"),
    "1301.TW": ("台塑", "上市股"), "1303.TW": ("南亞", "上市股"), "1216.TW": ("統一", "上市股"),
    "2912.TW": ("統一超", "上市股"), "2395.TW": ("研華", "上市股"), "3034.TW": ("聯詠", "上市股"),
    "2327.TW": ("國巨", "上市股"), "2301.TW": ("光寶科", "上市股"), "4938.TW": ("和碩", "上市股"),
    "2376.TW": ("技嘉", "上市股"), "2377.TW": ("微星", "上市股"), "3017.TW": ("奇鋐", "上市股"),
    "3324.TWO": ("雙鴻", "上櫃股"), "3661.TW": ("世芯-KY", "上市股"), "3443.TW": ("創意", "上市股"),
    "3005.TW": ("神基", "上市股"), "6435.TWO": ("大中", "上櫃股"), "2356.TW": ("英業達", "上市股"),
    "2324.TW": ("仁寶", "上市股"), "2353.TW": ("宏碁", "上市股"), "2618.TW": ("長榮航", "上市股"),
    "2610.TW": ("華航", "上市股"), "1513.TW": ("中興電", "上市股"), "1519.TW": ("華城", "上市股"),
    "1503.TW": ("士電", "上市股"), "1504.TW": ("東元", "上市股"), "1605.TW": ("華新", "上市股"),
    "2408.TW": ("南亞科", "上市股"), "2344.TW": ("華邦電", "上市股"), "2337.TW": ("旺宏", "上市股"),
    "3702.TW": ("大聯大", "上市股"), "2347.TW": ("聯強", "上市股"), "8069.TWO": ("元太", "上櫃股"),
    "6488.TWO": ("環球晶", "上櫃股"), "5483.TWO": ("中美晶", "上櫃股"), "3105.TWO": ("穩懋", "上櫃股"),
    "8299.TWO": ("群聯", "上櫃股"), "5274.TWO": ("信驊", "上櫃股"), "3529.TWO": ("力旺", "上櫃股"),
    "6510.TWO": ("精測", "上櫃股"),
}

EXCEL_COLS = [
    "股票代碼", "股票名稱", "市場別", "營業概況(主力業務)", "資料日期", "收盤價", "單日漲幅(%)",
    "成交量(張)", "日成交金額(萬元)", "量能倍數(倍)", "日週轉率(%)", "昨日帶寬(%)", "今日帶寬(%)",
    "距20MA乖離(%)", "MACD柱狀體", "上影線/紅K實體比",
    "突破紅K低點(停損參考)", "20MA月線價", "20日VWAP資金成本", "60日籌碼密集區(POC)",
    "前波60日高點", "近5日法人合計(張)", "近20日法人合計(張)", "外資/投信連買天數", "法人20日買超成本", "近5日融資增減(張)", "券資比(%)", "資券籌碼訊號",
    "狀態與未達標細項", "實戰型態定調與優先級"
]


# ==========================================
# 統一 15:30 更新基準與有效交易日偵測
# ==========================================
def get_market_data_cutoff_info():
    """
    統一以台灣時間 15:30 為每日資料更新切點（收盤價＋證交所三大法人皆已公布）：
    - 15:30 前：排除今日，採用「前一個交易日」完整收盤資訊。
    - 15:30 後：採用「今日」最新完整收盤資訊。
    回傳：(exclude_today, today_tw_date, period_key, now_str)
    """
    now_tw = datetime.now(timezone(timedelta(hours=8)))
    today_date = now_tw.date()
    is_before_1530 = (now_tw.hour, now_tw.minute) < (15, 30)
    if is_before_1530:
        # 15:30 前的快取鑰匙固定為「昨日日期_post1530」，完美承接昨天 15:30 後的快取！
        prev_day = today_date - timedelta(days=1)
        period_key = f"{prev_day.strftime('%Y-%m-%d')}_post1530"
    else:
        period_key = f"{today_date.strftime('%Y-%m-%d')}_post1530"
    return is_before_1530, today_date, period_key, now_tw.strftime("%Y-%m-%d %H:%M")


@st.cache_data(ttl=300, show_spinner=False)
def get_effective_trading_date_str(period_key):
    """透過台股加權指數(^TWII)自動確認實際採用之最新完整交易日期（自動略過週末與國定假日）"""
    exclude_today, today_tw_date, _, _ = get_market_data_cutoff_info()
    try:
        df_idx = yf.download("2330.TW", period="15d", interval="1d", progress=False)
        if df_idx is not None and not df_idx.empty:
            idx_dates = pd.to_datetime(df_idx.index).date
            if exclude_today:
                valid_dates = [d for d in idx_dates if d < today_tw_date]
            else:
                valid_dates = list(idx_dates)
            if valid_dates:
                return valid_dates[-1].strftime("%Y-%m-%d")
    except Exception:
        pass
    fallback_dt = (today_tw_date - timedelta(days=1)) if exclude_today else today_tw_date
    return fallback_dt.strftime("%Y-%m-%d")


def _parse_num_str(val):
    if val is None:
        return 0.0
    s = str(val).replace(",", "").replace(" ", "").strip()
    if not s or s in ("--", "-", "N/A"):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


@st.cache_data(ttl=86400, show_spinner=False)
def fetch_tw_tickers_fast(mode="top"):
    """快取 24 小時抓取上市/上櫃/興櫃 4 碼普通股代碼表（含官方 OpenAPI JSON 雙重備援）"""
    if mode == "top":
        return DEFAULT_TOP_STOCKS

    mode_config = {
        "twse": (
            "https://isin.twse.com.tw/isin/C_public.jsp?strMode=2",
            "https://openapi.twse.com.tw/v1/opendata/t187ap03_L",
            ".TW",
            "上市股",
        ),
        "tpex": (
            "https://isin.twse.com.tw/isin/C_public.jsp?strMode=4",
            "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap03_O",
            ".TWO",
            "上櫃股",
        ),
        "emerging": (
            "https://isin.twse.com.tw/isin/C_public.jsp?strMode=5",
            None,
            ".TWO",
            "興櫃股",
        ),
    }

    if mode not in mode_config:
        return DEFAULT_TOP_STOCKS

    isin_url, openapi_url, suffix, market_label = mode_config[mode]
    stock_dict = {}
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    ssl_ctx = ssl._create_unverified_context()

    try:
        req = urllib.request.Request(isin_url, headers=headers)
        with urllib.request.urlopen(req, timeout=12, context=ssl_ctx) as resp:
            raw_bytes = resp.read()
        html_text = raw_bytes.decode("big5", errors="ignore")

        rows = html_text.split("<tr")
        pattern = re.compile(r">\s*(\d{4})[\s\u3000]+([^<\r\n\t]+?)\s*<")
        for row_html in rows:
            if "ESVUFR" in row_html:
                m = pattern.search(row_html)
                if m:
                    code, name = m.group(1).strip(), m.group(2).strip()
                    if len(code) == 4 and code.isdigit() and not code.startswith("00"):
                        stock_dict[f"{code}{suffix}"] = (name, market_label)

        if not stock_dict:
            for code, name in pattern.findall(html_text):
                code, name = code.strip(), name.strip()
                if len(code) == 4 and code.isdigit() and not code.startswith("00"):
                    stock_dict[f"{code}{suffix}"] = (name, market_label)
    except Exception:
        pass

    if not stock_dict and openapi_url:
        try:
            req = urllib.request.Request(openapi_url, headers=headers)
            with urllib.request.urlopen(req, timeout=12, context=ssl_ctx) as resp:
                data_list = json.loads(resp.read().decode("utf-8", errors="ignore"))
            for item in data_list:
                code = str(item.get("公司代號") or item.get("SecuritiesCompanyCode") or "").strip()
                name = str(item.get("公司簡稱") or item.get("CompanyAbbreviation") or "").strip()
                if len(code) == 4 and code.isdigit() and not code.startswith("00") and name:
                    stock_dict[f"{code}{suffix}"] = (name, market_label)
        except Exception:
            pass

    return stock_dict if stock_dict else DEFAULT_TOP_STOCKS


TW_INDUSTRY_MAP = {
    "01": "水泥工業", "02": "食品工業", "03": "塑膠工業", "04": "紡織纖維",
    "05": "電機機械", "06": "電器電纜", "07": "化學生技醫療", "08": "玻璃陶瓷",
    "09": "造紙工業", "10": "鋼鐵工業", "11": "橡膠工業", "12": "汽車工業",
    "13": "電子工業", "14": "建材營造", "15": "航運業", "16": "觀光餐旅",
    "17": "金融保險", "18": "貿易百貨", "19": "綜合", "20": "其他",
    "21": "化學工業", "22": "生技醫療業", "23": "油電燃氣業", "24": "半導體業",
    "25": "電腦及週邊設備業", "26": "光電業", "27": "通信網路業", "28": "電子零組件業",
    "29": "電子通路業", "30": "資訊服務業", "31": "其他電子業", "32": "文化創意業",
    "33": "農業科技業", "34": "電子商務", "35": "綠能環保", "36": "數位雲端",
    "37": "運動休閒", "38": "居家生活",
}


def _extract_business_info(item):
    raw_ind = str(item.get("產業別") or item.get("IndustryCategory") or item.get("Industry") or "").strip()
    ind_name = TW_INDUSTRY_MAP.get(raw_ind.zfill(2), raw_ind) if raw_ind.isdigit() else raw_ind

    biz = ""
    for k in ("主要經營業務", "主要業務", "營業項目", "MainBusiness", "BusinessDescription", "MajorBusiness"):
        v = str(item.get(k) or "").strip()
        if v and v not in ("--", "-", "N/A"):
            biz = v
            break
    if not biz:
        for k, v in item.items():
            kl = str(k).lower()
            if ("業務" in str(k) or "business" in kl) and str(v).strip() not in ("", "--", "-", "N/A"):
                biz = str(v).strip()
                break

    biz = re.sub(r"\s+", " ", biz).strip()
    if len(biz) > 95:
        biz = biz[:92] + "..."

    if ind_name and biz:
        return f"【{ind_name}】{biz}"
    if ind_name:
        return f"【{ind_name}】"
    if biz:
        return biz
    return ""


@st.cache_data(ttl=86400, show_spinner=False)
def fetch_company_profile_map():
    """從證交所與櫃買中心官方 OpenAPI 一次取得各公司發行總張數與營業概況(主力業務)"""
    profile_map = {}
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    ssl_ctx = ssl._create_unverified_context()
    for url in [
        "https://openapi.twse.com.tw/v1/opendata/t187ap03_L",
        "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap03_O",
        "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap03_R",
    ]:
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=12, context=ssl_ctx) as resp:
                data_list = json.loads(resp.read().decode("utf-8", errors="ignore"))
            for item in data_list:
                code = str(item.get("公司代號") or item.get("SecuritiesCompanyCode") or "").strip()
                if not code:
                    continue
                shares = _parse_num_str(item.get("已發行普通股數或TDR原股發行股數") or item.get("IssuedCommonShares") or "")
                if shares <= 0:
                    cap = _parse_num_str(item.get("實收資本額") or item.get("PaidInCapital") or "")
                    if cap > 0:
                        shares = cap / 10.0
                lots = (shares / 1000.0) if shares > 0 else 0.0
                biz_desc = _extract_business_info(item)
                profile_map[code] = {"lots": lots, "business": biz_desc}
        except Exception:
            continue
    return profile_map


@st.cache_data(ttl=7200, show_spinner=False)
def fetch_insti_day_cached(exchange, yyyymmdd):
    """快取 2 小時抓取單日上市(TWSE)或上櫃(TPEX)三大法人買賣超 JSON（單位：張）"""
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    ssl_ctx = ssl._create_unverified_context()
    day_map = {}

    if exchange == "TWSE":
        url = f"https://www.twse.com.tw/rwd/zh/fund/T86?date={yyyymmdd}&selectType=ALLBUT0999&response=json"
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=10, context=ssl_ctx) as resp:
                payload = json.loads(resp.read().decode("utf-8", errors="ignore"))
            if payload.get("stat") == "OK" and payload.get("data") and payload.get("fields"):
                fields = [str(f).strip() for f in payload["fields"]]
                idx_f = next((i for i, f in enumerate(fields) if "外陸資買賣超股數" in f or "外資買賣超股數" in f), None)
                idx_t = next((i for i, f in enumerate(fields) if "投信買賣超股數" in f), None)
                idx_d = next((i for i, f in enumerate(fields) if f == "自營商買賣超股數"), None)
                idx_tot = next((i for i, f in enumerate(fields) if "三大法人買賣超股數" in f), None)
                for r in payload["data"]:
                    code = str(r[0]).strip()
                    if not code:
                        continue
                    f_l = _parse_num_str(r[idx_f]) / 1000.0 if idx_f is not None else 0.0
                    t_l = _parse_num_str(r[idx_t]) / 1000.0 if idx_t is not None else 0.0
                    d_l = _parse_num_str(r[idx_d]) / 1000.0 if idx_d is not None else 0.0
                    tot_l = _parse_num_str(r[idx_tot]) / 1000.0 if idx_tot is not None else (f_l + t_l + d_l)
                    day_map[code] = (f_l, t_l, d_l, tot_l)
        except Exception:
            pass
    elif exchange == "TPEX":
        y, m, d = int(yyyymmdd[:4]), yyyymmdd[4:6], yyyymmdd[6:8]
        ad_date = f"{y}/{m}/{d}"
        roc_date = f"{y - 1911}/{m}/{d}"
        urls = [
            f"https://www.tpex.org.tw/www/zh-tw/insti/dailyTrade?type=Daily&sect=EW&date={ad_date}&id=&response=json",
            f"https://www.tpex.org.tw/web/stock/3insti/daily_trade/3itrade_hedge_result.php?l=zh-tw&o=json&se=EW&t=D&d={roc_date}",
        ]
        for url in urls:
            try:
                req = urllib.request.Request(url, headers=headers)
                with urllib.request.urlopen(req, timeout=10, context=ssl_ctx) as resp:
                    payload = json.loads(resp.read().decode("utf-8", errors="ignore"))
                if "tables" in payload and payload["tables"]:
                    tbl = payload["tables"][0]
                    fields = [str(f).strip() for f in tbl.get("fields", [])]
                    rows_data = tbl.get("data", [])
                    idx_f = next((i for i, f in enumerate(fields) if "外資及陸資" in f and "買賣超" in f and "自營" not in f), None)
                    if idx_f is None:
                        idx_f = next((i for i, f in enumerate(fields) if "外資" in f and "買賣超" in f), None)
                    idx_t = next((i for i, f in enumerate(fields) if "投信" in f and "買賣超" in f), None)
                    idx_d = next((i for i, f in enumerate(fields) if "自營商" in f and "買賣超" in f and "合計" in f), None)
                    if idx_d is None:
                        idx_d = next((i for i, f in enumerate(fields) if f == "自營商-買賣超股數"), None)
                    idx_tot = next((i for i, f in enumerate(fields) if "三大法人" in f and "買賣超" in f), None)
                    for r in rows_data:
                        code = str(r[0]).strip()
                        if not code:
                            continue
                        f_l = _parse_num_str(r[idx_f]) / 1000.0 if idx_f is not None else 0.0
                        t_l = _parse_num_str(r[idx_t]) / 1000.0 if idx_t is not None else 0.0
                        d_l = _parse_num_str(r[idx_d]) / 1000.0 if idx_d is not None else 0.0
                        tot_l = _parse_num_str(r[idx_tot]) / 1000.0 if idx_tot is not None else (f_l + t_l + d_l)
                        day_map[code] = (f_l, t_l, d_l, tot_l)
                    if day_map:
                        break
                elif "aaData" in payload and payload["aaData"]:
                    for r in payload["aaData"]:
                        if len(r) < 20:
                            continue
                        code = str(r[0]).strip()
                        f_l = _parse_num_str(r[10]) / 1000.0
                        t_l = _parse_num_str(r[13]) / 1000.0
                        d_l = _parse_num_str(r[22]) / 1000.0 if len(r) > 22 else (_parse_num_str(r[16]) + _parse_num_str(r[19])) / 1000.0
                        tot_l = _parse_num_str(r[-1]) / 1000.0
                        day_map[code] = (f_l, t_l, d_l, tot_l)
                    if day_map:
                        break
            except Exception:
                continue
    return day_map




# ==========================================
# 大盤與個股資券變化抓取引擎 (TWSE MI_MARGN & TPEx)
# ==========================================
@st.cache_data(ttl=7200, show_spinner=False)
def fetch_margin_day_cached(exchange, yyyymmdd):
    """
    抓取單日上市(TWSE MI_MARGN)或上櫃(TPEX margin/balance)資券資料。
    回傳 dict:
      {
        "summary": {"margin_amt_yi": float, "margin_amt_chg_yi": float, "margin_lots": int, "margin_lots_chg": int, "short_lots": int, "short_lots_chg": int} | None,
        "stocks": {code: (m_prev, m_curr, m_chg, s_prev, s_curr, s_chg)}
      }
    """
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    ssl_ctx = ssl._create_unverified_context()
    result = {"summary": None, "stocks": {}}

    if exchange == "TWSE":
        url = f"https://www.twse.com.tw/rwd/zh/marginTrading/MI_MARGN?date={yyyymmdd}&selectType=ALL&response=json"
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=10, context=ssl_ctx) as resp:
                payload = json.loads(resp.read().decode("utf-8", errors="ignore"))
            if payload.get("stat") != "OK":
                return result

            tables = payload.get("tables", [])
            # 1) 解析大盤信用交易統計表 (tables[0])
            for tbl in tables:
                rows_d = tbl.get("data", [])
                fields = [str(f).strip() for f in tbl.get("fields", [])]
                if not rows_d:
                    continue
                if any("融資金額" in str(r[0]) for r in rows_d[:5]):
                    idx_prev = next((i for i, f in enumerate(fields) if "前日餘額" in f), 4)
                    idx_curr = next((i for i, f in enumerate(fields) if "今日餘額" in f), 5)
                    summ = {
                        "margin_amt_yi": 0.0, "margin_amt_chg_yi": 0.0,
                        "margin_lots": 0, "margin_lots_chg": 0,
                        "short_lots": 0, "short_lots_chg": 0,
                    }
                    for r in rows_d:
                        if len(r) <= max(idx_prev, idx_curr):
                            continue
                        lbl = str(r[0]).strip()
                        v_prev = _parse_num_str(r[idx_prev])
                        v_curr = _parse_num_str(r[idx_curr])
                        if "融資金額" in lbl:
                            # 單位：仟元 -> 轉億元 (/ 100000)
                            summ["margin_amt_yi"] = round(v_curr / 100000.0, 2)
                            summ["margin_amt_chg_yi"] = round((v_curr - v_prev) / 100000.0, 2)
                        elif "融資" in lbl and "交易單位" in lbl:
                            summ["margin_lots"] = int(round(v_curr))
                            summ["margin_lots_chg"] = int(round(v_curr - v_prev))
                        elif "融券" in lbl and "交易單位" in lbl:
                            summ["short_lots"] = int(round(v_curr))
                            summ["short_lots_chg"] = int(round(v_curr - v_prev))
                    if summ["margin_amt_yi"] > 0:
                        result["summary"] = summ
                # 2) 解析個股融資融券明細表 (通常 tables[1])
                elif len(rows_d) > 30 and any("前日餘額" in f for f in fields):
                    prev_idxs = [i for i, f in enumerate(fields) if "前日餘額" in f]
                    curr_idxs = [i for i, f in enumerate(fields) if "今日餘額" in f]
                    m_p_i = prev_idxs[0] if len(prev_idxs) >= 1 else 5
                    m_c_i = curr_idxs[0] if len(curr_idxs) >= 1 else 6
                    s_p_i = prev_idxs[1] if len(prev_idxs) >= 2 else 11
                    s_c_i = curr_idxs[1] if len(curr_idxs) >= 2 else 12
                    for r in rows_d:
                        if len(r) <= max(m_p_i, m_c_i, s_p_i, s_c_i):
                            continue
                        code = str(r[0]).strip()
                        if not code:
                            continue
                        mp = int(round(_parse_num_str(r[m_p_i])))
                        mc = int(round(_parse_num_str(r[m_c_i])))
                        sp = int(round(_parse_num_str(r[s_p_i])))
                        sc = int(round(_parse_num_str(r[s_c_i])))
                        result["stocks"][code] = (mp, mc, mc - mp, sp, sc, sc - sp)
        except Exception:
            pass

    elif exchange == "TPEX":
        y, m, d = int(yyyymmdd[:4]), yyyymmdd[4:6], yyyymmdd[6:8]
        ad_date = f"{y}/{m}/{d}"
        roc_date = f"{y - 1911}/{m}/{d}"
        urls = [
            f"https://www.tpex.org.tw/www/zh-tw/margin/balance?date={ad_date}&id=&response=json",
            f"https://www.tpex.org.tw/web/stock/margin_trading/margin_balance/margin_bal_result.php?l=zh-tw&o=json&d={roc_date}",
        ]
        for url in urls:
            try:
                req = urllib.request.Request(url, headers=headers)
                with urllib.request.urlopen(req, timeout=10, context=ssl_ctx) as resp:
                    payload = json.loads(resp.read().decode("utf-8", errors="ignore"))
                rows_d = []
                fields = []
                if "tables" in payload and payload["tables"]:
                    tbl = payload["tables"][0]
                    fields = [str(f).strip() for f in tbl.get("fields", [])]
                    rows_d = tbl.get("data", [])
                elif "aaData" in payload and payload["aaData"]:
                    rows_d = payload["aaData"]

                if not rows_d:
                    continue

                # 櫃買標準欄位：0=代號, 2=前資餘額, 6=資餘額, 10=前券餘額, 14=券餘額
                m_p_i = next((i for i, f in enumerate(fields) if "前資餘額" in f), 2)
                m_c_i = next((i for i, f in enumerate(fields) if f in ("資餘額", "融資餘額") or ("資餘額" in f and "前" not in f)), 6)
                s_p_i = next((i for i, f in enumerate(fields) if "前券餘額" in f), 10)
                s_c_i = next((i for i, f in enumerate(fields) if f in ("券餘額", "融券餘額") or ("券餘額" in f and "前" not in f)), 14)

                for r in rows_d:
                    if len(r) <= max(m_p_i, m_c_i, s_p_i, s_c_i):
                        continue
                    code = str(r[0]).strip()
                    if not code:
                        continue
                    mp = int(round(_parse_num_str(r[m_p_i])))
                    mc = int(round(_parse_num_str(r[m_c_i])))
                    sp = int(round(_parse_num_str(r[s_p_i])))
                    sc = int(round(_parse_num_str(r[s_c_i])))
                    result["stocks"][code] = (mp, mc, mc - mp, sp, sc, sc - sp)
                if result["stocks"]:
                    break
            except Exception:
                continue
    return result


@st.cache_data(ttl=1800, show_spinner=False)
def fetch_twse_taiex_history_cached(period_key):
    """
    直接從證交所官方 FMTQIK (每日市場成交資訊) 抓取近 2 個月加權指數收盤紀錄。
    比 Yahoo Finance ^TWII 更準時且精確，避免 ^TWII 漏更新前一日大盤收盤。
    回傳 list of dict: [{"date": "YYYY-MM-DD", "yyyymmdd": "YYYYMMDD", "close": float, "chg_pts": float}, ...] (由舊到新)
    """
    exclude_today, today_tw_date, _, _ = get_market_data_cutoff_info()
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
    ssl_ctx = ssl._create_unverified_context()

    # 抓本月與前兩個月，確保至少有 25~40 個交易日可算 20MA
    months = []
    y, m = today_tw_date.year, today_tw_date.month
    for _ in range(3):
        months.append(f"{y:04d}{m:02d}01")
        m -= 1
        if m == 0:
            m = 12
            y -= 1

    rows_by_date = {}
    for ym01 in reversed(months):
        url = f"https://www.twse.com.tw/rwd/zh/afterTrading/FMTQIK?date={ym01}&response=json"
        try:
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=10, context=ssl_ctx) as resp:
                payload = json.loads(resp.read().decode("utf-8", errors="ignore"))
            if payload.get("stat") != "OK" or not payload.get("data"):
                continue
            fields = [str(f).strip() for f in payload.get("fields", [])]
            idx_close = next((i for i, f in enumerate(fields) if "加權股價指數" in f or "指數" in f), 4)
            idx_chg = next((i for i, f in enumerate(fields) if "漲跌點數" in f), 5)
            for r in payload["data"]:
                if len(r) <= max(idx_close, idx_chg):
                    continue
                roc_str = str(r[0]).strip()
                parts = roc_str.split("/")
                if len(parts) != 3:
                    continue
                ad_y = int(parts[0]) + 1911
                ad_m = int(parts[1])
                ad_d = int(parts[2])
                dt_obj = datetime(ad_y, ad_m, ad_d).date()
                if exclude_today and dt_obj >= today_tw_date:
                    continue
                close_v = _parse_num_str(r[idx_close])
                chg_v = _parse_num_str(r[idx_chg])
                if close_v <= 0:
                    continue
                d_iso = dt_obj.strftime("%Y-%m-%d")
                rows_by_date[d_iso] = {
                    "date": d_iso,
                    "yyyymmdd": dt_obj.strftime("%Y%m%d"),
                    "close": round(close_v, 2),
                    "chg_pts": round(chg_v, 2),
                }
        except Exception:
            continue

    return [rows_by_date[k] for k in sorted(rows_by_date.keys())]


def evaluate_market_margin_signal(taiex_chg_1d, taiex_chg_5d, taiex_bias_20ma, chg_1d_yi, chg_5d_yi, curr_amt_yi, short_curr, short_chg_1d, short_chg_5d, margin_date):
    """根據大盤指數位置與融資融券變化，輸出明確的「過熱警訊」或「落底訊號」白話說明（嚴格區分單日融資增減方向）"""
    # 1) 過熱警訊 A：大盤下跌但融資逆勢增加（散戶接刀）
    is_overheat_knife = (
        (taiex_chg_1d <= -0.5 or taiex_chg_5d <= -1.5)
        and (chg_1d_yi >= 15.0 or (chg_1d_yi > 0 and chg_5d_yi >= 30.0))
    )
    # 2) 過熱警訊 B：大盤高檔或單日急漲伴隨融資暴增追價
    is_overheat_chase = (
        ((taiex_bias_20ma >= 2.5 or taiex_chg_5d >= 1.5) and (chg_1d_yi >= 25.0 or (chg_1d_yi > 0 and chg_5d_yi >= 60.0)))
        or (taiex_chg_1d >= 1.5 and chg_1d_yi >= 50.0)
    )
    # 3) 落底訊號 A：止跌轉強＋價漲資減（必須當日融資未增加 chg_1d_yi <= 0）
    is_bottom_turn = (
        (taiex_chg_1d > 0.0)
        and (chg_1d_yi <= 0.0)
        and (chg_1d_yi <= -3.0 or chg_5d_yi <= -20.0)
    )
    # 4) 落底訊號 B：跌深殺融資清洗浮額（必須當日融資未增加 chg_1d_yi <= 0）
    is_bottom_flush = (
        (taiex_chg_1d <= -0.5 or taiex_bias_20ma <= 0.0 or taiex_chg_5d <= -1.0)
        and (chg_1d_yi <= 0.0)
        and (chg_1d_yi <= -25.0 or chg_5d_yi <= -60.0)
    )

    short_note = (
        f"融券餘額 {short_curr:,} 張（單日 {short_chg_1d:+,} 張、近5日 {short_chg_5d:+,} 張"
        + ("，空單增加具潛在軋空力道" if short_chg_5d >= 10000 else "，空單回補減壓" if short_chg_5d <= -10000 else "，變動平穩")
        + "）。"
    )

    if is_overheat_knife:
        signal_level = "OVERHEAT"
        signal_badge = "🚨 過熱警訊（大盤下跌＋融資逆勢接刀）"
        summary_text = (
            f"大盤走弱（單日 {taiex_chg_1d:+.2f}% / 近5日 {taiex_chg_5d:+.2f}%），但上市融資餘額卻逆勢增加"
            f"（單日 {chg_1d_yi:+.2f} 億、近5日 {chg_5d_yi:+.2f} 億，總餘額 {curr_amt_yi:,.2f} 億），"
            f"屬於典型「散戶逢低接刀、浮額未清」警訊！籌碼趨於凌亂、反彈易遇解套賣壓，建議降低持股水位並嚴守停損。{short_note}"
        )
    elif is_overheat_chase:
        signal_level = "OVERHEAT"
        signal_badge = "🚨 過熱警訊（大盤急漲／高檔＋融資暴增追價）"
        summary_text = (
            f"大盤上漲（單日 {taiex_chg_1d:+.2f}% / 距月線 {taiex_bias_20ma:+.2f}%），"
            f"上市融資餘額單日大幅回補增加（單日 {chg_1d_yi:+.2f} 億、近5日 {chg_5d_yi:+.2f} 億，總餘額 {curr_amt_yi:,.2f} 億），"
            f"顯示短線散戶追價積極、浮額快速回流，慎防隔日劇烈震盪或主力趁反彈調節，切勿過度追高。{short_note}"
        )
    elif is_bottom_turn:
        signal_level = "BOTTOM"
        signal_badge = "🟢 落底訊號（止跌轉強＋價漲資減）"
        summary_text = (
            f"大盤止跌回穩（單日 {taiex_chg_1d:+.2f}% / 近5日 {taiex_chg_5d:+.2f}%），"
            f"同時融資持續退場（單日 {chg_1d_yi:+.2f} 億、近5日 {chg_5d_yi:+.2f} 億，總餘額 {curr_amt_yi:,.2f} 億），"
            f"呈現最健康的「價漲資減」結構！代表浮額已清洗、籌碼流向法人與大戶手中，具備落底轉強契機。{short_note}"
        )
    elif is_bottom_flush:
        signal_level = "BOTTOM"
        signal_badge = "🟢 落底訊號（跌深殺融資清洗浮額）"
        summary_text = (
            f"大盤回檔修正（單日 {taiex_chg_1d:+.2f}% / 距月線 {taiex_bias_20ma:+.2f}%）伴隨明顯「殺融資」"
            f"（單日 {chg_1d_yi:+.2f} 億、近5日 {chg_5d_yi:+.2f} 億，總餘額 {curr_amt_yi:,.2f} 億），"
            f"恐慌浮額正加速出清、籌碼逐步沉澱，留意後續止跌收紅轉為「價漲資減」的落底切入點。{short_note}"
        )
    else:
        signal_level = "NEUTRAL"
        if taiex_chg_1d >= 0 and chg_1d_yi >= 0:
            sub_lbl = "價漲資溫和增（多方跟進）"
        elif taiex_chg_1d < 0 and chg_1d_yi < 0:
            sub_lbl = "價跌資同步減（良性回檔換手）"
        else:
            sub_lbl = "多空平衡整理"
        signal_badge = f"⚖️ 溫和常態（{sub_lbl}）"
        summary_text = (
            f"大盤收盤變動 {taiex_chg_1d:+.2f}%（近5日 {taiex_chg_5d:+.2f}%、距月線 {taiex_bias_20ma:+.2f}%），"
            f"融資餘額 {curr_amt_yi:,.2f} 億元（單日 {chg_1d_yi:+.2f} 億、近5日 {chg_5d_yi:+.2f} 億），"
            f"目前未出現極端過熱或恐慌斷頭，屬常態籌碼換手格局，可依個股技術面與法人動向操作。{short_note}"
        )

    return {
        "signal_level": signal_level,
        "signal_badge": signal_badge,
        "summary_text": summary_text,
        "margin_date": margin_date,
    }


@st.cache_data(ttl=1800, show_spinner=False)
def get_market_margin_analysis_cached(period_key):
    """優先使用證交所官方 FMTQIK 加權指數 + MI_MARGN 融資融券統計，確保日期與漲跌幅 100% 準確對齊"""
    exclude_today, today_tw_date, _, _ = get_market_data_cutoff_info()
    taiex_rows = fetch_twse_taiex_history_cached(period_key)

    # 備援：僅當證交所 FMTQIK 連線失敗時才退回 Yahoo Finance ^TWII
    if len(taiex_rows) < 2:
        try:
            df_idx = yf.download(tickers="^TWII", period="2mo", interval="1d", auto_adjust=False, progress=False)
            if df_idx is not None and not df_idx.empty:
                if isinstance(df_idx.columns, pd.MultiIndex):
                    df_idx.columns = [c[0] for c in df_idx.columns]
                df_idx = df_idx.dropna(subset=["Close"]).copy()
                if exclude_today:
                    df_idx = df_idx[pd.to_datetime(df_idx.index).date < today_tw_date]
                taiex_rows = []
                for dt, row in df_idx.iterrows():
                    d_iso = dt.strftime("%Y-%m-%d") if hasattr(dt, "strftime") else str(dt)[:10]
                    taiex_rows.append({
                        "date": d_iso,
                        "yyyymmdd": d_iso.replace("-", ""),
                        "close": round(float(row["Close"]), 2),
                        "chg_pts": 0.0,
                    })
        except Exception:
            pass

    if taiex_rows:
        candidate_dates = [r["yyyymmdd"] for r in taiex_rows[-10:]]
    else:
        base_d = today_tw_date - timedelta(days=1 if exclude_today else 0)
        candidate_dates = [(base_d - timedelta(days=i)).strftime("%Y%m%d") for i in range(10, -1, -1)]

    # 取得最近最多 5 個有公布資券的交易日
    valid_days = []
    for d_str in reversed(candidate_dates):
        res = fetch_margin_day_cached("TWSE", d_str)
        if res and res.get("summary"):
            valid_days.append((d_str, res["summary"]))
            if len(valid_days) >= 5:
                break

    # 預設取最新一個加權指數交易日；若資券日有對應日期則對齊同日指數計算資券關係
    def _calc_taiex_metrics(rows, target_date_iso=None):
        if not rows or len(rows) < 2:
            return "--", 0.0, 0.0, 0.0, 0.0
        idx = len(rows) - 1
        if target_date_iso:
            for k in range(len(rows) - 1, -1, -1):
                if rows[k]["date"] <= target_date_iso:
                    idx = k
                    break
        if idx < 1:
            idx = len(rows) - 1
        c_now = rows[idx]["close"]
        c_prev = rows[idx - 1]["close"]
        chg_1d = round((c_now / c_prev - 1.0) * 100.0, 2) if c_prev > 0 else 0.0
        c_5 = rows[max(0, idx - 5)]["close"]
        chg_5d = round((c_now / c_5 - 1.0) * 100.0, 2) if c_5 > 0 else chg_1d
        win20 = [x["close"] for x in rows[max(0, idx - 19): idx + 1]]
        ma20 = sum(win20) / len(win20) if win20 else c_now
        bias20 = round((c_now / ma20 - 1.0) * 100.0, 2) if ma20 > 0 else 0.0
        return rows[idx]["date"], c_now, chg_1d, chg_5d, bias20

    if not valid_days:
        t_date, t_close, t_1d, t_5d, t_bias = _calc_taiex_metrics(taiex_rows)
        return {
            "available": False,
            "signal_level": "NEUTRAL",
            "signal_badge": "ℹ️ 大盤資券資料暫未取得",
            "summary_text": "目前證交所融資融券統計暫無回應或尚未公告，請稍後重新整理。",
            "taiex_date": t_date,
            "margin_date": "--",
            "taiex_close": t_close,
            "taiex_chg_1d": t_1d,
            "taiex_chg_5d": t_5d,
            "taiex_bias_20ma": t_bias,
        }

    latest_d_raw, latest_s = valid_days[0]
    margin_date = f"{latest_d_raw[:4]}-{latest_d_raw[4:6]}-{latest_d_raw[6:8]}"
    taiex_date, taiex_close, taiex_chg_1d, taiex_chg_5d, taiex_bias_20ma = _calc_taiex_metrics(taiex_rows, target_date_iso=margin_date)

    curr_amt_yi = latest_s["margin_amt_yi"]
    chg_1d_yi = latest_s["margin_amt_chg_yi"]
    chg_5d_yi = round(sum(x[1]["margin_amt_chg_yi"] for x in valid_days), 2)
    short_curr = latest_s["short_lots"]
    short_chg_1d = latest_s["short_lots_chg"]
    short_chg_5d = int(sum(x[1]["short_lots_chg"] for x in valid_days))

    diag = evaluate_market_margin_signal(
        taiex_chg_1d, taiex_chg_5d, taiex_bias_20ma,
        chg_1d_yi, chg_5d_yi, curr_amt_yi,
        short_curr, short_chg_1d, short_chg_5d, margin_date,
    )
    diag.update({
        "available": True,
        "taiex_date": taiex_date,
        "taiex_close": taiex_close,
        "taiex_chg_1d": taiex_chg_1d,
        "taiex_chg_5d": taiex_chg_5d,
        "taiex_bias_20ma": taiex_bias_20ma,
        "curr_amt_yi": curr_amt_yi,
        "chg_1d_yi": chg_1d_yi,
        "chg_5d_yi": chg_5d_yi,
        "short_curr": short_curr,
        "short_chg_1d": short_chg_1d,
        "short_chg_5d": short_chg_5d,
        "days_counted": len(valid_days),
    })
    return diag


def evaluate_stock_margin_signal(pct_1d, pct_5d, bias20, vol_ratio, m_curr, m_chg_1d, m_chg_5d, s_curr, s_chg_5d, margin_date):
    """
    判斷個股資券變化（籌碼乾淨度與主力洗盤／出貨動向）：
    - 過熱警訊：股價高檔爆量且融資暴增，或股價下跌但融資一路增加（散戶接刀）
    - 落底訊號：股價跌深後出現「殺融資」清洗浮額，隨後止跌轉為「價漲資減」
    """
    m_base = max(m_curr - m_chg_5d, 1)
    m_chg_5d_pct = round((m_chg_5d / m_base) * 100.0, 2) if m_curr > 0 or m_chg_5d != 0 else 0.0
    m_chg_1d_pct = round((m_chg_1d / max(m_curr - m_chg_1d, 1)) * 100.0, 2) if m_curr > 0 or m_chg_1d != 0 else 0.0
    sm_ratio = round((s_curr / m_curr) * 100.0, 2) if m_curr > 0 else 0.0

    # 1) 過熱警訊 A：高檔爆量且融資暴增
    overheat_high_vol = (
        (bias20 >= 5.0 or pct_5d >= 6.0)
        and (vol_ratio >= 1.8)
        and (m_chg_5d > 0 and (m_chg_5d_pct >= 15.0 or (m_chg_5d >= 200 and m_chg_5d_pct >= 8.0) or m_chg_1d_pct >= 10.0))
    )
    # 2) 過熱警訊 B：股價下跌但融資一路增加（散戶接刀）
    overheat_catching_knife = (
        (pct_1d < 0 or pct_5d <= -2.0 or bias20 < 0)
        and (m_chg_5d > 0 and (m_chg_5d_pct >= 5.0 or (m_chg_1d > 0 and m_chg_5d >= 30)))
    )
    # 3) 落底訊號 A：止跌轉為「價漲資減」（清洗浮額後轉強）
    bottom_price_up_margin_down = (
        (pct_1d > 0 or pct_5d >= 0.5)
        and (m_chg_5d < 0 and (m_chg_5d_pct <= -3.0 or (m_chg_1d < 0 and m_chg_5d <= -20)))
    )
    # 4) 落底訊號 B：跌深出現「殺融資」清洗浮額
    bottom_margin_flush = (
        (pct_5d <= -3.0 or bias20 <= -2.0 or pct_1d <= -2.0)
        and (m_chg_5d < 0 and (m_chg_5d_pct <= -8.0 or m_chg_1d_pct <= -3.5))
    )

    squeeze_note = f"；🔥 券資比達 {sm_ratio:.1f}%（近5日融券 {s_chg_5d:+d} 張），具潛在軋空題材" if sm_ratio >= 15.0 else ""

    if overheat_high_vol:
        short_tag = "🚨 過熱警訊(高檔爆量資暴增)"
        detail = (
            f"🚨 過熱警訊（高檔爆量＋融資暴增）！股價位於相對高檔（距月線 {bias20:+.2f}%、量能 {vol_ratio} 倍），"
            f"近5日融資大增 {m_chg_5d:+d} 張（{m_chg_5d_pct:+.1f}%，單日 {m_chg_1d:+d} 張，餘額 {m_curr:,} 張），"
            f"散戶追價過熱、浮額加重，留意主力趁爆量出貨或隔日劇烈震盪{squeeze_note}（資券日 {margin_date}）。"
        )
    elif overheat_catching_knife:
        short_tag = "🚨 過熱警訊(價跌資增接刀)"
        detail = (
            f"🚨 過熱警訊（股價下跌＋融資一路增加）！股價走弱（單日 {pct_1d:+.2f}% / 近5日 {pct_5d:+.2f}%），"
            f"但近5日融資逆勢增加 {m_chg_5d:+d} 張（{m_chg_5d_pct:+.1f}%，單日 {m_chg_1d:+d} 張，餘額 {m_curr:,} 張），"
            f"屬於典型「散戶逢低接刀」，籌碼趨於凌亂、反彈易遭解套賣壓打壓，不建議急於搶短{squeeze_note}（資券日 {margin_date}）。"
        )
    elif bottom_price_up_margin_down:
        short_tag = "🟢 落底訊號(價漲資減)"
        detail = (
            f"🟢 落底訊號（止跌轉為價漲資減）！股價收紅轉強（單日 {pct_1d:+.2f}%），"
            f"同時近5日融資減少 {m_chg_5d:+d} 張（{m_chg_5d_pct:+.1f}%，單日 {m_chg_1d:+d} 張，餘額 {m_curr:,} 張），"
            f"顯示浮額已獲清洗、散戶籌碼轉移至主力大戶手中，籌碼乾淨度高、有利波段延續{squeeze_note}（資券日 {margin_date}）。"
        )
    elif bottom_margin_flush:
        short_tag = "🟢 落底訊號(殺融資洗浮額)"
        detail = (
            f"🟢 落底訊號（跌深殺融資清洗浮額）！股價回檔期間近5日融資大減 {m_chg_5d:+d} 張（{m_chg_5d_pct:+.1f}%，餘額 {m_curr:,} 張），"
            f"浮額正加速清洗、籌碼逐步沉澱；後續若出現止跌紅K轉為「價漲資減」，將是極佳的切入觀察訊號{squeeze_note}（資券日 {margin_date}）。"
        )
    elif m_curr == 0 and m_chg_5d == 0:
        short_tag = "⚪ 無融資餘額"
        detail = f"目前無信用交易融資餘額（或未開放信用交易），籌碼結構單純不受融資斷頭干擾（資券日 {margin_date}）。"
    else:
        short_tag = "⚖️ 資券溫和穩定"
        detail = (
            f"⚖️ 籌碼變動溫和：融資餘額 {m_curr:,} 張（單日 {m_chg_1d:+d} 張、近5日 {m_chg_5d:+d} 張 / {m_chg_5d_pct:+.1f}%），"
            f"融券 {s_curr:,} 張（券資比 {sm_ratio:.1f}%），無明顯散戶過熱或接刀異常，屬正常換手{squeeze_note}（資券日 {margin_date}）。"
        )

    return short_tag, detail, sm_ratio


def calc_streak(series_vals):
    if not series_vals:
        return 0
    last_val = series_vals[-1]
    if abs(last_val) < 0.5:
        return 0
    sign = 1 if last_val > 0 else -1
    streak = 0
    for v in reversed(series_vals):
        if (sign > 0 and v >= 0.5) or (sign < 0 and v <= -0.5):
            streak += 1
        else:
            break
    return streak * sign


def enrich_rows_with_insti(rows_list):
    """補齊證交所/櫃買中心三大法人買賣超、推估成本與日週轉率"""
    if not rows_list:
        return

    ex_date_pairs = set()
    for r in rows_list:
        m_lbl = r.get("市場別", "")
        ex = "TWSE" if "上市" in m_lbl else "TPEX" if "上櫃" in m_lbl else None
        r["_exchange"] = ex
        if ex:
            for dt_str, _ in r.get("_recent_20_prices", []):
                ex_date_pairs.add((ex, dt_str))

    insti_results = {}
    if ex_date_pairs:
        def _fetch(pair):
            time.sleep(0.05)
            return pair, fetch_insti_day_cached(pair[0], pair[1])

        with ThreadPoolExecutor(max_workers=4) as pool:
            futs = [pool.submit(_fetch, p) for p in ex_date_pairs]
            for fut in as_completed(futs):
                pair, dmap = fut.result()
                insti_results[pair] = dmap

    # 同步抓取最近 6 個交易日資券資料（上市 TWSE MI_MARGN / 上櫃 TPEx）
    margin_pairs = set()
    for r in rows_list:
        ex = r.get("_exchange")
        if ex:
            for dt_str, _ in r.get("_recent_20_prices", [])[-6:]:
                margin_pairs.add((ex, dt_str))

    margin_results = {}
    if margin_pairs:
        def _fetch_m(pair):
            time.sleep(0.05)
            return pair, fetch_margin_day_cached(pair[0], pair[1])

        with ThreadPoolExecutor(max_workers=4) as pool:
            futs = [pool.submit(_fetch_m, p) for p in margin_pairs]
            for fut in as_completed(futs):
                pair, mres = fut.result()
                margin_results[pair] = mres

    profile_map = fetch_company_profile_map()
    for r in rows_list:
        ex = r.get("_exchange")
        code = r.get("股票代碼", "")
        recent_prices = r.get("_recent_20_prices", [])
        c_today = r["收盤價"]
        vwap_20 = r["20日VWAP資金成本"]
        poc_60 = r["60日籌碼密集區(POC)"]
        vwap_gap = round((c_today / vwap_20 - 1.0) * 100.0, 2)
        poc_pos = "下方（形成底部換手鐵板支撐）" if c_today >= poc_60 else "上方（留意前波套牢解套賣壓）"

        prof = profile_map.get(code, {})
        total_lots = prof.get("lots", 0.0)
        biz_text = prof.get("business", "")
        if not biz_text:
            # 若 OpenAPI 無資料，退回嘗試抓取 Yahoo Finance 產業別
            try:
                suf = ".TW" if ex == "TWSE" else ".TWO"
                yf_info = yf.Ticker(f"{code}{suf}").info or {}
                ind_en = yf_info.get("industry") or yf_info.get("sector") or ""
                biz_text = f"【{ind_en}】" if ind_en else "尚無公開營業概況資料"
            except Exception:
                biz_text = "尚無公開營業概況資料"
        r["營業概況(主力業務)"] = biz_text
        if total_lots > 0:
            turnover_pct = round(r["成交量(張)"] / total_lots * 100.0, 2)
            r["日週轉率(%)"] = turnover_pct
            if turnover_pct >= 8.0:
                to_note = f"（⚠️ 日週轉率達 {turnover_pct}% >= 8%，當沖或換手極度過熱，留意隔日震盪！）"
            elif turnover_pct >= 0.5:
                to_note = f"（日週轉率 {turnover_pct}%，屬健康換手點火區間）"
            else:
                to_note = f"（日週轉率 {turnover_pct}%，屬溫和低換手或大型權值股）"
        else:
            r["日週轉率(%)"] = "N/A"
            to_note = ""

        vwap_line = (
            f"• 【量價主力成本】20日量價加權資金成本(VWAP)為 『{vwap_20} 元』（現價距資金成本 {vwap_gap:+.2f}%）；"
            f"近60日最大量籌碼密集區(POC)落在 『{poc_60} 元』，位於現價{poc_pos}；"
            f"今日成交金額 {r['日成交金額(萬元)']} 萬元{to_note}。"
        )

        if not ex or not recent_prices:
            r["近5日法人合計(張)"] = "興櫃無揭露"
            r["近20日法人合計(張)"] = "興櫃無揭露"
            r["外資/投信連買天數"] = "興櫃無揭露"
            r["法人20日買超成本"] = "無揭露(看VWAP)"
            insti_note = "• 【三大法人】興櫃市場未逐日揭露三大法人買賣超，建議以 20日VWAP 與 60日POC 判斷主力成本。"
        else:
            f_list, t_list, d_list, tot_list = [], [], [], []
            buy_val_tot, buy_vol_tot = 0.0, 0.0
            buy_val_f, buy_vol_f = 0.0, 0.0
            buy_val_t, buy_vol_t = 0.0, 0.0
            valid_days = 0

            for dt_str, tp_price in recent_prices:
                day_map = insti_results.get((ex, dt_str), {})
                if day_map:
                    valid_days += 1
                f_l, t_l, d_l, tot_l = day_map.get(code, (0.0, 0.0, 0.0, 0.0))
                f_list.append(f_l)
                t_list.append(t_l)
                d_list.append(d_l)
                tot_list.append(tot_l)

                if tot_l > 0:
                    buy_val_tot += tot_l * tp_price
                    buy_vol_tot += tot_l
                if f_l > 0:
                    buy_val_f += f_l * tp_price
                    buy_vol_f += f_l
                if t_l > 0:
                    buy_val_t += t_l * tp_price
                    buy_vol_t += t_l

            if valid_days == 0:
                r["近5日法人合計(張)"] = "連線逾時"
                r["近20日法人合計(張)"] = "連線逾時"
                r["外資/投信連買天數"] = "連線逾時"
                r["法人20日買超成本"] = "暫無法取得"
                insti_note = "• 【三大法人】證交所/櫃買中心暫時連線逾時，請以 20日VWAP 與 60日POC 作為成本參考。"
            else:
                tot_5d = int(round(sum(tot_list[-5:])))
                tot_20d = int(round(sum(tot_list)))
                f_5d = int(round(sum(f_list[-5:])))
                t_5d = int(round(sum(t_list[-5:])))
                d_5d = int(round(sum(d_list[-5:])))
                f_20d = int(round(sum(f_list)))
                t_20d = int(round(sum(t_list)))

                f_streak = calc_streak(f_list)
                t_streak = calc_streak(t_list)
                cost_tot = round(buy_val_tot / buy_vol_tot, 2) if buy_vol_tot >= 1 else None
                cost_f = round(buy_val_f / buy_vol_f, 2) if buy_vol_f >= 1 else None
                cost_t = round(buy_val_t / buy_vol_t, 2) if buy_vol_t >= 1 else None

                def _fmt_s(n, s):
                    return f"{n}連買{s}日" if s > 0 else f"{n}連賣{abs(s)}日" if s < 0 else f"{n}持平"

                streak_str = f"{_fmt_s('外資', f_streak)} / {_fmt_s('投信', t_streak)}"
                r["近5日法人合計(張)"] = tot_5d
                r["近20日法人合計(張)"] = tot_20d
                r["外資/投信連買天數"] = streak_str
                r["法人20日買超成本"] = cost_tot if cost_tot is not None else "近20日無買超"

                if cost_tot is not None:
                    gap_pct = round((c_today / cost_tot - 1.0) * 100.0, 2)
                    if -3.0 <= gap_pct <= 5.0:
                        cost_note = f"三大法人近20日買超推估成本約 『{cost_tot} 元』（現價僅差 {gap_pct:+.2f}%，與主力成本高度同步、具護盤動機！）"
                    elif gap_pct > 5.0:
                        cost_note = f"三大法人近20日買超推估成本約 『{cost_tot} 元』（現價已高出 +{gap_pct:.2f}%，留意法人獲利調節）。"
                    else:
                        cost_note = f"三大法人近20日買超推估成本約 『{cost_tot} 元』（現價低於成本 {gap_pct:.2f}%，處於主力攤平或套牢區）。"
                else:
                    cost_note = "近 20 日三大法人合計未出現明顯淨買超日。"

                sub = []
                if cost_f is not None:
                    sub.append(f"外資均價 {cost_f} 元(20日 {f_20d:+d} 張)")
                if cost_t is not None:
                    sub.append(f"投信均價 {cost_t} 元(20日 {t_20d:+d} 張)")
                sub_str = ("；" + "、".join(sub)) if sub else ""

                if tot_5d > 0 and tot_20d > 0:
                    verdict = f"🔥 法人偏多助攻！近5日買超 {tot_5d:+d} 張（外資 {f_5d:+d} / 投信 {t_5d:+d} / 自營 {d_5d:+d}），近20日累計 {tot_20d:+d} 張（{streak_str}）。"
                elif tot_5d > 0 and tot_20d <= 0:
                    verdict = f"🌱 法人由賣轉買點火！近5日轉買超 {tot_5d:+d} 張（外資 {f_5d:+d} / 投信 {t_5d:+d}），開始回補（{streak_str}）。"
                elif tot_5d < 0 and tot_20d > 0:
                    verdict = f"⚠️ 短線逢高調節！近20日雖買超 {tot_20d:+d} 張，但近5日轉賣超 {tot_5d:+d} 張（{streak_str}），留意是否換手。"
                else:
                    verdict = f"❄️ 法人籌碼偏保守！近5日賣超 {tot_5d:+d} 張、近20日 {tot_20d:+d} 張（{streak_str}），若技術面突破多屬內資主力主導。"

                insti_note = f"• 【三大法人動向】{verdict}\n• 【法人推估成本】{cost_note}{sub_str}"

        # 計算個股資券變化（近 5 個已公布資券交易日）
        if not ex or not recent_prices:
            r["近5日融資增減(張)"] = "興櫃無信用"
            r["券資比(%)"] = "N/A"
            r["資券籌碼訊號"] = "⚪ 興櫃無信用交易"
            margin_note = "• 【個股資券變化】興櫃市場未開放信用交易（融資／融券），請以 20日VWAP 與 60日POC 觀察籌碼。"
        else:
            m_days = []
            for dt_str, _ in recent_prices[-6:]:
                mres = margin_results.get((ex, dt_str), {})
                stks = mres.get("stocks", {}) if isinstance(mres, dict) else {}
                if stks:
                    m_days.append((dt_str, stks.get(code)))
            if not m_days:
                r["近5日融資增減(張)"] = "暫無資料"
                r["券資比(%)"] = "N/A"
                r["資券籌碼訊號"] = "ℹ️ 資券暫未公告"
                margin_note = "• 【個股資券變化】證交所／櫃買中心資券統計暫未公告或連線逾時（資券通常於晚間更新）。"
            else:
                m_days_5 = m_days[-5:]
                last_dt_raw, last_tup = m_days_5[-1]
                m_date_fmt = f"{last_dt_raw[:4]}-{last_dt_raw[4:6]}-{last_dt_raw[6:8]}"
                if last_tup is None:
                    m_curr, m_chg_1d, s_curr = 0, 0, 0
                else:
                    _, m_curr, m_chg_1d, _, s_curr, _ = last_tup
                m_chg_5d = sum((t[2] if t is not None else 0) for _, t in m_days_5)
                s_chg_5d = sum((t[5] if t is not None else 0) for _, t in m_days_5)
                p_5d_base = recent_prices[-6][1] if len(recent_prices) >= 6 else recent_prices[0][1]
                pct_5d = round((c_today / p_5d_base - 1.0) * 100.0, 2) if p_5d_base > 0 else float(r.get("單日漲幅(%)", 0.0))
                short_tag, m_detail, sm_ratio = evaluate_stock_margin_signal(
                    float(r.get("單日漲幅(%)", 0.0)), pct_5d,
                    float(r.get("距20MA乖離(%)", 0.0)), float(r.get("量能倍數(倍)", 1.0)),
                    m_curr, m_chg_1d, m_chg_5d, s_curr, s_chg_5d, m_date_fmt,
                )
                r["近5日融資增減(張)"] = int(m_chg_5d)
                r["券資比(%)"] = sm_ratio
                r["資券籌碼訊號"] = short_tag
                margin_note = f"• 【個股資券變化】{m_detail}"

        r["_deep_report"]["chip_cost_text"] = f"{vwap_line}\n{insti_note}\n{margin_note}"


def resolve_user_queries(query_str):
    tokens = [t.strip() for t in re.split(r"[,，、;\s]+", query_str.strip()) if t.strip()]
    if not tokens:
        return [], []

    all_map = {}
    for m in ("twse", "tpex", "emerging"):
        all_map.update(fetch_tw_tickers_fast(mode=m))

    resolved, not_found = [], []
    seen = set()
    for tk_in in tokens:
        found = None
        clean_in = tk_in.upper().replace(".TW", "").replace(".TWO", "")
        if clean_in.isdigit():
            for suffix in (".TW", ".TWO"):
                sym = f"{clean_in}{suffix}"
                if sym in all_map:
                    name, m_lbl = all_map[sym]
                    found = (sym, clean_in, name, m_lbl)
                    break
            if not found:
                found = (f"{clean_in}.TW", clean_in, f"代號{clean_in}", "上市股")
        else:
            for sym, (name, m_lbl) in all_map.items():
                if name == tk_in:
                    found = (sym, sym.replace(".TW", "").replace(".TWO", ""), name, m_lbl)
                    break
            if not found:
                for sym, (name, m_lbl) in all_map.items():
                    if tk_in in name:
                        found = (sym, sym.replace(".TW", "").replace(".TWO", ""), name, m_lbl)
                        break

        if found and found[0] not in seen:
            seen.add(found[0])
            resolved.append(found)
        elif not found:
            not_found.append(tk_in)

    return resolved, not_found


def calc_volume_profile_poc(high_s, low_s, close_s, vol_s, lookback=60, bins=20):
    sub_h, sub_l, sub_c, sub_v = high_s.iloc[-lookback:], low_s.iloc[-lookback:], close_s.iloc[-lookback:], vol_s.iloc[-lookback:]
    p_min, p_max = float(sub_l.min()), float(sub_h.max())
    if p_max <= p_min:
        return round(float(sub_c.iloc[-1]), 2)
    edges = np.linspace(p_min, p_max, bins + 1)
    tp = (sub_h + sub_l + sub_c) / 3.0
    hist, _ = np.histogram(tp.values, bins=edges, weights=sub_v.values)
    max_bin = int(np.argmax(hist))
    return round(float((edges[max_bin] + edges[max_bin + 1]) / 2.0), 2)


def generate_deep_breakdown(info, cond_flags, extra_ctx, params):
    bw_prev, bw_today = info["昨日帶寬(%)"], info["今日帶寬(%)"]
    vol_mul, vol_lots = info["量能倍數(倍)"], info["成交量(張)"]
    bias, pct_chg = info["距20MA乖離(%)"], info["單日漲幅(%)"]
    osc_val, shadow_ratio = info["MACD柱狀體"], info["上影線/紅K實體比"]
    c_today, stop_low, ma20 = info["收盤價"], info["突破紅K低點(停損參考)"], info["20MA月線價"]

    h_today, o_today = extra_ctx["h_today"], extra_ctx["o_today"]
    days_since_cross, osc_prev = extra_ctx["days_since_cross"], extra_ctx["osc_prev"]
    high_60d, vcp_min_ratio = extra_ctx["high_60d"], extra_ctx["vcp_min_ratio"]
    matched_count = sum(int(v) for v in cond_flags.values())
    shadow_num = shadow_ratio if isinstance(shadow_ratio, (int, float)) else 99.9

    if matched_count == 4:
        if bw_prev < 10.0 and pct_chg >= 3.0:
            archetype = "🌟 型態最標準的「極度壓縮帶量突破股」（四條件完美共振）"
            badge = "★★★ 最優先關注（4/4 標準第一拍起漲）"
        else:
            archetype = "🔥 動能、波動與趨勢同步到位的「四條件共振突破股」"
            badge = "★★★ 優先關注（4/4 條件全數達標）"
    elif matched_count == 3:
        if not cond_flags["macd"]:
            if days_since_cross is not None and days_since_cross > params["macd_tolerance"] and osc_val > 0:
                archetype = "💎 型態極佳的「先翻紅蓄勢、後爆量壓縮突破股」（豐達科型）"
                badge = "★★★ 優先關注（3/4 強勢蓄勢突破）"
            else:
                archetype = "📈 布林壓縮帶量突破、惟 MACD 動能微幅縮腳之準突破股"
                badge = "★★☆ 列入觀察（3/4 留意隔日動能續強）"
        elif not cond_flags["vol_price"]:
            if shadow_num > params["upper_shadow_ratio"] and pct_chg < 1.0 and vol_mul >= 1.5:
                archetype = "🛑 盤中衝高卻幾乎吐回漲幅的「假突破警戒股」（可成型）"
                badge = "★☆☆ 警戒避開（3/4 展現第4道濾網保護作用）"
            elif shadow_num > params["upper_shadow_ratio"]:
                archetype = "⚠️ 動能與壓縮皆到位，但「盤中逢高遇壓留上影線」之準突破股（億豐型）"
                badge = "★★☆ 觀察等待（3/4 等隔日站穩上影線確認換手）"
            else:
                archetype = "📊 指標與布林開口到位、惟攻擊量能尚未放大之溫和轉強股"
                badge = "★★☆ 列入觀察（3/4 等待補量至 1.5 倍以上）"
        elif not cond_flags["bb"]:
            archetype = "🚀 動能與量價強勢，但突破前布林帶寬已偏寬之「波段延續股」"
            badge = "★★☆ 謹慎操作（3/4 非第一根壓縮起漲）"
        else:
            archetype = "🔎 符合三項條件之「準突破觀察股」"
            badge = "★★☆ 列入觀察（3/4 達標）"
    elif matched_count == 2:
        if cond_flags["ma"]:
            archetype = "🌤️ 站穩上揚月線、但尚未點火發動的「多頭整理蓄勢股」"
            badge = "★☆☆ 追蹤等待（2/4 達標，尚未出現突破訊號）"
        else:
            archetype = "🌥️ 局部指標轉強、但主趨勢或關鍵條件尚未到位之「過渡盤整股」"
            badge = "★☆☆ 暫時觀望（2/4 達標，結構尚未成熟）"
    else:
        if c_today < ma20:
            archetype = "❄️ 股價位於月線之下、處於「空頭修正或弱勢打底期」"
            badge = f"✖ 不建議介入（僅達標 {matched_count}/4，逆勢風險高）"
        else:
            archetype = "⏸️ 均線與波動尚未共振、處於「無方向盤整觀察期」"
            badge = f"✖ 暫不進場（僅達標 {matched_count}/4，缺乏攻擊訊號）"

    pros = []
    if cond_flags["bb"]:
        sq = "極度壓縮（籌碼沉澱極乾淨）" if bw_prev < 10.0 else "健康收斂狀態"
        pros.append(f"• 【✅ 壓縮後真開口】昨日帶寬僅 {bw_prev}% 屬{sq}，今日擴張至 {bw_today}% 且上下軌反向張開。")
    if cond_flags["vol_price"]:
        pros.append(f"• 【✅ 攻擊量＋紮實紅K】量能放大至 {vol_mul} 倍（{vol_lots} 張），收漲 +{pct_chg}% 且上影線極短（比值 {shadow_ratio}）。")
    if cond_flags["macd"]:
        pros.append(f"• 【✅ MACD剛翻紅點火】MACD 柱狀體（{osc_val}）近 {days_since_cross or 1} 日內剛翻正且持續遞增。")
    if cond_flags["ma"]:
        pros.append(f"• 【✅ 位階安全順勢】收盤 {c_today} 元站穩上揚 20MA（{ma20} 元），乖離率僅 {bias}%。")
    pros_text = "\n".join(pros) if pros else "• 目前四項核心突破條件皆尚未達標（缺乏主力點火訊號）。"

    diag = []
    if matched_count == 4:
        diag.append(f"• 【四項濾網全過】今日開盤 {o_today} → 最高 {h_today} → 收盤 {c_today} 元，兼具壓縮、帶量長紅與指標翻正。")
    else:
        if not cond_flags["macd"]:
            if osc_val <= 0:
                diag.append(f"• 【❌ MACD未翻多】柱狀體為 {osc_val}，短線動能仍弱。")
            elif osc_val <= osc_prev:
                diag.append(f"• 【❌ MACD微縮腳】柱狀體雖為正（{osc_val}），但較昨日（{osc_prev:.3f}）收斂。")
            elif days_since_cross is not None and days_since_cross > params["macd_tolerance"]:
                diag.append(f"• 【❌ MACD提前翻紅（屬正向蓄勢）】早在 {days_since_cross} 天前即翻紅，今日續增配合爆量突破。")
            else:
                diag.append("• 【❌ MACD已翻紅逾20天】屬中繼或高檔延續段。")
        if not cond_flags["bb"]:
            diag.append(f"• 【❌ 布林開口未達標】{extra_ctx['bb_miss']}（昨日 {bw_prev}% → 今日 {bw_today}%）。")
        if not cond_flags["ma"]:
            diag.append(f"• 【❌ 20MA趨勢未達標】{extra_ctx['ma_miss']}（現價 {c_today} / 20MA {ma20} / 乖離 {bias}%）。")
        if not cond_flags["vol_price"]:
            if shadow_num > params["upper_shadow_ratio"] and pct_chg < 1.0 and vol_mul >= 1.5 and (c_today > o_today):
                diag.append(f"• 【❌ 假突破警戒！】爆出 {vol_mul} 倍量能衝高至 {h_today} 元，卻幾乎吐回漲幅僅收 +{pct_chg}%，賣壓沉重！")
            elif shadow_num > params["upper_shadow_ratio"] and (c_today > o_today):
                mid_s = round((c_today + h_today) / 2.0, 2)
                diag.append(f"• 【❌ 留長上影線】量能 {vol_mul} 倍但衝高 {h_today} 元遇壓拉回（上影線比 {shadow_ratio}）；等隔日站穩 {mid_s} 元確認消化賣壓。")
            else:
                diag.append(f"• 【❌ 量價型態未達標】{extra_ctx['vp_miss']}（量倍數 {vol_mul}x，漲幅 {pct_chg}%）。")

    if c_today >= high_60d * 0.995:
        p_note = f"今日已突破近 60 日高點（{high_60d} 元），上方無近季套牢賣壓！"
    else:
        room = round((high_60d / c_today - 1.0) * 100.0, 1)
        p_note = f"前波 60 日高點壓力區約在 {high_60d} 元（距現價約 +{room}%）。"
    v_note = f"近 10 日最低量為月均量 {int(vcp_min_ratio*100)}%" + ("（具極度量縮惜售洗盤訊號）。" if vcp_min_ratio <= 0.55 else "。")
    diag.append(f"• 【前波壓力與洗盤】{p_note}{v_note}")

    stop_pct = round((1.0 - stop_low / c_today) * 100.0, 2)
    ma_pct = round((1.0 - ma20 / c_today) * 100.0, 2)
    overlap = "（✨ 紅K低點與月線重疊，形成雙重支撐鐵板區！）" if abs(stop_low - ma20) / ma20 <= 0.015 else ""

    if matched_count == 4 or (matched_count == 3 and not cond_flags["macd"] and days_since_cross and days_since_cross > params["macd_tolerance"]):
        act = "• 【操作建議】型態具備起漲優勢，可順勢分批布局；隔日若開高 >3% 勿過度追價。"
    elif matched_count == 3 and shadow_num > params["upper_shadow_ratio"] and pct_chg < 1.0:
        act = f"• 【操作建議】衝高回落留長上影線，暫不進場！等日後長紅收盤站穩今日高點 {h_today} 元再考慮。"
    elif matched_count == 3 and shadow_num > params["upper_shadow_ratio"]:
        act = f"• 【操作建議】觀察隔日能否帶量站穩上影線一半（{round((c_today + h_today) / 2.0, 2)} 元）確認換手。"
    elif c_today < ma20:
        act = f"• 【操作建議】現價仍在月線（{ma20} 元）之下，不符合右側順勢原則，等待重新帶量站上月線再評估。"
    else:
        act = "• 【操作建議】列入自選追蹤，等待量能 >= 1.5 倍且紮實紅K撐開布林上軌的發動日。"

    if c_today >= ma20:
        dfn = f"• 【雙層防守價】短線停損看今日K棒低點 『{stop_low} 元』（-{stop_pct}%）；波段改守 20MA 『{ma20} 元』（-{ma_pct}%，跌破三日站不回出場）{overlap}"
    else:
        dfn = f"• 【轉強與止跌價】上方轉強壓力看 20MA 『{ma20} 元』，下檔止跌暫看今日低點 『{stop_low} 元』。"

    return {
        "archetype_title": archetype,
        "rating_badge": badge,
        "pros_text": pros_text,
        "diagnosis_text": "\n".join(diag),
        "chip_cost_text": "",
        "strategy_text": f"{act}\n{dfn}",
        "one_line_summary": f"{badge} | {archetype}",
    }


def evaluate_single_stock(df, params, force_report=False):
    if df is None or df.empty:
        return None, None
    if isinstance(df.columns, pd.MultiIndex):
        df = df.copy()
        df.columns = [c[0] if isinstance(c, tuple) else c for c in df.columns]

    req_cols = ["Open", "High", "Low", "Close", "Volume"]
    for c in req_cols:
        if c not in df.columns:
            return None, None

    df = df.dropna(subset=req_cols).copy()
    df = df[df["Volume"] > 0]

    # 統一 15:30 嚴格收盤保護：15:30 前排除今日資料，採用前一個完整交易日
    exclude_today, today_tw_date, _, _ = get_market_data_cutoff_info()
    if exclude_today and len(df) > 0:
        idx_dates = pd.to_datetime(df.index).date
        df = df[idx_dates < today_tw_date]

    if len(df) < 30:
        return None, None

    close = pd.Series(df["Close"].squeeze()).astype(float)
    open_p = pd.Series(df["Open"].squeeze()).astype(float)
    high = pd.Series(df["High"].squeeze()).astype(float)
    low = pd.Series(df["Low"].squeeze()).astype(float)
    vol_lots = pd.Series(df["Volume"].squeeze()).astype(float) / 1000.0
    tp = (high + low + close) / 3.0

    w_close = (high + low + close * 2.0) / 4.0
    dif = w_close.ewm(span=params["fast_len"], adjust=False).mean() - w_close.ewm(span=params["slow_len"], adjust=False).mean()
    osc = dif - dif.ewm(span=params["sig_len"], adjust=False).mean()

    mid = close.rolling(params["bb_len"]).mean()
    std = close.rolling(params["bb_len"]).std(ddof=1)
    up, dn = mid + params["bb_k"] * std, mid - params["bb_k"] * std
    bw = (up - dn) / mid * 100.0

    ma = close.rolling(params["ma_len"]).mean()
    bias20 = (close / ma - 1.0) * 100.0
    vwap_20_s = (tp * vol_lots).rolling(20).sum() / vol_lots.rolling(20).sum()
    vol_ma5 = vol_lots.shift(1).rolling(5).mean()
    vol_ma20 = vol_lots.shift(1).rolling(20).mean()

    ma_up = params["ma_up_bars"]
    if len(df) <= ma_up + 5:
        return None, None

    c_t, c_p = float(close.iloc[-1]), float(close.iloc[-2])
    o_t, h_t, l_t = float(open_p.iloc[-1]), float(high.iloc[-1]), float(low.iloc[-1])
    v_t, v_ma5_t = float(vol_lots.iloc[-1]), float(vol_ma5.iloc[-1])
    osc_t, osc_p = float(osc.iloc[-1]), float(osc.iloc[-2])
    bw_t, bw_p = float(bw.iloc[-1]), float(bw.iloc[-2])
    up_t, up_p = float(up.iloc[-1]), float(up.iloc[-2])
    dn_t, dn_p = float(dn.iloc[-1]), float(dn.iloc[-2])
    ma_t, ma_pn = float(ma.iloc[-1]), float(ma.iloc[-1 - ma_up])
    bias_t = float(bias20.iloc[-1])
    vwap_20_t = round(float(vwap_20_s.iloc[-1]), 2) if not np.isnan(float(vwap_20_s.iloc[-1])) else round(ma_t, 2)

    if np.isnan(v_ma5_t) or v_ma5_t <= 0 or np.isnan(bw_t) or np.isnan(ma_t):
        return None, None

    cross_up = (osc > 0) & (osc.shift(1) <= 0)
    days_cross = next((d for d in range(1, min(25, len(osc))) if bool(cross_up.iloc[-d])), None)

    cond_macd = (days_cross is not None and days_cross <= params["macd_tolerance"]) and (osc_t > 0) and (osc_t > osc_p)
    cond_bb = (bw_p <= params["max_pre_bw"]) and (bw_t > bw_p) and (up_t > up_p) and (dn_t < dn_p)
    cond_ma = (c_t > ma_t) and (ma_t > ma_pn) and (bias_t <= params["max_bias20"])

    body = c_t - o_t
    upper_s = h_t - c_t
    s_ratio = round(upper_s / body, 2) if body > 0 else 99.9
    amount_wan = round(c_t * v_t / 10.0, 1)
    cond_vp = (amount_wan >= params["min_amount_wan"]) and (v_t >= v_ma5_t * params["vol_ratio"]) and (body > 0) and (upper_s <= body * params["upper_shadow_ratio"])

    matched = int(cond_macd) + int(cond_bb) + int(cond_ma) + int(cond_vp)
    if (not force_report) and matched < 3:
        return None, None

    poc_60 = calc_volume_profile_poc(high, low, close, vol_lots, lookback=min(60, len(df)), bins=20)
    lb60 = min(60, len(high) - 1)
    high_60d = round(float(high.iloc[-1 - lb60 : -1].max()), 2) if lb60 > 5 else round(h_t, 2)
    v_ma20_v = float(vol_ma20.iloc[-1]) if not np.isnan(float(vol_ma20.iloc[-1])) and float(vol_ma20.iloc[-1]) > 0 else v_ma5_t
    vcp_min = round(float(vol_lots.iloc[-11:-1].min()) / v_ma20_v, 2) if len(vol_lots) >= 12 and v_ma20_v > 0 else 1.0

    recent_20 = [
        (dt.strftime("%Y%m%d") if hasattr(dt, "strftime") else str(dt)[:10].replace("-", ""), float(p))
        for dt, p in zip(df.index[-20:], tp.iloc[-20:])
    ]

    # 建構近 30 個交易日價量圖表序列（供 PWA 與 Streamlit 繪製雙圖表）
    vol_ma5_chart = vol_lots.rolling(5).mean()
    n_pts = min(30, len(df))
    chart_30d = {
        "dates": [dt.strftime("%Y-%m-%d") if hasattr(dt, "strftime") else str(dt)[:10] for dt in df.index[-n_pts:]],
        "close": [round(float(x), 2) for x in close.iloc[-n_pts:]],
        "ma20": [round(float(x), 2) if not np.isnan(float(x)) else None for x in ma.iloc[-n_pts:]],
        "vol": [int(round(float(x))) for x in vol_lots.iloc[-n_pts:]],
        "vol_ma5": [int(round(float(x))) if not np.isnan(float(x)) else None for x in vol_ma5_chart.iloc[-n_pts:]],
        "up": [bool(float(c) >= float(o)) for c, o in zip(close.iloc[-n_pts:], open_p.iloc[-n_pts:])],
    }

    bb_miss = ""
    if not cond_bb:
        if bw_p > params["max_pre_bw"]:
            bb_miss = f"昨日帶寬({bw_p:.1f}%)高於壓縮上限{params['max_pre_bw']}%"
        elif not (bw_t > bw_p):
            bb_miss = "今日帶寬仍在收斂未放大"
        else:
            bb_miss = "上下軌未反向張開(屬通道平移)"

    ma_miss = ""
    if not cond_ma:
        if c_t <= ma_t:
            ma_miss = "收盤未站上20MA月線"
        elif ma_t <= ma_pn:
            ma_miss = "20MA月線尚未翻揚向上"
        else:
            ma_miss = f"距20MA乖離({bias_t:.1f}%)超過上限"

    vp_miss = ""
    if not cond_vp:
        if body <= 0:
            vp_miss = "今日非紅K(收黑或平盤)"
        elif amount_wan < params["min_amount_wan"]:
            vp_miss = f"日成交金額({amount_wan}萬元)未達{params['min_amount_wan']}萬元門檻"
        elif v_t < v_ma5_t * params["vol_ratio"]:
            vp_miss = f"量能倍數({v_t/v_ma5_t:.2f}倍)未達{params['vol_ratio']}倍"
        else:
            vp_miss = f"上影線/紅K實體比達 {s_ratio} 倍(超過 {params['upper_shadow_ratio']} 上限)"

    macd_miss = ""
    if not cond_macd:
        if osc_t <= 0:
            macd_miss = "MACD未翻紅"
        elif osc_t <= osc_p:
            macd_miss = "MACD柱狀體縮腳"
        elif days_cross:
            macd_miss = f"MACD已提前{days_cross}天翻紅"
        else:
            macd_miss = "MACD已翻紅逾20天"

    miss_all = [x for x in (macd_miss, bb_miss, ma_miss, vp_miss) if x]
    last_dt = df.index[-1]
    dt_str = last_dt.strftime("%Y-%m-%d") if hasattr(last_dt, "strftime") else str(last_dt)[:10]

    info = {
        "營業概況(主力業務)": "",
        "資料日期": dt_str,
        "收盤價": round(c_t, 2),
        "單日漲幅(%)": round((c_t / c_p - 1.0) * 100.0, 2),
        "成交量(張)": int(round(v_t)),
        "日成交金額(萬元)": amount_wan,
        "量能倍數(倍)": round(v_t / v_ma5_t, 2),
        "日週轉率(%)": "",
        "昨日帶寬(%)": round(bw_p, 2),
        "今日帶寬(%)": round(bw_t, 2),
        "距20MA乖離(%)": round(bias_t, 2),
        "MACD柱狀體": round(osc_t, 3),
        "上影線/紅K實體比": s_ratio if body > 0 else "非紅K",
        "突破紅K低點(停損參考)": round(l_t, 2),
        "20MA月線價": round(ma_t, 2),
        "20日VWAP資金成本": vwap_20_t,
        "60日籌碼密集區(POC)": poc_60,
        "前波60日高點": high_60d,
        "近5日法人合計(張)": "",
        "近20日法人合計(張)": "",
        "外資/投信連買天數": "",
        "法人20日買超成本": "",
        "狀態與未達標細項": "🎯 4/4 條件完全符合" if matched == 4 else f"達標 {matched}/4 項；未達標：{'; '.join(miss_all)}",
    }

    flags = {"macd": cond_macd, "bb": cond_bb, "ma": cond_ma, "vol_price": cond_vp}
    ctx = {
        "h_today": round(h_t, 2), "o_today": round(o_t, 2), "days_since_cross": days_cross,
        "osc_prev": osc_p, "high_60d": high_60d, "vcp_min_ratio": vcp_min,
        "bb_miss": bb_miss, "ma_miss": ma_miss, "vp_miss": vp_miss,
    }
    dr = generate_deep_breakdown(info, flags, ctx, params)
    info["實戰型態定調與優先級"] = dr["one_line_summary"]
    info["_deep_report"] = dr
    info["_recent_20_prices"] = recent_20
    info["_chart_30d"] = chart_30d
    info["_matched_count"] = matched

    if force_report:
        return "CUSTOM", info
    if matched == 4:
        return "FULL", info
    if matched == 3 and cond_ma and body > 0 and amount_wan >= params["min_amount_wan"]:
        return "NEAR", info
    return None, None


# ==========================================
# 雙層快取：第一層讀取每日 15:30 預抓檔，第二層為全體訪客共享記憶體快取
# ==========================================
def is_default_params(mode_choice, scan_params):
    expected_amt = 500.0 if mode_choice == "emerging" else 2000.0
    return (
        scan_params["macd_tolerance"] == 3
        and abs(scan_params["max_pre_bw"] - 15.0) < 1e-6
        and abs(scan_params["max_bias20"] - 8.0) < 1e-6
        and abs(scan_params["min_amount_wan"] - expected_amt) < 1e-6
        and abs(scan_params["vol_ratio"] - 1.5) < 1e-6
        and abs(scan_params["upper_shadow_ratio"] - 0.5) < 1e-6
    )


def try_load_precomputed_cache(mode_choice, effective_date_str):
    """嘗試讀取由每日 15:30 GitHub Actions 主動預抓產生的 JSON 靜態快取檔"""
    cache_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "daily_cache", f"{mode_choice}.json")
    if not os.path.exists(cache_path):
        return None
    try:
        with open(cache_path, "r", encoding="utf-8") as f:
            payload = json.load(f)
        if payload.get("data_date") == effective_date_str:
            return payload
    except Exception:
        pass
    return None


@st.cache_data(ttl=43200, show_spinner=False)
def run_shared_market_scan_cached(mode_choice, period_key, param_items_tuple):
    """
    全體訪客共享快取：綁定 (市場別, 15:30切點日期鑰匙, 參數組合)。
    同一時段內首位訪客掃描後，所有後續訪客在 0.1 秒內直接共享結果，直到下一個 15:30 自動換新！
    """
    scan_params = dict(param_items_tuple)
    stock_dict = fetch_tw_tickers_fast(mode=mode_choice)
    tickers = list(stock_dict.keys())
    total = len(tickers)
    market_title = MODE_LABELS[mode_choice]

    full_list, near_list = [], []
    valid_cnt = 0
    batch_size = 40

    for start_i in range(0, total, batch_size):
        batch = tickers[start_i : start_i + batch_size]
        try:
            data = yf.download(tickers=batch, period="5mo", interval="1d", group_by="ticker", auto_adjust=False, threads=True, progress=False)
        except Exception:
            continue

        for sym in batch:
            try:
                sub_df = data if len(batch) == 1 else (data[sym] if isinstance(data.columns, pd.MultiIndex) and sym in data.columns.get_level_values(0) else None)
                if sub_df is None or sub_df.dropna(how="all").empty:
                    continue
                valid_cnt += 1
                status, res = evaluate_single_stock(sub_df, scan_params, force_report=False)
                if status in ("FULL", "NEAR") and res:
                    code = sym.replace(".TW", "").replace(".TWO", "")
                    name, m_lbl = stock_dict.get(sym, ("", market_title))
                    row = {"股票代碼": code, "股票名稱": name, "市場別": m_lbl}
                    row.update(res)
                    (full_list if status == "FULL" else near_list).append(row)
            except Exception:
                continue
        time.sleep(0.08)

    near_top = sorted(near_list, key=lambda x: x["單日漲幅(%)"], reverse=True)[:30]
    full_sorted = sorted(full_list, key=lambda x: x["今日帶寬(%)"])

    if full_sorted or near_top:
        enrich_rows_with_insti(full_sorted + near_top)

    data_date = ""
    if full_sorted:
        data_date = full_sorted[0]["資料日期"]
    elif near_top:
        data_date = near_top[0]["資料日期"]

    return {
        "full_sorted": full_sorted,
        "near_top": near_top,
        "valid_cnt": valid_cnt,
        "total": total,
        "data_date": data_date,
        "generated_at": datetime.now(timezone(timedelta(hours=8))).strftime("%Y-%m-%d %H:%M"),
        "source": "shared_live_cache",
    }



def evaluate_holding_position(df, code, name, market_lbl, buy_price, lots, buy_date_str=""):
    """三段式自動移動停利／停損引擎（與電腦版 v2.4 規則 100% 同步）"""
    if df is None or df.empty:
        return None
    if isinstance(df.columns, pd.MultiIndex):
        df = df.copy()
        df.columns = [c[0] if isinstance(c, tuple) else c for c in df.columns]

    req_cols = ["Open", "High", "Low", "Close", "Volume"]
    for c in req_cols:
        if c not in df.columns:
            return None
    df = df.dropna(subset=req_cols).copy()
    df = df[df["Volume"] > 0]

    exclude_today, today_tw_date, _, _ = get_market_data_cutoff_info()
    if exclude_today and len(df) > 0:
        idx_dates = pd.to_datetime(df.index).date
        df = df[idx_dates < today_tw_date]

    if len(df) < 25:
        return None

    close = pd.Series(df["Close"].squeeze()).astype(float)
    open_p = pd.Series(df["Open"].squeeze()).astype(float)
    high = pd.Series(df["High"].squeeze()).astype(float)
    low = pd.Series(df["Low"].squeeze()).astype(float)

    ma10 = close.rolling(10).mean()
    ma20 = close.rolling(20).mean()
    prev_close = close.shift(1)
    tr = pd.concat([(high - low), (high - prev_close).abs(), (low - prev_close).abs()], axis=1).max(axis=1)
    atr14 = tr.rolling(14).mean()

    c_t, l_t = float(close.iloc[-1]), float(low.iloc[-1])
    ma10_t, ma20_t = round(float(ma10.iloc[-1]), 2), round(float(ma20.iloc[-1]), 2)
    atr_t = float(atr14.iloc[-1]) if not np.isnan(float(atr14.iloc[-1])) else (c_t * 0.03)

    days_below_ma20 = 0
    for flag in reversed((close < ma20).values):
        if bool(flag):
            days_below_ma20 += 1
        else:
            break

    idx_dates_str = [dt.strftime("%Y-%m-%d") if hasattr(dt, "strftime") else str(dt)[:10] for dt in df.index]
    init_low = l_t
    period_high = float(high.iloc[-20:].max())

    if buy_date_str and buy_date_str.strip():
        b_clean = buy_date_str.strip().replace("/", "-")
        m_idx = [i for i, ds in enumerate(idx_dates_str) if ds >= b_clean]
        if m_idx:
            init_low = float(low.iloc[m_idx[0]])
            period_high = float(high.iloc[m_idx[0]:].max())
    else:
        red_mask = (close.iloc[-10:] > open_p.iloc[-10:]).values
        red_lows = low.iloc[-10:].values[red_mask]
        if len(red_lows) > 0:
            init_low = float(red_lows[-1])

    cost_amt = buy_price * lots * 1000.0
    mkt_val = c_t * lots * 1000.0
    est_fee_tax = cost_amt * 0.001425 + mkt_val * (0.001425 + 0.003)
    net_pnl = int(round(mkt_val - cost_amt - est_fee_tax))
    pnl_pct = round((c_t / buy_price - 1.0) * 100.0, 2)

    hard_stop_7 = round(buy_price * 0.93, 2)
    breakeven_price = round(buy_price * 1.006, 2)
    atr_trail_price = round(period_high - 2.0 * atr_t, 2)

    if pnl_pct < 5.0:
        stage = "🛡️ 第一階段：剛進場防守期 (<+5%)"
        main_stop = round(max(init_low, hard_stop_7), 2)
        lock_line = "尚未進入鎖利期"
        if c_t < main_stop:
            alert = f"🚨 觸發初始停損！收盤({c_t})已跌破防守價({main_stop})"
            tip = f"已跌破第一階段防守線（紅K低點 {round(init_low,2)} / -7%硬停損 {hard_stop_7} 取高），建議依紀律執行停損。"
        else:
            gap = round((1.0 - main_stop / c_t) * 100.0, 2)
            alert = f"✅ 持有防守中（距停損緩衝 {gap}%）"
            tip = f"嚴守初始停損 『{main_stop} 元』（突破紅K低點 {round(init_low,2)} 與 -7% 硬停損 {hard_stop_7} 取高者）。"
    elif pnl_pct <= 15.0:
        stage = "🔒 第二階段：脫離成本保本期 (+5%~+15%)"
        main_stop = round(max(breakeven_price, ma20_t), 2)
        lock_line = ma10_t
        if c_t < main_stop:
            alert = f"⚠️ 觸及保本／月線防線({main_stop})，建議保本出場"
            tip = f"股價回測保本線 『{main_stop} 元』（含稅費保本 {breakeven_price} 與 20MA {ma20_t} 取高），建議保本離場。"
        else:
            alert = f"✅ 已鎖定保本不賠！（移動防線 {main_stop} 元）"
            tip = f"停損已自動上移至 『{main_stop} 元』（保本價 {breakeven_price} 與上揚月線 {ma20_t} 取高），安心續抱等主升段。"
    else:
        stage = "🚀 第三階段：主升段奔跑停利期 (>+15%)"
        lock_line = round(max(ma10_t, atr_trail_price), 2)
        main_stop = ma20_t
        if days_below_ma20 >= 3:
            alert = f"🚨 觸發波段全出！已連破 20MA ({ma20_t}) 達 {days_below_ma20} 日"
            tip = "已符合「跌破 20MA 三日站不回」出場條件，建議獲利了結全出！"
        elif c_t < lock_line:
            alert = f"⚠️ 跌破減碼鎖利線({lock_line})，建議先減碼一半"
            tip = f"跌破減碼鎖利線 『{lock_line} 元』（10MA {ma10_t} / 高點-2ATR {atr_trail_price} 取高），先鎖利一半、剩餘守 20MA ({ma20_t})。"
        else:
            alert = f"🔥 主升段強勢奔跑中！（減碼線 {lock_line} / 月線 {ma20_t}）"
            tip = f"強勢續抱！跌破減碼線 『{lock_line} 元』 先鎖利一半；跌破 20MA 『{ma20_t} 元』 三日站不回全出。"

    vol_lots_h = pd.Series(df["Volume"].squeeze()).astype(float) / 1000.0
    vol_ma5_h = vol_lots_h.rolling(5).mean()
    n_h = min(30, len(df))
    h_chart_30d = {
        "dates": idx_dates_str[-n_h:],
        "close": [round(float(x), 2) for x in close.iloc[-n_h:]],
        "ma20": [round(float(x), 2) if not np.isnan(float(x)) else None for x in ma20.iloc[-n_h:]],
        "vol": [int(round(float(x))) for x in vol_lots_h.iloc[-n_h:]],
        "vol_ma5": [int(round(float(x))) if not np.isnan(float(x)) else None for x in vol_ma5_h.iloc[-n_h:]],
        "up": [bool(float(c) >= float(o)) for c, o in zip(close.iloc[-n_h:], open_p.iloc[-n_h:])],
    }

    return {
        "股票代碼": code, "股票名稱": name, "市場別": market_lbl, "資料日期": idx_dates_str[-1],
        "_chart_30d": h_chart_30d,
        "買進日期": buy_date_str or "未填", "持有張數": lots, "買進均價": round(buy_price, 2),
        "最新收盤價": round(c_t, 2), "未實現報酬(%)": pnl_pct, "預估淨損益(元)": net_pnl,
        "目前防守階段": stage, "今日移動停損/全出價": main_stop, "短線減碼鎖利線": lock_line,
        "20MA月線價": ma20_t, "10MA均線價": ma10_t, "連破月線天數": days_below_ma20,
        "盤後風控警示狀態": alert, "明日操作建議": tip,
    }


def build_holdings_excel_bytes(rows):
    cols = [
        "股票代碼", "股票名稱", "市場別", "營業概況(主力業務)", "資料日期", "買進日期", "持有張數",
        "買進均價", "最新收盤價", "未實現報酬(%)", "預估淨損益(元)",
        "目前防守階段", "今日移動停損/全出價", "短線減碼鎖利線",
        "20MA月線價", "10MA均線價", "連破月線天數", "盤後風控警示狀態", "明日操作建議"
    ]
    df_out = pd.DataFrame([{k: r[k] for k in cols} for r in rows]) if rows else pd.DataFrame(columns=cols)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_out.to_excel(writer, sheet_name="每日持倉移動停利停損", index=False)
        ws = writer.sheets["每日持倉移動停利停損"]
        h_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        h_font = Font(name="Microsoft JhengHei", size=10, bold=True, color="FFFFFF")
        for c_i in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=c_i)
            cell.fill, cell.font = h_fill, h_font
            ws.column_dimensions[get_column_letter(c_i)].width = 16 if c_i < 17 else 44
    buf.seek(0)
    return buf.getvalue()


COMPACT_COLS = [
    ("股票代碼", "代號"), ("股票名稱", "名稱"), ("收盤價", "收盤價"), ("單日漲幅(%)", "漲幅%"),
    ("日成交金額(萬元)", "成交額\n(萬元)"), ("量能倍數(倍)", "量倍數"), ("日週轉率(%)", "週轉%"),
    ("今日帶寬(%)", "今日\n帶寬%"), ("距20MA乖離(%)", "20MA\n乖離%"),
    ("20日VWAP資金成本", "20日\nVWAP"), ("近5日法人合計(張)", "法人5日\n(張)"), ("突破紅K低點(停損參考)", "紅K停損"),
]


def _calc_wrap_h(text, chars_per_line=48, base_h=24, line_h=17):
    if not text:
        return base_h
    lines = 0
    for p in str(text).split("\n"):
        w = sum(2 if ord(ch) > 127 else 1 for ch in p)
        lines += max(1, int((w - 1) // (chars_per_line * 2)) + 1)
    return max(base_h, lines * line_h + 10)


def build_excel_bytes(sheets_spec, summary_rows, is_single_check=False):
    """方案 A+B 零橫軸 Excel 匯出（單檔健檢採 A~H 直式券商報告；全市場採 A~L 12欄單螢幕總表＋末頁27欄原始資料）"""
    buf = io.BytesIO()
    thin_b = Border(
        left=Side(style="thin", color="BFBFBF"), right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"), bottom=Side(style="thin", color="BFBFBF"),
    )
    h_font = Font(name="Microsoft JhengHei", size=9.5, bold=True)
    d_font = Font(name="Microsoft JhengHei", size=9.5)
    sec_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    sec_font = Font(name="Microsoft JhengHei", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    chip_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    stop_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

    all_raw = []
    for _, r_list, _ in sheets_spec:
        all_raw.extend(r_list)
    df_raw = pd.DataFrame([{k: r.get(k, "") for k in EXCEL_COLS} for r in all_raw]) if all_raw else pd.DataFrame(columns=EXCEL_COLS)

    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        if is_single_check and sheets_spec:
            # 方案 B：A~H 欄直式券商健檢報告
            sheet_name, rows_data, _ = sheets_spec[0]
            ws = writer.book.create_sheet(title=sheet_name, index=0)
            for c_i, w in {1: 16, 2: 14, 3: 13, 4: 13, 5: 17, 6: 14, 7: 13, 8: 13}.items():
                ws.column_dimensions[get_column_letter(c_i)].width = w
            cur = 1
            for idx, row in enumerate(rows_data, start=1):
                dr = row.get("_deep_report", {})
                ws.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=8)
                ct = ws.cell(row=cur, column=1, value=f"📋 【個股深度健檢報告 #{idx}】 {row.get('股票代碼','')} {row.get('股票名稱','')} ({row.get('市場別','')}) ｜ 資料日：{row.get('資料日期','')}")
                ct.font, ct.fill = sec_font, sec_fill
                ws.row_dimensions[cur].height = 26
                cur += 1

                ws.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=8)
                ca = ws.cell(row=cur, column=1, value=f"{dr.get('archetype_title','')} ｜ {dr.get('rating_badge','')}")
                ca.font, ca.fill = h_font, header_fill
                ws.row_dimensions[cur].height = 24
                cur += 1

                pairs = [
                    ("收盤 / 漲幅", f"{row.get('收盤價','')}元 ({row.get('單日漲幅(%)','')}%)", "20日VWAP成本", f"{row.get('20日VWAP資金成本','')}元"),
                    ("量 / 成交額", f"{row.get('成交量(張)','')}張 / {row.get('日成交金額(萬元)','')}萬", "60日POC密集區", f"{row.get('60日籌碼密集區(POC)','')}元"),
                    ("量倍 / 週轉", f"{row.get('量能倍數(倍)','')}x / {row.get('日週轉率(%)','')}%", "法人20日成本", f"{row.get('法人20日買超成本','')}"),
                    ("帶寬(昨→今)", f"{row.get('昨日帶寬(%)','')}%→{row.get('今日帶寬(%)','')}%", "法人5日 / 20日", f"{row.get('近5日法人合計(張)','')} / {row.get('近20日法人合計(張)','')}張"),
                    ("20MA / 乖離", f"{row.get('20MA月線價','')} ({row.get('距20MA乖離(%)','')}%)", "紅K停損 / 60高", f"{row.get('突破紅K低點(停損參考)','')} / {row.get('前波60日高點','')}"),
                ]
                for ll, lv, rl, rv in pairs:
                    ws.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=2)
                    ws.merge_cells(start_row=cur, start_column=3, end_row=cur, end_column=4)
                    ws.merge_cells(start_row=cur, start_column=5, end_row=cur, end_column=6)
                    ws.merge_cells(start_row=cur, start_column=7, end_row=cur, end_column=8)
                    ws.cell(row=cur, column=1, value=ll).font = h_font
                    ws.cell(row=cur, column=1).fill = header_fill
                    ws.cell(row=cur, column=3, value=lv).font = d_font
                    ws.cell(row=cur, column=5, value=rl).font = h_font
                    ws.cell(row=cur, column=5).fill = chip_fill
                    ws.cell(row=cur, column=7, value=rv).font = d_font
                    ws.row_dimensions[cur].height = 22
                    for ci in range(1, 9):
                        ws.cell(row=cur, column=ci).border = thin_b
                    cur += 1

                for sl, stxt, sbg in [
                    ("🏢 營業概況", row.get("營業概況(主力業務)", "尚無資料"), header_fill),
                    ("✅ 達標優勢", dr.get("pros_text", ""), header_fill),
                    ("🔍 缺口與壓力", dr.get("diagnosis_text", ""), header_fill),
                    ("💰 主力與法人", dr.get("chip_cost_text", ""), chip_fill),
                    ("🛡️ 操作與防守", dr.get("strategy_text", ""), header_fill),
                ]:
                    ws.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=2)
                    ws.merge_cells(start_row=cur, start_column=3, end_row=cur, end_column=8)
                    ws.cell(row=cur, column=1, value=sl).font = h_font
                    ws.cell(row=cur, column=1).fill = sbg
                    cv = ws.cell(row=cur, column=3, value=stxt)
                    cv.font = d_font
                    cv.alignment = Alignment(wrap_text=True, vertical="center")
                    ws.row_dimensions[cur].height = _calc_wrap_h(stxt, chars_per_line=44)
                    for ci in range(1, 9):
                        ws.cell(row=cur, column=ci).border = thin_b
                    cur += 1
                cur += 1
        else:
            # 方案 A：12 欄單螢幕總表 + A~L 欄同寬卡片
            for sheet_name, rows_data, title_label in sheets_spec:
                df_c = pd.DataFrame([{lbl: r.get(k, "") for k, lbl in COMPACT_COLS} for r in rows_data]) if rows_data else pd.DataFrame(columns=[lbl for _, lbl in COMPACT_COLS])
                df_c.to_excel(writer, sheet_name=sheet_name, index=False)
                ws = writer.sheets[sheet_name]
                for r_i in range(1, ws.max_row + 1):
                    ws.row_dimensions[r_i].height = 30 if r_i == 1 else 23
                    for c_i in range(1, 13):
                        cell = ws.cell(row=r_i, column=c_i)
                        cell.border = thin_b
                        if r_i == 1:
                            cell.fill = stop_fill if c_i == 12 else (chip_fill if c_i in (10, 11) else header_fill)
                            cell.font = h_font
                            cell.alignment = Alignment(vertical="center", horizontal="center", wrap_text=True)
                        else:
                            cell.font = d_font
                for c_i, w in {1: 9, 2: 12, 3: 9.5, 4: 9, 5: 11.5, 6: 9, 7: 9, 8: 9.5, 9: 9.5, 10: 11, 11: 11, 12: 11}.items():
                    ws.column_dimensions[get_column_letter(c_i)].width = w

                cur = ws.max_row + 2
                ws.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=12)
                ct = ws.cell(row=cur, column=1, value=f"📋 【{title_label} — 個股實戰拆解與主力成本報告（免拉橫軸）】")
                ct.font, ct.fill = sec_font, sec_fill
                cur += 1
                for idx, row in enumerate(rows_data, start=1):
                    dr = row.get("_deep_report", {})
                    ws.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=12)
                    ch = ws.cell(row=cur, column=1, value=f"#{idx} {row.get('股票代碼','')} {row.get('股票名稱','')} — {dr.get('archetype_title','')} | {dr.get('rating_badge','')}")
                    ch.font, ch.fill = h_font, header_fill
                    cur += 1
                    for sl, stxt in [
                        ("🏢 營業概況", row.get("營業概況(主力業務)", "尚無資料")),
                        ("✅ 達標優勢", dr.get("pros_text", "")),
                        ("🔍 缺口與壓力", dr.get("diagnosis_text", "")),
                        ("💰 主力與法人", dr.get("chip_cost_text", "")),
                        ("🛡️ 操作與防守", dr.get("strategy_text", "")),
                    ]:
                        ws.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=2)
                        ws.merge_cells(start_row=cur, start_column=3, end_row=cur, end_column=12)
                        ws.cell(row=cur, column=1, value=sl).font = h_font
                        cv = ws.cell(row=cur, column=3, value=stxt)
                        cv.font = d_font
                        cv.alignment = Alignment(wrap_text=True, vertical="center")
                        ws.row_dimensions[cur].height = _calc_wrap_h(stxt, chars_per_line=48)
                        for ci in range(1, 13):
                            ws.cell(row=cur, column=ci).border = thin_b
                        cur += 1
                    cur += 1

        df_raw.to_excel(writer, sheet_name="原始完整數據(27欄)", index=False)
        pd.DataFrame(summary_rows, columns=["項目", "數值 / 設定"]).to_excel(writer, sheet_name="掃描統計與參數", index=False)

    buf.seek(0)
    return buf.getvalue()


def build_chart_guidance_text(chart_data, stop_price, ref_date=None):
    """依據最新收盤價、20MA、紅K停損線與 5 日均量，自動產生股價與成交量白話技術判讀"""
    if not chart_data or not chart_data.get("close"):
        return "暫無足夠歷史資料。", "暫無足夠成交量資料。"

    c_last = chart_data["close"][-1]
    c_prev = chart_data["close"][-2] if len(chart_data["close"]) >= 2 else c_last
    ma20_last = chart_data["ma20"][-1] or c_last
    ma20_prev = chart_data["ma20"][-2] if len(chart_data["ma20"]) >= 2 and chart_data["ma20"][-2] else ma20_last
    v_last = chart_data["vol"][-1]
    v_ma5_last = chart_data["vol_ma5"][-1] or max(v_last, 1)

    # 1. 上方股價折線圖判讀（收盤價 vs 20MA & 紅K停損線 + 停損/停利建議）
    bias_ma20 = round((c_last / ma20_last - 1.0) * 100.0, 2) if ma20_last > 0 else 0.0
    ma_dir = "上揚" if ma20_last >= ma20_prev else "下彎"
    stop_val = float(stop_price) if stop_price else round(c_last * 0.93, 2)

    if c_last < stop_val and c_last < ma20_last:
        price_advice = f"🚨 **停損警示**：最新收盤（{c_last}元）已同時跌破紅K停損線（{stop_val}元）與 20MA 月線（{ma20_last}元），短線轉弱，建議依紀律執行停損出場。"
    elif c_last < stop_val:
        price_advice = f"⚠️ **短線破線警戒**：收盤（{c_last}元）已跌破紅K停損線（{stop_val}元），雖仍在 20MA（{ma20_last}元）附近，但屬於短線假突破警訊，建議先減碼或嚴格停損。"
    elif c_last < ma20_last:
        price_advice = f"⚠️ **跌破月線觀察**：收盤（{c_last}元）低於 20MA 月線（{ma20_last}元，乖離 {bias_ma20:+.2f}%），若連續三日站不回月線建議波段出場。"
    elif bias_ma20 >= 12.0:
        price_advice = f"🔥 **主升段分批停利區**：收盤（{c_last}元）遠高於{ma_dir} 20MA（{ma20_last}元，正乖離 +{bias_ma20:.2f}%），短線漲幅已大，可考慮分批獲利了結一半，剩餘部位改守 10MA 或 20MA 移動停利。"
    else:
        gap_s = round((1.0 - stop_val / c_last) * 100.0, 2)
        price_advice = f"✅ **多頭順勢續抱**：收盤（{c_last}元）穩站於{ma_dir} 20MA（{ma20_last}元，乖離 +{bias_ma20:.2f}%）與紅K停損線（{stop_val}元）之上（距停損緩衝 {gap_s}%），趨勢健康建議續抱觀察。"

    # 2. 下方成交量柱狀圖判讀（成交量 vs 5日均量線之價量關係分析）
    vol_ratio = round(v_last / v_ma5_last, 2) if v_ma5_last > 0 else 1.0
    is_up = c_last >= c_prev
    if is_up and vol_ratio >= 1.3:
        vol_advice = f"🔥 **價漲量增（健康攻擊盤）**：今日收漲且成交量（{v_last:,}張）為 5 日均量（{v_ma5_last:,}張）的 **{vol_ratio} 倍**，買盤資金積極推升，屬強勢多方表態。"
    elif is_up and vol_ratio < 0.85:
        vol_advice = f"🌤️ **價漲量縮（籌碼鎖定／惜售）**：今日收紅但成交量僅為 5 日均量的 **{vol_ratio} 倍**（{v_last:,}張 vs 均量 {v_ma5_last:,}張），代表賣壓極輕，但後續上攻仍需補量確認。"
    elif (not is_up) and vol_ratio < 0.85:
        vol_advice = f"🛡️ **價跌量縮（良性洗盤換手）**：今日小幅回檔且量能萎縮至 5 日均量的 **{vol_ratio} 倍**（{v_last:,}張），屬於正常浮額清洗、非主力出貨，只要不破支撐無須過度恐慌。"
    elif (not is_up) and vol_ratio >= 1.3:
        vol_advice = f"🚨 **價跌量增（高檔出貨警訊）**：今日收跌且爆出 5 日均量 **{vol_ratio} 倍** 的大量（{v_last:,}張），代表上方籌碼鬆動或主力調節，務必提高警覺並嚴守停損線！"
    else:
        vol_advice = f"📊 **量價溫和平穩**：今日成交量（{v_last:,}張）接近 5 日均量（{v_ma5_last:,}張，比值 {vol_ratio} 倍），多空處於常態換手整理。"

    return price_advice, vol_advice


def render_price_volume_charts(title_name, chart_data, stop_price, mark_date=None):
    """上方股價折線圖 + 中間紅綠成交量柱狀圖 + 下方技術分析與價量關係引導說明"""
    if not chart_data or not chart_data.get("dates"):
        st.caption("暫無足夠歷史資料繪製價量圖表。")
        return

    df_c = pd.DataFrame({
        "日期": chart_data["dates"],
        "收盤價": chart_data["close"],
        "20MA月線": chart_data["ma20"],
        "成交量": chart_data["vol"],
        "5日均量": chart_data["vol_ma5"],
        "漲跌": ["漲" if u else "跌" for u in chart_data["up"]],
    })
    df_c["紅K停損線"] = float(stop_price) if stop_price else None

    # 上方圖：股價折線圖 + 20MA月線(橘虛線) + 紅K停損線(綠虛線)
    y_min = min([x for x in df_c["收盤價"].tolist() + [stop_price] if x is not None]) * 0.97
    y_max = max([x for x in df_c["收盤價"].tolist() + [stop_price] if x is not None]) * 1.03

    base = alt.Chart(df_c).encode(x=alt.X("日期:N", axis=alt.Axis(labelAngle=-45, title=None)))
    line_close = base.mark_line(color="#38bdf8", strokeWidth=2.5, point=True).encode(
        y=alt.Y("收盤價:Q", scale=alt.Scale(domain=[y_min, y_max]), title="股價 (元)"),
        tooltip=["日期", "收盤價", "20MA月線", "紅K停損線"],
    )
    line_ma20 = base.mark_line(color="#f59e0b", strokeDash=[5, 4], strokeWidth=1.8).encode(y="20MA月線:Q")
    line_stop = base.mark_line(color="#22c55e", strokeDash=[3, 3], strokeWidth=1.6).encode(y="紅K停損線:Q")

    price_layers = [line_close, line_ma20, line_stop]
    if mark_date and mark_date in df_c["日期"].values:
        rule_df = pd.DataFrame({"日期": [mark_date]})
        v_rule = alt.Chart(rule_df).mark_rule(color="#ef4444", strokeDash=[4, 2], strokeWidth=2).encode(x="日期:N")
        price_layers.append(v_rule)

    price_chart = alt.layer(*price_layers).properties(height=230, title=f"📈 {title_name} — 每日股價走勢 (藍:收盤 / 橘虛:20MA / 綠虛:停損線)")
    st.altair_chart(price_chart, use_container_width=True)

    # 中間圖：紅綠成交量柱狀圖 + 5日均量線
    bars = base.mark_bar().encode(
        y=alt.Y("成交量:Q", title="成交量 (張)"),
        color=alt.Color("漲跌:N", scale=alt.Scale(domain=["漲", "跌"], range=["#ef4444", "#22c55e"]), legend=None),
        tooltip=["日期", "成交量", "5日均量"],
    )
    line_vma5 = base.mark_line(color="#fbbf24", strokeWidth=2).encode(y="5日均量:Q")
    vol_chart = alt.layer(bars, line_vma5).properties(height=150, title="📊 對應每日成交量變化 (紅:收漲 / 綠:收跌 / 黃線:5日均量)")
    st.altair_chart(vol_chart, use_container_width=True)

    # 下方：連帶的技術分析與價量關係引導說明
    p_guide, v_guide = build_chart_guidance_text(chart_data, stop_price, mark_date)
    st.markdown(
        f"""<div class="card-box card-watch">
            <div class="sec-head">📈 股價折線圖判讀（收盤價 vs. 20MA 月線與紅K停損線）</div>
            <div class="sec-body">{p_guide}</div>
            <div class="sec-head">📊 成交量柱狀圖判讀（成交量 vs. 5日均量價量關係）</div>
            <div class="sec-body">{v_guide}</div>
        </div>""",
        unsafe_allow_html=True,
    )


def render_mobile_stock_card(row):
    dr = row["_deep_report"]
    badge = dr["rating_badge"]
    css_cls = "card-star" if "★★★" in badge else "card-warn" if ("警戒" in badge or "✖" in badge) else "card-watch"

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("收盤價", f"{row['收盤價']}", f"{row['單日漲幅(%)']:+.2f}% | 週轉 {row['日週轉率(%)']}%")
    c2.metric("20日VWAP成本", f"{row['20日VWAP資金成本']}", f"乖離 {row['距20MA乖離(%)']}%")
    c3.metric("法人20日成本", f"{row['法人20日買超成本']}", f"5日 {row['近5日法人合計(張)']}張")
    c4.metric("短線停損 / 月線", f"{row['突破紅K低點(停損參考)']}", f"20MA {row['20MA月線價']}")

    html_card = f"""
    <div class="card-box {css_cls}">
        <div class="card-title">{row['股票代碼']} {row['股票名稱']} ({row['市場別']}) ｜ 資料日：{row['資料日期']}</div>
        <div class="card-badge">{dr['archetype_title']}<br>{badge} ｜ 資券：{row.get("資券籌碼訊號", "待同步")}</div>
        <div class="sec-head">🏢 營業概況（主力業務）</div>
        <div class="sec-body">{row.get("營業概況(主力業務)", "尚無資料")}</div>
        <div class="sec-head">✅ 達標指標優勢</div>
        <div class="sec-body">{dr['pros_text']}</div>
        <div class="sec-head">🔍 關鍵缺口／盤中軌跡與前波壓力</div>
        <div class="sec-body">{dr['diagnosis_text']}</div>
        <div class="sec-head">💰 主力進場成本與三大法人籌碼</div>
        <div class="sec-body">{dr['chip_cost_text']}</div>
        <div class="sec-head">🛡️ 實戰進出場策略與雙層防守價</div>
        <div class="sec-body">{dr['strategy_text']}</div>
    </div>
    """
    st.markdown(html_card, unsafe_allow_html=True)
    if row.get("_chart_30d"):
        with st.expander(f"📈 查看 {row['股票代碼']} {row['股票名稱']} 股價與成交量走勢圖 ＋ 技術判讀"):
            render_price_volume_charts(
                f"{row['股票代碼']} {row['股票名稱']}",
                row["_chart_30d"],
                row.get("突破紅K低點(停損參考)"),
                mark_date=row.get("資料日期"),
            )


# ==========================================
# 主介面排版 (統一 15:30 基準提示 + 資料日期標註)
# ==========================================
st.title("📈 台股選股 & 主力籌碼健檢")

is_before_1530, today_tw_date, period_key, now_str = get_market_data_cutoff_info()
effective_date_str = get_effective_trading_date_str(period_key)

if is_before_1530:
    st.info(
        f"📅 **目前採用資料日期：`{effective_date_str}`**（台灣時間 {now_str}）\n\n"
        f"🕒 **提醒**：本工具每日 **15:30** 統一更新（同步收盤價與證交所三大法人籌碼）；"
        f"**15:30 前抓取的資料為前一個交易日之完整收盤資訊**，不受盤中跳動影響。"
    )
else:
    st.success(
        f"📅 **目前採用資料日期：`{effective_date_str}`**（台灣時間 {now_str}）\n\n"
        f"🕒 已過每日 **15:30** 更新時段，採用最新完整收盤價與三大法人籌碼資料。"
    )


@st.cache_data(ttl=300, show_spinner=False)
def load_premarket_daily_cached():
    """優先讀取本機 daily_cache/premarket_daily.json，次之讀取 GitHub Raw 雲端盤前狀態檔"""
    local_p = os.path.join(os.path.dirname(os.path.abspath(__file__)), "daily_cache", "premarket_daily.json")
    if os.path.exists(local_p):
        try:
            with open(local_p, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    try:
        ssl_ctx = ssl._create_unverified_context()
        url = f"https://raw.githubusercontent.com/emin0426/tw-stock-app/main/daily_cache/premarket_daily.json?t={int(time.time())}"
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=8, context=ssl_ctx) as resp:
            return json.loads(resp.read().decode("utf-8", errors="ignore"))
    except Exception:
        return None


# 🌅 盤前自動分析看板（固定顯示 08:15 完整簡報 & 08:58 終版開盤定案 2 張卡片）
pm_data = load_premarket_daily_cached()
if pm_data and (pm_data.get("card_0815") or pm_data.get("card_0858")):
    c15 = pm_data.get("card_0815")
    c58 = pm_data.get("card_0858")
    col_pm1, col_pm2 = st.columns(2)
    with col_pm1:
        if c15:
            tone_cls = "card-warn" if c15.get("divergence_flag") == "BEAR_DIVERGENCE" or c15.get("tone") == "偏空" else "card-star" if c15.get("tone") == "偏多" else "card-watch"
            st.markdown(
                f"""<div class="card-box {tone_cls}">
                    <div class="card-title">🌅 {c15['title']}（{c15['tone']} {c15['score']:+d}/5）</div>
                    <div class="sec-body">
                        <b>美股與 ADR</b>：{c15['us_line']}<br>
                        <b>夜盤與籌碼</b>：{c15['chip_line']}<br>
                        <b>08:15 亞股校正</b>：{c15['asia_line']}<br>
                        <b>初步預估</b>：<b>{c15['est_gap_text']}</b><br>
                        <b>觀察劇本</b>：{c15['scenario']}
                    </div>
                </div>""",
                unsafe_allow_html=True,
            )
        else:
            st.info("🌅 **08:15 盤前完整簡報**：等待今日 08:15 自動更新...")
    with col_pm2:
        if c58:
            st.markdown(
                f"""<div class="card-box card-watch">
                    <div class="card-title">🔔 {c58['title']}</div>
                    <div class="sec-body">
                        <b>方向判定</b>：{c58['direction']}<br>
                        <b>台指期日盤</b>：{c58['futures_line']}<br>
                        <b>權值股試撮</b>：{c58['auction_line']}<br>
                        <b>終版預估跳空</b>：<b>{c58['est_gap_text']}</b><br>
                        <b>操作提醒</b>：{c58['reminder']}
                    </div>
                </div>""",
                unsafe_allow_html=True,
            )
        else:
            st.info("🔔 **08:58 終版開盤定案**：Phase 2 接入券商 API 後將於 08:58 自動定案（目前參考左側 08:15 簡報）。")
else:
    st.caption("🌅 盤前自動分析看板：尚未偵測到 `daily_cache/premarket_daily.json`（執行一次 `premarket_scan.yml` 後將於此自動顯示 08:15 與 08:58 兩張盤前卡片）。")

# 📊 每日大盤資券變化與多空溫度計
mkt_margin = get_market_margin_analysis_cached(period_key)
if mkt_margin.get("available"):
    m_lvl = mkt_margin["signal_level"]
    m_cls = "card-warn" if m_lvl == "OVERHEAT" else "card-star" if m_lvl == "BOTTOM" else "card-watch"
    st.markdown(
        f"""<div class="card-box {m_cls}">
            <div class="card-title">📊 大盤資券變化 ｜ {mkt_margin['signal_badge']}</div>
            <div class="sec-body">
                <b>加權指數 ({mkt_margin.get('taiex_date', mkt_margin['margin_date'])})</b>：{mkt_margin['taiex_close']:,.2f}（單日 {mkt_margin['taiex_chg_1d']:+.2f}% ／ 近5日 {mkt_margin['taiex_chg_5d']:+.2f}% ／ 距20MA {mkt_margin['taiex_bias_20ma']:+.2f}%）<br>
                <b>上市融資餘額</b>：{mkt_margin['curr_amt_yi']:,.2f} 億元（單日 <b>{mkt_margin['chg_1d_yi']:+.2f} 億</b> ／ 近{mkt_margin.get('days_counted',5)}日 <b>{mkt_margin['chg_5d_yi']:+.2f} 億</b>）｜ 資券日：{mkt_margin['margin_date']}<br>
                <b>盤勢解讀</b>：{mkt_margin['summary_text']}
            </div>
        </div>""",
        unsafe_allow_html=True,
    )
else:
    st.caption(f"📊 大盤資券變化：{mkt_margin.get('summary_text', '暫無資料')}")

with st.expander("⚙️ 調整四大選股條件參數（預設已為進階實戰最佳值）", expanded=False):
    col_a, col_b = st.columns(2)
    macd_tol = col_a.number_input("MACD 翻紅容錯天數", min_value=1, max_value=10, value=3)
    max_pre_bw = col_b.number_input("突破前布林帶寬上限(%)", min_value=5.0, max_value=50.0, value=15.0, step=1.0)
    max_bias20 = col_a.number_input("距 20MA 最大乖離率(%)", min_value=2.0, max_value=30.0, value=8.0, step=1.0)
    min_amount_wan = col_b.number_input("最低日成交金額(萬元，0=不限)", min_value=0, max_value=50000, value=2000, step=500)
    vol_ratio = col_a.number_input("攻擊量倍數(較前5日均量)", min_value=1.0, max_value=5.0, value=1.5, step=0.1)
    upper_shadow = col_b.number_input("上影線/紅K實體比例上限", min_value=0.1, max_value=2.0, value=0.5, step=0.1)

params = {
    "fast_len": 12, "slow_len": 26, "sig_len": 9, "macd_tolerance": int(macd_tol),
    "bb_len": 20, "bb_k": 2.0, "max_pre_bw": float(max_pre_bw),
    "ma_len": 20, "ma_up_bars": 1, "max_bias20": float(max_bias20),
    "min_amount_wan": float(min_amount_wan), "vol_ratio": float(vol_ratio), "upper_shadow_ratio": float(upper_shadow),
}

tab1, tab2, tab3 = st.tabs(["🩺 個股即時健檢", "🚀 全市場選股", "💼 持倉移動停利/停損"])

# ---------- 分頁 1：手動輸入個股即時健檢 ----------
with tab1:
    query_text = st.text_input(
        "輸入台股代號或中文名稱（支援多檔，以逗號或空白隔開）",
        placeholder="例如：世禾, 3004, 億豐, 台積電",
    )
    if st.button("🩺 立即執行個股健檢 ＋ 證交所法人成本分析", type="primary"):
        if not query_text.strip():
            st.warning("請先輸入股票代號或名稱！")
        else:
            with st.spinner("正在載入市場代碼表、抓取完整收盤日 K 與證交所／櫃買中心三大法人買賣超..."):
                resolved, not_found = resolve_user_queries(query_text)
                if not_found:
                    st.warning(f"找不到相符標的：{', '.join(not_found)}")

                results = []
                for sym, code, name, m_lbl in resolved:
                    df = yf.download(tickers=sym, period="5mo", interval="1d", auto_adjust=False, progress=False)
                    if (df is None or df.empty) and sym.endswith(".TW"):
                        df = yf.download(tickers=sym.replace(".TW", ".TWO"), period="5mo", interval="1d", auto_adjust=False, progress=False)
                        if df is not None and not df.empty:
                            m_lbl = "上櫃股"
                    if df is None or df.empty:
                        st.error(f"無法取得 {code} {name} 的 Yahoo 日 K 資料。")
                        continue
                    _, res = evaluate_single_stock(df, params, force_report=True)
                    if res:
                        row = {"股票代碼": code, "股票名稱": name, "市場別": m_lbl}
                        row.update(res)
                        results.append(row)

                if results:
                    enrich_rows_with_insti(results)
                    st.success(f"✅ 完成 {len(results)} 檔個股深度健檢！（資料日期：{results[0]['資料日期']}）")
                    for r in results:
                        render_mobile_stock_card(r)

                    excel_data = build_excel_bytes(
                        [("個股深度健檢報告", results, "手動指定個股四條件＋主力成本深度健檢")],
                        [("資料日期", results[0]["資料日期"]), ("產出時間", now_str), ("健檢檔數", f"{len(results)} 檔")],
                        is_single_check=True,
                    )
                    tag = f"{results[0]['股票代碼']}{results[0]['股票名稱']}" if len(results) == 1 else f"自選{len(results)}檔"
                    dt_tag = datetime.now(timezone(timedelta(hours=8))).strftime("%Y%m%d_%H%M")
                    st.download_button(
                        label="📥 下載完整 Excel 健檢報告",
                        data=excel_data,
                        file_name=f"個股健檢報告_{tag}_{dt_tag}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

# ---------- 分頁 2：全市場選股掃描（支援 15:30 預抓快取秒開＋訪客共享快取） ----------
with tab2:
    mode_choice = st.radio(
        "選擇掃描市場範圍：",
        options=["top", "twse", "tpex", "emerging"],
        format_func=lambda k: {
            "top": "熱門權值股 (64檔)",
            "twse": "上市股 (~1000檔)",
            "tpex": "上櫃股 (~800檔)",
            "emerging": "興櫃股 (~350檔)",
        }[k],
        horizontal=True,
    )

    if mode_choice == "emerging":
        st.caption("💡 提醒：掃描【興櫃股】時，若成交金額門檻維持預設 2000 萬元，程式會自動降為 **500 萬元** 以免誤殺興櫃績優股。")

    if st.button("🚀 查看／掃描選股結果（自動套用每日 15:30 共享快取）", type="primary"):
        market_title = MODE_LABELS[mode_choice]
        scan_params = dict(params)
        if mode_choice == "emerging" and scan_params["min_amount_wan"] == 2000.0:
            scan_params["min_amount_wan"] = 500.0

        # 第一層：優先嘗試秒載每日 15:30 主動預抓之靜態快取檔
        result_pkg = None
        if is_default_params(mode_choice, scan_params):
            result_pkg = try_load_precomputed_cache(mode_choice, effective_date_str)

        # 第二層：若無預抓檔或使用者自訂參數，則呼叫全體訪客共享記憶體快取（首次執行需數十秒，後續訪客秒開）
        if result_pkg is None:
            with st.spinner(f"正在讀取【{market_title}】共享快取或執行首次掃描（完成後將自動快取至下一個 15:30，供後續使用者秒開）..."):
                result_pkg = run_shared_market_scan_cached(
                    mode_choice, period_key, tuple(sorted(scan_params.items()))
                )

        full_sorted = result_pkg["full_sorted"]
        near_top = result_pkg["near_top"]
        valid_cnt = result_pkg["valid_cnt"]
        total = result_pkg["total"]
        actual_date = result_pkg.get("data_date") or effective_date_str
        gen_time = result_pkg.get("generated_at", now_str)
        src_badge = "⚡ 來自每日 15:30 自動預抓快取（極速秒載）" if result_pkg.get("source") == "daily_precomputed" else "⚡ 已寫入全體訪客共享快取（保留至下個 15:30）"

        st.success(
            f"🎉 【{market_title}】資料日期：**{actual_date}** ｜ 掃描 {valid_cnt}/{total} 檔 ｜ "
            f"完全符合：{len(full_sorted)} 檔 ｜ 準突破觀察：{len(near_top)} 檔\n\n{src_badge}（快取建立時間：{gen_time}）"
        )

        st.subheader(f"🎯 4 項條件完全符合（{len(full_sorted)} 檔）")
        if full_sorted:
            for r in full_sorted:
                render_mobile_stock_card(r)
        else:
            st.caption("本次掃描無同時符合 4 項嚴格條件的個股，請參考下方準突破觀察名單。")

        st.subheader(f"👀 準突破觀察名單：符合 3 項（前 {len(near_top)} 檔）")
        if near_top:
            for r in near_top:
                render_mobile_stock_card(r)
        else:
            st.caption("無符合 3 項條件之觀察股。")

        excel_bytes = build_excel_bytes(
            [
                ("1_完全符合選股", full_sorted, "分頁1：4項條件完全符合選股"),
                ("2_準突破觀察名單", near_top, "分頁2：準突破觀察名單(符合3項)"),
            ],
            [
                ("掃描市場範圍", market_title),
                ("資料日期", actual_date),
                ("快取建立時間", gen_time),
                ("成功讀取檔數", f"{valid_cnt}/{total}"),
                ("完全符合檔數", f"{len(full_sorted)} 檔"),
                ("準突破觀察檔數", f"{len(near_top)} 檔"),
            ],
        )
        dt_tag = actual_date.replace("-", "")
        st.download_button(
            label=f"📥 下載【{market_title}】Excel 完整報表 ({actual_date})",
            data=excel_bytes,
            file_name=f"進階選股結果_{market_title}_{dt_tag}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ---------- 分頁 3：我的持倉三段式自動移動停利／停損管理 ----------
with tab3:
    st.caption("💡 規則：第1段(<+5%)守紅K低點/-7% ｜ 第2段(+5~15%)自動上移保本/20MA ｜ 第3段(>+15%)減碼線(10MA/2ATR)＋波段全出(破20MA三日)。")

    # 從網址參數自動還原先前記憶的多檔持倉 (格式: code:price:lots:date|...)
    if "holdings" not in st.session_state:
        init_h = []
        raw_h = st.query_params.get("h", "")
        if raw_h:
            for item in raw_h.split("|"):
                parts = item.split(":")
                if len(parts) >= 3:
                    try:
                        init_h.append({"code": parts[0], "buy_price": float(parts[1]), "lots": int(parts[2]), "buy_date": parts[3] if len(parts) > 3 else ""})
                    except ValueError:
                        pass
        st.session_state["holdings"] = init_h

    def _sync_holdings_to_url():
        encoded = "|".join([f"{h['code']}:{h['buy_price']}:{h['lots']}:{h.get('buy_date','')}" for h in st.session_state["holdings"]])
        if encoded:
            st.query_params["h"] = encoded
        elif "h" in st.query_params:
            del st.query_params["h"]

    with st.form("add_holding_form", clear_on_submit=True):
        hc1, hc2 = st.columns(2)
        h_code = hc1.text_input("股票代號或名稱", placeholder="如：3004 或 豐達科")
        h_price = hc2.number_input("買進均價 (元)", min_value=0.1, value=100.0, step=0.5)
        hc3, hc4 = st.columns(2)
        h_lots = hc3.number_input("持有張數", min_value=1, value=1, step=1)
        h_date = hc4.text_input("買進日期 (選填)", placeholder="如：2026-11-18")
        if st.form_submit_button("➕ 加入持倉清單（自動記憶於網址）"):
            if h_code.strip():
                st.session_state["holdings"].append({
                    "code": h_code.strip(), "buy_price": float(h_price), "lots": int(h_lots), "buy_date": h_date.strip()
                })
                _sync_holdings_to_url()
                st.rerun()

    if st.session_state["holdings"]:
        st.markdown("**📋 目前追蹤持倉（把更新後的網址加入書籤即可永久記憶）**")
        for idx, h in enumerate(st.session_state["holdings"]):
            r1, r2 = st.columns([5, 1])
            r1.write(f"{idx+1}. **{h['code']}** ｜ 買進價 `{h['buy_price']}` 元 × `{h['lots']}` 張" + (f" ({h['buy_date']})" if h.get("buy_date") else ""))
            if r2.button("刪", key=f"del_h_{idx}"):
                st.session_state["holdings"].pop(idx)
                _sync_holdings_to_url()
                st.rerun()

        if st.button("🔄 一鍵計算全部持倉今日移動停利／停損", type="primary"):
            with st.spinner("正在抓取最新收盤日 K 並計算三段式移動防守線..."):
                h_results = []
                for h in st.session_state["holdings"]:
                    res_q, _ = resolve_user_queries(h["code"])
                    if not res_q:
                        st.warning(f"找不到標的：{h['code']}")
                        continue
                    sym, code, name, m_lbl = res_q[0]
                    df_h = yf.download(tickers=sym, period="6mo", interval="1d", auto_adjust=False, progress=False)
                    if (df_h is None or df_h.empty) and sym.endswith(".TW"):
                        df_h = yf.download(tickers=sym.replace(".TW", ".TWO"), period="6mo", interval="1d", auto_adjust=False, progress=False)
                        if df_h is not None and not df_h.empty:
                            m_lbl = "上櫃股"
                    item_res = evaluate_holding_position(df_h, code, name, m_lbl, h["buy_price"], h["lots"], h.get("buy_date", ""))
                    if item_res:
                        h_results.append(item_res)

                if h_results:
                    tot_pnl = sum(r["預估淨損益(元)"] for r in h_results)
                    tot_cost = sum(r["買進均價"] * r["持有張數"] * 1000 for r in h_results)
                    tot_pct = round(tot_pnl / tot_cost * 100.0, 2) if tot_cost > 0 else 0.0
                    st.success(f"✅ 完成 {len(h_results)} 檔持倉風控計算！總預估淨損益：{tot_pnl:+,d} 元 ({tot_pct:+.2f}%)")

                    for r in h_results:
                        cls = "card-warn" if ("🚨" in r["盤後風控警示狀態"] or "⚠️" in r["盤後風控警示狀態"]) else "card-star" if "🚀" in r["目前防守階段"] else "card-watch"
                        m1, m2, m3, m4 = st.columns(4)
                        m1.metric("最新收盤", f"{r['最新收盤價']}", f"{r['未實現報酬(%)']:+.2f}%")
                        m2.metric("預估淨損益", f"{r['預估淨損益(元)']:+,d}", f"{r['持有張數']}張(買{r['買進均價']})")
                        m3.metric("移動停損/全出", f"{r['今日移動停損/全出價']}", f"20MA {r['20MA月線價']}")
                        m4.metric("短線減碼鎖利", f"{r['短線減碼鎖利線']}", f"10MA {r['10MA均線價']}")

                        st.markdown(
                            f"""<div class="card-box {cls}">
                                <div class="card-title">{r['股票代碼']} {r['股票名稱']} ({r['市場別']}) ｜ 資料日：{r['資料日期']}</div>
                                <div class="card-badge">{r['目前防守階段']}<br>{r['盤後風控警示狀態']}</div>
                                <div class="sec-body">👉 {r['明日操作建議']}</div>
                            </div>""",
                            unsafe_allow_html=True,
                        )
                        if r.get("_chart_30d"):
                            with st.expander(f"📈 查看 {r['股票代碼']} {r['股票名稱']} 追蹤價量走勢圖 ＋ 停損/停利判讀"):
                                render_price_volume_charts(
                                    f"{r['股票代碼']} {r['股票名稱']}",
                                    r["_chart_30d"],
                                    r.get("今日移動停損/全出價"),
                                    mark_date=r.get("買進日期") if r.get("買進日期") != "未填" else None,
                                )

                    h_excel = build_holdings_excel_bytes(h_results)
                    dt_t = h_results[0]["資料日期"].replace("-", "")
                    st.download_button(
                        label=f"📥 下載每日持倉移動停利／停損 Excel ({dt_t})",
                        data=h_excel,
                        file_name=f"每日持倉移動停利停損報表_{dt_t}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
    else:
        st.info("請於上方表單輸入您買進的股票代號與買進均價，即可開始每日盤後追蹤三段式移動停利／停損！")

# 底部指標判讀速查與免責聲明
with st.expander("📖 各欄位指標意義與數值高低判讀指南"):
    st.markdown(
        """
        - **最低日成交金額（上市/上櫃預設 `2,000 萬`、興櫃 `500 萬`）**：兼顧高價股與小型績優股的流動性底線，避免買到滑價冷門股。
        - **日週轉率（量佔股本 %）**：中小型股起漲甜美區約 `0.5% ~ 5%`；若單日衝破 **`8%`** 以上會亮出過熱警戒。
        - **量能倍數（甜美區 `1.5 ~ 3.5 倍`）**：代表主力資金進場點火；並非愈高愈好，暴增 5~10 倍易伴隨隔日沖倒貨。
        - **昨日帶寬（愈低愈好 `<10%` 極佳）與今日帶寬放大**：壓縮蓄力後正式釋放波動；昨日帶寬若已 >20% 易追在高點。
        - **距 20MA 乖離（甜美區 `2% ~ 6%`）**：愈低離月線支撐愈近、下檔風險愈低。
        - **20日 VWAP 資金成本**：比 20MA 更貼近近期爆量進場主力的真實平均成本。
        - **60日籌碼密集區 (POC)**：在現價下方代表底部換手支撐；在上方代表前波套牢壓力。
        - **法人 20 日買超成本**：依證交所／櫃買中心真實每日買超加權推算，現價緊貼成本上方 0%~5% 時主力最有意願護盤。
        """
    )

st.markdown(
    """
    <div class="disclaimer-box">
        ⚠️ <b>免責聲明</b>：本工具僅針對公開市場歷史價量資料與證交所／櫃買中心三大法人籌碼進行自動化統計分析，
        所有指標數據與文字拆解均為程式規則推算結果，<b>不構成任何形式之投資買賣建議或投顧推薦</b>。
        股市投資具有風險，使用者應獨立判斷並審慎評估自身風險承受能力，盈虧自負。
    </div>
    """,
    unsafe_allow_html=True,
)
