# -*- coding: utf-8 -*-
"""
台股進階三條件選股 & 主力籌碼成本健檢 — 手機網頁版 App (Streamlit)
支援 iPhone / Android 瀏覽器「加入主畫面」，跨平台免安裝使用！
核心功能（同步電腦版 v1.8）：
1. 嚴格收盤資料保護：台灣時間 14:00 前自動排除今日盤中未定型 K 棒，只用前一個完整交易日收盤資料。
2. 單檔/多檔個股即時深度健檢（支援輸入代號或中文股名，如：世禾、3004、億豐）。
3. 全市場分類掃描（熱門權值股 / 上市股 / 上櫃股 / 興櫃股）＋即時進度條與預估剩餘時間。
4. 量價主力成本（20日VWAP、60日POC）＋證交所／櫃買中心真實三大法人近20日買超與推估成本。
5. 一鍵匯出與電腦版相同格式之 Excel 報表（含同頁顧問級實戰拆解卡片）。
"""

import io
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
    </style>
    """,
    unsafe_allow_html=True,
)

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
    "股票代碼", "股票名稱", "市場別", "資料日期", "收盤價", "單日漲幅(%)",
    "成交量(張)", "量能倍數(倍)", "昨日帶寬(%)", "今日帶寬(%)",
    "距20MA乖離(%)", "MACD柱狀體", "上影線/紅K實體比",
    "突破紅K低點(停損參考)", "20MA月線價", "20日VWAP資金成本", "60日籌碼密集區(POC)",
    "前波60日高點", "近5日法人合計(張)", "近20日法人合計(張)", "外資/投信連買天數", "法人20日買超成本",
    "狀態與未達標細項", "實戰型態定調與優先級"
]


# ==========================================
# 時間基準與快取資料抓取
# ==========================================
def get_market_data_cutoff_info():
    """台灣時間 14:00 前排除今日盤中未定型 K 棒，14:00 後採用今日完整收盤資料"""
    now_tw = datetime.now(timezone(timedelta(hours=8)))
    today_date = now_tw.date()
    if now_tw.hour < 14:
        msg = f"台灣時間 {now_tw.strftime('%H:%M')}（未滿 14:00）：已自動排除今日盤中資料，嚴格採用【前一個完整交易日收盤資料】。"
        return True, today_date, msg
    else:
        tip = "（當日三大法人約 15:30 後完整公布）" if now_tw.hour < 16 else ""
        msg = f"台灣時間 {now_tw.strftime('%H:%M')}（已過 14:00）：採用【今日({today_date})最新完整收盤資料】{tip}"
        return False, today_date, msg


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


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_insti_day_cached(exchange, yyyymmdd):
    """快取 1 小時抓取單日上市(TWSE)或上櫃(TPEX)三大法人買賣超 JSON（單位：張）"""
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
    """平行抓取證交所/櫃買中心近20日三大法人資料並計算推估買超成本"""
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

    for r in rows_list:
        ex = r.get("_exchange")
        code = r.get("股票代碼", "")
        recent_prices = r.get("_recent_20_prices", [])
        c_today = r["收盤價"]
        vwap_20 = r["20日VWAP資金成本"]
        poc_60 = r["60日籌碼密集區(POC)"]
        vwap_gap = round((c_today / vwap_20 - 1.0) * 100.0, 2)
        poc_pos = "下方（形成底部換手鐵板支撐）" if c_today >= poc_60 else "上方（留意前波套牢解套賣壓）"
        vwap_line = (
            f"• 【量價主力成本】20日量價加權資金成本(VWAP)為 『{vwap_20} 元』（現價距資金成本 {vwap_gap:+.2f}%）；"
            f"近60日最大量籌碼密集區(POC)落在 『{poc_60} 元』，位於現價{poc_pos}。"
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

        r["_deep_report"]["chip_cost_text"] = f"{vwap_line}\n{insti_note}"


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


# ==========================================
# 核心技術指標計算與顧問級拆解
# ==========================================
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
    overlap = f"（✨ 紅K低點與月線重疊，形成雙重支撐鐵板區！）" if abs(stop_low - ma20) / ma20 <= 0.015 else ""

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

    # 嚴格收盤保護：14:00 前剔除今日未收盤 K 棒
    exclude_today, today_tw_date, _ = get_market_data_cutoff_info()
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
    cond_vp = (v_t >= params["min_vol"]) and (v_t >= v_ma5_t * params["vol_ratio"]) and (body > 0) and (upper_s <= body * params["upper_shadow_ratio"])

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
        elif v_t < params["min_vol"]:
            vp_miss = f"成交量({int(v_t)}張)未達{params['min_vol']}張門檻"
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
        "資料日期": dt_str,
        "收盤價": round(c_t, 2),
        "單日漲幅(%)": round((c_t / c_p - 1.0) * 100.0, 2),
        "成交量(張)": int(round(v_t)),
        "量能倍數(倍)": round(v_t / v_ma5_t, 2),
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
    info["_matched_count"] = matched

    if force_report:
        return "CUSTOM", info
    if matched == 4:
        return "FULL", info
    if matched == 3 and cond_ma and body > 0 and v_t >= params["min_vol"]:
        return "NEAR", info
    return None, None


# ==========================================
# 匯出與電腦版相同格式之 Excel (記憶體位元流)
# ==========================================
def build_excel_bytes(sheets_spec, summary_rows):
    """產出可供手機直接下載的 Excel 檔案位元流（含同頁下方診斷卡片）"""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sheet_name, rows_data, title_label in sheets_spec:
            df_sheet = pd.DataFrame([{k: r[k] for k in EXCEL_COLS} for r in rows_data]) if rows_data else pd.DataFrame(columns=EXCEL_COLS)
            df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.sheets[sheet_name]

            header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
            chip_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
            h_font = Font(name="Microsoft JhengHei", size=10, bold=True)
            d_font = Font(name="Microsoft JhengHei", size=10)

            for r_i in range(1, ws.max_row + 1):
                ws.row_dimensions[r_i].height = 24
                for c_i in range(1, ws.max_column + 1):
                    cell = ws.cell(row=r_i, column=c_i)
                    if r_i == 1:
                        cell.fill = chip_fill if 16 <= c_i <= 22 else header_fill
                        cell.font = h_font
                        cell.alignment = Alignment(vertical="center", horizontal="center")
                    else:
                        cell.font = d_font

            # 同頁下方寫入診斷卡片
            cur = ws.max_row + 3
            max_c = len(EXCEL_COLS)
            sec_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            sec_font = Font(name="Microsoft JhengHei", size=11, bold=True, color="FFFFFF")
            ws.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=max_c)
            ct = ws.cell(row=cur, column=1, value=f"📋 【{title_label} — 個股深度實戰拆解與主力成本報告】")
            ct.font, ct.fill = sec_font, sec_fill
            cur += 1

            for idx, row in enumerate(rows_data, start=1):
                dr = row.get("_deep_report", {})
                ws.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=max_c)
                ch = ws.cell(row=cur, column=1, value=f"#{idx} {row['股票代碼']} {row['股票名稱']} — {dr.get('archetype_title', '')} | {dr.get('rating_badge', '')}")
                ch.font = h_font
                ch.fill = header_fill
                cur += 1
                for lbl, txt in [
                    ("✅ 達標優勢", dr.get("pros_text", "")),
                    ("🔍 缺口與壓力", dr.get("diagnosis_text", "")),
                    ("💰 主力成本與法人", dr.get("chip_cost_text", "")),
                    ("🛡️ 操作與防守價", dr.get("strategy_text", "")),
                ]:
                    ws.merge_cells(start_row=cur, start_column=1, end_row=cur, end_column=3)
                    ws.merge_cells(start_row=cur, start_column=4, end_row=cur, end_column=max_c)
                    ws.cell(row=cur, column=1, value=lbl).font = h_font
                    cv = ws.cell(row=cur, column=4, value=txt)
                    cv.font = d_font
                    cv.alignment = Alignment(wrap_text=True, vertical="center")
                    ws.row_dimensions[cur].height = max(36, (txt.count("\n") + 1) * 18 + 8)
                    cur += 1
                cur += 1

            for c_i in range(1, max_c + 1):
                ws.column_dimensions[get_column_letter(c_i)].width = 18 if c_i < 23 else 48

        pd.DataFrame(summary_rows, columns=["項目", "數值 / 設定"]).to_excel(writer, sheet_name="掃描統計與參數", index=False)

    buf.seek(0)
    return buf.getvalue()


# ==========================================
# 手機友善診斷卡片渲染元件
# ==========================================
def render_mobile_stock_card(row):
    dr = row["_deep_report"]
    badge = dr["rating_badge"]
    css_cls = "card-star" if "★★★" in badge else "card-warn" if ("警戒" in badge or "✖" in badge) else "card-watch"

    # 關鍵數字 4 格卡片 (手機上自動 2x2 排列)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("收盤價", f"{row['收盤價']}", f"{row['單日漲幅(%)']:+.2f}%")
    c2.metric("20日VWAP成本", f"{row['20日VWAP資金成本']}", f"乖離 {row['距20MA乖離(%)']}%")
    c3.metric("法人20日成本", f"{row['法人20日買超成本']}", f"5日 {row['近5日法人合計(張)']}張")
    c4.metric("短線停損 / 月線", f"{row['突破紅K低點(停損參考)']}", f"20MA {row['20MA月線價']}")

    html_card = f"""
    <div class="card-box {css_cls}">
        <div class="card-title">{row['股票代碼']} {row['股票名稱']} ({row['市場別']}) ｜ 資料日：{row['資料日期']}</div>
        <div class="card-badge">{dr['archetype_title']}<br>{badge}</div>
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


# ==========================================
# 主介面排版 (手機優先)
# ==========================================
st.title("📈 台股選股 & 主力籌碼健檢")
_, _, cutoff_msg = get_market_data_cutoff_info()
st.info(f"🕒 {cutoff_msg}")

# 參數設定收合區（預設收合讓手機畫面乾淨）
with st.expander("⚙️ 調整四大選股條件參數（預設已為進階實戰最佳值）", expanded=False):
    col_a, col_b = st.columns(2)
    macd_tol = col_a.number_input("MACD 翻紅容錯天數", min_value=1, max_value=10, value=3)
    max_pre_bw = col_b.number_input("突破前布林帶寬上限(%)", min_value=5.0, max_value=50.0, value=15.0, step=1.0)
    max_bias20 = col_a.number_input("距 20MA 最大乖離率(%)", min_value=2.0, max_value=30.0, value=8.0, step=1.0)
    min_vol = col_b.number_input("最低成交量門檻(張)", min_value=0, max_value=10000, value=500, step=100)
    vol_ratio = col_a.number_input("攻擊量倍數(較前5日均量)", min_value=1.0, max_value=5.0, value=1.5, step=0.1)
    upper_shadow = col_b.number_input("上影線/紅K實體比例上限", min_value=0.1, max_value=2.0, value=0.5, step=0.1)

params = {
    "fast_len": 12, "slow_len": 26, "sig_len": 9, "macd_tolerance": int(macd_tol),
    "bb_len": 20, "bb_k": 2.0, "max_pre_bw": float(max_pre_bw),
    "ma_len": 20, "ma_up_bars": 1, "max_bias20": float(max_bias20),
    "min_vol": int(min_vol), "vol_ratio": float(vol_ratio), "upper_shadow_ratio": float(upper_shadow),
}

tab1, tab2 = st.tabs(["🩺 單檔／自選個股即時健檢", "🚀 全市場選股掃描"])

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
            with st.spinner("正在載入市場代碼表、抓取日 K 與證交所／櫃買中心三大法人買賣超..."):
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
                    st.success(f"✅ 完成 {len(results)} 檔個股深度健檢！")
                    for r in results:
                        render_mobile_stock_card(r)

                    # 提供 Excel 完整報告下載
                    excel_data = build_excel_bytes(
                        [("個股深度健檢報告", results, "手動指定個股四條件＋主力成本深度健檢")],
                        [("產出時間", datetime.now(timezone(timedelta(hours=8))).strftime("%Y-%m-%d %H:%M")), ("健檢檔數", f"{len(results)} 檔")],
                    )
                    tag = f"{results[0]['股票代碼']}{results[0]['股票名稱']}" if len(results) == 1 else f"自選{len(results)}檔"
                    dt_tag = datetime.now(timezone(timedelta(hours=8))).strftime("%Y%m%d_%H%M")
                    st.download_button(
                        label="📥 下載完整 Excel 健檢報告",
                        data=excel_data,
                        file_name=f"個股健檢報告_{tag}_{dt_tag}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

# ---------- 分頁 2：全市場選股掃描 ----------
with tab2:
    mode_choice = st.radio(
        "選擇掃描市場範圍：",
        options=["top", "twse", "tpex", "emerging"],
        format_func=lambda k: {
            "top": "熱門權值股 (64檔，約10秒)",
            "twse": "上市股 (~1000檔，約1.5分鐘)",
            "tpex": "上櫃股 (~800檔，約1分鐘)",
            "emerging": "興櫃股 (~350檔，約30秒)",
        }[k],
        horizontal=True,
    )

    if st.button("🚀 開始掃描選股", type="primary"):
        market_title = MODE_LABELS[mode_choice]
        stock_dict = fetch_tw_tickers_fast(mode=mode_choice)
        tickers = list(stock_dict.keys())
        total = len(tickers)

        pbar = st.progress(0.0, text=f"準備掃描【{market_title}】共 {total} 檔...")
        full_list, near_list = [], []
        valid_cnt = 0
        batch_size = 40
        t0 = time.time()

        for start_i in range(0, total, batch_size):
            batch = tickers[start_i : start_i + batch_size]
            end_i = min(start_i + batch_size, total)
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
                    status, res = evaluate_single_stock(sub_df, params, force_report=False)
                    if status in ("FULL", "NEAR") and res:
                        code = sym.replace(".TW", "").replace(".TWO", "")
                        name, m_lbl = stock_dict.get(sym, ("", market_title))
                        row = {"股票代碼": code, "股票名稱": name, "市場別": m_lbl}
                        row.update(res)
                        (full_list if status == "FULL" else near_list).append(row)
                except Exception:
                    continue

            elapsed = time.time() - t0
            eta = int(round((elapsed / end_i) * (total - end_i))) if end_i < total else 0
            pbar.progress(end_i / total, text=f"📊 進度 {end_i}/{total} 檔 ({end_i/total*100:.0f}%) ｜ 已用 {int(elapsed)}秒 ｜ 預估剩餘 {eta}秒")
            time.sleep(0.1)

        near_top = sorted(near_list, key=lambda x: x["單日漲幅(%)"], reverse=True)[:30]
        full_sorted = sorted(full_list, key=lambda x: x["今日帶寬(%)"])

        if full_sorted or near_top:
            pbar.progress(1.0, text="🏦 正在為入選股補抓證交所／櫃買中心三大法人買超與成本...")
            enrich_rows_with_insti(full_sorted + near_top)

        pbar.empty()
        st.success(f"🎉 【{market_title}】掃描完成！（讀取 {valid_cnt}/{total} 檔）完全符合：{len(full_sorted)} 檔 ｜ 準突破觀察：{len(near_top)} 檔")

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
                ("產出時間", datetime.now(timezone(timedelta(hours=8))).strftime("%Y-%m-%d %H:%M")),
                ("成功讀取檔數", f"{valid_cnt}/{total}"),
                ("完全符合檔數", f"{len(full_sorted)} 檔"),
                ("準突破觀察檔數", f"{len(near_top)} 檔"),
            ],
        )
        dt_tag = datetime.now(timezone(timedelta(hours=8))).strftime("%Y%m%d_%H%M")
        st.download_button(
            label=f"📥 下載【{market_title}】Excel 完整報表",
            data=excel_bytes,
            file_name=f"進階選股結果_{market_title}_{dt_tag}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# 底部指標判讀速查
with st.expander("📖 各欄位指標意義與數值高低判讀指南"):
    st.markdown(
        """
        - **量能倍數（甜美區 `1.5 ~ 3.5 倍`）**：代表主力資金進場點火；並非愈高愈好，暴增 5~10 倍易伴隨隔日沖倒貨。
        - **昨日帶寬（愈低愈好 `<10%` 極佳）與今日帶寬放大**：壓縮蓄力後正式釋放波動；昨日帶寬若已 >20% 易追在高點。
        - **距 20MA 乖離（甜美區 `2% ~ 6%`）**：愈低離月線支撐愈近、下檔風險愈低。
        - **20日 VWAP 資金成本**：比 20MA 更貼近近期爆量進場主力的真實平均成本。
        - **60日籌碼密集區 (POC)**：在現價下方代表底部換手支撐；在上方代表前波套牢壓力。
        - **法人 20 日買超成本**：依證交所／櫃買中心真實每日買超加權推算，現價緊貼成本上方 0%~5% 時主力最有意願護盤。
        """
    )